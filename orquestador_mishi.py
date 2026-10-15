"""Orquestador local productor–consumidor ($0: no depende de API de pago para operar).

 PRODUCTOR  AgenteReceptor     long polling de Telegram (timeout 25 s). Descarga en paralelo (hasta 8), calcula SHA-256, guarda
                               en `Facturas Mishiguene/`, encola en coordinador.py (RECIBIDO) y confirma en < 10 s:
                               «✅ N comprobantes recibidos y encolados». El offset solo avanza DESPUÉS de guardar y encolar.
 CONSUMIDOR AgenteProcesador   un único worker. Toma hasta BATCH_MAX trabajos y corre procesar_lote_wpp.ejecutar sobre el
                               lote: caché SHA / dHash / aprendizaje primero; Gemini (API) y luego Tesseract solo si falta
                               (secuencial: no satura VRAM ni rate-limit). La escritura a Sheets es EN BLOQUE: una
                               values_batch_update por planilla y lote, con backoff ante 429/5xx. Un fallo del lote se
                               aísla archivo por archivo; nada se pierde ni se marca ESCRITO tras un error.
 Auxiliares AgenteFileWatcher (carpetas locales → misma cola), AgenteAprobaciones (Q:X), AgenteRecalculoCascada (solo con --con-cascada).
Estado: estado/orquestador_estado.json (latido). Parar: Ctrl+C o crear estado/STOP. Un solo consumidor de getUpdates.
Uso: python orquestador_mishi.py [--push-drive] [--sin-telegram] [--reemplazar]   (o RUN_BOT.bat)
"""
from __future__ import annotations

import argparse
import hashlib
import json
import logging
import os
import queue
import shutil
import subprocess
import sys
import threading
import time
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from types import SimpleNamespace

BASE = Path(__file__).resolve().parent
sys.path.insert(0, str(BASE))
ESTADO = Path(os.environ.get("MISHI_ESTADO_DIR") or BASE / "estado")
LATIDO = ESTADO / "orquestador_estado.json"
PIDFILE = ESTADO / "orquestador.pid"
STOP = ESTADO / "STOP"
ENTRADA = BASE / "Facturas Mishiguene"
LOG = logging.getLogger("orquestador")
BATCH_MAX = int(os.environ.get("MISHI_BATCH_MAX", "20"))
DESCARGAS_PARALELAS = 8
PARAR = threading.Event()
ESTADO_LOCK = threading.Lock()              # telegram_state.json lo tocan productor y consumidor
BUS: "queue.Queue[tuple[str, dict]]" = queue.Queue()
COMANDOS: "queue.Queue[dict]" = queue.Queue()   # /asignar y /deshacer: los ejecuta el worker (único escritor)
METRICAS: dict[str, dict] = {}


def _m(nombre: str) -> dict:
    return METRICAS.setdefault(nombre, {"ciclos": 0, "errores": 0, "ultimo_ms": None, "ultimo_error": None, "eventos": 0})


class Agente(threading.Thread):
    intervalo = 3.0

    def __init__(self, nombre: str):
        super().__init__(name=nombre, daemon=True)
        self.nombre = nombre

    def ciclo(self) -> None:
        raise NotImplementedError

    def run(self) -> None:
        m = _m(self.nombre)
        espera_error = 2.0
        while not STOP.exists() and not PARAR.is_set():
            t0 = time.perf_counter()
            try:
                self.ciclo()
                espera_error = 2.0
            except Exception as e:                            # un fallo no tumba al orquestador; backoff creciente
                m["errores"] += 1
                m["ultimo_error"] = f"{type(e).__name__}: {str(e)[:160]}"
                try:
                    import registro_fallos as rf
                    rf.evento(f"orquestador:{self.nombre}", "ciclo_agente", "ERROR", error=e)
                except Exception:
                    pass
                LOG.exception("%s falló", self.nombre)
                PARAR.wait(espera_error)
                espera_error = min(espera_error * 2, 60)
            m["ciclos"] += 1
            m["ultimo_ms"] = round((time.perf_counter() - t0) * 1000)
            PARAR.wait(self.intervalo)


# =========================================================================== #
# PRODUCTOR
# =========================================================================== #
class AgenteReceptor(Agente):
    intervalo = 0.5

    def __init__(self):
        super().__init__("AgenteReceptor")
        import telegram_ingestor as ti
        self.ti = ti
        token, self.chat = os.environ.get("TELEGRAM_BOT_TOKEN"), os.environ.get("TELEGRAM_CHAT_ID")
        if not token:
            raise RuntimeError("falta TELEGRAM_BOT_TOKEN (.env)")
        self.api = ti.TelegramAPI(token)

    def _descargar(self, t: dict) -> dict:
        destino = ENTRADA / f"{t['chat']}_{t['mid']}_{Path(t['nombre']).name}"
        try:
            ruta = self.api.descargar(t["file_id"], destino)
            return {**t, "ruta": ruta, "sha": hashlib.sha256(ruta.read_bytes()).hexdigest()}
        except Exception as e:                                # noqa: BLE001
            return {**t, "error": e}

    def ciclo(self) -> None:
        ti = self.ti
        with ESTADO_LOCK:
            est = ti.cargar_estado()
        ups = self.api.get_updates(est.get("offset"), timeout=25)      # long polling real
        if not ups:
            return
        t0 = time.perf_counter()
        ENTRADA.mkdir(exist_ok=True)
        tareas, ultimo = [], est.get("offset")
        import coordinador
        for u in ups:
            ultimo = u["update_id"] + 1
            msg = u.get("message") or u.get("channel_post")
            if not msg:
                continue
            chat = str(msg["chat"]["id"])
            if self.chat and chat != str(self.chat):
                continue
            with ESTADO_LOCK:
                e2 = ti.cargar_estado()
                ti.registrar_ids(e2, chat, [msg["message_id"]])
                ti.guardar_estado(e2)
            texto = (msg.get("text") or "").strip()
            cmd = texto.lower().split("@")[0].split(" ")[0]
            if cmd == "/pendientes":
                self.api.responder(chat, ti.texto_pendientes(), msg["message_id"])
            elif cmd in ("/asignar", "/deshacer"):
                COMANDOS.put({"cmd": cmd, "texto": texto.lower(), "chat": chat, "mid": msg["message_id"]})
            demo = bool(ti.DEMO_RE.search(msg.get("caption") or ""))
            for file_id, nombre in ti.adjuntos(msg):
                tareas.append({"file_id": file_id, "nombre": nombre, "chat": chat, "mid": msg["message_id"], "demo": demo})
        nuevos = dup = fallidos = 0
        cola = coordinador.Cola()
        if tareas:
            with ThreadPoolExecutor(max_workers=DESCARGAS_PARALELAS) as ex:
                res = list(ex.map(self._descargar, tareas))
            with ESTADO_LOCK:
                vistos = set(ti.cargar_estado().get("sha256", {}))
            for r in res:
                if "error" in r:
                    fallidos += 1
                    LOG.warning("descarga fallida (%s): %s", r["nombre"], str(r["error"])[:100])
                    continue
                previo = cola.obtener(r["sha"])
                if not r["demo"] and (r["sha"] in vistos or (previo and previo["estado"] in ("ESCRITO", "REVISAR"))):
                    dup += 1
                    r["ruta"].unlink(missing_ok=True)
                    continue
                if previo and previo["estado"] in ("RECIBIDO", "LEIDO", "VALIDADO") and not r["demo"]:
                    dup += 1                                     # ya encolado (reenvío de la misma foto)
                    continue
                cola.registrar(r["sha"], r["ruta"].name, "telegram", ruta=str(r["ruta"]), chat=r["chat"],
                               message_id=r["mid"], demo=r["demo"])
                if previo and r["demo"]:                         # DEMO re-procesa aunque ya estuviera
                    cola.transicion(r["sha"], "RECIBIDO", forzar=True)
                nuevos += 1
        if fallidos == 0:                                       # offset solo tras guardar y encolar TODO
            with ESTADO_LOCK:
                e3 = ti.cargar_estado()
                e3["offset"] = ultimo
                ti.guardar_estado(e3)
        if tareas and (nuevos or dup):
            chat0 = tareas[0]["chat"]
            txt = f"✅ {nuevos} comprobante(s) recibidos y encolados." + (f" ({dup} ya estaban procesados o en cola.)" if dup else "")
            if fallidos:
                txt += f" ⚠ {fallidos} no se pudieron descargar: Telegram los reenvía en el próximo ciclo."
            if fallidos == 0 or nuevos:
                self.api.responder(chat0, txt)
        _m(self.nombre)["eventos"] += nuevos
        LOG.info("receptor: %d archivos (%d nuevos, %d duplicados, %d fallidos) en %.1f s", len(tareas), nuevos, dup,
                 fallidos, time.perf_counter() - t0)


# =========================================================================== #
# CONSUMIDOR (único escritor)
# =========================================================================== #
class AgenteProcesador(Agente):
    intervalo = 0.5

    def __init__(self, push: bool):
        super().__init__("AgenteProcesador")
        self.push = push
        self.api = None
        token = os.environ.get("TELEGRAM_BOT_TOKEN")
        if token:
            import telegram_ingestor as ti
            self.ti, self.api = ti, ti.TelegramAPI(token)

    # ---- comandos que escriben ------------------------------------------------
    def _comandos(self) -> None:
        while True:
            try:
                c = COMANDOS.get_nowait()
            except queue.Empty:
                return
            try:
                if c["cmd"] == "/asignar":
                    r = self.ti.comando_asignar(c["texto"], self.push)
                else:
                    with ESTADO_LOCK:
                        est = self.ti.cargar_estado()
                    d = self.ti.deshacer_push(est.get("ultima_inyeccion"))
                    r = (f"↩ Deshecho: {d['filas']} fila(s) de «{d['planilla']}»." if d.get("deshecho")
                         else f"No deshice nada: {d['motivo']}")
                    if d.get("deshecho"):
                        with ESTADO_LOCK:
                            e2 = self.ti.cargar_estado()
                            e2["ultima_inyeccion"] = None
                            self.ti.guardar_estado(e2)
            except Exception as e:                            # noqa: BLE001
                r = f"No pude ejecutar {c['cmd']}: {type(e).__name__}"
            if self.api:
                self.api.responder(c["chat"], r, c["mid"])

    # ---- un lote ------------------------------------------------------------
    def _ejecutar(self, jobs: list[dict], demo: bool):
        import procesar_lote_wpp as lote
        tmp = BASE / "Carpeta Revision" / ("lote demo orquestador" if demo else "lote orquestador")
        shutil.rmtree(tmp, ignore_errors=True)
        tmp.mkdir(parents=True, exist_ok=True)
        try:
            for j in jobs:
                shutil.copy2(j["ruta"], tmp / Path(j["ruta"]).name)
            ns = SimpleNamespace(push_drive=self.push, confirmo_escribir_produccion=self.push, dry_run=False, json=None)
            # CAMINO RÁPIDO: catálogos en RAM, sin reescribir los .xlsx de Copias_Gestion; escritura directa a Sheets
            return lote.ejecutar(tmp, not demo, ns, indexar_informes=False, forzar_nuevas=demo, demo=demo,
                                 escribir_xlsx_local=False)[0]
        finally:
            shutil.rmtree(tmp, ignore_errors=True)

    def _lote(self, jobs: list[dict], demo: bool) -> list[tuple[dict, str, dict]]:
        """Devuelve [(job, estado_final, detalle)]. Si el lote entero falla, se aísla archivo por archivo."""
        import coordinador
        cola = coordinador.Cola()
        salida = []
        try:
            reporte = self._ejecutar(jobs, demo)
            reportes = [(jobs, reporte)]
        except Exception as e:                                # noqa: BLE001
            LOG.warning("lote de %d falló (%s): se aísla archivo por archivo", len(jobs), type(e).__name__)
            reportes = []
            for j in jobs:
                try:
                    reportes.append(([j], self._ejecutar([j], demo)))
                except Exception as e1:                       # noqa: BLE001
                    import registro_fallos as rf
                    rf.evento(j["clave"], "procesar_lote", "ERROR", error=e1, archivo=j["archivo"])
                    cola.transicion(j["clave"], "ERROR", error=f"{type(e1).__name__}: {str(e1)[:200]}", forzar=True)
                    salida.append((j, "ERROR", {"error": type(e1).__name__}))
        for grupo, rep in reportes:
            pushes = {k: v for k, v in (rep.get("push_drive") or {}).items() if v}
            errores_push = {v.get("local") for v in pushes.values() if v.get("error") or v.get("bloqueado")}
            aud = {a["archivo"].split(" [")[0]: a for a in rep.get("auditoria", [])}
            for j in grupo:
                a = aud.get(Path(j["ruta"]).name, {})
                est = str(a.get("estado", ""))
                if a.get("local") in errores_push:
                    cola.transicion(j["clave"], "ERROR", error="falló la escritura en la planilla (se reintenta)", forzar=True)
                    salida.append((j, "ERROR", {"estado_pipeline": est}))
                    continue
                for e in ("LEIDO", "VALIDADO"):
                    cola.transicion(j["clave"], e, forzar=True)
                final = "REVISAR" if ("REVISAR" in est or "INDETERMINADO" in est) else "ESCRITO"
                cola.transicion(j["clave"], final, {"estado_pipeline": est, "local": a.get("local"),
                                                    "push": [p.get("planilla") for p in pushes.values() if p.get("escrito")]},
                                forzar=True)
                salida.append((j, final, {"estado_pipeline": est, "local": a.get("local"), "archivo": a.get("archivo")}))
            with ESTADO_LOCK:
                est_t = self.ti.cargar_estado() if self.api else {}
                if self.api:
                    for j, fin, _ in salida:
                        if fin in ("ESCRITO", "REVISAR") and not demo:
                            est_t.setdefault("sha256", {})[j["clave"]] = {"archivo": j["archivo"], "chat": j.get("chat"),
                                                                           "message_id": j.get("message_id")}
                    escritos = [r for r in pushes.values() if r.get("escrito") and r.get("detalle") and r.get("tipo") != "remitos_proveedor"]
                    if escritos:
                        est_t["ultima_inyeccion"] = escritos[-1]
                    self.ti.guardar_estado(est_t)
            self._avisar(grupo, salida, rep, pushes, demo)
        return salida

    def _avisar(self, grupo, salida, rep, pushes, demo) -> None:
        if not self.api:
            return
        por_chat: dict[str, list] = {}
        for j in grupo:
            if j.get("chat"):
                por_chat.setdefault(str(j["chat"]), []).append(j)
        res = {j["clave"]: (f, d) for j, f, d in salida}
        for chat, js in por_chat.items():
            esc = [res[j["clave"]] for j in js if res.get(j["clave"], ("",))[0] == "ESCRITO"]
            rev = [res[j["clave"]] for j in js if res.get(j["clave"], ("",))[0] == "REVISAR"]
            err = [res[j["clave"]] for j in js if res.get(j["clave"], ("",))[0] == "ERROR"]
            lineas = [f"📦 Lote procesado ({len(js)}): {len(esc)} OK · {len(rev)} para revisar · {len(err)} con error"
                      + (" · DEMO" if demo else "")]
            for f, d in rev[:10]:
                lineas.append(f"  ⚠ {d.get('archivo') or ''} → {d.get('estado_pipeline', '')}")
            for rr in rep.get("remitos_obrador") or []:
                lineas.append(f"  🚚 Remito Obrador {rr['remito']} → {rr['receptor']}: {len(rr['renglones'])} renglones, ${rr['total']:,.2f}")
            for r in pushes.values():
                if r.get("bloqueado"):
                    lineas.append(f"  🔒 {r.get('local')}: pestaña {r.get('hoja')} protegida/llena; respaldo en Remitos_Pendientes")
                elif r.get("escrito"):
                    lineas.append("  " + self.ti.texto_push(r).split("\n")[0])
            if err:
                lineas.append("  Los errores quedan en cola y se reintentan solos.")
            try:
                self.api.responder(chat, "\n".join(lineas))
            except Exception:                                 # noqa: BLE001
                LOG.exception("no pude avisar el resultado del lote")

    def ciclo(self) -> None:
        import coordinador
        self._comandos()
        cola = coordinador.Cola()
        jobs = cola.tomar(BATCH_MAX, solo_demo=False)
        demo = False
        if not jobs:
            jobs, demo = cola.tomar(BATCH_MAX, solo_demo=True), True
        buenos = []
        for j in jobs:
            if j.get("ruta") and Path(j["ruta"]).exists():
                buenos.append(j)
            else:                                                # sin archivo no se puede procesar: ERROR visible, no bucle
                cola.transicion(j["clave"], "ERROR", error="archivo no encontrado en disco", forzar=True)
        jobs = buenos
        if not jobs:
            return
        if len(jobs) < BATCH_MAX and not demo:                  # agrupa lo que siga llegando (hasta 2 s)
            time.sleep(2)
            extra = cola.tomar(BATCH_MAX - len(jobs), solo_demo=False)
            jobs += [j for j in extra if j.get("ruta") and Path(j["ruta"]).exists()]
        t0 = time.perf_counter()
        salida = self._lote(jobs, demo)
        c = {k: sum(1 for _, f, _ in salida if f == k) for k in ("ESCRITO", "REVISAR", "ERROR")}
        _m(self.nombre)["eventos"] += len(salida)
        BUS.put(("facturas_procesadas", {"n": c["ESCRITO"] + c["REVISAR"]}))
        LOG.info("lote de %d en %.1f s: %s", len(jobs), time.perf_counter() - t0, c)


# =========================================================================== #
# AUXILIARES
# =========================================================================== #
class AgenteFileWatcher(Agente):
    """Archivos que entran por carpeta (no por Telegram): se encolan igual que los de Telegram."""
    intervalo = 3.0
    CARPETAS = ["Facturas Mishiguene", "Precios/Listas_Mail"]
    EXT = {".jpg", ".jpeg", ".png", ".webp", ".pdf"}

    def __init__(self):
        super().__init__("AgenteFileWatcher")
        self.vistos: dict[str, tuple[float, int]] = {}
        self.primera = True

    def ciclo(self) -> None:
        import coordinador
        act = {}
        for c in self.CARPETAS:
            d = BASE / c
            if d.is_dir():
                for p in d.glob("*"):
                    if p.is_file() and p.suffix.lower() in (self.EXT | {".xlsx", ".xls", ".csv"}):
                        st = p.stat()
                        act[str(p)] = (st.st_mtime, st.st_size)
        if self.primera:
            self.vistos, self.primera = act, False
            return
        nuevos = [Path(k) for k, v in act.items() if self.vistos.get(k) != v and time.time() - v[0] > 20]
        self.vistos.update({str(p): act[str(p)] for p in nuevos})
        cola = coordinador.Cola()
        listas = [p for p in nuevos if p.parent.name == "Listas_Mail"]
        for p in nuevos:
            if p.parent.name == "Facturas Mishiguene" and p.suffix.lower() in self.EXT:
                h = hashlib.sha256(p.read_bytes()).hexdigest()
                if cola.obtener(h) is None:
                    cola.registrar(h, p.name, "carpeta", ruta=str(p), chat=None, message_id=None, demo=False)
                    _m(self.nombre)["eventos"] += 1
        if listas:
            BUS.put(("lista_precios_nueva", {"archivos": [p.name for p in listas]}))


class AgenteAprobaciones(Agente):
    intervalo = 120.0

    def __init__(self, vivo: bool):
        super().__init__("AgenteAprobaciones")
        self.vivo = vivo

    def ciclo(self) -> None:
        import aprobaciones
        import conocimiento
        import imputacion_facturas as imp
        if not self.vivo:
            return
        gc, _ = imp._gspread()
        for local, sid in imp.IDS_PLANILLAS.items():
            sh = gc.open_by_key(sid)
            ws = next((w for w in sh.worksheets() if w.title in ("Remitos Facturas", "Remitos_Pendientes")), None)
            if ws is None:
                continue
            vals = ws.get_all_values()
            if not vals or len(vals[0]) <= aprobaciones.PRIMERA_META:
                continue
            for d in aprobaciones.planear(aprobaciones.leer(vals)):
                if d["accion"] == "CERRAR_SIN_EFECTO":
                    r = aprobaciones.celdas_resultado("CERRAR_SIN_EFECTO", True)
                    ws.update(f"V{d['fila']}:W{d['fila']}", [[r["V"], r["W"]]], value_input_option="RAW")
                elif d["accion"] == "CARGAR":
                    conocimiento.registrar_dato(
                        f"aprobacion.{local}.{d['id']}", {"proveedor": d["proveedor"], "producto": d["producto"],
                                                          "motivo": d["motivo"], "evidencia": d["evidencia"]},
                        "validado", f"{sh.title}/{ws.title} fila {d['fila']} (decisión humana)")
                    BUS.put(("aprobado_pendiente_carga", {"local": local, "id": d["id"]}))
                if d["accion"] in ("CARGAR", "CERRAR_SIN_EFECTO"):
                    _m(self.nombre)["eventos"] += 1


class AgenteRecalculoCascada(Agente):
    intervalo = 5.0
    COMANDOS = {"costos_bom": [["bom_diferencias.py"]],
                "comparativa_precios": [["comparador_precios.py", "--sin-nogada"], ["scouting_sin_referencia.py"]]}

    def __init__(self):
        super().__init__("AgenteRecalculoCascada")
        self.pendiente: set[str] = set()
        self.ultimo = 0.0

    def ciclo(self) -> None:
        import conocimiento
        try:
            while True:
                ev, _ = BUS.get_nowait()
                if ev in ("facturas_procesadas", "lista_precios_nueva"):
                    fuente = "planilla_facturas" if ev == "facturas_procesadas" else "listas_proveedores"
                    huella = hashlib.sha256(f"{ev}{time.time()}".encode()).hexdigest()[:16]
                    conocimiento.registrar_fuente(fuente, huella)
                    self.pendiente.update(conocimiento.dependientes(fuente))
                    _m(self.nombre)["eventos"] += 1
        except queue.Empty:
            pass
        if self.pendiente and time.time() - self.ultimo > 600:      # debounce 10 min, en segundo plano
            self.ultimo = time.time()
            hechos = []
            for d in ("costos_bom", "comparativa_precios"):
                if d in self.pendiente:
                    for cmd in self.COMANDOS[d]:
                        r = subprocess.run([sys.executable, *cmd], cwd=str(BASE), capture_output=True, text=True, timeout=900,
                                           creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
                        hechos.append((cmd[0], r.returncode))
            LOG.info("cascada recalculada: %s", hechos)
            self.pendiente.clear()


# =========================================================================== #
def _entorno() -> str:
    import entornos
    return entornos.entorno()


def otros_consumidores() -> list[int]:
    r = subprocess.run(["powershell", "-NoProfile", "-Command",
                        "Get-CimInstance Win32_Process | Where-Object { $_.CommandLine -match 'telegram_ingestor.py|orquestador_mishi.py' "
                        "-and $_.ProcessId -ne " + str(os.getpid()) + " -and $_.Name -match 'python' } | "
                        "Select-Object -ExpandProperty ProcessId"], capture_output=True, text=True,
                       creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    return [int(x) for x in r.stdout.split() if x.strip().isdigit()]


def latido(agentes: list[Agente], t0: float) -> None:
    ESTADO.mkdir(exist_ok=True)
    try:
        import coordinador
        cola = coordinador.Cola().resumen()
    except Exception:                                          # noqa: BLE001
        cola = {}
    try:
        import invoice_parser
        stats = dict(invoice_parser.STATS)
    except Exception:                                          # noqa: BLE001
        stats = {}
    datos = {"pid": os.getpid(), "inicio": time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(t0)),
             "uptime_s": round(time.time() - t0), "ahora": time.strftime("%Y-%m-%d %H:%M:%S"),
             "cola": cola, "lectura": stats, "batch_max": BATCH_MAX,
             "agentes": {a.nombre: {**_m(a.nombre), "vivo": a.is_alive()} for a in agentes},
             "entorno": _entorno()}
    tmp = LATIDO.with_suffix(f".{os.getpid()}.tmp")
    tmp.write_text(json.dumps(datos, ensure_ascii=False, indent=1), encoding="utf-8")
    for n in range(8):
        try:
            tmp.replace(LATIDO)
            break
        except PermissionError:                                # estado_bot.py puede estar leyéndolo
            time.sleep(0.05 * (n + 1))


def main(argv=None) -> int:
    logging.basicConfig(level=logging.INFO, format="%(asctime)s | %(levelname)s | %(name)s | %(message)s", datefmt="%H:%M:%S")
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--push-drive", action="store_true", help="escribe en la planilla real del entorno activo")
    ap.add_argument("--con-cascada", action="store_true",
                    help="activa AgenteRecalculoCascada (BOM + comparador). Apagado por defecto: no corre con facturas de Telegram")
    ap.add_argument("--sin-telegram", action="store_true")
    ap.add_argument("--sin-aprobaciones", action="store_true")
    ap.add_argument("--reemplazar", action="store_true", help="ignora otros consumidores detectados (no recomendado)")
    a = ap.parse_args(argv)
    if not a.push_drive and _entorno() == "PRUEBA":
        a.push_drive = True                    # en PRUEBA la escritura viva en las copias NO CARGAR es el comportamiento por defecto
    ESTADO.mkdir(exist_ok=True)
    otros = otros_consumidores()
    if otros and not a.reemplazar:
        print(f"Ya hay otro proceso consumiendo Telegram/orquestando (PID {otros}). Detenelo primero o usá --reemplazar.")
        return 3
    if STOP.exists():
        STOP.unlink()
    PIDFILE.write_text(str(os.getpid()))
    try:                                                       # token/chat desde .env (sin imprimirlos)
        for l in (BASE / ".env").read_text(encoding="utf-8").splitlines():
            if "=" in l and not l.startswith("#"):
                k, v = l.split("=", 1)
                os.environ.setdefault(k.strip(), v.strip().strip('"'))
    except OSError:
        pass
    agentes: list[Agente] = []
    if not a.sin_telegram:
        agentes.append(AgenteReceptor())
    agentes += [AgenteProcesador(a.push_drive), AgenteFileWatcher()]
    if a.con_cascada:
        agentes.append(AgenteRecalculoCascada())
    if not a.sin_aprobaciones:
        agentes.append(AgenteAprobaciones(vivo=a.push_drive))
    # trabajos que quedaron reclamados por un proceso anterior: se liberan al arrancar
    try:
        import coordinador
        coordinador.Cola()._mutar(lambda t: [j.pop("en_proceso", None) for j in t.values()])
    except Exception:                                          # noqa: BLE001
        pass
    def _precarga():
        try:
            import imputacion_facturas as imp
            r = imp.precargar_ram()
            LOG.info("catálogos en RAM: %s en %.1f s", ", ".join(r["locales"]), r["segundos"])
        except Exception as e:                                 # noqa: BLE001
            LOG.warning("precarga en RAM falló (%s): se cargará bajo demanda", str(e)[:120])
    threading.Thread(target=_precarga, name="precarga-ram", daemon=True).start()
    t0 = time.time()
    for g in agentes:
        g.start()
    LOG.info("orquestador activo (PID %d, entorno %s, escritura=%s, lote=%d): %s", os.getpid(),
             _entorno(), a.push_drive, BATCH_MAX, ", ".join(g.nombre for g in agentes))
    try:
        while not STOP.exists():
            latido(agentes, t0)
            time.sleep(5)
    except KeyboardInterrupt:
        pass
    PARAR.set()
    for g in agentes:
        g.join(timeout=60)
    latido(agentes, t0)
    PIDFILE.unlink(missing_ok=True)
    LOG.info("orquestador detenido")
    return 0


if __name__ == "__main__":
    sys.exit(main())
