"""Ingesta de comprobantes por Telegram (Bot API oficial, sin dependencias nuevas: solo `requests`).

Cada foto/PDF que entra al grupo se baja, se identifica por SHA-256 (nunca se procesa dos veces) y pasa por todo el
pipeline de procesar_lote_wpp.py (local por sello/edificio, historial, remitos con birome, NC, prorrateo, imputación
en las COPIAS). El bot responde en el chat qué detectó.

Costo cero en reposo: `getUpdates` con `offset` guardado en telegram_state.json; sin mensajes nuevos vuelve en una
sola llamada HTTP (<100 ms de CPU) y no consume ningún modelo ni token.

Configuración (variables de entorno, nunca en el código):
    TELEGRAM_BOT_TOKEN   token del bot (@BotFather)
    TELEGRAM_CHAT_ID     (opcional) solo se atiende ese chat/grupo

Uso:
    python telegram_ingestor.py --once                    # una pasada (Tarea Programada cada 15 min)
    python telegram_ingestor.py --watch --interval 60     # modo vivo
    python telegram_ingestor.py --setup-task [--minutos 15]   # crea la Tarea Programada de Windows
    python telegram_ingestor.py --once --dry-run          # procesa sin escribir en las copias
    python telegram_ingestor.py --limpiar                 # borra para todos los mensajes del grupo, salvo el fijado
    (en el grupo, escribir /limpiar hace lo mismo)
    python telegram_ingestor.py --conectar <TOKEN>        # una sola vez: guarda el token en .env, detecta el grupo
                                                          # "Facturas Mishiguene", lo guarda y manda el mensaje de prueba

Las fotos leídas por Tesseract (sin la revisión visual de Claude) nunca se inyectan solas: quedan como REVISAR.
"""

from __future__ import annotations

import argparse
import re
import hashlib
import json
import logging
import os
import shutil
import subprocess
import sys
import time
from pathlib import Path
from types import SimpleNamespace
from typing import Any, Callable

BASE_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(BASE_DIR))
STATE_JSON = BASE_DIR / "telegram_state.json"
ENTRADA = BASE_DIR / "Facturas Mishiguene"          # carpeta diaria (sin guiones); lo que llega por Telegram cae acá
ENV_FILE = BASE_DIR / ".env"
NOMBRE_GRUPO = "Facturas Mishiguene"
API = "https://api.telegram.org"
LOG = logging.getLogger("telegram_ingestor")
DEMO_RE = re.compile(r"\b(demo|prueba)\b", re.I)


def deshacer_push(registro):
    import imputacion_facturas as imp
    return imp.deshacer_push(registro)


def texto_push(res: dict) -> str:
    """Resumen legible de una escritura en la planilla real (fila exacta, valores y color)."""
    if not res:
        return ""
    if not res.get("escrito"):
        return f"ℹ Planilla real {res.get('local', '')}: no escribí ({res.get('motivo', '')})."
    hoja = res.get("hoja", "Facturas")
    filas = (f"fila {res['desde_fila']}" if res["desde_fila"] == res["hasta_fila"]
             else f"filas {res['desde_fila']} a {res['hasta_fila']}")
    l = [f"✍ Escribí en Google Sheets «{res.get('planilla')}», pestaña {hoja}, {filas}"
         + (" (DEMO: se borra con /deshacer)" if res.get("demo") else "") + ":"]
    for d in res["detalle"][:8]:
        if res.get("tipo") == "remitos":
            l.append(f"   fila {d['fila']}: {d['FECHA']} · {d['PRODUCTO']} · cant {d['CANT']} · ${d['PRECIO U']} · total ${d['TOTAL']}")
        else:
            l.append(f"   fila {d['fila']}: N° {d['N']} · {d['proveedor']} · {d['E']} · cant {d['G']} · ${d['H']} · {d['color']}")
    if res.get("omitidas"):
        l.append(f"   (no dupliqué {len(res['omitidas'])} fila(s) que ya estaban cargadas)")
    return "\n".join(l)
EXT_OK = {".jpg", ".jpeg", ".png", ".webp", ".pdf"}


# --------------------------------------------------------------------------- #
# API mínima (transporte inyectable para tests)
# --------------------------------------------------------------------------- #
class TelegramAPI:
    def __init__(self, token: str, http: Callable[..., Any] | None = None):
        self.token = token
        self.enviados: list[tuple[str, int]] = []
        if http is None:
            import requests
            http = requests.request
        self._http = http

    def _call(self, metodo: str, **params) -> dict:
        r = self._http("POST", f"{API}/bot{self.token}/{metodo}", data=params, timeout=30)
        r.raise_for_status()
        d = r.json()
        if not d.get("ok"):
            raise RuntimeError(f"Telegram {metodo}: {d.get('description')}")
        return d["result"]

    def get_updates(self, offset: int | None, timeout: int = 0) -> list[dict]:
        p: dict[str, Any] = {"timeout": timeout, "allowed_updates": json.dumps(["message", "channel_post"])}
        if offset is not None:
            p["offset"] = offset
        return self._call("getUpdates", **p)

    def descargar(self, file_id: str, destino: Path) -> Path:
        info = self._call("getFile", file_id=file_id)
        r = self._http("GET", f"{API}/file/bot{self.token}/{info['file_path']}", timeout=60)
        r.raise_for_status()
        destino = destino.with_suffix(destino.suffix or Path(info["file_path"]).suffix)
        destino.write_bytes(r.content)
        return destino

    def responder(self, chat_id: int | str, texto: str, reply_to: int | None = None) -> dict:
        p: dict[str, Any] = {"chat_id": chat_id, "text": texto[:4000], "disable_web_page_preview": "true"}
        if reply_to:
            p["reply_to_message_id"] = reply_to
        r = self._call("sendMessage", **p)
        if isinstance(r, dict) and r.get("message_id"):
            self.enviados.append((str(chat_id), int(r["message_id"])))     # se persiste en telegram_state.json
        return r

    def get_chat(self, chat_id: int | str) -> dict:
        return self._call("getChat", chat_id=chat_id)

    def borrar(self, chat_id: int | str, message_id: int) -> bool:
        """deleteMessage (para todos los participantes). False si Telegram lo rechaza (ya no existe, sin permiso)."""
        try:
            self._call("deleteMessage", chat_id=chat_id, message_id=message_id)
            return True
        except Exception:
            return False


# --------------------------------------------------------------------------- #
# Estado
# --------------------------------------------------------------------------- #
def cargar_estado() -> dict:
    try:
        return json.loads(STATE_JSON.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {"offset": None, "sha256": {}}


def registrar_ids(est: dict, chat: int | str, ids: list[int]) -> None:
    lista = est.setdefault("mensajes", {}).setdefault(str(chat), [])
    vistos = set(lista)
    lista.extend(i for i in ids if i not in vistos)


def limpiar(api: "TelegramAPI", chat_id: int | str, est: dict, extra_ids: list[int] | None = None,
            max_barrido: int = 5000) -> dict:
    """Borra para todos los participantes todos los mensajes del grupo salvo el fijado (`pinned_message` de getChat).
    Usa los message_id guardados (entrantes + los del bot) y además barre 1..máximo conocido, para llevarse también lo
    anterior al registro. Telegram devuelve en getChat solo el último mensaje fijado."""
    fijado = (api.get_chat(chat_id).get("pinned_message") or {}).get("message_id")
    conocidos = set(est.get("mensajes", {}).get(str(chat_id), [])) | set(extra_ids or [])
    tope = min(max(conocidos | {0}), max_barrido) if conocidos else 0
    objetivo = sorted((conocidos | set(range(1, tope + 1))) - ({fijado} if fijado else set()), reverse=True)
    borrados = 0
    for mid in objetivo:
        if api.borrar(chat_id, mid):
            borrados += 1
        time.sleep(0.03)                       # sin saturar el límite de la API
    est.setdefault("mensajes", {})[str(chat_id)] = [fijado] if fijado else []
    return {"chat": str(chat_id), "fijado": fijado, "intentados": len(objetivo), "borrados": borrados}


def guardar_estado(e: dict) -> None:
    temporal = STATE_JSON.with_suffix(".tmp")
    temporal.write_text(json.dumps(e, ensure_ascii=False, indent=1), encoding="utf-8")
    temporal.replace(STATE_JSON)


def adjuntos(msg: dict) -> list[tuple[str, str]]:
    """[(file_id, nombre sugerido)] de un mensaje: foto (la de mayor tamaño) o documento imagen/PDF."""
    out = []
    if msg.get("photo"):
        mejor = max(msg["photo"], key=lambda p: p.get("file_size", 0) or p.get("width", 0))
        out.append((mejor["file_id"], f"tg_{msg['message_id']}_{mejor['file_unique_id']}.jpg"))
    doc = msg.get("document")
    if doc and (Path(doc.get("file_name", "")).suffix.lower() in EXT_OK
                or str(doc.get("mime_type", "")).startswith(("image/", "application/pdf"))):
        out.append((doc["file_id"], doc.get("file_name") or f"tg_{msg['message_id']}_{doc['file_unique_id']}"))
    return out


# --------------------------------------------------------------------------- #
# Texto de respuesta
# --------------------------------------------------------------------------- #
def texto_respuesta(reporte: dict, archivos: list[str]) -> str:
    lineas = []
    nombres = set(archivos)
    for a in reporte["auditoria"]:
        if a["archivo"] not in nombres and not any(a["archivo"] == n for n in nombres):
            continue
        l = [f"📄 {a['tipo']} {a.get('numero') or ''} — {a.get('proveedor') or '?'}"]
        l.append(f"   Local: {a.get('local') or 'INDETERMINADO'}" + (f"  ({a['regla']})" if a.get("regla") else ""))
        l.append(f"   Estado: {a['estado']}")
        apr = a.get("aprendido") or {}
        if apr.get("regla"):
            l.append(f"   Regla aprendida: {apr['regla']}")
        for al in (a.get("alertas") or [])[:4]:
            l.append(f"   ⚠ {al}")
        lineas += l
    for rem in reporte.get("remitos", []):
        if rem.get("archivo") in nombres:
            lineas.append(f"🧾 Remito {rem['remito']}: ${rem['total_impreso']:,.0f} impreso → ${rem['total_efectivo']:,.0f} efectivo")
            for r in rem["renglones"]:
                if r["anulado"]:
                    lineas.append(f"   ✖ {r['descripcion']}: anulado a mano con 'NO'")
                elif r.get("peso_real_recibido_kg"):
                    lineas.append(f"   ⚖ {r['nota']}")
            for al in rem.get("alertas") or []:
                lineas.append(f"   🔔 {al}")
    for n, plan in (reporte.get("planes") or {}).items():
        for f in plan:
            av = "; ".join(f.get("avisos") or [])
            if "prorrate" in av.lower():
                lineas.append(f"➗ {n}: {av}")
    return "\n".join(lineas) or "Recibido, sin novedades."


# --------------------------------------------------------------------------- #
# Pasada
# --------------------------------------------------------------------------- #
def pasada(api: TelegramAPI, escribir: bool = True, chat_permitido: str | None = None, timeout: int = 0,
           push: bool = False) -> dict:
    """Una consulta a Telegram + procesamiento de lo nuevo. Devuelve estadísticas."""
    t0 = time.perf_counter()
    push = push and escribir
    est = cargar_estado()
    offset_inicial = est.get("offset")
    fallo = False
    ups = api.get_updates(est.get("offset"), timeout=timeout)
    stats = {"updates": len(ups), "archivos": 0, "duplicados": 0, "respondidos": 0, "ms": 0}
    if not ups:
        if api.enviados:                                            # mensajes del bot enviados antes de esta consulta
            for ch, mid in api.enviados:
                registrar_ids(est, ch, [mid])
            api.enviados.clear()
            if escribir:
                guardar_estado(est)
        stats["ms"] = round((time.perf_counter() - t0) * 1000)
        return stats
    ENTRADA.mkdir(parents=True, exist_ok=True)
    pendientes: list[tuple[Path, dict]] = []
    ultimo = est.get("offset")
    for u in ups:
        ultimo = max(ultimo or 0, u["update_id"] + 1)
        msg = u.get("message") or u.get("channel_post")
        if not msg:
            continue
        chat = msg["chat"]["id"]
        if chat_permitido and str(chat) != str(chat_permitido):
            continue
        registrar_ids(est, chat, [msg["message_id"]])                # todo lo que entra al grupo
        texto_msg = (msg.get("text") or "").strip().lower()
        if not escribir and texto_msg.split("@")[0] in ("/limpiar", "/deshacer"):
            continue
        if texto_msg.split("@")[0] == "/limpiar":
            r = limpiar(api, chat, est, [msg["message_id"]])
            stats["limpiado"] = r
            est["offset"] = ultimo
            api.responder(chat, f"🧹 Listo: {r['borrados']} mensajes borrados"
                                + (" (el fijado se conserva)." if r["fijado"] else "."))
            continue
        if texto_msg.split("@")[0] == "/deshacer":
            r = deshacer_push(est.get("ultima_inyeccion"))
            if r.get("deshecho"):
                est["ultima_inyeccion"] = None
                api.responder(chat, f"↩ Deshecho: borré {r['filas']} fila(s) ({r['desde_fila']}-{r['hasta_fila']}) de "
                                    f"«{r['planilla']}».", msg["message_id"])
            else:
                api.responder(chat, f"No deshice nada: {r['motivo']}", msg["message_id"])
            continue
        demo = bool(DEMO_RE.search(msg.get("caption") or ""))
        for file_id, nombre in adjuntos(msg):
            nombre = Path(nombre).name
            ruta = api.descargar(file_id, ENTRADA / f"{chat}_{msg['message_id']}_{nombre}")
            h = hashlib.sha256(ruta.read_bytes()).hexdigest()
            if demo:                                              # DEMO / PRUEBA: se re-procesa aunque ya esté en caché
                pendientes.append((ruta, {"chat": chat, "mid": msg["message_id"], "demo": True}))
                stats["archivos"] += 1
                continue
            if h in est["sha256"]:
                stats["duplicados"] += 1
                if escribir:
                    api.responder(chat, "Ya lo tenía procesado (mismo contenido).", msg["message_id"])
                continue
            pendientes.append((ruta, {"chat": chat, "mid": msg["message_id"]}))
            stats["archivos"] += 1
    if pendientes:
        import procesar_lote_wpp as lote
        for demo_grupo in (False, True):
            grupo = [(p, m) for p, m in pendientes if bool(m.get("demo")) == demo_grupo]
            if not grupo:
                continue
            tmp = BASE_DIR / "Carpeta Revision" / ("lote demo temporal" if demo_grupo else "lote actual temporal")
            shutil.rmtree(tmp, ignore_errors=True)
            tmp.mkdir(parents=True, exist_ok=True)
            for p, _ in grupo:
                shutil.copy2(p, tmp / p.name)
            ns = SimpleNamespace(push_drive=push, confirmo_escribir_produccion=push, dry_run=not escribir, json=None)
            try:
                reporte, *_ = lote.ejecutar(tmp, escribir and not demo_grupo, ns, indexar_informes=False,
                                            forzar_nuevas=demo_grupo, demo=demo_grupo)
                pushes = [r for r in (reporte.get("push_drive") or {}).values() if r]
                errores_push = [r for r in pushes if r.get("error")]
                if errores_push:                              # no se registra SHA como éxito ni se avanza el offset
                    class _PushFallido(RuntimeError):
                        pass
                    e0 = errores_push[0]["error"]
                    raise _PushFallido(f"{e0.get('tipo')} HTTP {e0.get('http_status')} {e0.get('http_mensaje', '')}")
                escritos = [r for r in pushes if r.get("escrito") and r.get("detalle") and r.get("tipo") != "remitos_proveedor"]
                if escritos:
                    est["ultima_inyeccion"] = escritos[-1]       # para /deshacer
                resumen_push = "\n".join(texto_push(r) for r in pushes)
                for p, meta in grupo:
                    txt = texto_respuesta(reporte, [p.name])
                    for rr in reporte.get("remitos_obrador") or []:
                        txt += f"\n🚚 Remito Obrador {rr['remito']} → {rr['receptor']}: {len(rr['renglones'])} renglones, ${rr['total']:,.2f}"
                    if resumen_push:
                        txt += "\n" + resumen_push
                    if escribir:
                        api.responder(meta["chat"], txt, meta["mid"])
                        stats["respondidos"] += 1
                    if escribir and not demo_grupo:
                        h = hashlib.sha256(p.read_bytes()).hexdigest()
                        est["sha256"][h] = {"archivo": p.name, "chat": meta["chat"], "message_id": meta["mid"]}
            except Exception as e:                               # un archivo roto no frena el resto
                LOG.exception("falló el procesamiento")
                fallo = True
                import registro_fallos as rf
                det = rf.detalle_excepcion(e)
                for p_, _m in grupo:
                    rf.evento(hashlib.sha256(p_.read_bytes()).hexdigest(), "procesar_lote", "ERROR", error=e,
                              archivo=p_.name, demo=demo_grupo)
                cod = f" HTTP {det['http_status']} {det.get('http_status_texto', '')}".rstrip() if det.get("http_status") else ""
                for _, meta in grupo:
                    if not escribir:
                        continue
                    api.responder(meta["chat"], f"No pude procesar el comprobante ({det['tipo']}{cod}). Queda en cola "
                                                "para reintentar; detalle en logs/pipeline_eventos.jsonl.", meta["mid"])
            finally:
                shutil.rmtree(tmp, ignore_errors=True)
    est["offset"] = offset_inicial if fallo else ultimo
    for ch, mid in api.enviados:                                   # los mensajes del bot también se registran
        registrar_ids(est, ch, [mid])
    api.enviados.clear()
    if escribir:
        guardar_estado(est)
    stats["ms"] = round((time.perf_counter() - t0) * 1000)
    return stats


# --------------------------------------------------------------------------- #
# .env y auto-conexión
# --------------------------------------------------------------------------- #
def cargar_env(path: Path | None = None) -> None:
    """Carga KEY=VALUE de .env en os.environ (sin pisar lo que ya esté definido)."""
    p = path or ENV_FILE
    if not p.is_file():
        return
    for linea in p.read_text(encoding="utf-8").splitlines():
        linea = linea.strip()
        if linea and not linea.startswith("#") and "=" in linea:
            k, v = linea.split("=", 1)
            os.environ.setdefault(k.strip(), v.strip().strip('"').strip("'"))


def guardar_env(clave: str, valor: str, path: Path | None = None) -> None:
    p = path or ENV_FILE
    lineas = p.read_text(encoding="utf-8").splitlines() if p.is_file() else []
    nuevas, visto = [], False
    for l in lineas:
        if l.split("=", 1)[0].strip() == clave:
            nuevas.append(f"{clave}={valor}")
            visto = True
        else:
            nuevas.append(l)
    if not visto:
        nuevas.append(f"{clave}={valor}")
    p.write_text("\n".join(nuevas) + "\n", encoding="utf-8")
    os.environ[clave] = valor


def _norm_titulo(t: str) -> str:
    return " ".join(str(t or "").lower().replace("_", " ").replace("-", " ").split())


def conectar(token: str, http: Callable[..., Any] | None = None, escribir: bool = True, push: bool = False) -> dict:
    """--conectar TOKEN: guarda el token en .env, detecta el chat_id del grupo 'Facturas Mishiguene' (o el último grupo
    donde hayan hablado), lo guarda como TELEGRAM_CHAT_ID, manda el mensaje de prueba y procesa lo pendiente."""
    api = TelegramAPI(token, http)
    me = api._call("getMe")
    guardar_env("TELEGRAM_BOT_TOKEN", token)
    ups = api.get_updates(None, timeout=0)          # sin confirmar: pasada() vuelve a leerlos para procesarlos
    grupos: list[dict] = []
    for u in ups:
        for k in ("message", "channel_post", "my_chat_member", "edited_message"):
            ch = (u.get(k) or {}).get("chat")
            if ch and ch.get("type") in ("group", "supergroup", "channel"):
                grupos.append(ch)
    if not grupos:
        return {"conectado": False, "bot": me.get("username"),
                "motivo": f"el bot todavía no vio ningún grupo: agregá @{me.get('username')} al grupo "
                          f"'{NOMBRE_GRUPO}', escribí cualquier mensaje y volvé a correr el comando "
                          "(si el bot tiene 'Group Privacy' activada, desactivala en @BotFather o hacelo admin)"}
    elegido = next((g for g in reversed(grupos) if _norm_titulo(g.get("title")) == _norm_titulo(NOMBRE_GRUPO)), grupos[-1])
    chat_id = elegido["id"]
    guardar_env("TELEGRAM_CHAT_ID", str(chat_id))
    api.responder(chat_id, f"✅ Agente Mishiguene conectado a {NOMBRE_GRUPO}. "
                           "Mandá fotos o PDFs acá y los proceso solos.")
    stats = pasada(api, escribir=escribir, chat_permitido=str(chat_id), push=push)
    return {"conectado": True, "bot": me.get("username"), "grupo": elegido.get("title"), "chat_id": chat_id,
            "coincide_nombre": _norm_titulo(elegido.get("title")) == _norm_titulo(NOMBRE_GRUPO), "pasada": stats}


def crear_tarea(minutos: int = 15) -> list[str]:
    cmd = ["schtasks", "/Create", "/SC", "MINUTE", "/MO", str(minutos), "/TN", "Mishiguene_Telegram_Ingestor",
           "/TR", f'"{sys.executable}" "{Path(__file__).resolve()}" --once --push-drive', "/F"]
    return cmd


def main(argv: list[str] | None = None) -> int:
    logging.basicConfig(level=logging.INFO, format="%(asctime)s | %(levelname)s | %(message)s", datefmt="%H:%M:%S")
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    g = ap.add_mutually_exclusive_group()
    g.add_argument("--once", action="store_true")
    g.add_argument("--watch", action="store_true")
    g.add_argument("--setup-task", action="store_true", help="Crea la Tarea Programada de Windows (cada --minutos)")
    ap.add_argument("--interval", type=int, default=60, help="Segundos entre consultas en --watch")
    ap.add_argument("--minutos", type=int, default=15)
    ap.add_argument("--dry-run", action="store_true", help="Procesa sin escribir en las copias")
    ap.add_argument("--push-drive", action="store_true",
                    help="Escribe EN VIVO en la planilla real del local (amarillo #FFF8E1 OK / rojo #FFEBEE revisar)")
    ap.add_argument("--limpiar", action="store_true",
                    help="Borra para todos los mensajes del grupo salvo el fijado (también: /limpiar en el grupo)")
    ap.add_argument("--conectar", metavar="TOKEN", help="Guarda el token en .env, detecta el grupo y manda el mensaje de prueba")
    args = ap.parse_args(argv)
    cargar_env()

    if args.conectar:
        r = conectar(args.conectar.strip(), escribir=not args.dry_run, push=args.push_drive)
        if r["conectado"]:
            print(f"Conectado: bot @{r['bot']} -> grupo '{r['grupo']}' (chat_id {r['chat_id']}); .env actualizado. "
                  f"Pendientes procesados: {r['pasada']}")
            if not r["coincide_nombre"]:
                print(f"Aviso: el grupo detectado no se llama '{NOMBRE_GRUPO}' (se tomó el último grupo con actividad).")
            return 0
        print("No conectado:", r["motivo"])
        return 3

    if args.setup_task:
        cmd = crear_tarea(args.minutos)
        print("Ejecutando:", " ".join(cmd))
        return subprocess.call(cmd)
    token = os.environ.get("TELEGRAM_BOT_TOKEN")
    if not token:
        print("Falta TELEGRAM_BOT_TOKEN en el entorno (setx TELEGRAM_BOT_TOKEN \"123:ABC\"; luego abrir otra terminal).")
        return 2
    api = TelegramAPI(token)
    chat = os.environ.get("TELEGRAM_CHAT_ID")
    if args.limpiar:
        if not chat:
            print("Falta TELEGRAM_CHAT_ID: correr primero --conectar <TOKEN>.")
            return 2
        est = cargar_estado()
        r = limpiar(api, chat, est)
        guardar_estado(est)
        print(f"Limpieza: {r['borrados']} mensajes borrados de {r['intentados']} intentados "
              f"(fijado conservado: {r['fijado']}).")
        return 0
    if args.once or not args.watch:
        s = pasada(api, escribir=not args.dry_run, chat_permitido=chat, push=args.push_drive)
        LOG.info("pasada: %s", s)
        return 0
    LOG.info("modo vivo cada %d s (Ctrl+C para salir)", args.interval)
    while True:
        try:
            s = pasada(api, escribir=not args.dry_run, chat_permitido=chat, push=args.push_drive)
            if s["updates"]:
                LOG.info("pasada: %s", s)
        except Exception:
            LOG.exception("error en la pasada; se reintenta")
        time.sleep(args.interval)


if __name__ == "__main__":
    sys.exit(main())
