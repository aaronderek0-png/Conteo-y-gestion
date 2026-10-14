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
    python telegram_ingestor.py --conectar <TOKEN>        # una sola vez: guarda el token en .env, detecta el grupo
                                                          # "Facturas Mishiguene", lo guarda y manda el mensaje de prueba

Las fotos leídas por Tesseract (sin la revisión visual de Claude) nunca se inyectan solas: quedan como REVISAR.
"""

from __future__ import annotations

import argparse
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
STATE_JSON = BASE_DIR / "telegram_state.json"
ENTRADA = BASE_DIR / "Facturas Mishiguene"          # carpeta diaria (sin guiones); lo que llega por Telegram cae acá
ENV_FILE = BASE_DIR / ".env"
NOMBRE_GRUPO = "Facturas Mishiguene"
API = "https://api.telegram.org"
LOG = logging.getLogger("telegram_ingestor")
EXT_OK = {".jpg", ".jpeg", ".png", ".webp", ".pdf"}


# --------------------------------------------------------------------------- #
# API mínima (transporte inyectable para tests)
# --------------------------------------------------------------------------- #
class TelegramAPI:
    def __init__(self, token: str, http: Callable[..., Any] | None = None):
        self.token = token
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

    def responder(self, chat_id: int | str, texto: str, reply_to: int | None = None) -> None:
        p: dict[str, Any] = {"chat_id": chat_id, "text": texto[:4000], "disable_web_page_preview": "true"}
        if reply_to:
            p["reply_to_message_id"] = reply_to
        self._call("sendMessage", **p)


# --------------------------------------------------------------------------- #
# Estado
# --------------------------------------------------------------------------- #
def cargar_estado() -> dict:
    try:
        return json.loads(STATE_JSON.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {"offset": None, "sha256": {}}


def guardar_estado(e: dict) -> None:
    STATE_JSON.write_text(json.dumps(e, ensure_ascii=False, indent=1), encoding="utf-8")


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
def pasada(api: TelegramAPI, escribir: bool = True, chat_permitido: str | None = None, timeout: int = 0) -> dict:
    """Una consulta a Telegram + procesamiento de lo nuevo. Devuelve estadísticas."""
    t0 = time.perf_counter()
    est = cargar_estado()
    ups = api.get_updates(est.get("offset"), timeout=timeout)
    stats = {"updates": len(ups), "archivos": 0, "duplicados": 0, "respondidos": 0, "ms": 0}
    if not ups:
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
        for file_id, nombre in adjuntos(msg):
            ruta = api.descargar(file_id, ENTRADA / nombre)
            h = hashlib.sha256(ruta.read_bytes()).hexdigest()
            if h in est["sha256"]:
                stats["duplicados"] += 1
                ruta.unlink(missing_ok=True)
                api.responder(chat, "Ya lo tenía procesado (mismo contenido).", msg["message_id"])
                continue
            est["sha256"][h] = {"archivo": ruta.name, "chat": chat, "message_id": msg["message_id"]}
            pendientes.append((ruta, {"chat": chat, "mid": msg["message_id"]}))
            stats["archivos"] += 1
    if pendientes:
        import procesar_lote_wpp as lote
        tmp = BASE_DIR / "Carpeta Revision" / "lote actual temporal"
        tmp.mkdir(exist_ok=True)
        for p, _ in pendientes:
            shutil.copy2(p, tmp / p.name)
        ns = SimpleNamespace(push_drive=False, confirmo_escribir_produccion=False, dry_run=not escribir, json=None)
        try:
            reporte, *_ = lote.ejecutar(tmp, escribir, ns, indexar_informes=False)
            for p, meta in pendientes:
                api.responder(meta["chat"], texto_respuesta(reporte, [p.name]), meta["mid"])
                stats["respondidos"] += 1
        except Exception as e:                                   # un archivo roto no frena el resto
            LOG.exception("falló el procesamiento")
            for _, meta in pendientes:
                api.responder(meta["chat"], f"No pude procesar el comprobante ({type(e).__name__}). Queda para revisión.",
                              meta["mid"])
        finally:
            shutil.rmtree(tmp, ignore_errors=True)
    est["offset"] = ultimo
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


def conectar(token: str, http: Callable[..., Any] | None = None, escribir: bool = True) -> dict:
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
    stats = pasada(api, escribir=escribir, chat_permitido=str(chat_id))
    return {"conectado": True, "bot": me.get("username"), "grupo": elegido.get("title"), "chat_id": chat_id,
            "coincide_nombre": _norm_titulo(elegido.get("title")) == _norm_titulo(NOMBRE_GRUPO), "pasada": stats}


def crear_tarea(minutos: int = 15) -> list[str]:
    cmd = ["schtasks", "/Create", "/SC", "MINUTE", "/MO", str(minutos), "/TN", "Mishiguene_Telegram_Ingestor",
           "/TR", f'"{sys.executable}" "{Path(__file__).resolve()}" --once', "/F"]
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
    ap.add_argument("--conectar", metavar="TOKEN", help="Guarda el token en .env, detecta el grupo y manda el mensaje de prueba")
    args = ap.parse_args(argv)
    cargar_env()

    if args.conectar:
        r = conectar(args.conectar.strip(), escribir=not args.dry_run)
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
    if args.once or not args.watch:
        s = pasada(api, escribir=not args.dry_run, chat_permitido=chat)
        LOG.info("pasada: %s", s)
        return 0
    LOG.info("modo vivo cada %d s (Ctrl+C para salir)", args.interval)
    while True:
        try:
            s = pasada(api, escribir=not args.dry_run, chat_permitido=chat)
            if s["updates"]:
                LOG.info("pasada: %s", s)
        except Exception:
            LOG.exception("error en la pasada; se reintenta")
        time.sleep(args.interval)


if __name__ == "__main__":
    sys.exit(main())
