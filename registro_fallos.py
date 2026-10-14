"""Registro persistente y saneado de fallos del pipeline (Telegram -> lectura -> validación -> escritura).

Una línea JSON por evento en `logs/pipeline_eventos.jsonl`: ID de trabajo (sha del comprobante), etapa, tipo de
fallo, intento, resultado y mensaje saneado. Nunca guarda tokens, claves, URLs con credenciales ni el documento.
"""
from __future__ import annotations

import json
import os
import re
import time
from datetime import datetime
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
LOG_DIR = Path(os.environ.get("MISHI_LOG_DIR") or BASE_DIR / "logs")
EVENTOS = LOG_DIR / "pipeline_eventos.jsonl"

_SECRETOS = [
    (re.compile(r"bot\d+:[A-Za-z0-9_-]{20,}"), "bot<TOKEN>"),
    (re.compile(r"(?i)(api[_-]?key|token|authorization|bearer|password|private_key)[\"'=: ]+[^\s\"',}]+"), r"\1=<oculto>"),
    (re.compile(r"https?://[^\s\"']*[?&](key|token|access_token|apikey)=[^\s\"']+"), "<url con credenciales>"),
    (re.compile(r"-----BEGIN [A-Z ]+-----.*?-----END [A-Z ]+-----", re.S), "<clave privada>"),
]


def sanear(texto, largo: int = 400) -> str:
    t = str(texto)
    for rx, rep in _SECRETOS:
        t = rx.sub(rep, t)
    return t.replace("\n", " ")[:largo]


def detalle_excepcion(e: BaseException) -> dict:
    """Tipo, módulo y, si es un APIError de gspread/Google, status HTTP + mensaje del servidor + recurso saneados."""
    d = {"tipo": type(e).__name__, "modulo": type(e).__module__, "mensaje": sanear(e)}
    resp = getattr(e, "response", None)
    if resp is not None:
        try:
            d["http_status"] = resp.status_code
            err = resp.json().get("error", {})
            d["http_status_texto"] = err.get("status")
            d["http_mensaje"] = sanear(err.get("message", ""))
            req = getattr(resp, "request", None)
            if req is not None:
                d["metodo"] = req.method
                d["endpoint"] = re.sub(r"/spreadsheets/[A-Za-z0-9_-]+", "/spreadsheets/<id>", str(req.url).split("?")[0])
        except Exception:
            pass
    return d


def evento(trabajo: str, etapa: str, resultado: str, *, intento: int = 1, error: BaseException | None = None,
           **extra) -> dict:
    reg = {"ts": datetime.now().isoformat(timespec="seconds"), "trabajo": str(trabajo)[:16], "etapa": etapa,
           "intento": intento, "resultado": resultado}
    if error is not None:
        reg["error"] = detalle_excepcion(error)
    reg.update({k: sanear(v, 200) if isinstance(v, str) else v for k, v in extra.items()})
    try:
        LOG_DIR.mkdir(exist_ok=True)
        with EVENTOS.open("a", encoding="utf-8") as f:
            f.write(json.dumps(reg, ensure_ascii=False) + "\n")
    except OSError:
        pass
    return reg


def con_reintentos(fn, *, trabajo: str, etapa: str, intentos: int = 3, espera: float = 2.0,
                   transitorios=(429, 500, 502, 503, 504)):
    """Ejecuta fn() reintentando solo fallos transitorios (HTTP 429/5xx, timeouts). Registra cada intento."""
    for n in range(1, intentos + 1):
        try:
            r = fn()
            evento(trabajo, etapa, "OK", intento=n)
            return r
        except Exception as e:  # noqa: BLE001
            st = getattr(getattr(e, "response", None), "status_code", None)
            transitorio = st in transitorios or isinstance(e, (TimeoutError, ConnectionError))
            evento(trabajo, etapa, "ERROR_TRANSITORIO" if transitorio and n < intentos else "ERROR", intento=n, error=e)
            if not transitorio or n >= intentos:
                raise
            time.sleep(espera * n)
