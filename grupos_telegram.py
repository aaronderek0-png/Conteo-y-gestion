"""Mapa chat de Telegram -> local (1 grupo general o 5 grupos por local).

Fuentes (se fusionan, gana la última): `entornos.json["telegram_chats_locales"]`, variable `TELEGRAM_CHATS_LOCALES`
(JSON {"chat_id": "Local"}) y `estado/chats_locales.json` (lo escribe el comando `/vincular_local <Local>`).
Sin dependencias del resto del proyecto (lo importan el orquestador, el ingestor y procesar_lote_wpp).
"""
from __future__ import annotations

import json
import os
import re
from pathlib import Path

BASE = Path(__file__).resolve().parent
ARCHIVO = Path(os.environ.get("MISHI_ESTADO_DIR") or BASE / "estado") / "chats_locales.json"
LOCALES = ("Arcos", "Congreso", "Lafinur", "Cabello", "Chimidoha")
_SINONIMOS = {"obrador": "Congreso", "nunez": "Chimidoha", "núñez": "Chimidoha", "chimi": "Chimidoha"}
_ARCHIVO_CHAT = re.compile(r"^(-?\d{5,})_\d+_")


def normalizar_local(txt: str | None) -> str | None:
    t = (txt or "").strip().lower()
    if not t:
        return None
    if t in _SINONIMOS:
        return _SINONIMOS[t]
    for l in LOCALES:
        if l.lower() == t:
            return l
    return None


def _json(p: Path) -> dict:
    try:
        d = json.loads(p.read_text(encoding="utf-8"))
        return d if isinstance(d, dict) else {}
    except (OSError, ValueError):
        return {}


def chats_locales() -> dict[str, str]:
    out: dict[str, str] = {}
    out.update(_json(BASE / "entornos.json").get("telegram_chats_locales") or {})
    try:
        out.update(json.loads(os.environ.get("TELEGRAM_CHATS_LOCALES") or "{}"))
    except ValueError:
        pass
    out.update(_json(ARCHIVO))
    return {str(k): v for k, v in out.items() if normalizar_local(v)}


def vincular(chat_id: str | int, local: str) -> str:
    l = normalizar_local(local)
    if not l:
        raise ValueError(f"local desconocido: {local!r} (válidos: {', '.join(LOCALES)})")
    d = _json(ARCHIVO)
    d[str(chat_id)] = l
    ARCHIVO.parent.mkdir(exist_ok=True)
    tmp = ARCHIVO.with_suffix(".tmp")
    tmp.write_text(json.dumps(d, ensure_ascii=False, indent=1), encoding="utf-8")
    tmp.replace(ARCHIVO)
    return l


def local_de_chat(chat_id: str | int | None) -> str | None:
    return chats_locales().get(str(chat_id)) if chat_id is not None else None


def local_de_archivo(nombre: str | None) -> str | None:
    """Los adjuntos descargados se llaman `<chat>_<message_id>_<nombre>`: de ahí sale el grupo de origen."""
    m = _ARCHIVO_CHAT.match(Path(str(nombre or "")).name)
    return local_de_chat(m.group(1)) if m else None
