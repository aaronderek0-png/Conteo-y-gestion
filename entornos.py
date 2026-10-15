"""Ambientes PRUEBA / PRODUCCIÓN con control por ID (el nombre del archivo NO es el control).

- `entornos.json`: {"produccion": {local: id}, "prueba": {local: id}} (lo completa crear_trials.py o el dueño).
- Variable MISHI_ENTORNO = PRODUCCION (defecto) | PRUEBA.
- En PRUEBA: `ids()` devuelve solo IDs de prueba, y `validar_destino(id)` lanza `EscrituraProhibida` si el ID es de
  producción o no está registrado como prueba. `imputacion_facturas._gspread()` envuelve open_by_key con este control.
- Al registrar una copia se exige: ID distinto de producción, título que empiece con «NO CARGAR - PRUEBA -» y editable.
"""
from __future__ import annotations

import json
import os
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
ARCHIVO = BASE_DIR / "entornos.json"
PREFIJO = "NO CARGAR - PRUEBA - "


class EscrituraProhibida(RuntimeError):
    pass


def _datos() -> dict:
    p = Path(os.environ.get("MISHI_ENTORNOS_JSON") or ARCHIVO)
    try:
        return json.loads(p.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {"produccion": {}, "prueba": {}}


def entorno() -> str:
    e = (os.environ.get("MISHI_ENTORNO") or "PRODUCCION").upper()
    if e not in ("PRUEBA", "PRODUCCION"):
        raise ValueError(f"MISHI_ENTORNO inválido: {e}")
    return e


def ids(produccion_defecto: dict | None = None) -> dict:
    d = _datos()
    if entorno() == "PRUEBA":
        if not d.get("prueba"):
            raise EscrituraProhibida("modo PRUEBA sin IDs de prueba registrados en entornos.json")
        return dict(d["prueba"])
    return dict(d.get("produccion") or produccion_defecto or {})


def ids_produccion(produccion_defecto: dict | None = None) -> set[str]:
    return set((_datos().get("produccion") or produccion_defecto or {}).values())


def validar_destino(sheet_id: str, produccion_defecto: dict | None = None) -> None:
    """Llamar antes de abrir/escribir una planilla. En PRUEBA solo pasan IDs de prueba registrados."""
    if entorno() != "PRUEBA":
        return
    if sheet_id in ids_produccion(produccion_defecto):
        raise EscrituraProhibida(f"modo PRUEBA: el ID {sheet_id[:6]}… es de PRODUCCIÓN; escritura rechazada")
    if sheet_id not in set(_datos().get("prueba", {}).values()):
        raise EscrituraProhibida(f"modo PRUEBA: el ID {sheet_id[:6]}… no está registrado como copia de prueba")


def registrar_prueba(local: str, sheet_id: str, titulo: str, editable: bool, ruta: Path | None = None,
                     produccion_defecto: dict | None = None) -> None:
    d = _datos()
    prod = set((d.get("produccion") or produccion_defecto or {}).values())
    if sheet_id in prod:
        raise EscrituraProhibida("el ID de la copia coincide con uno de producción")
    if not titulo.startswith(PREFIJO):
        raise ValueError(f"el título debe empezar con «{PREFIJO}»")
    if not editable:
        raise ValueError("la copia no es editable por la Service Account")
    d.setdefault("prueba", {})[local] = sheet_id
    d.setdefault("produccion", dict(produccion_defecto or {}))
    p = Path(ruta or os.environ.get("MISHI_ENTORNOS_JSON") or ARCHIVO)
    tmp = p.with_suffix(".tmp")
    tmp.write_text(json.dumps(d, ensure_ascii=False, indent=1), encoding="utf-8")
    tmp.replace(p)
