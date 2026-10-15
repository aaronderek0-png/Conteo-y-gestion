"""Registro único de aprobación por circuito (pestaña `Remitos Facturas` / `Remitos_Pendientes`).

El estado humano manda (no el color). Las 16 columnas originales (A:P, ENCABEZADO_PENDIENTES) no se tocan: la
metadata va a partir de la columna Q. Python lee decisiones periódicamente; no hace falta Apps Script.

  Q ESTADO HUMANO  (REVISIÓN | APROBADO | DENEGADO)      R MOTIVO / OBSERVACIÓN     S EVIDENCIA / REFERENCIA
  T FECHA DECISIÓN                                       U RESPONSABLE
  V ESTADO TÉCNICO (PENDIENTE | ESCRITO | ERROR | DENEGADO_CERRADO)
  W ESCRITO EN (planilla/pestaña/rango verificado)       X ID DESTINO (id estable de documento/renglón → destino)

Reglas: APROBADO no vuelve a cargar si V = ESCRITO; DENEGADO conserva evidencia y no toca inventario ni gasto;
REVISIÓN espera. Un ERROR se reintenta solo si sigue APROBADO.
"""
from __future__ import annotations

from datetime import datetime

COLUMNAS_META = ["ESTADO HUMANO", "MOTIVO / OBSERVACIÓN", "EVIDENCIA / REFERENCIA", "FECHA DECISIÓN",
                 "RESPONSABLE", "ESTADO TÉCNICO", "ESCRITO EN", "ID DESTINO"]
PRIMERA_META = 16                      # índice 0-based de la columna Q
HUMANOS = ("REVISIÓN", "APROBADO", "DENEGADO")
TECNICOS = ("PENDIENTE", "ESCRITO", "ERROR", "DENEGADO_CERRADO")


def _norm_humano(x: str) -> str:
    t = (x or "").strip().upper().replace("REVISION", "REVISIÓN")
    return t if t in HUMANOS else ("REVISIÓN" if not t else t)


def leer(valores: list[list]) -> list[dict]:
    """valores = get_all_values() de la pestaña (fila 1 = encabezado). Devuelve un dict por renglón con su n° de fila."""
    out = []
    for n, fila in enumerate(valores[1:], 2):
        fila = list(fila) + [""] * (PRIMERA_META + len(COLUMNAS_META) - len(fila))
        if not fila[0]:
            continue
        out.append({"fila": n, "id": fila[0], "remito": fila[1], "proveedor": fila[3], "producto": fila[4],
                    "humano": _norm_humano(fila[PRIMERA_META]), "motivo": fila[PRIMERA_META + 1],
                    "evidencia": fila[PRIMERA_META + 2], "tecnico": (fila[PRIMERA_META + 5] or "PENDIENTE").upper(),
                    "escrito_en": fila[PRIMERA_META + 6]})
    return out


def planear(decisiones: list[dict]) -> list[dict]:
    """Acción por renglón: CARGAR | CERRAR_SIN_EFECTO | ESPERAR | NADA | INVALIDO (nunca duplica una carga escrita)."""
    plan = []
    for d in decisiones:
        h, t = d["humano"], d["tecnico"]
        if h not in HUMANOS:
            a = "INVALIDO"
        elif h == "APROBADO":
            a = "NADA" if t == "ESCRITO" else "CARGAR"
        elif h == "DENEGADO":
            a = "NADA" if t == "DENEGADO_CERRADO" else "CERRAR_SIN_EFECTO"
        else:
            a = "ESPERAR"
        if a == "CARGAR" and not (d.get("motivo") or d.get("evidencia")):
            a = "ESPERAR"            # aprobar sin motivo ni evidencia deja el caso en revisión
        plan.append({**d, "accion": a})
    return plan


def celdas_resultado(accion: str, ok: bool, escrito_en: str = "", error: str = "", responsable: str = "agente") -> dict:
    """Valores a escribir en V:W (técnico) tras ejecutar `accion`. No cambia el estado humano."""
    if accion == "CERRAR_SIN_EFECTO":
        return {"V": "DENEGADO_CERRADO", "W": "sin efecto en inventario ni gasto"}
    if accion == "CARGAR":
        return {"V": "ESCRITO" if ok else "ERROR", "W": escrito_en if ok else error[:200]}
    return {}


def requests_formato(sheet_id: int, ultima_fila: int = 5000) -> list[dict]:
    """batchUpdate: validación de lista en Q y formato condicional por ESTADO (no por color manual)."""
    rng = {"sheetId": sheet_id, "startRowIndex": 1, "endRowIndex": ultima_fila,
           "startColumnIndex": PRIMERA_META, "endColumnIndex": PRIMERA_META + 1}
    fila_completa = {"sheetId": sheet_id, "startRowIndex": 1, "endRowIndex": ultima_fila,
                     "startColumnIndex": 0, "endColumnIndex": PRIMERA_META + len(COLUMNAS_META)}

    def regla(texto, rgb):
        return {"addConditionalFormatRule": {"index": 0, "rule": {"ranges": [fila_completa], "booleanRule": {
            "condition": {"type": "CUSTOM_FORMULA", "values": [{"userEnteredValue": f'=$Q2="{texto}"'}]},
            "format": {"backgroundColor": {"red": rgb[0], "green": rgb[1], "blue": rgb[2]}}}}}}
    return [
        {"setDataValidation": {"range": rng, "rule": {"condition": {"type": "ONE_OF_LIST", "values": [
            {"userEnteredValue": h} for h in HUMANOS]}, "showCustomUi": True, "strict": True}}},
        regla("REVISIÓN", (1.0, 0.973, 0.882)), regla("APROBADO", (0.9, 0.96, 0.9)), regla("DENEGADO", (1.0, 0.922, 0.933)),
    ]


def ahora() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M")
