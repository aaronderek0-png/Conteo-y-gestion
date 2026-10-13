"""Conector a Google Sheets / Google Drive para el agente de consumos Mishiguene.

Se autentica con una Service Account (credentials.json), lee la pestaña
'Facturas' de un Google Sheet y devuelve un DataFrame de Pandas con todas las
filas cuya 'Categoría Gasto' sea 'Materia prima' o 'Descartables/packaging' (incluye todas
las subcategorías de 'Categoría 2': Almacén, Carnes, Frutas y Verduras, etc.).

Las columnas se toman por POSICIÓN (no por nombre de encabezado), porque los
encabezados del Sheet pueden variar o repetirse:
    D (3)  -> Proveedor
    E (4)  -> Productos
    N (13) -> Cantidad Inventario
    O (14) -> Precio Unitario
    P (15) -> Mes factura (entero; se descartan filas sin mes)
    T (19) -> Categoría Gasto
    U (20) -> Categoría 2
    X (23) -> Código SKU

Uso:
    python drive_connector.py <SPREADSHEET_ID>
    python drive_connector.py <SPREADSHEET_ID> --creds credentials.json --out facturas.xlsx
"""

from __future__ import annotations

import argparse
import logging
import re
import sys
import unicodedata
from pathlib import Path

import gspread
import pandas as pd
from google.oauth2.service_account import Credentials

LOG = logging.getLogger("drive_connector")

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/drive.readonly",
]

DEFAULT_CREDS = Path(__file__).with_name("credentials.json")
SHEET_TAB = "Facturas"

# Índice de columna (0-based) -> nombre normalizado en el DataFrame
COLUMN_MAP = {
    3: "Proveedor",
    4: "Productos",
    11: "Unidad Medida",
    13: "Cantidad Inventario",
    14: "Precio Unitario",
    15: "Mes factura",
    19: "Categoría Gasto",
    20: "Categoría 2",
    23: "Código SKU",
}

# Valores exactos de la columna T en los Sheets
CATEGORIAS_GASTO_OBJETIVO = ["Materia prima", "Descartables/packaging"]


# --------------------------------------------------------------------------- #
# Autenticación
# --------------------------------------------------------------------------- #
def get_client(creds_path: str | Path = DEFAULT_CREDS) -> gspread.Client:
    """Devuelve un cliente gspread autenticado con la Service Account."""
    creds_path = Path(creds_path)
    if not creds_path.exists():
        raise FileNotFoundError(
            f"No se encontró {creds_path}. Descargá la clave JSON de la Service Account."
        )
    creds = Credentials.from_service_account_file(str(creds_path), scopes=SCOPES)
    return gspread.authorize(creds)


# --------------------------------------------------------------------------- #
# Helpers
# --------------------------------------------------------------------------- #
def _normalize(text: str) -> str:
    """Minúsculas y sin acentos: 'Lácteos' -> 'lacteos'."""
    text = unicodedata.normalize("NFKD", str(text))
    return "".join(c for c in text if not unicodedata.combining(c)).lower().strip()


def _to_number(value: str) -> float | None:
    """Convierte '1.234,56' / '1,234.56' / '12' a float. Devuelve None si no es numérico."""
    s = re.sub(r"[^\d,.\-]", "", str(value))
    if not s or s in {"-", ".", ","}:
        return None
    if "," in s and "." in s:
        # El último separador es el decimal
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        s = s.replace(",", ".")
    elif re.fullmatch(r"-?\d{1,3}(\.\d{3})+", s):
        # Formato AR: '$12.096' -> 12096 (el punto es separador de miles)
        s = s.replace(".", "")
    try:
        return float(s)
    except ValueError:
        return None


# --------------------------------------------------------------------------- #
# Lectura + filtrado
# --------------------------------------------------------------------------- #
def read_facturas(
    spreadsheet_id: str,
    creds_path: str | Path = DEFAULT_CREDS,
    categorias_gasto: list[str] | None = None,
) -> pd.DataFrame:
    """Lee la pestaña 'Facturas' y filtra por 'Categoría Gasto' (columna T).

    Se conserva una fila si 'Categoría Gasto' coincide con alguna de las
    categorías objetivo (comparación exacta, sin acentos ni mayúsculas).
    'Categoría 2' no se filtra: entran todas sus subcategorías.
    """
    client = get_client(creds_path)
    ws = client.open_by_key(spreadsheet_id).worksheet(SHEET_TAB)
    # get_all_values() en lugar de get_all_records(): tolera encabezados duplicados/vacíos
    return filas_a_dataframe(ws.get_all_values(), categorias_gasto)


def read_facturas_xlsx(
    path: str | Path,
    sheet: str = SHEET_TAB,
    categorias_gasto: list[str] | None = None,
) -> pd.DataFrame:
    """Igual que read_facturas pero desde un .xlsx descargado (copia local de la planilla)."""
    raw = pd.read_excel(path, sheet_name=sheet, header=None, dtype=object).fillna("")

    def _txt(v: object) -> str:
        # Los decimales se pasan con coma: _to_number leería "0.125" como 125 (punto de miles)
        if isinstance(v, float):
            return str(int(v)) if v.is_integer() else repr(v).replace(".", ",")
        return str(v)

    return filas_a_dataframe([[_txt(v) for v in r] for r in raw.values.tolist()], categorias_gasto)


def filas_a_dataframe(rows: list[list[str]], categorias_gasto: list[str] | None = None) -> pd.DataFrame:
    """Filas crudas de la pestaña 'Facturas' (la primera es el encabezado) -> DataFrame filtrado."""
    categorias_gasto = categorias_gasto or CATEGORIAS_GASTO_OBJETIVO
    if len(rows) < 2:
        LOG.warning("La pestaña '%s' está vacía.", SHEET_TAB)
        return pd.DataFrame(columns=list(COLUMN_MAP.values()))

    width = max(COLUMN_MAP) + 1
    data = [(list(r) + [""] * width)[:width] for r in rows[1:]]  # rellena filas cortas
    df = pd.DataFrame(data)[list(COLUMN_MAP)].rename(columns=COLUMN_MAP)
    df.index = range(2, len(df) + 2)  # índice = nº de fila en el Sheet
    df.index.name = "Fila Sheet"

    objetivo = {_normalize(c) for c in categorias_gasto}
    mask = df["Categoría Gasto"].map(_normalize).isin(objetivo)

    out = df[mask].copy()
    out["Cantidad Inventario"] = out["Cantidad Inventario"].map(_to_number)
    out["Precio Unitario"] = out["Precio Unitario"].map(_to_number)
    out["Código SKU"] = out["Código SKU"].str.strip()
    out["Unidad Medida"] = out["Unidad Medida"].str.strip()
    out["Mes factura"] = out["Mes factura"].map(_to_number)
    out = out.dropna(subset=["Mes factura"])
    out["Mes factura"] = out["Mes factura"].astype(int)
    LOG.info("Filas leídas: %d | filtradas: %d", len(df), len(out))
    return out


# --------------------------------------------------------------------------- #
# CLI
# --------------------------------------------------------------------------- #
def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("spreadsheet_id", help="ID del Google Sheet (el tramo entre /d/ y /edit en la URL)")
    ap.add_argument("--creds", default=str(DEFAULT_CREDS), help="Ruta a credentials.json")
    ap.add_argument("--out", help="Exportar resultado a .xlsx o .csv")
    args = ap.parse_args()

    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")
    try:
        df = read_facturas(args.spreadsheet_id, args.creds)
    except gspread.exceptions.SpreadsheetNotFound:
        LOG.error("Sheet no encontrado o sin acceso. ¿Lo compartiste con el email de la Service Account?")
        return 1
    except gspread.exceptions.WorksheetNotFound:
        LOG.error("No existe la pestaña '%s'.", SHEET_TAB)
        return 1

    if args.out:
        if args.out.lower().endswith(".csv"):
            df.to_csv(args.out, encoding="utf-8-sig")
        else:
            df.to_excel(args.out)
        LOG.info("Exportado a %s", args.out)
    else:
        with pd.option_context("display.max_rows", 50, "display.width", 200):
            print(df)
    return 0


if __name__ == "__main__":
    sys.exit(main())
