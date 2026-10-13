"""Genera el Dossier Mensual de compras de Mishiguene.

Lee las facturas filtradas de los 5 locales (drive_connector.read_facturas),
y arma un Excel Dossier_Mensual_Mishiguene.xlsx con 4 pestañas:

    Congreso Septiembre -> solo Congreso, Mes factura = 9
    Congreso Octubre    -> solo Congreso, Mes factura = 10
    Cadena Septiembre   -> los 5 locales sumados, Mes factura = 9
    Cadena Octubre      -> los 5 locales sumados, Mes factura = 10

Cada pestaña agrupa por [Código SKU, Productos, Proveedor, Categoría Gasto,
Categoría 2] con:
    Cantidad Inventario -> suma
    Monto Total         -> suma de (Cantidad x Precio Unitario) fila a fila
    Precio Ponderado    -> Monto Total / Cantidad Inventario

Uso:
    python consolidador.py
"""

from __future__ import annotations

import sys
from pathlib import Path

import gspread
import pandas as pd

from drive_connector import read_facturas

# Completá cada valor con el ID del Google Sheet del local
# (el tramo entre /d/ y /edit en la URL).
LOCALES = {
    "Lafinur": "1tv0uI4KsSEPRUecYfnwcDqpGbUDelGqWwe3Hp-fikOc",
    "Arcos": "1frh93qB8Od5mrtAfMuObKwRW93p7kev6DNgmERqz6ho",
    "Cabello": "1zAxYyNB98Sb1-D0HtfokESXc6FrDS96EbFYyMmsseIE",
    "Congreso": "1udLEQpQ29avK4c6fNnQig-w02xIV82IgLe85yVJXaOY",
    "Chimidoha": "1gSHbDSgICeTy6lZpyIPRskFbEs0JB5mxw1e71oOkIzU",
}

GROUP_COLS = ["Código SKU", "Productos", "Proveedor", "Categoría Gasto", "Categoría 2"]

# Orden de columnas en el Excel: categorías junto a la descripción para filtrar fácil
OUTPUT_COLS = [
    "Código SKU", "Productos", "Categoría Gasto", "Categoría 2", "Proveedor",
    "Cantidad Inventario", "Monto Total", "Precio Ponderado",
]

# Pestaña -> (local o None para toda la cadena, mes)
VISTAS = {
    "Congreso Septiembre": ("Congreso", 9),
    "Congreso Octubre": ("Congreso", 10),
    "Cadena Septiembre": (None, 9),
    "Cadena Octubre": (None, 10),
}

OUTPUT = Path(__file__).with_name("Dossier_Mensual_Mishiguene.xlsx")


def cargar_locales() -> pd.DataFrame | None:
    """Lee los 5 locales y devuelve un único DataFrame crudo con la columna 'Local'."""
    frames = []

    for local, sheet_id in LOCALES.items():
        print(f"Procesando {local}...")

        if not sheet_id or sheet_id.startswith("PEGAR_ID"):
            print(f"  [!] {local}: falta completar el ID del Sheet, lo salteo.")
            continue

        try:
            df = read_facturas(sheet_id)
        except gspread.exceptions.SpreadsheetNotFound:
            print(f"  [ERROR] {local}: Sheet no encontrado o sin acceso "
                  "(¿lo compartiste con la Service Account?).")
            continue
        except gspread.exceptions.WorksheetNotFound:
            print(f"  [ERROR] {local}: no existe la pestaña 'Facturas'.")
            continue

        df = df.reset_index()  # conserva 'Fila Sheet' como columna
        df.insert(0, "Local", local)
        frames.append(df)
        print(f"  [OK] {local}: {len(df)} filas.")

    return pd.concat(frames, ignore_index=True) if frames else None


def resumen_mes(df: pd.DataFrame, mes: int) -> pd.DataFrame:
    """Filtra el DataFrame crudo por 'Mes factura' y lo agrupa por GROUP_COLS."""
    df = df[df["Mes factura"] == mes].copy()
    df["Monto Total"] = df["Cantidad Inventario"] * df["Precio Unitario"]

    resumen = df.groupby(GROUP_COLS, as_index=False, dropna=False).agg(
        **{
            "Cantidad Inventario": ("Cantidad Inventario", "sum"),
            "Monto Total": ("Monto Total", "sum"),
        }
    )
    # Precio ponderado; si la cantidad es 0 queda vacío en vez de dividir por cero
    cantidad = resumen["Cantidad Inventario"].where(resumen["Cantidad Inventario"] != 0)
    resumen["Precio Ponderado"] = resumen["Monto Total"] / cantidad

    return (
        resumen[OUTPUT_COLS]
        .round({"Monto Total": 2, "Precio Ponderado": 2})
        .sort_values("Monto Total", ascending=False)
        .reset_index(drop=True)
    )


def main() -> int:
    crudo = cargar_locales()
    if crudo is None:
        print("No se obtuvo data de ningún local. No se generó el Excel.")
        return 1

    print("\nArmando pestañas...")
    with pd.ExcelWriter(OUTPUT) as writer:
        for hoja, (local, mes) in VISTAS.items():
            base = crudo if local is None else crudo[crudo["Local"] == local]
            vista = resumen_mes(base, mes)
            vista.to_excel(writer, sheet_name=hoja, index=False)
            print(f"  [OK] {hoja}: {len(vista)} SKUs | "
                  f"${vista['Monto Total'].sum():,.2f}")

    print(f"\n¡Listo! Exportado a {OUTPUT.name}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
