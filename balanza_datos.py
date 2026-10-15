"""Datos unificados para la pestaña 'Balanza de Negociación & Benchmarking' de la web.

Une, en un solo libro (`Balanza_Datos.xlsx`):
  * CONSUMOS: todas las hojas por local y mes del Dossier de compras (`Congreso Septiembre`, `Arcos Octubre`,
    ...). La vista 'Cadena' se calcula en la app sumando los locales.
  * MERCADO: cada insumo del Dossier cruzado con las listas de proveedores (El Criollo, Posta, El Nuevo Emporio,
    Alyser, Villares, Nogada, La Mesón, Raíces, Open Fruit) con el motor de `comparador_precios.comparar`, en
    formato largo: una fila por (insumo, proveedor, producto de lista) con el precio llevado a la unidad del
    insumo. Así todos los insumos consumidos en cualquier local/mes tienen sus competidores mapeados.

Uso:
    python balanza_datos.py                      # listas de Drive + catálogos locales (Criollo / Nogada)
    python balanza_datos.py --listas-dir carpeta # listas ya descargadas, sin Drive
Salida: Balanza_Datos.xlsx en la raíz (publicar_web.py lo copia a WebTest/).
"""
from __future__ import annotations

import argparse
import logging
import re
import sys
import tempfile
from datetime import datetime
from pathlib import Path

import pandas as pd

import comparador_precios as cp

LOG = logging.getLogger("balanza_datos")
BASE_DIR = Path(__file__).resolve().parent
(BASE_DIR / "Reportes_Generados").mkdir(exist_ok=True)
DOSSIER = BASE_DIR / "Reportes_Generados" / "Dossier_Mensual_Mishiguene.xlsx"
SALIDA = BASE_DIR / "Reportes_Generados" / "Balanza_Datos.xlsx"
MESES = {"enero": 1, "febrero": 2, "marzo": 3, "abril": 4, "mayo": 5, "junio": 6, "julio": 7, "agosto": 8,
         "septiembre": 9, "octubre": 10, "noviembre": 11, "diciembre": 12}
COLS_DOSSIER = ["Código SKU", "Productos", "Categoría Gasto", "Categoría 2", "Proveedor", "Cantidad Inventario",
                "Monto Total", "Precio Ponderado"]
_BASE_RE = re.compile(r"\s*\[([^\]]*?)(?:,\s*score\s*([\d.]+))?\]\s*$")


def leer_consumos(dossier: Path) -> pd.DataFrame:
    """Hojas '<Local> <Mes>' del Dossier -> tabla larga con Local y Mes. Se omite 'Cadena' (es la suma)."""
    partes = []
    for hoja, df in pd.read_excel(dossier, sheet_name=None).items():
        m = re.match(r"^(?P<local>.+?)\s+(?P<mes>[A-Za-záéíóúñ]+)$", hoja.strip())
        if not m or m["mes"].lower() not in MESES or not set(COLS_DOSSIER) <= set(df.columns):
            LOG.warning("Hoja '%s' omitida (nombre o columnas no reconocidos)", hoja)
            continue
        if m["local"].lower() == "cadena":
            continue
        d = df[COLS_DOSSIER].copy()
        d.insert(0, "Local", m["local"].strip())
        d.insert(1, "Mes", MESES[m["mes"].lower()])
        partes.append(d)
    c = pd.concat(partes, ignore_index=True)
    c["Código SKU"] = c["Código SKU"].astype(str).str.strip()
    c = c[~c["Código SKU"].str.lower().isin({"", "nan", "none", "sin sku", "no encontrado"})]
    for col in ("Cantidad Inventario", "Monto Total", "Precio Ponderado"):
        c[col] = pd.to_numeric(c[col], errors="coerce")
    return c.reset_index(drop=True)


def dossier_unificado(consumos: pd.DataFrame) -> pd.DataFrame:
    """Un renglón por (SKU, producto, proveedor) sumando locales y meses, con Precio Ponderado = monto/cantidad."""
    claves = ["Código SKU", "Productos", "Categoría Gasto", "Categoría 2", "Proveedor"]
    g = (consumos.groupby(claves, as_index=False, dropna=False)
         .agg(**{"Cantidad Inventario": ("Cantidad Inventario", "sum"), "Monto Total": ("Monto Total", "sum")}))
    g["Precio Ponderado"] = (g["Monto Total"] / g["Cantidad Inventario"]).where(g["Cantidad Inventario"] > 0)
    return g.dropna(subset=["Precio Ponderado"]).reset_index(drop=True)


def cargar_listas(listas_dir: Path | None) -> pd.DataFrame:
    extra = [p for p in (cp.CRIOLLO_CSV, cp.NOGADA_CSV) if Path(p).is_file()]
    if listas_dir:
        paths = sorted(p for p in listas_dir.iterdir() if p.is_file())
        if not any("criollo" in p.name.lower() for p in paths):
            paths += [p for p in extra if "criollo" in Path(p).name.lower()]
        if not any("nogada" in p.name.lower() for p in paths):
            paths += [p for p in extra if "nogada" in Path(p).name.lower()]
        return cp.cargar_listas(paths)
    with tempfile.TemporaryDirectory(prefix="listas_proveedores_") as tmp:
        paths = cp.descargar_listas(Path(tmp), cp.DRIVE_FOLDER_ID, cp.DEFAULT_CREDS)
        if not any("criollo" in p.name.lower() for p in paths):
            paths += [Path(p) for p in extra if "criollo" in Path(p).name.lower()]
        if not any("nogada" in p.name.lower() for p in paths):
            paths += [Path(p) for p in extra if "nogada" in Path(p).name.lower()]
        return cp.cargar_listas(paths)


def ofertas_largas(res: pd.DataFrame, proveedores: list[str]) -> pd.DataFrame:
    """Matriz del comparador -> una fila por oferta comparable (precio numérico > 0)."""
    filas = []
    for _, r in res.iterrows():
        for prov in proveedores:
            precio = r.get(prov)
            if not isinstance(precio, (int, float)) or pd.isna(precio) or precio <= 0:
                continue
            match = str(r.get(f"Match {prov}", ""))
            m = _BASE_RE.search(match)
            filas.append({
                "Código SKU": str(r["Código SKU"]).strip(), "Producto (Dossier)": r["Producto (Dossier)"],
                "Opción": r.get("Opción"), "Proveedor mercado": prov,
                "Producto mercado": _BASE_RE.sub("", match).strip(),
                "Base": (m.group(1).removeprefix("$") if m else ""),
                "Score": float(m.group(2)) if m and m.group(2) else None,
                "Precio unitario": float(precio),
            })
    o = pd.DataFrame(filas, columns=["Código SKU", "Producto (Dossier)", "Opción", "Proveedor mercado",
                                     "Producto mercado", "Base", "Score", "Precio unitario"])
    # una oferta por (SKU, proveedor, producto de lista): la de menor precio
    o = o.sort_values("Precio unitario").drop_duplicates(["Código SKU", "Proveedor mercado", "Producto mercado"])
    return o.sort_values(["Código SKU", "Precio unitario"]).reset_index(drop=True)


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--dossier", default=str(DOSSIER))
    ap.add_argument("--listas-dir", help="listas de proveedores ya descargadas (sin Drive)")
    ap.add_argument("--salida", default=str(SALIDA))
    args = ap.parse_args()
    logging.basicConfig(level=logging.INFO, format="%(asctime)s | %(levelname)s | %(message)s", datefmt="%H:%M:%S")

    consumos = leer_consumos(Path(args.dossier))
    LOG.info("Consumos: %d filas · locales %s · meses %s", len(consumos), sorted(consumos["Local"].unique()),
             sorted(consumos["Mes"].unique()))
    dos = cp.aplicar_exclusiones_dossier(cp.aplicar_lista_negra(dossier_unificado(consumos)))
    LOG.info("Insumos a mapear contra el mercado: %d filas únicas (SKU+producto+proveedor)", len(dos))

    listas = cargar_listas(Path(args.listas_dir) if args.listas_dir else None)
    proveedores = list(dict.fromkeys(listas["proveedor"]))
    LOG.info("Listas de mercado: %s", ", ".join(f"{p} ({int((listas.proveedor == p).sum())})" for p in proveedores))
    lineas = cp.leer_productos_txt(Path(cp.PRODUCTOS_TXT)) if Path(cp.PRODUCTOS_TXT).is_file() else []
    res = cp.comparar(dos, listas, lineas)
    ofertas = ofertas_largas(res, proveedores)
    cobertura = ofertas["Código SKU"].nunique()
    LOG.info("Mercado: %d ofertas · %d de %d SKU con al menos un proveedor mapeado", len(ofertas), cobertura,
             dos["Código SKU"].nunique())

    meta = pd.DataFrame([
        ("Generado", datetime.now().strftime("%Y-%m-%d %H:%M")),
        ("Fuente de consumos", Path(args.dossier).name),
        ("Proveedores de mercado", ", ".join(proveedores)),
        ("SKU con competidores", f"{cobertura} de {dos['Código SKU'].nunique()}"),
    ], columns=["Dato", "Valor"])
    with pd.ExcelWriter(args.salida) as w:
        consumos.to_excel(w, sheet_name="Consumos", index=False)
        ofertas.to_excel(w, sheet_name="Mercado", index=False)
        meta.to_excel(w, sheet_name="Meta", index=False)
    LOG.info("Escrito %s", args.salida)
    return 0


if __name__ == "__main__":
    sys.exit(main())
