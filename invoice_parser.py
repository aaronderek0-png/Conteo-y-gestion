"""Parser de facturas PDF de proveedores satélite (Alyser, La Meson, ...).

Extrae de cada factura las columnas:
    [Proveedor, Producto, Cantidad, Precio Unitario, Fecha]
y devuelve un DataFrame de Pandas listo para cruzar con inventarios/compras.

Estrategia (en cascada):
  1. pdfplumber.extract_tables()  -> si el PDF tiene grilla, se mapean columnas por encabezado.
  2. Regex línea a línea sobre el texto -> fallback para PDFs sin bordes de tabla.
  Cada ítem se valida con Cantidad x Precio ≈ Total (cuando la línea trae total).

Uso:
    python invoice_parser.py factura.pdf
    python invoice_parser.py carpeta_facturas/ --out compras.xlsx
"""

from __future__ import annotations

import argparse
import logging
import re
import sys
import unicodedata
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

import pandas as pd
import pdfplumber

LOG = logging.getLogger("invoice_parser")

COLUMNS = ["Proveedor", "Producto", "Cantidad", "Precio Unitario", "Fecha"]

# --------------------------------------------------------------------------- #
# Regex base
# --------------------------------------------------------------------------- #
# Número en formato AR (1.234,56), US (1,234.56) o simple (12 / 12,5 / 12.5)
NUM = r"\d{1,3}(?:[.,]\d{3})*(?:[.,]\d{1,4})?|\d+(?:[.,]\d{1,4})?"

UNIDADES = r"KG|KGS|KILO|GR|UN|UND|UNI|UNID|U|LT|LTS|L|CJ|CAJA|BUL|BULTO|PAQ|PACK|DOC|BOT|LATA|BAL|BALDE"

DATE_LABELED = re.compile(
    r"fecha(?:\s+de)?(?:\s+emisi[oó]n)?\s*[:.]?\s*(\d{1,2}[/.-]\d{1,2}[/.-]\d{2,4})",
    re.IGNORECASE,
)
DATE_ANY = re.compile(r"\b(\d{1,2}[/.-]\d{1,2}[/.-](?:\d{4}|\d{2}))\b")

# Línea de detalle genérica:
#   [codigo] DESCRIPCION  CANTIDAD [UNIDAD]  [$]PRECIO_UNIT  [BONIF%]  [$]TOTAL
GENERIC_LINE = re.compile(
    rf"""^\s*
    (?:(?P<codigo>[A-Z0-9-]{{2,}})\s+)?                 # código de artículo (opcional)
    (?P<producto>[A-Za-zÁÉÍÓÚÑáéíóúñ].*?)\s+            # descripción (empieza con letra)
    (?P<cantidad>{NUM})\s*
    (?:(?P<unidad>{UNIDADES})\.?\s+)?                   # unidad de medida (opcional)
    \$?\s*(?P<precio>{NUM})\s+
    (?:(?P<bonif>{NUM})\s*%\s+)?                        # bonificación (opcional)
    \$?\s*(?P<total>{NUM})\s*$""",
    re.VERBOSE | re.IGNORECASE,
)

# Líneas que nunca son ítems (totales, impuestos, pie de factura)
SKIP_LINE = re.compile(
    r"\b(sub\s*total|total|iva|percep|ingresos\s+brutos|iibb|neto\s+gravado|"
    r"cae|vto|vencimiento|saldo|importe|son\s+pesos|transporte|p[aá]gina)\b",
    re.IGNORECASE,
)

# Alias de encabezados para mapear tablas de pdfplumber
HEADER_ALIASES = {
    "producto": ("descripcion", "detalle", "producto", "articulo", "concepto"),
    "cantidad": ("cantidad", "cant", "cant.", "unidades", "kilos"),
    "precio": ("precio unitario", "p. unit", "p.unit", "precio unit", "unitario", "precio", "p.u."),
    "total": ("importe", "total", "subtotal", "monto"),
}


# --------------------------------------------------------------------------- #
# Perfiles por proveedor
# --------------------------------------------------------------------------- #
@dataclass
class SupplierProfile:
    """Configuración específica de un proveedor.

    `keywords` / `cuit` sirven para detectarlo en el texto del PDF.
    `line_regex` permite sobreescribir el patrón genérico si su layout es distinto
    (debe exponer los grupos `producto`, `cantidad`, `precio` y opcionalmente `total`/`bonif`).
    """

    name: str
    keywords: tuple[str, ...]
    cuit: tuple[str, ...] = ()
    line_regex: re.Pattern | None = None
    date_regex: re.Pattern | None = None
    extra_skip: tuple[str, ...] = field(default_factory=tuple)


# TODO: calibrar `line_regex` y `cuit` con una factura real de cada proveedor.
SUPPLIERS: list[SupplierProfile] = [
    SupplierProfile(name="Alyser", keywords=("alyser",)),
    SupplierProfile(name="La Meson", keywords=("la meson", "la mesón", "meson")),
]


# --------------------------------------------------------------------------- #
# Helpers
# --------------------------------------------------------------------------- #
def _normalize(text: str) -> str:
    text = unicodedata.normalize("NFKD", text or "")
    return "".join(c for c in text if not unicodedata.combining(c)).lower().strip()


def parse_number(raw: str | None) -> float | None:
    """Convierte '1.234,56' / '1,234.56' / '12,5' / '$ 300' a float."""
    if raw is None:
        return None
    s = re.sub(r"[^\d.,-]", "", str(raw))
    if not s or s in {"-", ".", ","}:
        return None
    if "," in s and "." in s:
        # El separador que aparece último es el decimal
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        # Coma sola: decimal (formato AR), salvo patrón de miles '1,234,567'
        s = s.replace(",", "") if re.fullmatch(r"\d{1,3}(,\d{3}){2,}", s) else s.replace(",", ".")
    elif "." in s:
        # Punto solo: miles si es '1.234' / '1.234.567', decimal en otro caso
        if re.fullmatch(r"\d{1,3}(\.\d{3})+", s):
            s = s.replace(".", "")
    try:
        return float(s)
    except ValueError:
        return None


def parse_date(raw: str) -> pd.Timestamp | None:
    for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%d.%m.%Y", "%d/%m/%y", "%d-%m-%y", "%d.%m.%y"):
        try:
            return pd.Timestamp(datetime.strptime(raw, fmt))
        except ValueError:
            continue
    return None


def detect_supplier(text: str, filename: str = "") -> SupplierProfile | None:
    haystack = _normalize(text[:3000] + " " + filename)  # el encabezado alcanza
    digits = re.sub(r"\D", "", text[:3000])
    for prof in SUPPLIERS:
        if any(c and re.sub(r"\D", "", c) in digits for c in prof.cuit):
            return prof
        if any(_normalize(k) in haystack for k in prof.keywords):
            return prof
    return None


def extract_date(text: str, profile: SupplierProfile | None) -> pd.Timestamp | None:
    patterns = [p for p in (profile.date_regex if profile else None, DATE_LABELED, DATE_ANY) if p]
    for pat in patterns:
        for m in pat.finditer(text):
            d = parse_date(m.group(1))
            if d is not None:
                return d
    return None


def _is_consistent(qty: float, price: float, total: float | None, bonif: float | None) -> bool:
    """Valida qty * precio * (1 - bonif) ≈ total (tolerancia 2% por redondeos)."""
    if total is None or total == 0:
        return True
    expected = qty * price * (1 - (bonif or 0) / 100)
    return abs(expected - total) <= max(1.0, 0.02 * total)


# --------------------------------------------------------------------------- #
# Estrategia 1: tablas
# --------------------------------------------------------------------------- #
def _map_header(row: list[str | None]) -> dict[str, int]:
    mapping: dict[str, int] = {}
    cells = [_normalize(c or "") for c in row]
    for key, aliases in HEADER_ALIASES.items():
        for idx, cell in enumerate(cells):
            if idx in mapping.values():
                continue
            if any(cell == a or cell.startswith(a) for a in aliases):
                mapping[key] = idx
                break
    return mapping


def items_from_tables(pdf: pdfplumber.PDF) -> list[dict]:
    items: list[dict] = []
    for page in pdf.pages:
        for table in page.extract_tables():
            if not table:
                continue
            header_idx, mapping = None, {}
            for i, row in enumerate(table[:5]):  # el encabezado suele estar arriba
                m = _map_header(row)
                if {"producto", "cantidad", "precio"} <= m.keys():
                    header_idx, mapping = i, m
                    break
            if header_idx is None:
                continue
            for row in table[header_idx + 1:]:
                producto = (row[mapping["producto"]] or "").replace("\n", " ").strip()
                qty = parse_number(row[mapping["cantidad"]])
                price = parse_number(row[mapping["precio"]])
                total = parse_number(row[mapping["total"]]) if "total" in mapping else None
                if not producto or qty is None or price is None or SKIP_LINE.search(producto):
                    continue
                items.append({"Producto": producto, "Cantidad": qty, "Precio Unitario": price,
                              "_valido": _is_consistent(qty, price, total, None)})
    return items


# --------------------------------------------------------------------------- #
# Estrategia 2: regex línea a línea
# --------------------------------------------------------------------------- #
def items_from_text(text: str, profile: SupplierProfile | None) -> list[dict]:
    pattern = (profile.line_regex if profile and profile.line_regex else GENERIC_LINE)
    extra_skip = profile.extra_skip if profile else ()
    items: list[dict] = []
    for line in text.splitlines():
        if SKIP_LINE.search(line) or any(s.lower() in line.lower() for s in extra_skip):
            continue
        m = pattern.match(line)
        if not m:
            continue
        g = m.groupdict()
        qty, price = parse_number(g.get("cantidad")), parse_number(g.get("precio"))
        if qty is None or price is None or qty <= 0:
            continue
        total, bonif = parse_number(g.get("total")), parse_number(g.get("bonif"))
        items.append({
            "Producto": re.sub(r"\s{2,}", " ", g["producto"]).strip(),
            "Cantidad": qty,
            "Precio Unitario": price,
            "_valido": _is_consistent(qty, price, total, bonif),
        })
    return items


# --------------------------------------------------------------------------- #
# API pública
# --------------------------------------------------------------------------- #
def parse_invoice(pdf_path: str | Path, supplier: str | None = None) -> pd.DataFrame:
    """Parsea una factura PDF y devuelve un DataFrame con COLUMNS (+ Archivo, Validado)."""
    pdf_path = Path(pdf_path)
    with pdfplumber.open(pdf_path) as pdf:
        text = "\n".join(page.extract_text() or "" for page in pdf.pages)
        if not text.strip():
            LOG.warning("%s: sin texto extraíble (¿PDF escaneado? requiere OCR).", pdf_path.name)
            return pd.DataFrame(columns=COLUMNS)

        profile = detect_supplier(text, pdf_path.stem)
        proveedor = supplier or (profile.name if profile else "DESCONOCIDO")
        fecha = extract_date(text, profile)

        items = items_from_tables(pdf)
        if not items:
            items = items_from_text(text, profile)

    if not items:
        LOG.warning("%s: no se detectaron ítems.", pdf_path.name)
        return pd.DataFrame(columns=COLUMNS)

    df = pd.DataFrame(items)
    df.insert(0, "Proveedor", proveedor)
    df["Fecha"] = fecha
    df["Archivo"] = pdf_path.name
    df = df.rename(columns={"_valido": "Validado"})
    n_bad = int((~df["Validado"]).sum())
    LOG.info("%s | %s | %s | %d ítems (%d sin validar)", pdf_path.name, proveedor,
             fecha.date() if fecha is not None else "sin fecha", len(df), n_bad)
    return df[COLUMNS + ["Validado", "Archivo"]]


def parse_folder(folder: str | Path) -> pd.DataFrame:
    """Parsea todos los PDF de una carpeta y concatena los resultados."""
    frames = []
    for pdf_path in sorted(Path(folder).glob("*.pdf")):
        try:
            frames.append(parse_invoice(pdf_path))
        except Exception as exc:  # un PDF roto no debe frenar el lote
            LOG.error("%s: error al parsear -> %s", pdf_path.name, exc)
    frames = [f for f in frames if not f.empty]
    if not frames:
        return pd.DataFrame(columns=COLUMNS + ["Validado", "Archivo"])
    return pd.concat(frames, ignore_index=True)


def main(argv: list[str] | None = None) -> int:
    logging.basicConfig(level=logging.INFO, format="%(asctime)s | %(levelname)s | %(message)s",
                        datefmt="%H:%M:%S")
    ap = argparse.ArgumentParser(description="Extrae ítems de facturas PDF a un DataFrame.")
    ap.add_argument("path", help="Archivo PDF o carpeta con PDFs")
    ap.add_argument("--supplier", help="Forzar nombre de proveedor (solo para un archivo)")
    ap.add_argument("--out", help="Exportar a .csv o .xlsx")
    args = ap.parse_args(argv)

    path = Path(args.path)
    df = parse_folder(path) if path.is_dir() else parse_invoice(path, args.supplier)

    with pd.option_context("display.max_rows", 200, "display.width", 200):
        print(df)

    if args.out:
        out = Path(args.out)
        if out.suffix.lower() == ".xlsx":
            df.to_excel(out, index=False)
        else:
            df.to_csv(out, index=False, encoding="utf-8-sig")
        LOG.info("Exportado a %s", out)
    return 0


if __name__ == "__main__":
    sys.exit(main())
