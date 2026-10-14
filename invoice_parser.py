"""Parser de facturas PDF de proveedores satélite (Alyser, La Meson, ...).

Extrae de cada factura las columnas:
    [Proveedor, Producto, Cantidad, Precio Unitario, Fecha]
y devuelve un DataFrame de Pandas listo para cruzar con inventarios/compras.

Estrategia (en cascada):
  1. pdfplumber.extract_tables()  -> si el PDF tiene grilla, se mapean columnas por encabezado.
  2. Regex línea a línea sobre el texto -> fallback para PDFs sin bordes de tabla.
  Cada ítem se valida con Cantidad x Precio ≈ Total (cuando la línea trae total).

Motor multimodal (procesar_comprobante): además de PDFs digitales lee fotos de WhatsApp
(.jpg/.jpeg/.png/.webp) y PDFs escaneados:
  imagen -> pre-proceso (fondo/iluminación, contraste CLAHE, reducción de ruido, enderezado y
  rotación automática) -> OCR (Tesseract / RapidOCR / EasyOCR, el primero disponible) -> texto ->
  JSON {local_destino, numero_factura, fecha_factura, fecha_ingreso, proveedor_detectado,
  items[{descripcion_original, cantidad, unidad_detectada, precio_unitario, descuento_pct,
  total_linea}], total_comprobante, validacion}. Caché por SHA-256 en cache_facturas.json y
  validación aritmética (suma de renglones vs total del comprobante, ±1 %).

Uso:
    python invoice_parser.py factura.pdf
    python invoice_parser.py carpeta_facturas/ --out compras.xlsx
    python invoice_parser.py foto_remito.jpg --json salida.json      # motor multimodal
"""

from __future__ import annotations

import argparse
import hashlib
import json
import logging
import os
import re
import shutil
import sys
import unicodedata
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Callable

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


# --------------------------------------------------------------------------- #
# Motor multimodal: fotos de WhatsApp, escaneos y PDFs -> JSON de comprobante
# --------------------------------------------------------------------------- #
BASE_DIR = Path(__file__).resolve().parent
CACHE_FACTURAS = BASE_DIR / "cache_facturas.json"
IMG_EXT = {".jpg", ".jpeg", ".png", ".webp", ".bmp", ".tif", ".tiff"}
TOLERANCIA_TOTAL = 0.01          # suma de renglones vs total del comprobante (±1 %)
ALICUOTAS_IVA = (0.21, 0.105, 0.27)

# Local de destino: lo que dice el encabezado ('Obrador' == 'Congreso')
LOCALES_RE = [
    ("Arcos", re.compile(r"\barcos\b", re.I)),
    ("Congreso", re.compile(r"\b(?:congreso|obrador)\b", re.I)),
    ("Cabello", re.compile(r"\bcabello\b", re.I)),
    ("Lafinur", re.compile(r"\blafinur\b", re.I)),
    ("Chimidoha", re.compile(r"\bchimi(?:doha)?\b", re.I)),
]

_NUM_FACT = re.compile(
    r"(?:factura|fact\.?|fc|remito|rto\.?|comprobante|comp\.?|n[°ºo]\.?|nro\.?|n\s*[°º])\s*"
    r"(?:[abcxm]\s*)?[:#.]?\s*(\d{4,5}\s*[-–]\s*\d{6,8}|\d{5,13})", re.I)
_FECHA_FACT = re.compile(r"fecha(?:\s+(?:de\s+)?(?:emisi[oó]n|factura|comprobante))?\s*[:.]?\s*"
                         r"(\d{1,2}\s*[/.-]\s*\d{1,2}\s*[/.-]\s*\d{2,4})", re.I)
_FECHA_INGRESO = re.compile(r"(?:ingreso|recepci[oó]n|entrega|recibido)\s*[:.]?\s*"
                            r"(\d{1,2}\s*[/.-]\s*\d{1,2}\s*[/.-]\s*\d{2,4})", re.I)
_TOTAL = re.compile(r"(?<!sub)(?<!sub\s)\btotal(?:\s+(?:a\s+pagar|general|comprobante|final|factura))?"
                    r"\s*[:$]*\s*\$?\s*(" + NUM + r")\s*$", re.I | re.M)
# CANT [UNIDAD] DESCRIPCION PRECIO [BONIF%] TOTAL  (remitos donde la cantidad va primero)
QTY_FIRST_LINE = re.compile(
    rf"""^\s*(?P<cantidad>{NUM})\s+(?:(?P<unidad>{UNIDADES})\.?\s+)?
    (?P<producto>[A-Za-zÁÉÍÓÚÑáéíóúñ].*?)\s+\$?\s*(?P<precio>{NUM})\s+
    (?:(?P<bonif>{NUM})\s*%\s+)?\$?\s*(?P<total>{NUM})\s*$""",
    re.VERBOSE | re.IGNORECASE,
)
_UNIDAD_CANON = {"kg": "KG", "kgs": "KG", "kilo": "KG", "gr": "GR", "lt": "LT", "lts": "LT", "l": "LT",
                 "un": "UN", "und": "UN", "uni": "UN", "unid": "UN", "u": "UN", "cj": "CAJA", "caja": "CAJA",
                 "bul": "BULTO", "bulto": "BULTO", "paq": "PAQ", "pack": "PACK", "doc": "DOC", "bot": "BOT",
                 "lata": "LATA", "bal": "BALDE", "balde": "BALDE"}


class OCRNoDisponible(RuntimeError):
    """No hay ningún motor OCR instalado (ver mensaje para instalar uno)."""


def sha256_archivo(path: Path) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for bloque in iter(lambda: f.read(1 << 20), b""):
            h.update(bloque)
    return h.hexdigest()


# ---- pre-procesamiento de imagen ------------------------------------------ #
def _cargar_bgr(path: Path):
    import cv2
    import numpy as np
    img = cv2.imdecode(np.fromfile(str(path), dtype=np.uint8), cv2.IMREAD_COLOR)   # rutas con tildes/espacios
    if img is None:
        from PIL import Image
        img = cv2.cvtColor(np.array(Image.open(path).convert("RGB")), cv2.COLOR_RGB2BGR)
    return img


def estimar_inclinacion(gris) -> float:
    """Ángulo (grados, + = antihorario) con que hay que girar la imagen para enderezar el texto.
    Se busca el que maximiza la nitidez del perfil horizontal de tinta, en ±15°."""
    import cv2
    import numpy as np
    h, w = gris.shape[:2]
    escala = 800 / max(h, w) if max(h, w) > 800 else 1.0
    chico = cv2.resize(gris, None, fx=escala, fy=escala, interpolation=cv2.INTER_AREA) if escala < 1 else gris
    _, bn = cv2.threshold(chico, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)
    if bn.mean() < 1:
        return 0.0
    ch, cw = bn.shape
    centro = (cw / 2, ch / 2)

    def nitidez(ang: float) -> float:
        m = cv2.getRotationMatrix2D(centro, ang, 1.0)
        rot = cv2.warpAffine(bn, m, (cw, ch), flags=cv2.INTER_NEAREST)
        perfil = rot.sum(axis=1).astype(np.float64)
        return float(np.sum(np.diff(perfil) ** 2))

    mejor = max(np.arange(-15, 15.01, 1.0), key=nitidez)
    fino = max(np.arange(mejor - 1, mejor + 1.01, 0.25), key=nitidez)
    return float(fino)


def _orientacion_90(gris) -> int:
    """0 o 90 grados: un texto horizontal da un perfil de filas mucho más 'rayado' que uno vertical."""
    import cv2
    import numpy as np
    _, bn = cv2.threshold(gris, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)
    filas, cols = bn.sum(axis=1).astype(float), bn.sum(axis=0).astype(float)
    return 0 if filas.var() * 1.0 >= cols.var() else 90


def preprocesar_imagen(img_bgr):
    """Foto de celular -> imagen de grises lista para OCR.
    Pasos: ancho de trabajo 1800-2600 px, aplanado de iluminación (divide por el fondo), reducción de
    ruido, contraste adaptativo (CLAHE), giro de 90° si el texto está de costado y enderezado fino."""
    import cv2
    import numpy as np
    img = img_bgr
    h, w = img.shape[:2]
    objetivo = 2000
    if w < 1800 or w > 2600:
        f = objetivo / w
        img = cv2.resize(img, None, fx=f, fy=f, interpolation=cv2.INTER_CUBIC if f > 1 else cv2.INTER_AREA)
    gris = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    fondo = cv2.medianBlur(cv2.dilate(gris, np.ones((7, 7), np.uint8)), 51)       # sombras / flash desparejo
    gris = cv2.divide(gris, fondo, scale=255)
    gris = cv2.fastNlMeansDenoising(gris, None, h=9, templateWindowSize=7, searchWindowSize=21)
    gris = cv2.createCLAHE(clipLimit=2.5, tileGridSize=(8, 8)).apply(gris)
    if _orientacion_90(gris) == 90:
        gris = cv2.rotate(gris, cv2.ROTATE_90_CLOCKWISE)
    ang = estimar_inclinacion(gris)
    if abs(ang) >= 0.3:
        h2, w2 = gris.shape[:2]
        m = cv2.getRotationMatrix2D((w2 / 2, h2 / 2), ang, 1.0)
        gris = cv2.warpAffine(gris, m, (w2, h2), flags=cv2.INTER_CUBIC, borderMode=cv2.BORDER_REPLICATE)
    return gris


# ---- motores OCR ------------------------------------------------------------ #
def _motor_ocr() -> tuple[str, Callable]:
    """Primer motor OCR disponible: Tesseract (pytesseract + binario), RapidOCR, EasyOCR."""
    try:
        import pytesseract
        exe = os.environ.get("TESSERACT_CMD") or shutil.which("tesseract") or next(
            (p for p in (r"C:\Program Files\Tesseract-OCR\tesseract.exe",
                         r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe") if Path(p).is_file()), None)
        if exe:
            pytesseract.pytesseract.tesseract_cmd = exe
            langs = pytesseract.get_languages(config="")
            lang = "spa+eng" if "spa" in langs else "eng"
            return "tesseract", lambda im: pytesseract.image_to_string(im, lang=lang, config="--oem 3 --psm 6")
    except Exception as e:   # pragma: no cover - depende del equipo
        LOG.debug("Tesseract no disponible: %s", e)
    for modulo, clase in (("rapidocr_onnxruntime", "RapidOCR"), ("rapidocr", "RapidOCR")):
        try:
            motor = getattr(__import__(modulo, fromlist=[clase]), clase)()

            def _rapid(im, _m=motor):
                res, _ = _m(im)
                return "\n".join(_renglones_ocr([(b[0][1], b[0][0], t) for b, t, _c in (res or [])]))
            return "rapidocr", _rapid
        except Exception:
            continue
    try:
        import easyocr
        lector = easyocr.Reader(["es", "en"], gpu=False, verbose=False)
        return "easyocr", lambda im: "\n".join(
            _renglones_ocr([(b[0][1], b[0][0], t) for b, t, _c in lector.readtext(im)]))
    except Exception:
        pass
    raise OCRNoDisponible(
        "No hay motor OCR. Instalar UNO de: (a) Tesseract para Windows (choco install tesseract, o el instalador de "
        "UB-Mannheim; con idioma 'spa') + `pip install pytesseract`; (b) `pip install rapidocr-onnxruntime`; "
        "(c) `pip install easyocr`. Opcional: variable TESSERACT_CMD con la ruta del ejecutable.")


def _renglones_ocr(cajas: list[tuple[float, float, str]]) -> list[str]:
    """Cajas (y, x, texto) de OCR por palabra/fragmento -> renglones de texto en orden de lectura."""
    if not cajas:
        return []
    cajas = sorted(cajas)
    alto = 18.0      # tolerancia vertical (px) para considerar dos fragmentos del mismo renglón
    renglones, actual, y_ref = [], [], cajas[0][0]
    for y, x, t in cajas:
        if abs(y - y_ref) > alto:
            renglones.append(" ".join(t2 for _, t2 in sorted(actual)))
            actual, y_ref = [], y
        actual.append((x, t))
    if actual:
        renglones.append(" ".join(t2 for _, t2 in sorted(actual)))
    return renglones


def _calidad_texto(texto: str) -> float:
    """Heurística 0-1: proporción de renglones con precio/cantidad + palabras de comprobante."""
    lineas = [l for l in texto.splitlines() if l.strip()]
    if not lineas:
        return 0.0
    con_num = sum(bool(re.search(r"\d+[.,]\d{2}\b", l)) for l in lineas)
    claves = 1 if re.search(r"total|factura|remito|fecha|cuit|cant", texto, re.I) else 0
    return min(1.0, con_num / max(len(lineas), 1) + 0.2 * claves)


def ocr_imagen(img_bgr, ocr_fn: Callable | None = None) -> tuple[str, str, list[str]]:
    """OCR de una imagen ya cargada. Devuelve (texto, nombre_motor, advertencias).
    Si el texto sale muy pobre se reintenta girada 180° (foto cabeza abajo) y se queda con el mejor."""
    import cv2
    avisos: list[str] = []
    nombre = "inyectado"
    if ocr_fn is None:
        nombre, ocr_fn = _motor_ocr()
    gris = preprocesar_imagen(img_bgr)
    texto = ocr_fn(gris)
    if _calidad_texto(texto) < 0.25:
        alt = ocr_fn(cv2.rotate(gris, cv2.ROTATE_180))
        if _calidad_texto(alt) > _calidad_texto(texto):
            texto = alt
            avisos.append("imagen girada 180° automáticamente")
    return texto, nombre, avisos


# ---- texto -> comprobante ----------------------------------------------------- #
def _unidad_canon(u: str | None) -> str | None:
    return _UNIDAD_CANON.get(re.sub(r"[.\s]", "", (u or "").lower())) if u else None


def _fecha_iso(raw: str | None) -> str | None:
    if not raw:
        return None
    d = parse_date(re.sub(r"\s+", "", raw))
    return d.date().isoformat() if d is not None else None


def _items_comprobante(texto: str, perfil: SupplierProfile | None) -> list[dict]:
    patrones = [QTY_FIRST_LINE, (perfil.line_regex if perfil and perfil.line_regex else GENERIC_LINE)]
    extra_skip = perfil.extra_skip if perfil else ()
    items: list[dict] = []
    for linea in texto.splitlines():
        if SKIP_LINE.search(linea) or any(s.lower() in linea.lower() for s in extra_skip):
            continue
        linea = re.sub(r"[|¦_]+", " ", linea)
        mejor = None
        for pat in patrones:             # puede casar más de un patrón: gana el que cierra cant x precio = total
            m = pat.match(linea)
            if not m:
                continue
            g = m.groupdict()
            cant, precio = parse_number(g.get("cantidad")), parse_number(g.get("precio"))
            if cant is None or precio is None or cant <= 0:
                continue
            total, bonif = parse_number(g.get("total")), parse_number(g.get("bonif")) or 0.0
            esperado = round(cant * precio * (1 - bonif / 100), 2)
            item = {
                "descripcion_original": re.sub(r"\s{2,}", " ", g["producto"]).strip(" -.:"),
                "cantidad": cant,
                "unidad_detectada": _unidad_canon(g.get("unidad")),
                "precio_unitario": precio,
                "descuento_pct": bonif,
                "total_linea": total if total is not None else esperado,
                "linea_consistente": _is_consistent(cant, precio, total, bonif),
            }
            if mejor is None or (item["linea_consistente"] and not mejor["linea_consistente"]):
                mejor = item
        if mejor is not None:
            items.append(mejor)
    return items


def validar_aritmetica(items: list[dict], total_comprobante: float | None) -> dict:
    """Suma de renglones vs total del comprobante (±1 %); acepta que el total incluya IVA no itemizado."""
    suma = round(sum(i["total_linea"] for i in items), 2)
    out = {"suma_renglones": suma, "total_comprobante": total_comprobante, "tolerancia_pct": TOLERANCIA_TOTAL * 100,
           "renglones_inconsistentes": sum(not i.get("linea_consistente", True) for i in items)}
    if not total_comprobante or not items:
        return {**out, "ok": False, "motivo": "sin total o sin renglones"}
    dif = abs(suma - total_comprobante) / total_comprobante
    out["diferencia_pct"] = round(dif * 100, 2)
    if dif <= TOLERANCIA_TOTAL:
        return {**out, "ok": True, "motivo": "suma de renglones = total"}
    for a in ALICUOTAS_IVA:
        if abs(suma * (1 + a) - total_comprobante) / total_comprobante <= TOLERANCIA_TOTAL:
            return {**out, "ok": True, "motivo": f"total incluye IVA {a * 100:g} % sobre los renglones"}
    return {**out, "ok": False, "motivo": f"la suma difiere {dif * 100:.1f} % del total (> {TOLERANCIA_TOTAL * 100:g} %)"}


def comprobante_desde_texto(texto: str, nombre_archivo: str = "", local: str | None = None) -> dict:
    """Texto plano (PDF digital u OCR) -> dict de comprobante con el esquema documentado arriba."""
    texto = texto.replace("\r", "")
    perfil = detect_supplier(texto, Path(nombre_archivo).stem)
    cabecera = texto[:1500]
    m_num = _NUM_FACT.search(cabecera)
    f_fact = _FECHA_FACT.search(texto)
    f_ing = _FECHA_INGRESO.search(texto)
    fecha_fact = _fecha_iso(f_fact.group(1)) if f_fact else None
    if fecha_fact is None:
        d = extract_date(texto, perfil)
        fecha_fact = d.date().isoformat() if d is not None else None
    totales = [parse_number(m.group(1)) for m in _TOTAL.finditer(texto)]
    totales = [t for t in totales if t]
    local_det = local
    if local_det is None:
        pos = [(m.start(), nombre) for nombre, rx in LOCALES_RE for m in [rx.search(cabecera)] if m]
        local_det = min(pos)[1] if pos else None
    if perfil:
        proveedor = perfil.name
    else:
        proveedor = next((l.strip() for l in cabecera.splitlines()
                          if len(re.findall(r"[A-Za-zÁÉÍÓÚÑ]", l)) >= 5 and not re.search(
                              r"factura|remito|fecha|cuit|original|duplicado|mishiguene|se[ñn]or|domicilio|tel", l, re.I)),
                         None)
    items = _items_comprobante(texto, perfil)
    total = totales[-1] if totales else None
    avisos = []
    if not items:
        avisos.append("no se detectaron renglones de detalle")
    return {
        "local_destino": local_det,
        "numero_factura": re.sub(r"\s+", "", m_num.group(1)) if m_num else None,
        "fecha_factura": fecha_fact,
        "fecha_ingreso": _fecha_iso(f_ing.group(1)) if f_ing else None,
        "proveedor_detectado": proveedor,
        "texto_encabezado": cabecera[:600],
        "items": items,
        "total_comprobante": total,
        "validacion": validar_aritmetica(items, total),
        "advertencias": avisos,
    }


def _leer_cache() -> dict:
    try:
        return json.loads(CACHE_FACTURAS.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}


def procesar_comprobante(path: str | Path, local: str | None = None, usar_cache: bool = True,
                         ocr_fn: Callable | None = None) -> dict:
    """PDF digital, PDF escaneado o foto (.jpg/.jpeg/.png/.webp) -> dict de comprobante.
    `ocr_fn(gris_ndarray) -> str` permite inyectar otro motor OCR (tests)."""
    path = Path(path)
    h = sha256_archivo(path)
    cache = _leer_cache() if usar_cache else {}
    if h in cache:
        LOG.info("%s: en caché (sha256 %s…)", path.name, h[:10])
        res = dict(cache[h])
        if local and not res.get("local_destino"):
            res["local_destino"] = local
        return {**res, "desde_cache": True}

    ext = path.suffix.lower()
    avisos: list[str] = []
    motor = None
    if ext == ".pdf":
        with pdfplumber.open(path) as pdf:
            texto = "\n".join(p.extract_text() or "" for p in pdf.pages)
            motor = "pdf-texto"
            if len(texto.strip()) < 40:                       # PDF escaneado: no tiene capa de texto
                import numpy as np
                import cv2
                partes = []
                for pg in pdf.pages:
                    im = pg.to_image(resolution=220).original.convert("RGB")
                    t, nombre, av = ocr_imagen(cv2.cvtColor(np.array(im), cv2.COLOR_RGB2BGR), ocr_fn)
                    partes.append(t)
                    avisos += av
                texto, motor = "\n".join(partes), f"pdf-escaneado+{nombre}"
    elif ext in IMG_EXT:
        texto, nombre, avisos = ocr_imagen(_cargar_bgr(path), ocr_fn)
        motor = f"imagen+{nombre}"
    else:
        raise ValueError(f"formato no soportado: {ext}")

    res = comprobante_desde_texto(texto, path.name, local)
    res.update({"archivo": path.name, "sha256": h, "motor": motor, "advertencias": res["advertencias"] + avisos})
    if usar_cache:
        cache[h] = {k: v for k, v in res.items() if k != "desde_cache"}
        CACHE_FACTURAS.write_text(json.dumps(cache, ensure_ascii=False, indent=1), encoding="utf-8")
    return res


def _resumen(res: dict) -> str:
    v = res.get("validacion", {})
    return (f"{res.get('archivo', '?')} | {res.get('motor')} | local={res.get('local_destino')} | "
            f"N°={res.get('numero_factura')} | {res.get('fecha_factura')} | {res.get('proveedor_detectado')} | "
            f"{len(res.get('items', []))} ítems | total={res.get('total_comprobante')} | "
            f"validación={'OK' if v.get('ok') else 'REVISAR'} ({v.get('motivo')})")


def main(argv: list[str] | None = None) -> int:
    logging.basicConfig(level=logging.INFO, format="%(asctime)s | %(levelname)s | %(message)s",
                        datefmt="%H:%M:%S")
    ap = argparse.ArgumentParser(description="Extrae comprobantes de PDFs, escaneos y fotos de WhatsApp.")
    ap.add_argument("path", help="Archivo (.pdf/.jpg/.jpeg/.png/.webp) o carpeta")
    ap.add_argument("--supplier", help="(modo clásico) forzar proveedor, solo para un PDF")
    ap.add_argument("--local", help="Forzar local de destino (Arcos/Congreso/Cabello/Lafinur/Chimidoha)")
    ap.add_argument("--json", help="Guardar los comprobantes en este .json")
    ap.add_argument("--out", help="(modo clásico) exportar el DataFrame de ítems a .csv/.xlsx")
    ap.add_argument("--sin-cache", action="store_true", help="No usar ni escribir cache_facturas.json")
    args = ap.parse_args(argv)

    path = Path(args.path)
    if args.out or args.supplier:          # modo clásico (DataFrame), solo PDFs digitales
        df = parse_folder(path) if path.is_dir() else parse_invoice(path, args.supplier)
        with pd.option_context("display.max_rows", 200, "display.width", 200):
            print(df)
        if args.out:
            out = Path(args.out)
            df.to_excel(out, index=False) if out.suffix.lower() == ".xlsx" else df.to_csv(
                out, index=False, encoding="utf-8-sig")
            LOG.info("Exportado a %s", out)
        return 0

    archivos = sorted(p for p in path.iterdir() if p.suffix.lower() in IMG_EXT | {".pdf"}) if path.is_dir() else [path]
    resultados, rc = [], 0
    for a in archivos:
        try:
            res = procesar_comprobante(a, args.local, usar_cache=not args.sin_cache)
        except OCRNoDisponible as e:
            LOG.error("%s: %s", a.name, e)
            rc = 2
            continue
        except Exception as e:                      # un archivo roto no frena el lote
            LOG.error("%s: error al procesar -> %s", a.name, e)
            rc = 1
            continue
        resultados.append(res)
        print(_resumen(res))
    if args.json:
        Path(args.json).write_text(json.dumps(resultados, ensure_ascii=False, indent=1), encoding="utf-8")
        LOG.info("JSON guardado en %s", args.json)
    return rc


if __name__ == "__main__":
    sys.exit(main())
