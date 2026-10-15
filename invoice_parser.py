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
STATS = {"cache_sha": 0, "dhash": 0, "vision_gemini": 0, "ocr": 0}   # contadores del proceso

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
            tdata = BASE_DIR / "tessdata"          # spa.traineddata local (la instalación de Program Files solo trae eng)
            if (tdata / "spa.traineddata").is_file():
                os.environ["TESSDATA_PREFIX"] = str(tdata)
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


def validar_aritmetica(items: list[dict], total_comprobante: float | None, extra: dict | None = None) -> dict:
    """Suma de renglones vs total del comprobante (±1 %). Acepta que el total incluya IVA no itemizado (21/10,5/27 %),
    impuestos internos y percepciones (`extra`: imp_internos_total, perc_iva, perc_iibb, perc_3, otros)."""
    extra = extra or {}
    suma = round(sum(i["total_linea"] for i in items), 2)
    imp = float(extra.get("imp_internos_total") or 0)
    percs = sum(float(extra.get(k) or 0) for k in ("perc_iva", "perc_iibb", "perc_3", "otros"))
    out = {"suma_renglones": suma, "total_comprobante": total_comprobante, "tolerancia_pct": TOLERANCIA_TOTAL * 100,
           "renglones_inconsistentes": sum(not i.get("linea_consistente", True) for i in items)}
    if not total_comprobante or not items:
        return {**out, "ok": False, "motivo": "sin total o sin renglones"}
    dif = abs(suma - total_comprobante) / abs(total_comprobante)
    out["diferencia_pct"] = round(dif * 100, 2)
    if dif <= TOLERANCIA_TOTAL:
        return {**out, "ok": True, "motivo": "suma de renglones = total"}
    for a in (0.0, *ALICUOTAS_IVA):
        esperado = suma * (1 + a) + imp + percs
        if abs(esperado - total_comprobante) / abs(total_comprobante) <= TOLERANCIA_TOTAL:
            det = f"IVA {a * 100:g} %" if a else "sin IVA"
            if imp:
                det += " + impuestos internos"
            if percs:
                det += " + percepciones"
            return {**out, "ok": True, "motivo": f"total = renglones + {det}"}
    # IVA mixto 21 % / 10,5 % (Villares, Raíces con varios rubros): se acepta si el IVA declarado cierra el total
    iva = float(extra.get("iva_21") or 0) + float(extra.get("iva_105") or 0) + float(extra.get("iva") or 0)
    if iva and abs(suma + imp + iva + percs - total_comprobante) / abs(total_comprobante) <= TOLERANCIA_TOTAL:
        return {**out, "ok": True, "motivo": "total = renglones + IVA declarado + cargos"}
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


# --------------------------------------------------------------------------- #
# Comprobantes reales de WhatsApp: tipo, copias, formatos de proveedor y LOCAL DE DESTINO
# --------------------------------------------------------------------------- #
# Razón social de los 4 locales (y de Chimi Doha): NO identifica el local por sí sola.
RAZON_SOCIAL_GRUPO = re.compile(r"lafinur\s*3368", re.I)
CUIT_LAFINUR_3368 = "30-71456509-1"
CUIT_CHIMI_DOHA = "30-71567213-4"


def _plano(t: str | None) -> str:
    return re.sub(r"\s+", " ", unicodedata.normalize("NFKD", t or "").encode("ascii", "ignore").decode().lower()).strip()


def detectar_local(sello: str | None = None, direccion_entrega: str | None = None, razon_social: str | None = None,
                   manuscrito: str | None = None, texto_cliente: str | None = None) -> dict:
    """Local de destino de un comprobante según la geografía real de Mishiguene.

    Prioridad: 1) sello (o 'RECIBIDO' manuscrito) · 2) anotación manuscrita del local · 3) dirección de entrega ·
    4) razón social propia de un local (Chimi Doha). 'LAFINUR 3368 SRL' es la razón social de los 4 locales
    (trampa fiscal): por sí sola NO asigna Lafinur y deja el local INDETERMINADO.

    Edificio Av. Congreso esq. Montañeses 1602 (Núñez): Planta Baja = Café Mishiguene Núñez (Chimidoha);
    Planta Alta = Centro de Producción (Congreso = Obrador, sello 'MISHIGUENE CERVIÑO').
    Devuelve {'local', 'regla', 'alertas'}."""
    # cada anotación manuscrita se evalúa por separado: 'Laffa 1 dev - mal estado' no anula 'Congreso (arriba a la derecha)'
    notas_ok = [n_ for n_ in re.split(r"\s\|\s", manuscrito or "") if not re.search(r"\bdev", _plano(n_))]
    s, d, r, m = _plano(sello), _plano(direccion_entrega), _plano(razon_social), _plano(" ".join(notas_ok))
    alertas: list[str] = []
    local, regla = None, None

    def por_texto(t: str, origen: str):
        if not t:
            return None, None
        if "cervi" in t:
            return "Congreso", f"{origen}: Mishiguene Cerviño = Centro de Producción/Obrador (PA del 1602)"
        if re.search(r"\bcongreso\b|\bobrador\b", t) and not re.search(r"nunez|chimi", t):
            return "Congreso", f"{origen}: anotación 'Congreso' / Obrador (Planta Alta del 1602)"
        if re.search(r"cafe\s+mishiguene\s+nunez|\bnunez\b|\bchimi\b", t):
            return "Chimidoha", f"{origen}: Café Mishiguene Núñez (PB del 1602)"
        if re.search(r"cafe\s+mishiguene\s+cabello|\bcabello\b", t):
            return "Cabello", f"{origen}: Cabello"
        if re.search(r"\barcos\b", t):
            return "Arcos", f"{origen}: Arcos"
        if re.search(r"\blafi\b|\blafinur\b", t) and not RAZON_SOCIAL_GRUPO.search(t):
            return "Lafinur", f"{origen}: Lafinur"
        return None, None

    local, regla = por_texto(s, "sello")
    if local is None and m and not re.search(r"\bdev", m):     # 'DEVO ... ARCOS' es una devolución, no un destino
        local, regla = por_texto(m, "anotación manuscrita")
    if local is None and d:
        if re.search(r"cabello\s*3181", d):
            local, regla = "Cabello", "dirección de entrega: Cabello 3181"
        elif re.search(r"\barcos\b", d):
            local, regla = "Arcos", "dirección de entrega: calle Arcos"
        elif re.search(r"cervi[nñ]o", d):
            local, regla = "Congreso", "dirección de entrega: Cerviño (Obrador)"
        elif re.search(r"congreso.*1602|montanes\w*\s*1602|1602.*montanes|congreso av\.? 1602", d):
            if re.search(r"chimi|nunez", r):
                local, regla = "Chimidoha", "dirección Congreso 1602 + razón social Chimi Doha / Café Núñez"
            else:
                local, regla = "Congreso", ("dirección Av. Congreso esq. Montañeses 1602 + Lafinur 3368 SRL sin sello de "
                                            "Café Núñez → Planta Alta / Obrador")
    if local is None and re.search(r"chimi\s*doha|cafe mishiguene nunez", r):
        local, regla = "Chimidoha", "razón social Chimi Doha / Café Mishiguene Núñez"
    # alertas de conflicto
    if local == "Congreso" and re.search(r"chimi\s*doha", r + " " + _plano(texto_cliente)):
        alertas.append("INFO compra centralizada: razón social Chimi Doha pero sello Cerviño → Congreso (Obrador) compra "
                       "y luego transfiere a Chimidoha con la matriz de recargos")
    if local == "Chimidoha" and re.search(r"\barcos\b", m):
        alertas.append("anotación 'Arcos' en un comprobante de Chimidoha: posible derivación/devolución")
    if local is None:
        alertas.append("local INDETERMINADO: 'Lafinur 3368 SRL' es la razón social de los 4 locales; sin sello ni "
                       "dirección de entrega no se asigna Lafinur")
    return {"local": local, "regla": regla, "alertas": alertas}


def separar_copias(texto: str) -> str:
    """Los PDF de AFIP traen ORIGINAL / DUPLICADO / TRIPLICADO: se usa solo la primera copia.
    Se corta donde reaparece el encabezado de la primera copia (o en la leyenda DUPLICADO/TRIPLICADO)."""
    lineas = [l for l in texto.splitlines() if l.strip()]
    if len(lineas) >= 6:
        cab = chr(10).join(lineas[:3])
        i = texto.find(cab, len(cab))
        if i > 0:
            texto = texto[:i]
    corte = re.search(r"^\s*(DUPLICADO|TRIPLICADO)\s*$", texto, re.M)
    if corte:
        return texto[:corte.start()]
    m = re.search(r"\bDUPLICADO\b", texto)
    if m and re.search(r"\bORIGINAL\b", texto[:m.start()]):
        return texto[:m.start()]
    return texto


def tipo_comprobante(texto: str) -> str:
    """Clasifica por evidencia del encabezado; no convierte un desconocido en factura."""
    cab = unicodedata.normalize("NFKD", texto[:1800]).encode("ascii", "ignore").decode().lower()
    if re.search(r"nota\s+de\s+credito", cab):
        return "NC"
    if re.search(r"nota\s+de\s+debito", cab):
        return "ND"
    if re.search(r"lista\s+(naturales|de precios)", cab):
        return "Lista"
    if re.search(r"(?:documento\s+)?no\s+valido\s+como\s+factura", cab):
        return "Remito"
    # Una factura puede mencionar el remito asociado: esa referencia no cambia su tipo.
    if re.search(r"\bfactura\s*(?:electronica\b|[abcme]\b|n[°ºo]|nro|numero|[-/]?\s*\d)", cab) or re.search(
            r"(?m)^\s*(?:[abcme]\s+)?factura\s*$", cab) or re.search(r"cod\.?\s*0*0[16]\b", cab):
        return "Factura"
    if re.search(r"(?m)^\s*(?:remito|rto\.?)\b", cab):
        return "Remito"
    return "Desconocido"


_FMT_AFIP = re.compile(
    r"^(?:\d+\s+(?=[A-Za-zÁÉÍÓÚÑ]))?(?P<d>.+?)\s+(?P<c>\d+(?:,\d+)?)\s+(?P<u>kg|kgs|unidades|unidad|otras unidades|lts?|u)\s+"
    r"(?P<p>\d+(?:,\d+)?)\s+(?P<b>\d+,\d{2})\s+(?:(?P<ib>\d+,\d{2})\s+)?(?P<s>\d+,\d{2})"
    r"(?:\s+(?P<iva>\d+(?:,\d+)?)%\s+(?P<si>\d+,\d{2}))?\s*$", re.I)
_FMT_RAICES = re.compile(
    r"^(?P<h>(?:\d+\s+)?\d+\s+[A-ZÁÉÍÓÚÑ0-9][A-ZÁÉÍÓÚÑ0-9 .]*?)\s+(?P<c>\d+\.\d{2})\s+(?P<p>[\d,]+\.\d{2})\s+(?P<t>[\d,]+\.\d{2})\s*$")
_FMT_FRIMAQ = re.compile(r"^(?P<c>\d+)\s+(?P<d>.+?)\s+(?P<p>\d+\.\d{2})\s+(?P<t>[\d,]+\.\d{2})\s*$")
_FMT_RODOFELI = re.compile(
    r"^(?P<cod>\S+(?:\s/\s\S+)?)\s+(?P<c>[\d,]+\.\d{2})\s+(?P<d>.+?)\s+(?P<p>[\d.]+)\s+\$\s+(?P<t>[\d,]+\.\d{2})\s*$")
_FMT_ALVEOLO = re.compile(
    r"^(?P<c>\d+(?:,\d+)?)\s+(?P<cod>\S+)\s+(?P<d>.+?)\s+(?P<p>[\d.]+,\d{2})\s+(?P<iva>[\d,]+)\s*%\s+(?P<b>[\d,]+)\s*%\s+"
    r"(?P<t>[\d.]+,\d{2})\s*$")
_FMT_NC_FREE = re.compile(
    r"^(?P<d>.+?)\s+(?P<c>[\d.,]+)\s+(?P<u>Kgs?|Un\w*)\s+(?P<p>[\d.]+,\d{2})\s+(?P<b>[\d,]+)\s+IVA\s+[\d.,]+%\s+\$\s*(?P<t>[\d.]+,\d{2})\s*$")
_FMT_NC_POSTA = re.compile(
    r"^(?P<d>.+?)\s+(?P<cod>[A-Z]{2}\d+)\s+(?P<u>UN|CAJ|KG)\s+(?P<c>[\d,]+)\s+(?P<p>[\d.]+,\d{2})\s+(?P<b>[\d,]+)%\s+(?P<t>[\d.]+,\d{2})\s*$")
_FMT_NC_VILLARES = re.compile(r"^(?P<d>NC MERC BONIFICADA.*?)\s+\([\d.]+%\)\s+(?P<p>[\d,]+\.\d{2})\s+(?P<t>[\d,]+\.\d{2})\s*$")

_UNID_RAICES = {"KILOS": "KG", "KILO": "KG", "UNIDAD": "UN", "CAJON": "CAJON", "BOLSAS": "BOLSA", "BOLSA": "BOLSA",
                "PAQUETE": "PAQ", "CAJA": "CAJA", "TORITO": "CAJON", "ATADO": "ATADO"}


def _item_pdf(d, c, u, p, b, t) -> dict:
    d = re.sub(r"\s{2,}", " ", d).strip(" .-")
    return {"descripcion_original": d, "cantidad": c, "unidad_detectada": u, "precio_unitario": p,
            "descuento_pct": b or 0.0, "total_linea": t,
            "linea_consistente": abs(c * p * (1 - (b or 0) / 100) - t) <= max(1.0, 0.02 * abs(t))}


def items_pdf_comprobante(texto: str) -> list[dict]:
    """Renglones de los formatos de proveedor vistos en Facturas_WPP/ (AFIP/ARCA genérico, Raíces, Frimaq,
    Rodofeli, Alveolo, NC de Free / Posta / Villares)."""
    items: list[dict] = []
    for ln in texto.splitlines():
        ln = ln.strip()
        if not ln:
            continue
        m = _FMT_AFIP.match(ln)
        if m:
            g = m.groupdict()
            u = {"kg": "KG", "kgs": "KG", "lt": "LT", "lts": "LT"}.get(g["u"].lower(), "UN")
            items.append(_item_pdf(g["d"], parse_number(g["c"]), u, parse_number(g["p"]), parse_number(g["b"]),
                                   parse_number(g["s"])))
            continue
        m = _FMT_RAICES.match(ln)
        if m:
            h = re.sub(r"\.{2,}", " ", m["h"])
            h = re.sub(r"^\d+\s+\d+\s+|^\d+\s+", "", h.strip(), count=1).strip()
            toks = h.split()
            u = _UNID_RAICES.get(toks[-1].upper()) if toks else None
            items.append(_item_pdf(" ".join(toks), float(m["c"]), u, float(m["p"].replace(",", "")), 0.0,
                                   float(m["t"].replace(",", ""))))
            continue
        for rx, kind in ((_FMT_RODOFELI, "rod"), (_FMT_ALVEOLO, "alv"), (_FMT_NC_FREE, "free"),
                         (_FMT_NC_POSTA, "posta"), (_FMT_NC_VILLARES, "vil"), (_FMT_FRIMAQ, "fri")):
            m = rx.match(ln)
            if not m:
                continue
            g = m.groupdict()
            if kind == "rod":
                items.append(_item_pdf(g["d"], float(g["c"].replace(",", "")), "UN", float(g["p"]), 0.0,
                                       float(g["t"].replace(",", ""))))
            elif kind == "alv":
                items.append(_item_pdf(g["d"], parse_number(g["c"]), "UN", parse_number(g["p"]), parse_number(g["b"]),
                                       parse_number(g["t"])))
            elif kind in ("free", "posta"):
                u = "KG" if g["u"].lower().startswith("k") else ("CAJA" if g["u"] == "CAJ" else "UN")
                items.append(_item_pdf(g["d"], parse_number(g["c"]), u, parse_number(g["p"]), parse_number(g["b"]),
                                       parse_number(g["t"])))
            elif kind == "vil":
                t = float(g["t"].replace(",", ""))
                items.append(_item_pdf(g["d"], 1.0, "UN", t, 0.0, t))
            else:
                t = float(g["t"].replace(",", ""))
                items.append(_item_pdf(g["d"], float(g["c"]), "UN", float(g["p"]), 0.0, t))
            break
    return items


def _numero_comprobante(texto: str) -> str | None:
    pats = [r"Punto de Venta:\s*(\d+)\s+Comp\.\s*Nro:\s*(\d+)",
            r"N[ºo°]:?\s*(\d{4})\s*-\s*(\d{8})", r"FACTURA\s*-?\s*(\d{4})/(\d{8})", r"\b(A\d{5})-(\d{8})\b",
            r"N[úu]mero:\s*(\d{5})-(\d{8})", r"NOTA DE CREDITO\s+(\d{5})-(\d{8})", r"\b(0010)-(\d{8})\b",
            r"\b(0[0-9]{3,4})-(\d{8})\b"]
    for p in pats:
        m = re.search(p, texto, re.I)
        if m:
            return f"{m.group(1)}-{m.group(2)}"
    return None


def _total_comprobante(texto: str) -> float | None:
    pats = [r"Importe Total:\s*\$\s*([\d.,]+)", r"TOTAL COMPROBANTE\.*\s*([\d,]+\.\d{2})", r"TOTAL PES\s*([\d.]+)",
            r"TOTAL Pesos\s*:\s*([\d,]+\.\d{2})", r"^Total:\s*([\d.,]+)\s*$", r"^Total\s+\$\s*([\d.,]+)\s*$",
            r"^TOTAL \$\s*([\d.,]+)", r"Importe Total:\s*\$([\d.,]+)"]
    for p in pats:
        m = re.search(p, texto, re.I | re.M)
        if m:
            raw = m.group(1)
            if re.fullmatch(r"[\d,]+\.\d{2}", raw):
                return float(raw.replace(",", ""))
            return parse_number(raw)
    return None


_PROVEEDORES_CONOCIDOS = (
    (r"RAICES|HENRIQUEZ", "RAICES"), (r"postamayorista|POSTA EXPRESS", "POSTA"), (r"Free Vegetales", "FREE"),
    (r"alimentosvillares|Villares", "VILLARES"), (r"ALVEOLO", "ALVEOLO SA (Levando)"), (r"FRIMAQ", "FRIMAQ S.R.L."),
    (r"RODOFELI", "ROBERTO O. RODOFELI Y CIA S.R.L."), (r"DISTRIBUIDORA KOSTA", "DISTRIBUIDORA KOSTA"),
    (r"CARBALLO", "CARBALLO FERNANDO ARIEL"), (r"SORRENTINO", "SORRENTINO LUCAS JAVIER"),
    (r"VILLARINO", "VILLARINO LORENA ALEJANDRA"), (r"PROCAMP", "PROCAMP (Sola Ogliese Patricio)"),
    (r"Grupo Free|grupofree", "GRUPO FREE"))


def leer_pdf_comprobante(path: str | Path) -> dict:
    """PDF digital de proveedor -> dict de comprobante (solo la copia ORIGINAL)."""
    path = Path(path)
    with pdfplumber.open(path) as pdf:
        texto = "\n".join(p.extract_text() or "" for p in pdf.pages)
    t = separar_copias(texto)
    tipo = tipo_comprobante(t)
    f = (re.search(r"(?:Fecha de Emisi[oó]n|Fecha emisi[oó]n|FECHA|Fecha)\s*:?\s*(\d{2}/\d{2}/\d{4})", t)
         or re.search(r"(\d{2}/\d{2}/\d{4})", t))
    items = items_pdf_comprobante(t)
    total = _total_comprobante(t)
    prov = next((nom for rx, nom in _PROVEEDORES_CONOCIDOS if re.search(rx, t, re.I)), None)
    entrega = None
    if re.search(r"AV\. CONGRESO ESQ\. MONTA", t):
        entrega = "AV. CONGRESO ESQ. MONTAÑESES 1602"
    elif re.search(r"Entregar:\s*CONGRESO 1602", t):
        entrega = "CONGRESO 1602 NUÑEZ"
    elif re.search(r"Direcci[oó]n de entrega:\s*Congreso 1602", t):
        entrega = "Congreso 1602 esquina Montañeses"
    else:
        cli = t.split("CUIT: 30714565091")[-1] if "CUIT: 30714565091" in t else t
        me = re.search(r"Domicilio(?: Comercial)?:\s*([^\n]+)", cli)
        entrega = me.group(1).strip() if me else None
    razon = "LAFINUR 3368 SRL" if re.search(r"LAFINUR 3368", t, re.I) else None
    ref = None
    if tipo == "NC":
        lineas = [l.strip() for l in t.splitlines()
                  if re.search(r"rechaza|error de carga|mal\s+estado|bonific|fact\.|FC\s*\d+|Reversi|Referencia", l, re.I)]
        ref = " | ".join(lineas)[:300] or None
    remito = re.search(r"REMITO:\s*(\d{4}/\d{8})", t)
    return {
        "archivo": path.name, "tipo": tipo, "numero_factura": _numero_comprobante(t),
        "fecha_factura": _fecha_iso(f.group(1)) if f else None, "proveedor_detectado": prov,
        "items": items, "total_comprobante": total, "razon_social_cliente": razon, "direccion_entrega": entrega,
        "sello": None, "firmas": [], "referencia": ref, "remito_asociado": remito.group(1) if remito else None,
        "texto_encabezado": t[:600], "motor": "pdf-texto (primera copia)",
        "validacion": validar_aritmetica(items, total) if items else {"ok": False, "motivo": "sin renglones"},
    }


def dhash_imagen(path: str | Path, lado: int = 16) -> str:
    """Hash perceptual (dHash 16x16 = 256 bits, hex). Sobrevive a la recompresión y al reescalado de Telegram/WhatsApp."""
    import cv2
    import numpy as np
    img = cv2.cvtColor(_cargar_bgr(Path(path)), cv2.COLOR_BGR2GRAY)
    r = cv2.resize(img, (lado + 1, lado), interpolation=cv2.INTER_AREA)
    bits = (r[:, 1:] > r[:, :-1]).flatten()
    return format(int("".join("1" if b else "0" for b in bits), 2), "x")


def _hamming(a: str, b: str) -> int:
    return bin(int(a, 16) ^ int(b, 16)).count("1")


def _digitos_en_ocr(path: Path, numeros: list[str]) -> bool:
    """Verificación rápida: el OCR (imagen reducida) tiene que ver los N° de los comprobantes candidatos."""
    import cv2
    nombre, ocr = _motor_ocr()
    img = _cargar_bgr(path)
    h, w = img.shape[:2]
    if w > 1100:
        img = cv2.resize(img, None, fx=1100 / w, fy=1100 / w, interpolation=cv2.INTER_AREA)
    texto = re.sub(r"\D+", " ", ocr(cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)))
    return all(re.search(rf"(?<!\d)0*{n}(?!\d)", texto) for n in numeros if n)


def reutilizar_lectura_por_parecido(path: str | Path, umbral: int = 40) -> list[dict]:
    """Telegram/WhatsApp recomprimen las fotos y cambian el SHA-256. Busca en cache_facturas.json una lectura visual de la
    misma foto por hash perceptual y la CONFIRMA con OCR rápido (los N° de comprobante tienen que figurar en la imagen).
    Si coincide, registra el alias por SHA-256 y devuelve los comprobantes: respuesta en ~1-3 s sin llamar a la API de visión."""
    path = Path(path)
    cache = _leer_cache()
    try:
        mio = dhash_imagen(path)
    except Exception:
        return []
    h = sha256_archivo(path)
    por_img: dict[str, list[tuple[str, dict]]] = {}
    for k, v in cache.items():
        if v.get("dhash") and _hamming(mio, v["dhash"]) <= umbral:
            por_img.setdefault(str(v.get("sha256", k.split("#")[0])), []).append((k, v))
    if not por_img:
        return []
    mejor = min(por_img.values(), key=lambda lst: _hamming(mio, lst[0][1]["dhash"]))
    numeros = [str(int(re.findall(r"\d+", str(v.get("numero_factura") or ""))[-1])) for _, v in mejor
               if re.findall(r"\d+", str(v.get("numero_factura") or ""))]
    distancia = _hamming(mio, mejor[0][1]["dhash"])
    if distancia > 6:                    # casi idéntica (<=6 de 256 bits) se acepta; si no, la confirma el OCR
        if not numeros or not _digitos_en_ocr(path, numeros):
            return []
    out = []
    for i, (_, v) in enumerate(mejor, 1):
        c = dict(v)
        c["sha256"] = h
        c["archivo"] = path.name if len(mejor) == 1 else f"{path.name} [comprobante {i}]"
        c["motor"] = (v.get("motor") or "") + " (reutilizada por parecido visual + OCR)"
        cache[h if len(mejor) == 1 else f"{h}#{i}"] = c
        out.append(c)
    CACHE_FACTURAS.write_text(json.dumps(cache, ensure_ascii=False, indent=1), encoding="utf-8")
    return out


def indexar_dhash(carpeta: str | Path) -> int:
    """Agrega 'dhash' a las lecturas de la caché cuyo archivo está en `carpeta` (se corre una vez por imagen)."""
    carpeta = Path(carpeta)
    cache = _leer_cache()
    por_sha = {}
    for p in carpeta.iterdir():
        if p.is_file() and p.suffix.lower() in IMG_EXT | {""}:
            try:
                por_sha[sha256_archivo(p)] = p
            except OSError:
                continue
    n = 0
    for k, v in cache.items():
        s = str(v.get("sha256", k.split("#")[0]))
        if not v.get("dhash") and s in por_sha:
            try:
                v["dhash"] = dhash_imagen(por_sha[s])
                n += 1
            except Exception:
                continue
    if n:
        CACHE_FACTURAS.write_text(json.dumps(cache, ensure_ascii=False, indent=1), encoding="utf-8")
    return n


_PROMPT_VISION = (
    "Sos un lector de comprobantes argentinos (facturas, remitos, notas de crédito). Leé la imagen adjunta ({ruta}). "
    "Puede haber MÁS DE UN comprobante en la misma foto (p. ej. dos facturas lado a lado): separalos. "
    "Respetá las correcciones a mano en birome (cantidades tachadas o reescritas, 'NO', pesos reales): la cantidad "
    "vigente es la manuscrita. Devolvé SOLO un JSON (sin texto extra) con una lista de objetos: "
    '{{"tipo":"Factura|Remito|NC|RemitoObrador","numero_factura":"PPPPP-NNNNNNNN","fecha_factura":"YYYY-MM-DD",'
    '"local_destino_base":"local de entrega base o null","proveedor_detectado":"...","sello":"texto del sello o null","direccion_entrega":"...|null",'
    '"razon_social_cliente":"...","firmas":[...],"anotaciones_manuscritas":[...],"total_comprobante":0,'
    '"items":[{{"local_destino":"local solo si anotado para este renglon, o null","descripcion_original":"...","cantidad":0,"unidad_detectada":"UN|KG|...","precio_unitario":0,'
    '"descuento_pct":0,"total_linea":0}}]}}. Una flecha en un renglón afecta solo ese renglón, con local_destino; local_destino_base conserva la entrega del resto. Usá RemitoObrador si lo emite Mishiguene (Av. Cerviño 4417) a un local.')


_GEMINI_BLOQ_HASTA = 0.0      # corte de circuito: clave rechazada / cuota agotada -> OCR sin reintentar cada foto
_GEMINI_ULTIMA = 0.0


def _gemini_modelos() -> list[str]:
    """GEMINI_MODEL (si está) y luego alias vigentes; gemini-1.5-flash fue dado de baja por Google (404)."""
    pref = (os.environ.get("GEMINI_MODEL") or "").strip()
    return [m for m in dict.fromkeys([pref, "gemini-3.8-flash", "gemini-flash-latest"]) if m]


def _cargar_env() -> None:
    try:
        from dotenv import load_dotenv
        load_dotenv(Path(__file__).with_name(".env"))
    except ImportError:                                      # sin python-dotenv: lectura directa
        try:
            for ln in Path(__file__).with_name(".env").read_text(encoding="utf-8").splitlines():
                k, _, v = ln.partition("=")
                if k.strip() and not k.startswith("#"):
                    os.environ.setdefault(k.strip(), v.strip().strip('"').strip("'"))
        except OSError:
            pass


def _vision_gemini(path: Path, timeout: int = 60) -> str | None:
    """Lee la foto con Gemini (API, sin ventanas ni procesos hijos). Devuelve el texto JSON o None si no hay clave,
    red, cuota o el modelo falla: el llamador cae a Tesseract. Nunca lanza ni registra la clave."""
    import time
    global _GEMINI_BLOQ_HASTA, _GEMINI_ULTIMA
    _cargar_env()
    key = os.environ.get("GEMINI_API_KEY")
    if not key or time.time() < _GEMINI_BLOQ_HASTA:
        return None
    try:
        import warnings
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            import google.generativeai as genai
        genai.configure(api_key=key)
        mime = {".png": "image/png", ".webp": "image/webp"}.get(path.suffix.lower(), "image/jpeg")
        datos = {"mime_type": mime, "data": path.read_bytes()}
        espera = float(os.environ.get("GEMINI_MIN_INTERVALO", "4.2")) - (time.time() - _GEMINI_ULTIMA)
        if espera > 0:                                       # nivel gratuito: ~15 pedidos/min
            time.sleep(espera)
        ultimo = ""
        for modelo in _gemini_modelos():
            try:
                _GEMINI_ULTIMA = time.time()
                r = genai.GenerativeModel(modelo).generate_content(
                    [_PROMPT_VISION.format(ruta=path.name), datos],
                    generation_config={"temperature": 0, "response_mime_type": "application/json"},
                    request_options={"timeout": timeout})
                return r.text
            except Exception as e:                           # noqa: BLE001
                ultimo = f"{modelo}: {type(e).__name__} {str(e)[:100]}".replace(key, "<clave>")
                nombre = type(e).__name__
                if nombre in ("PermissionDenied", "Unauthenticated", "InvalidArgument"):
                    _GEMINI_BLOQ_HASTA = time.time() + 600   # clave inválida: no insistir 10 min
                    break
                if nombre in ("ResourceExhausted", "TooManyRequests"):
                    _GEMINI_BLOQ_HASTA = time.time() + 90
                    break
        LOG.warning("Gemini falló (%s): se usa OCR local", ultimo)
    except Exception as e:                                   # noqa: BLE001
        LOG.warning("Gemini no disponible (%s): se usa OCR local", f"{type(e).__name__} {e}"[:120].replace(key, "<clave>"))
    return None


def leer_imagen_vision(path: str | Path, timeout: int = 60) -> list[dict]:
    """Foto que no está en cache_facturas.json: Gemini (API) separa los comprobantes, respeta la birome y devuelve
    la estructura. Si no responde lanza OCRNoDisponible y el llamador usa Tesseract. Cada comprobante se guarda en
    cache_facturas.json por SHA-256 ('<sha>' o '<sha>#n' si son varios)."""
    path = Path(path)
    h = sha256_archivo(path)
    cache = _leer_cache()
    if h in cache:
        STATS["cache_sha"] += 1
        return [cache[h]]
    previos = [v for k, v in cache.items() if k.startswith(h + "#")]
    if previos:
        STATS["cache_sha"] += 1
        return previos
    motor = "vision-gemini"
    txt, rc, err = _vision_gemini(path, timeout), 0, ""
    if txt:
        STATS["vision_gemini"] += 1
    else:
        raise OCRNoDisponible("Gemini sin respuesta (clave, red o cuota): se usa OCR local")
    m = re.search(r"\[.*\]|\{.*\}", txt or "", re.S)
    if rc != 0 or not m:
        raise OCRNoDisponible(f"visión sin respuesta utilizable ({rc}): {(err or txt or '')[:160]}")
    datos = json.loads(m.group(0))
    datos = datos if isinstance(datos, list) else [datos]
    out = []
    for i, c in enumerate(datos):
        for it in c.get("items", []):
            it.setdefault("descuento_pct", 0.0)
            it.setdefault("total_linea", round(it["cantidad"] * it["precio_unitario"] * (1 - it["descuento_pct"] / 100), 2))
            it["linea_consistente"] = abs(it["cantidad"] * it["precio_unitario"] * (1 - it["descuento_pct"] / 100)
                                          - it["total_linea"]) <= max(1.0, 0.02 * abs(it["total_linea"]))
        c.update({"archivo": path.name if len(datos) == 1 else f"{path.name} [comprobante {i + 1}]", "sha256": h,
                  "motor": motor, "validacion": validar_aritmetica(
                      c.get("items", []), c.get("total_comprobante"), c.get("extra")), "texto_encabezado": ""})
        cache[h if len(datos) == 1 else f"{h}#{i + 1}"] = c
        out.append(c)
    CACHE_FACTURAS.write_text(json.dumps(cache, ensure_ascii=False, indent=1), encoding="utf-8")
    return out


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
