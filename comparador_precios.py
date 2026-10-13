"""Comparador de precios: Dossier Mensual vs listas de proveedores en Google Drive.

1. Descarga las listas crudas de la carpeta 'Listas proveedores' de Drive
   (Google Sheets se exportan a .xlsx; PDFs se bajan tal cual).
2. Parsea cada lista con su lector específico (Posta, Emporio, Villares,
   Alyser, Raices, La Mesón, Open Fruit; El Criollo desde el CSV del scraper;
   Nogada scrapeando su tienda WooCommerce) y lleva cada precio a una base
   comparable ($/KG, $/LT o $/UN). El proveedor se reconoce por el nombre del archivo.
3. Cruza cada fila del Dossier (Precio Ponderado) contra el mejor match de
   CADA proveedor: matriz con una columna de precio por proveedor ('-' si no
   lo tiene), Mejor Opción (el de menor precio comparable) y Ahorro Unitario.
   Si 'Precios/productos a buscar.txt' pide una MARCA ('Tahina kanater',
   'Manteca Milkaut'), esa marca es obligatoria en el match: no gana una
   marca genérica.
   Si un insumo tiene varias variantes válidas (marca, variedad o tamaño de
   envase) se generan hasta 3 filas: 'Insumo (Opción 1)', '(Opción 2)'...; cada
   proveedor cotiza esa variante o lleva '-'. El ahorro total cuenta solo la
   Opción 1 de cada insumo (no suma dos veces el mismo consumo).
   El Nuevo Emporio se compara con su precio C/DTO (lista - 15%).
   Los insumos de LISTA_NEGRA no se procesan ni aparecen en los reportes.
   Miel: se cotiza todo el mercado (no solo orgánica) en 3 opciones fijas: baldes
   mayoristas, frasco de hasta 2 kg y orgánica, siempre en $/KG.
   Sardo: 'Queso Sardo Estacionado' (Gauchito/La Paulina) y 'Queso Sardo Semi
   Gauchito' salen como filas separadas con su nombre.
4. Genera un Excel por 'Categoría 2' (Reporte_Ahorro_Almacen.xlsx, ...) con
   las hojas Comparativa / Oportunidades / Sin match (filas en orden alfabético
   por producto), más un Reporte_Ahorro_Resumen.xlsx, en Reportes_Ahorro/.

La lógica de normalización y los filtros de match (FORBIDDEN_MATCHES,
CRITICAL_QUALIFIERS, abreviaturas) vienen de Precios/run_pipeline_pricing.py;
la lectura del PDF de Raices, de Verduleria/extraer_precios.py.

Uso:
    python comparador_precios.py
    python comparador_precios.py --hoja "Cadena Septiembre"
    python comparador_precios.py --listas-dir carpeta_local   # sin Drive
"""

from __future__ import annotations

import argparse
import html
import logging
import math
import re
import sys
import tempfile
import unicodedata
from pathlib import Path
from typing import Any, Callable

import pandas as pd
import pdfplumber
import requests
from google.auth.transport.requests import AuthorizedSession
from google.oauth2.service_account import Credentials
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from rapidfuzz import fuzz

BASE_DIR = Path(__file__).resolve().parent
DEFAULT_CREDS = BASE_DIR / "credentials.json"
DOSSIER = BASE_DIR / "Dossier_Mensual_Mishiguene.xlsx"
OUTPUT_DIR = BASE_DIR / "Reportes_Ahorro"

# Carpeta 'Listas proveedores' en Drive (compartida con la Service Account)
DRIVE_FOLDER_ID = "1s4RlRFVO-1NpW0sxyG-i0B2pHML9JU9P"
DRIVE_SCOPES = ["https://www.googleapis.com/auth/drive.readonly"]
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

# Catálogo de El Criollo scrapeado por Precios/run_pipeline_pricing.py (no está en Drive)
CRIOLLO_CSV = BASE_DIR / "Precios" / "catalogo_criollo_raw.csv"

# Nogada (nogada.com.ar): tienda WooCommerce; se scrapea su API pública de productos
# y se deja una copia en CSV (se usa de respaldo si la web no responde).
NOGADA_API = "https://nogada.com.ar/wp-json/wc/store/v1/products"
NOGADA_CSV = BASE_DIR / "Precios" / "catalogo_nogada_raw.csv"

# Lo que pedís comprar, una línea por producto ('Tahina kanater'): de ahí salen las marcas obligatorias
PRODUCTOS_TXT = BASE_DIR / "Precios" / "productos a buscar.txt"

# Marcas prioritarias de los quesos: el txt los pide sin marca ('Queso Muzzarella'),
# pero solo se cotizan estas (alcanza con cumplir una). Clave = línea del txt normalizada.
MARCAS_PRIORITARIAS = {
    "queso muzzarella": ["gauchito"],                                   # plancha/barra 10 kg El Gauchito (Emporio)
    "queso pategras": ["paulina", "tregar", "melincue", "gauchito"],
    "queso sardo estacionado": ["gauchito", "paulina"],
    "queso sardo semi estacionado": ["gauchito"],
}

# Umbrales de match / reporte
SCORE_MINIMO = 65            # score 0-100 para aceptar un match (los filtros duros ya descartan lo absurdo)
SCORE_MARCA = 50             # umbral cuando el match además cumple una marca exigida
SCORE_ALTA = 85              # desde acá la confianza es 'Alta'
TOLERANCIA_UNIDAD = 3.0      # si dossier/lista difiere más de 3x, la unidad no es comparable
SOBREPRECIO_SOSPECHOSO = 40.0  # más de 40% de diferencia suele ser otro producto/presentación
AHORRO_MINIMO_PCT = 3.0      # no reportar diferencias menores a este %
# Productos con proveedor real de compra y marca fija. Se muestran y se comparan con todos los proveedores
# (solo con la marca indicada); si el mejor precio es el del proveedor real, no hay delta ("proveedor actual").
# El precio siempre se lleva a $/KG: una bolsa institucional (p. ej. 10 KG) se cotiza como $ bolsa / 10
# (parse_presentacion + precio_alineado).
PRODUCTOS_ESPECIALES = [
    {"nombre": "PANKO BLANCO GOCHIS", "producto": {"panko"}, "marca": "gochiso", "proveedor_real": "Posta"},
]
# Filas del Dossier que son un error de origen (bulto/caja cargado como unidad) y no deben compararse.
# 'precio_max': por debajo de ese Precio Ponderado la fila es el error. Se informa en el log.
DOSSIER_EXCLUIR = [
    {"sku": "TAH", "precio_max": 5000.0,
     "motivo": "Tahina Frasco 900 g cargada por bulto/caja de 12 (precio/12): no es un frasco"},
]
# Envases de tamaño fijo: si el Dossier nombra el tamaño ('Frasco X 900Grs'), un candidato de otro tamaño
# (pote de 400 g) no es el mismo producto aunque el precio por KG se parezca.
ENVASE_FIJO_RE = re.compile(r"\b(frasco|pote|lata|botella|sachet|tarro|doypack)s?\b", re.I)
TOLERANCIA_TAMANO_ENVASE = 0.25
SEMILLAS_ENTERAS = {"sesamo", "chia", "lino", "amapola", "girasol", "zapallo"}
MAX_OPCIONES = 3             # variantes (Opción 1..3) que se muestran por insumo
CERCANIA_BUCKET = 0.15       # opciones a menos de ~15% de distancia al precio del Dossier empatan en cercanía
DESCUENTO_EMPORIO = 0.85     # El Nuevo Emporio: la columna C/DTO es lista - 15%

# Insumos del Dossier que se excluyen por completo (ruido / duplicados). Se comparan
# por nombre normalizado (sin tildes, mayúsculas ni espacios de más).
LISTA_NEGRA = [
    "Azucar caja x 5 kg",
    "Aceite De Oliva Zuelo",
    "Aceite girasol ALSAMAR",      # queda solo el 'Aceite girasol' genérico
    "Almidon De Maiz x kg",        # ya está 'Fécula de maíz'
]

LOG = logging.getLogger("comparador_precios")


# =========================================================================== #
# Normalización y match (base: Precios/run_pipeline_pricing.py)
# =========================================================================== #
STOP_WORDS = {
    "un", "und", "unidad", "unidades", "uni", "kg", "kilo", "kilos", "gr", "g",
    "gramo", "gramos", "ml", "lt", "l", "litro", "litros", "cc", "pack", "packs",
    "caja", "cajas", "x", "por", "de", "del", "la", "el", "los", "las", "y", "con",
    "en", "sin", "gluten", "tacc", "s", "tableta", "cajon", "balde", "of", "aprox",
    "bolsa", "tipo", "imp", "nuevo", "nueva", "oferta", "frasco", "bidon", "bot",
}

# Nombres de proveedor: en el Dossier indican a quién se le compra ('Garbanzos
# Villares (1kg)'), no son parte del producto, así que no se buscan en la lista.
PROVEEDOR_WORDS = {"villares", "criollo", "posta", "emporio", "alyser", "raices", "nogada"}

CRITICAL_QUALIFIERS = {
    "gochiso", "nutripanko", "naccato", "gauchito", "vaquero",
    "fontina", "pategras", "sardo", "semi", "estacionado",
    "vacalin", "repostero", "familiar", "tradicional",
    "entrefina", "gruesa", "fina", "celusal", "anclas",
    "milkaut", "paulina", "tregar", "calsa", "fenix", "garcia",
    "dolce", "dolca", "nescafe", "lucchetti", "favinco", "menoyo", "morixe", "44", "amargo",
    "tinto", "blanco", "blanca", "muzzarella", "grano", "molido",
    "verdes", "negras", "descarozada", "oliva", "oleico", "trufa", "organico", "organica",
    "000", "0000", "integral", "descremada", "entera", "ahumado", "dijon", "antigua",
    # varietales de vino
    "malbec", "pinot", "cabernet", "merlot", "chardonnay", "torrontes", "sauvignon", "syrah", "bonarda",
}

FORBIDDEN_MATCHES = {
    "sal": {"aceite", "oliva", "vinagre"},
    "fontina": {"blanco", "crema", "unt", "untable", "untar", "chedd", "cheddar", "sardo", "pategras", "muzzarella"},
    "muzzarella": {"raviol", "ravioles", "sorrentino", "sorrentinos", "pasta", "pastas", "pizza", "pizzas",
                   "empanada", "empanadas", "disco", "discos", "saltena", "canuelas"},
    "repostero": {"familiar", "tradicional", "san ignacio"},
    "familiar": {"repostero"},
    "tradicional": {"repostero"},
    "entrefina": {"gruesa", "fina"},
    "gruesa": {"entrefina", "fina"},
    "mostaza": {"arytza", "savora", "aderezo", "dips", "salsa", "pasta"},
    "avellana": {"syrup", "sirope", "jarabe", "pehuenia", "salsa", "pasta", "bano", "relleno",
                 "lheritier", "licor", "aroma", "esencia", "topping", "harina",
                 "chocolate", "cacao", "leche", "garrapinada"},
    "avellanas": {"syrup", "sirope", "jarabe", "pehuenia", "salsa", "pasta", "bano", "relleno",
                  "lheritier", "licor", "aroma", "esencia", "topping", "harina",
                  "chocolate", "cacao", "leche", "garrapinada"},
    "amargo": {"alcalino", "dulce", "leche"},
    "crema": {"39", "repostera", "vegetal", "chantilly"},
    "aceite": {"rocio", "ind", "abedul"},
    "girasol": {"oliva", "maiz", "soja", "oleico"},
    "ricota": {"raviol", "ravioles", "sorrentino", "sorrentinos", "tapa", "tapas"},
    "vino": {"rosado"},
    "tinto": {"blanco", "rosado"},
    "blanco": {"tinto", "rosado"},
    "grano": {"molido", "licor"},
    "molido": {"grano", "entero"},
    "verdes": {"negras", "negra", "fileteada", "fileteadas", "rellena", "rellenas", "morron"},
    "negras": {"verdes", "verde", "fileteada", "fileteadas", "rellena", "rellenas", "morron"},
    "descarozada": {"fileteada", "fileteadas", "rellena", "rellenas", "morron", "pasta", "rodaja", "hueso"},
    "azucar": {"muneca", "mascabo", "rubia", "impalpable", "negra", "talco", "organico", "coco", "datil"},
    "manteca": {"mani", "cacao", "karite"},
    "feta": {"emmenthal", "danbo", "dambo", "tybo", "cheddar", "pategras", "fontina", "jamon"},  # 'en fetas' = feteado
    "vaso": {"ensaladera"},
    "chocolate": {"bano", "cobertura", "relleno", "leche"},
    "almendras": {"harina", "bebida", "beb", "pasta"},
    "almendra": {"harina", "bebida", "beb", "pasta"},
    "harina": {"premezcla", "rebozador"},
}

# Marcas que, si aparecen en 'productos a buscar.txt' (o en el nombre del Dossier),
# pasan a ser OBLIGATORIAS en el match. Tokens ya normalizados (match_key).
# Una marca nueva en el txt solo se exige si está en este set.
BRANDS = {
    "melincue",
    "kanater", "milkaut", "gauchito", "paulina", "vacalin", "fenix", "calsa", "tregar", "lucchetti",
    "favinco", "menoyo", "celusal", "anclas", "garcia", "emeth", "hurlingham", "nescafe", "dolca",
    "gochiso", "naccato", "nutripanko", "vaquero", "areco", "dos",  # "dos" = Dos Anclas
}
# Grafías alternativas que una lista puede traer para la misma marca
BRAND_ALIASES = {"dolca": {"dolce"}}

# Qué distingue una variante de otra del mismo insumo (además de CRITICAL_QUALIFIERS):
# marcas de consumo y variedades/cortes. Lo que no figure acá no separa variantes.
VARIANT_BRANDS = BRANDS | {"alsamar", "legitimo", "natura", "canuelas", "ullum", "emporio", "laur", "rdl",
                           "zuelo", "agromar", "inalpa", "cumana", "argenfrut", "canale", "banda", "sytari"}
VARIANT_WORDS = {"sauco", "lechoso", "mediano", "mediana", "chico", "chica", "grande", "fileteada", "laminada",
                 "piel", "tostada", "salada", "cascara", "entero", "entera", "light", "premium",
                 "parboil", "suave", "tradicion", "clasico", "rayado", "partido", "cubierto"}

ABBREVS = {
    r"\bddl\b": "dulce leche", r"\brep\b": "repostero", r"\bfam\b": "familiar",
    r"\brall\b": "rallado", r"\bcrem\b": "crema", r"\blech\b": "leche", r"\bente\b": "entera",
    r"\bpel\b": "pelada", r"\btto\b": "tinto", r"\bbco\b": "blanco", r"\bbca\b": "blanca",
    r"\bgochis\b": "gochiso", r"\bestac\b": "estacionado", r"\b44%": "44", r"\b39%": "39",
    r"\ben polvo\b": "molido", r"\bpolvo\b": "molido", r"\bmolida\b": "molido",
    r"\bsemillas\b": "grano", r"\bsemilla\b": "grano", r"\bgra\b": "grano",
    r"\be/fn\b": "entrefina", r"\be/f\b": "entrefina", r"\bentre fina\b": "entrefina", r"\bluchetti\b": "lucchetti", r"\bluccheti\b": "lucchetti",
    r"\bmuzz\b": "muzzarella", r"\bmozz\b": "muzzarella", r"\bmozzar\b": "muzzarella",
    r"\bmozzarella\b": "muzzarella", r"\bmuza\b": "muzzarella",
    r"\bpasc\b": "tarta", r"\bpascualina\b": "tarta", r"\btartas\b": "tarta", r"\btapas\b": "tapa",
    r"\bdesc\b": "descarozada", r"\bdescaroz\b": "descarozada", r"\bdescarozados?\b": "descarozada",
    r"\bdescarozadas\b": "descarozada", r"\bs/carozos?\b": "descarozada", r"\bs/c\b": "descarozada",
    # abreviaturas típicas de Emporio
    r"\bgira\b": "girasol", r"\boliv\b": "oliva", r"\bahu\b": "ahumado", r"\bchoco\b": "chocolate",
    r"\bhellm\b": "hellmanns", r"\bmayo\b": "mayonesa", r"\bcgl\b": "congelado", r"\bao\b": "oleico",
    r"\balto oleico\b": "oleico", r"\buntable\b": "untar", r"\bgall\b": "gallina", r"\bperi\b": "perita",
    r"\borganica\b": "organico", r"\bgarbanzos\b": "garbanzo", r"\bbn\b": "bano",
    r"\bazucar imp\b": "azucar impalpable",  # 'imp' suelto es 'importado' (stop word)
}

# Cantidad + unidad dentro del nombre ('x 10kg', '900Grs', '12 x 1,5 LT'): se
# quitan antes de matchear para que el tamaño del envase no bloquee el match.
_UNIDADES = r"(?:kgs?|k|kilos?|grs?|gramos?|g|lts?|litros?|l|ml|cc|u|un|und|uni|unid)\b"
_SIZE_RE = re.compile(rf"(?:\b\d+\s*x\s*)?\d+(?:[.,]\d+)?\s*{_UNIDADES}\.?", re.I)
# Medidas '30 X 40' / '45*500' -> un solo token '30x40' (identifica el producto).
# Va antes de _SIZE_RE; '12 x 1,5 LT' o '10 X 100 UN' no son medidas.
_DIM_RE = re.compile(rf"\b(\d+)\s*[x*]\s*(\d+)\b(?![.,]\d)(?!\s*{_UNIDADES})", re.I)
# Cantidad suelta 'x 12', 'x 1', 'CJ X 12': es el empaque, no la variedad. Antes
# quedaba el '12' como número obligatorio y 'Leche entera x 12' no matcheaba nada.
_PACK_RE = re.compile(r"\bx\s*\d+\b", re.I)


def strip_accents(text: Any) -> str:
    text = "" if text is None or (isinstance(text, float) and math.isnan(text)) else str(text)
    return unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode()


def normalize_text(value: Any) -> str:
    text = strip_accents(value).lower()
    text = re.sub(r"(\d+(?:[.,]\d+)?)\s*(kg|kgs|kilo|kilos|gr|grs|g|lt|lts|l|litro|litros|ml|cc)\b", r"\1 \2", text)
    for pattern, repl in ABBREVS.items():
        text = re.sub(pattern, repl, text)
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return " ".join(t for t in text.split() if t not in STOP_WORDS)


def match_key(value: Any) -> str:
    """Texto normalizado SIN tamaños de envase: es lo que se compara."""
    text = re.sub(r"\bof\s*x\s*\d+\b", " ", strip_accents(value), flags=re.I)  # 'OFx15' de Emporio
    text = _SIZE_RE.sub(" ", _DIM_RE.sub(r" \1x\2 ", text))
    return normalize_text(_PACK_RE.sub(" ", text))


def _stem(token: str) -> str:
    return token[:-2] if token.endswith("es") and len(token) > 5 else token.rstrip("s")


IGNORE_IN_SCORING = {"queso", "de", "en", "el", "la", "con"}

# Pimienta: los colores son productos distintos y mutuamente excluyentes. Si el insumo
# dice un color, el candidato tiene que decir exactamente ese (stems, ya sin plural).
PIMIENTA_COLORES = {"negra", "blanca", "verde", "rosa", "rosada", "roja"}
_DIM_TOKEN = re.compile(r"\d+x\d+")


def _medida_parecida(a: str, b: str) -> bool:
    """'33x33' vs '33x32': cada lado a menos de 10% de diferencia."""
    return all(abs(int(x) - int(y)) <= 0.1 * max(int(x), int(y)) for x, y in zip(a.split("x"), b.split("x")))


def _es_termino(token: str) -> bool:
    return len(token) >= 3 and not token.isdigit() and not _DIM_TOKEN.fullmatch(token) \
        and token not in IGNORE_IN_SCORING


def _terminos(key: str) -> set[str]:
    return {_stem(t) for t in key.split() if t not in IGNORE_IN_SCORING}


def _consulta(producto: Any) -> str:
    """match_key del Dossier sin el nombre del proveedor ('Garbanzos Villares' -> 'garbanzos')."""
    return " ".join(t for t in match_key(producto).split() if t not in PROVEEDOR_WORDS)


_BRAND_STEMS = {_stem(b) for b in BRANDS}


def leer_productos_txt(path: Path) -> list[dict]:
    """Líneas de 'productos a buscar.txt' con sus marcas detectadas. [] si no existe."""
    if not path.is_file():
        LOG.warning("No encuentro %s: no se exigen marcas", path)
        return []
    try:
        texto = path.read_text(encoding="utf-8-sig")
    except UnicodeDecodeError:
        texto = path.read_text(encoding="cp1252")
    lineas = []
    for linea in texto.splitlines():
        if not linea.strip():
            continue
        todos = _terminos(match_key(linea))
        marcas = todos & _BRAND_STEMS
        clave = _sin_simbolos(linea)
        if not marcas and clave in _PRIORITARIAS:
            # línea sin marca con marcas prioritarias: una línea por marca (basta cumplir una)
            for marca in MARCAS_PRIORITARIAS[_PRIORITARIAS[clave]]:
                con_marca = f"{linea.strip()} {marca}"
                t = _terminos(match_key(con_marca))
                lineas.append({"linea": con_marca, "key": match_key(con_marca), "todos": t,
                               "marcas": t & _BRAND_STEMS, "generico": t - _BRAND_STEMS,
                               "size": parse_presentacion(linea)})
            continue
        lineas.append({"linea": linea.strip(), "key": match_key(linea), "todos": todos, "marcas": marcas,
                       "generico": todos - marcas, "size": parse_presentacion(linea)})
    return lineas


def marcas_requeridas(producto: Any, lineas: list[dict]) -> tuple[list[frozenset[str]], str | None]:
    """Marcas que debe traer un candidato para el producto del Dossier.

    Devuelve (alternativas, consulta). Alternativas: basta con cumplir una; []
    = sin exigencia de marca. `consulta` solo viene cuando la línea del txt es
    una marca pura ('NESCAFE DOLCA'): el producto se busca por ella y no por el
    nombre del Dossier ('cafe Nescafe personal'), que las listas no usan.
    El producto del Dossier corresponde a una línea del txt si su nombre (sin
    marca) coincide con lo genérico de la línea y está contenido en ella:
    'Tahina (Frasco X 900Grs)' <-> 'Tahina kanater' => exige 'kanater'.
    Si hay varias líneas ('Manteca 5 kg' <-> Milkaut x 5 kg / La Paulina x 5 kg)
    alcanza con cumplir la marca de alguna. Si el Dossier ya nombra la marca, esa
    marca se exige aunque el producto no esté en el txt.
    """
    d = _terminos(_consulta(producto))
    if not d:
        return [], None
    del_dossier = frozenset(d & _BRAND_STEMS)
    elegibles = [l for l in lineas if l["generico"] <= d <= l["todos"]]
    if not elegibles:
        # Línea que es solo marca ('NESCAFE DOLCA') y el Dossier nombra parte de ella
        puras = [l for l in lineas if not l["generico"] and l["marcas"] & d]
        if len(puras) == 1:
            return [frozenset(puras[0]["marcas"]) | del_dossier], puras[0]["key"]
        return ([del_dossier] if del_dossier else []), None
    # Línea cuya marca el Dossier ya nombra (o línea sin marca) tiene prioridad
    propias = [l for l in elegibles if l["marcas"] <= d]
    elegibles = propias or elegibles
    cant, base = parse_presentacion(str(producto))
    if base != "UN":
        mismo_tam = [l for l in elegibles if l["size"] == (cant, base)]
        elegibles = mismo_tam or elegibles
    opciones = {frozenset(l["marcas"]) | del_dossier for l in elegibles}
    return ([] if frozenset() in opciones else sorted(opciones, key=sorted)), None


def _cumple_marcas(opciones: list[frozenset[str]], cand_stems: set[str]) -> bool:
    def tiene(marca: str) -> bool:
        return bool(({marca} | {_stem(a) for a in BRAND_ALIASES.get(marca, ())}) & cand_stems)
    return not opciones or any(all(tiene(m) for m in op) for op in opciones)


def candidato_valido(query_tokens: list[str], cand_tokens_ordered: list[str],
                     marcas: list[frozenset[str]] | None = None, solo_marca: bool = False) -> bool:
    """Filtros duros heredados del pipeline anterior: si falla, no hay match posible.
    `marcas` (ver marcas_requeridas) es un filtro obligatorio extra; con `solo_marca`
    la consulta ES la marca y no se exige el sustantivo principal ('CAFE INS NESCAFE
    DOLCA' / 'NESCAFE DOLCA SUAVE' son el mismo producto)."""
    qset, cand_tokens = set(query_tokens), set(cand_tokens_ordered)
    if "pimienta" in qset:
        color_q = {_stem(t) for t in qset} & PIMIENTA_COLORES
        if color_q and {_stem(t) for t in cand_tokens} & PIMIENTA_COLORES != color_q:
            return False
    if marcas and not _cumple_marcas(marcas, {_stem(t) for t in cand_tokens}):
        return False

    for key, forbidden in FORBIDDEN_MATCHES.items():
        if key in qset and any(fb in cand_tokens and fb not in qset for fb in forbidden):
            return False

    required = qset & CRITICAL_QUALIFIERS
    if "semi" in qset or "semi" in cand_tokens:
        # 'Sardo estacionado' acepta como alternativa el 'Sardo semi (estacionado)'
        required.discard("estacionado")
    if "grano" in required and qset & SEMILLAS_ENTERAS and not cand_tokens & {"molido", "molida", "harina", "pasta"}:
        # 'Semillas de sésamo blanco' ('grano sesamo blanco'): las listas dicen 'SESAMO BLANCO LOS VALLES X 1 KG'
        required.discard("grano")
    if not required.issubset(cand_tokens):
        return False

    # Números que identifican variedad (harina 000 vs 0000, chocolate 87, crema 44)
    if not {t for t in qset if t.isdigit()}.issubset(cand_tokens):
        return False
    # Medidas: bolsa 30x40 != 40x60, pero servilleta 33x33 ~ 33x32
    cand_dims = [t for t in cand_tokens if _DIM_TOKEN.fullmatch(t)]
    if not all(any(_medida_parecida(q, c) for c in cand_dims) for q in qset if _DIM_TOKEN.fullmatch(q)):
        return False

    cand_stems = {_stem(t) for t in cand_tokens}
    if "semi" in cand_tokens:
        cand_stems.add(_stem("estacionado"))
    key_terms = [_stem(t) for t in query_tokens if _es_termino(t)]
    # 'semillas de X' se normaliza a 'grano X' (para 'comino en grano'); si la lista no lo dice ('SESAMO BLANCO
    # LOS VALLES'), 'grano' no puede ser el sustantivo principal: se ignora de los dos lados.
    _g = _stem("grano")
    if len(set(key_terms)) > 1 and _g in key_terms:
        key_terms = [t for t in key_terms if t != _g]
    # 'queso' se ignora de los dos lados: si no, 'Queso fontina' no podía matchear
    # con 'QUESO FONTINA PAULINA' (la primera palabra de la lista era 'queso').
    cand_terms = [_stem(t) for t in cand_tokens_ordered if _es_termino(t)]
    if len(set(cand_terms)) > 1 and _g in cand_terms and _g not in key_terms:
        cand_terms = [t for t in cand_terms if t != _g]
    if not key_terms or not cand_terms:
        return False
    # Sustantivo principal: la primera palabra de cada lado tiene que estar entre las
    # dos primeras del otro. Evita 'Hojaldre' -> 'Margarina Hojaldre' o
    # 'Chocolate blanco' -> 'Micro cereal chocolate blanco', pero acepta
    # 'Semillas de comino' -> 'Comino en grano'.
    if marcas:
        # La marca ya identifica el producto: 'PASTA DE SESAMO TAHINA AL KANATER' es 'Tahina kanater'
        if not solo_marca and key_terms[0] not in cand_terms:
            return False
    elif key_terms[0] not in cand_terms[:2] or cand_terms[0] not in key_terms[:2]:
        return False
    overlap = sum(1 for t in set(key_terms) if t in cand_stems)
    return overlap >= min(len(set(key_terms)), 2)


def score_match(query: str, cand: str) -> float:
    """token_set premia que la lista contenga todas nuestras palabras;
    token_sort castiga las palabras de más (otro producto con nombre parecido)."""
    return 0.6 * fuzz.token_set_ratio(query, cand) + 0.4 * fuzz.token_sort_ratio(query, cand)


# =========================================================================== #
# Precios y presentaciones
# =========================================================================== #
def parse_price(val: Any) -> float | None:
    if val is None or (isinstance(val, float) and math.isnan(val)):
        return None
    if isinstance(val, (int, float)):
        return float(val) if val > 0 else None
    s = str(val).strip().replace("$", "").replace(" ", "")
    m = re.search(r"[\d.,]+", s)
    if not m:
        return None
    raw = m.group(0)
    if "," in raw and "." in raw:
        raw = raw.replace(",", "") if raw.rfind(".") > raw.rfind(",") else raw.replace(".", "").replace(",", ".")
    elif "," in raw:
        raw = raw.replace(",", ".") if len(raw.split(",")[-1]) <= 2 else raw.replace(",", "")
    elif "." in raw:
        parts = raw.split(".")
        if len(parts[-1]) == 3:
            raw = raw.replace(".", "")
    try:
        res = float(raw)
    except ValueError:
        return None
    return res if res > 0 else None


_QTY_RE = re.compile(r"(?:(\d+)\s*x\s*)?(\d+(?:[.,]\d+)?)\s*(kgs?|k|kilos?|grs?|gramos?|g|lts?|litros?|l|ml|cc)\b", re.I)
# Caja sin peso/volumen: 'LECHE ... CJ X 12' -> el precio cubre 12 unidades
_CAJA_UN_RE = re.compile(r"\b(?:cj|caja|pack)\s*x\s*(\d+)\b", re.I)


def parse_presentacion(text: str, unidad_lista: str = "", precio_por_caja: bool = False) -> tuple[float, str]:
    """Devuelve (cantidad en unidad base, unidad base) que cubre el precio.

    - unidad_lista KG/LT  -> el precio ya es por kilo/litro.
    - '5 LT' en el nombre  -> el precio cubre 5 LT.
    - precio_por_caja y '12 X 1,5 LT' -> el precio cubre 18 LT (cajas de Alyser/Emporio).
    """
    u = strip_accents(unidad_lista).strip().upper().rstrip(".")
    if u in {"KG", "X KG", "KGS", "KILO", "KILOS"}:
        return 1.0, "KG"
    if u in {"LT", "L", "LTS", "LITRO", "LITROS"}:
        return 1.0, "LT"

    m = _QTY_RE.search(strip_accents(text))
    if not m:
        caja = _CAJA_UN_RE.search(strip_accents(text))
        return (float(caja.group(1)), "UN") if caja and int(caja.group(1)) > 1 else (1.0, "UN")
    amount = float(m.group(2).replace(",", "."))
    unit = m.group(3).lower()
    if unit.startswith(("kg", "kilo")) or unit == "k":
        base = "KG"
    elif unit.startswith(("gr", "gramo")) or unit == "g":
        amount, base = amount / 1000, "KG"
    elif unit in {"ml", "cc"}:
        amount, base = amount / 1000, "LT"
    else:
        base = "LT"
    if precio_por_caja and m.group(1):
        amount *= int(m.group(1))
    return (amount if amount > 0 else 1.0), base


_POR_UNIDAD_BASE = {"KG", "X KG", "KGS", "KILO", "KILOS", "LT", "L", "LTS", "LITRO", "LITROS"}


def _envase(texto: str, unidad_lista: str) -> float | None:
    """Tamaño de UN envase en unidad base (900 g -> 0.9), sin multiplicar por la
    caja. None si se cotiza por kilo/litro a granel o el nombre no trae tamaño."""
    if strip_accents(unidad_lista).strip().upper().rstrip(".") in _POR_UNIDAD_BASE:
        return None
    texto = strip_accents(texto)
    return parse_presentacion(texto)[0] if _QTY_RE.search(texto) else None


def _item(producto: str, precio: float, presentacion: str = "", unidad_lista: str = "",
          precio_por_caja: bool = False, factor: float | None = None, base: str | None = None,
          corto: str | None = None) -> dict:
    """`corto`: nombre sin 'extras' (origen, marca del productor...) que es lo que
    se puntúa; el nombre completo sigue sirviendo para los filtros duros y marcas."""
    if factor is None or base is None:
        factor, base = parse_presentacion(f"{producto} {presentacion}", unidad_lista, precio_por_caja)
    return {
        "producto": re.sub(r"\s+", " ", f"{producto} {presentacion}").strip(),
        "corto": re.sub(r"\s+", " ", corto if corto is not None else f"{producto} {presentacion}").strip(),
        "envase": _envase(f"{producto} {presentacion}", unidad_lista),
        "unidad_lista": unidad_lista,
        "precio_lista": round(precio, 2),
        "cantidad_base": factor,
        "unidad_base": base,
        "precio_base": round(precio / factor, 2),
    }


def _txt(v: Any) -> str:
    return "" if v is None or (isinstance(v, float) and math.isnan(v)) else str(v).strip()


# =========================================================================== #
# Lectores por proveedor
# =========================================================================== #
def leer_posta(path: Path) -> list[dict]:
    raw = pd.read_excel(path, header=None)
    hdr = next(i for i in range(min(15, len(raw)))
               if any("precio actual" in _txt(c).lower() for c in raw.iloc[i]))
    df = pd.read_excel(path, header=hdr)
    prod_col = next(c for c in df.columns if "nombre" in str(c).lower() or "descrip" in str(c).lower())
    price_col = next(c for c in df.columns if "precio actual" in str(c).lower())
    um_col = next((c for c in df.columns if str(c).strip().lower() in {"u.m.", "um", "unidad"}), None)

    items = []
    for _, r in df.iterrows():
        nombre, precio = _txt(r[prod_col]), parse_price(r[price_col])
        if not nombre or not precio:
            continue
        um = _txt(r[um_col]).upper() if um_col else "UN"
        if um == "MIL":  # precio por mil unidades (descartables)
            items.append(_item(nombre, precio, unidad_lista=um, factor=1000.0, base="UN"))
        else:
            items.append(_item(nombre, precio, unidad_lista=um, precio_por_caja=(um == "CAJ")))
    return items


def leer_emporio(path: Path) -> list[dict]:
    raw = pd.read_excel(path, header=None)
    items, en_tabla = [], False
    for _, r in raw.iterrows():
        c = [_txt(r[i]) if i < len(r) else "" for i in range(5)]
        if c[1].upper() == "DESCRIPCION":           # encabezado de sección (OFERTAS, ALMACEN, QUESOS...)
            en_tabla = c[0].upper() != "OFERTAS"
            continue
        if len(c[0]) < 3:
            continue
        if en_tabla:                                 # nombre | desc | unidad | PRECIO | C/DTO
            lista, dto, unidad = parse_price(c[3]), parse_price(c[4]), c[2].upper()
            # Se compara el precio con descuento (C/DTO = lista - 15%), nunca el de lista
            esperado = lista * DESCUENTO_EMPORIO if lista else None
            if dto and esperado and abs(dto - esperado) / esperado > 0.01:
                LOG.warning("Emporio: '%s' C/DTO %.2f no es lista x %.2f (%.2f); uso lista x %.2f",
                            c[0], dto, DESCUENTO_EMPORIO, esperado, DESCUENTO_EMPORIO)
                dto = None
            precio = dto or esperado
        else:                                        # ofertas: nombre | desc | PRECIO | nota (precio neto, sin C/DTO)
            precio = parse_price(c[2])
            unidad = "KG" if re.search(r"\bx\s*kg\b", c[1], re.I) else "UNI"
        if precio:
            items.append(_item(c[0], precio, c[1], unidad, precio_por_caja=unidad.startswith("CAJ")))
    return items


# Encabezado de la columna con 18% de descuento ('$550.000 (18%)'); no confundir
# con '$750.000 (18%+5%)'.
_VILLARES_COL_18 = re.compile(r"\(\s*18\s*%\s*\)")


def leer_villares(path: Path) -> list[dict]:
    """El precio que se compara es el de la columna '(18%)', no el PRECIO DE LISTA."""
    items = []
    for sheet, raw in pd.read_excel(path, sheet_name=None, header=None).items():
        hdr = next((i for i in range(min(10, len(raw)))
                    if any("precio de lista" in _txt(c).lower() for c in raw.iloc[i])), None)
        if hdr is None:
            LOG.debug("Villares: hoja '%s' sin columna 'PRECIO DE LISTA', se omite", sheet)
            continue
        encabezado = [_txt(c) for c in raw.iloc[hdr]]
        # PRECIO DE LISTA solo marca dónde terminan las columnas de descripción/unidad
        pcol = next(j for j, c in enumerate(encabezado) if "precio de lista" in c.lower())
        dcol = next((j for j, c in enumerate(encabezado) if _VILLARES_COL_18.search(c)), None)
        if dcol is None:
            LOG.warning("Villares: hoja '%s' sin columna de 18%% de descuento, se omite", sheet)
            continue
        for _, r in raw.iloc[hdr + 1:].iterrows():
            nombre, presentacion = _txt(r[2]), _txt(r[3])
            precio = parse_price(r[dcol])
            if not nombre or not precio:
                continue
            extras = " ".join(_txt(r[j]) for j in range(4, pcol))
            # Columna de unidad 'Kg.' => precio por kilo. Si no, el precio es por envase
            # individual ('20 x 100 gr.' cotiza el sobre de 100 gr, no el cartón).
            por_kilo = any(_txt(r[j]).lower().rstrip(".") == "kg" for j in range(4, pcol))
            items.append(_item(f"{nombre} {extras}", precio, presentacion, "KG" if por_kilo else "UN",
                               corto=f"{nombre} {presentacion}"))
    return items


_ALYSER_LINE = re.compile(r"^(\d{6,})\s+(.+?)\s+([A-Z]{2,4})\s+\$\s*([\d.,]+)\s*$")


def leer_alyser(path: Path) -> list[dict]:
    items = []
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            for line in (page.extract_text() or "").splitlines():
                m = _ALYSER_LINE.match(line.strip())
                if not m:
                    continue
                _, desc, un, precio = m.groups()
                precio = parse_price(precio)
                if precio:
                    items.append(_item(desc, precio, unidad_lista=un, precio_por_caja=(un == "CAJ")))
    return items


def leer_raices(path: Path) -> list[dict]:
    """Tabla de 2 bloques (izq/der): ARTICULO | CANTIDAD | Bulto/Paq | Kilo.
    Base: Verduleria/extraer_precios.py."""
    items = []
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            for tabla in page.extract_tables({"vertical_strategy": "lines", "horizontal_strategy": "lines",
                                               "intersection_tolerance": 5}):
                for fila in tabla:
                    celdas = [_txt(c) for c in fila] + [""] * 9
                    for art, cant, bulto, kilo in (celdas[0:4], celdas[5:9]):
                        if not art or art.upper() in {"ARTICULO", "HEAVY"}:
                            continue
                        p_bulto, p_kilo = parse_price(bulto), parse_price(kilo)
                        if p_kilo:
                            factor = p_bulto / p_kilo if p_bulto else 1.0
                            items.append(_item(art, p_bulto or p_kilo, cant, cant, factor=factor, base="KG"))
                        elif p_bulto:
                            items.append(_item(art, p_bulto, cant, cant, factor=1.0, base="UN"))
    return items


def leer_criollo(path: Path) -> list[dict]:
    """CSV que genera el scraper de Precios/run_pipeline_pricing.py; ya trae la
    presentación resuelta (factor_presentacion / unidad_base)."""
    df = pd.read_csv(path, sep=";", encoding="utf-8-sig")
    items = []
    for _, r in df.iterrows():
        nombre, precio = _txt(r["descripcion"]), parse_price(r["precio_nominal"])
        factor = parse_price(r["factor_presentacion"]) or 1.0
        if nombre and precio:
            items.append(_item(nombre, precio, factor=factor, base=_txt(r["unidad_base"]) or "UN"))
    return items


_NOGADA_MARCA = re.compile(r"^\s*(?:agroalimentos\s+)?(areco|reino de le[oó]n)\s*[–—-]*\s*", re.I)
_NOGADA_CAJA = re.compile(r"\(?\s*caja\s*x\s*(\d+)\s*u?\.?\s*\)?", re.I)


def scrapear_nogada(destino: Path) -> int:
    """Baja el catálogo de nogada.com.ar (API pública de WooCommerce, sin login) a
    un CSV. Solo escribe si la descarga completa salió bien. Devuelve la cantidad."""
    filas, pagina = [], 1
    while True:
        r = requests.get(NOGADA_API, params={"per_page": 100, "page": pagina},
                         headers={"User-Agent": "Mozilla/5.0"}, timeout=30)
        r.raise_for_status()
        for p in r.json():
            precios = p.get("prices") or {}
            try:
                precio = int(precios["price"]) / 10 ** int(precios.get("currency_minor_unit") or 0)
            except (KeyError, TypeError, ValueError):
                continue
            if precio > 0 and p.get("is_in_stock", True):
                filas.append({"nombre": html.unescape(p["name"]).strip(), "precio": precio,
                              "url": p.get("permalink", "")})
        if pagina >= int(r.headers.get("X-WP-TotalPages", 1)):
            break
        pagina += 1
    if not filas:
        raise RuntimeError("Nogada devolvió 0 productos")
    pd.DataFrame(filas).to_csv(destino, sep=";", index=False, encoding="utf-8-sig")
    return len(filas)


def leer_nogada(path: Path) -> list[dict]:
    """'Areco – Girasol Alto Oleico 5000 ml (caja x 4u.)' -> 'Aceite Girasol Alto Oleico
    5000 ml Areco (caja x4)': el nombre no trae 'Aceite' ni la marca al final, y el
    precio de las cajas cubre N envases (factor = N x tamaño)."""
    df = pd.read_csv(path, sep=";", encoding="utf-8-sig")
    items = []
    for _, r in df.iterrows():
        nombre, precio = _txt(r["nombre"]), parse_price(r["precio"])
        if not nombre or not precio:
            continue
        m = _NOGADA_MARCA.match(nombre)
        marca, resto = (m.group(1), nombre[m.end():]) if m else ("", nombre)
        caja = _NOGADA_CAJA.search(resto)
        envases = int(caja.group(1)) if caja else 1
        resto = re.sub(r"\s*[–—-]\s*", " ", _NOGADA_CAJA.sub(" ", resto)).strip()
        if not re.search(r"\baceite\b", strip_accents(resto), re.I):
            resto = f"Aceite {resto}"
        cantidad, base = parse_presentacion(resto)
        items.append(_item(f"{resto} {marca}", precio, f"(caja x{envases})" if caja else "",
                           factor=cantidad * envases, base=base))
    return items


# '<producto + presentación> <kg|unidad|lts> <precio>'
_LAMESON_LINEA = re.compile(
    r"^(?:>\s*)?(?P<desc>.+?)\s+(?P<un>kg\.?|unidad|unid|uni|lts?)\s+(?P<precio>\d+(?:\.\d+)?)\s*$", re.I)


def leer_lameson(path: Path) -> list[dict]:
    """PDF de La Mesón (Hoja 1..5: quesos, lácteos, fiambres, Arytza/Oliva Ilustre...).
    Cada línea es '<producto + presentación> <kg|unidad|lts> <precio>': 'kg'/'lts' =
    precio por kilo/litro; 'unidad' = precio del envase que dice el nombre
    ('frasco x 360grs'). Precio 0.00 = sin stock, se omite."""
    items, sin_leer = [], 0
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            for linea in (page.extract_text() or "").splitlines():
                m = _LAMESON_LINEA.match(linea.strip())
                if not m:
                    # renglón con precio que no se pudo leer (texto superpuesto en el PDF)
                    sin_leer += bool(re.search(r"\d+\.\d{2}\s*$", linea))
                    continue
                precio = parse_price(m["precio"])
                if precio:
                    un = m["un"].upper().rstrip(".")
                    items.append(_item(m["desc"], precio, unidad_lista={"LTS": "LT", "LT": "LT", "KG": "KG"}.get(un, "UN")))
    if sin_leer:
        LOG.warning("La Mesón: %d renglones con precio no se pudieron leer (texto superpuesto en el PDF)", sin_leer)
    return items


_OPENFRUIT_ENVASE = re.compile(
    r"\b(paq|paquete|bja|bandeja|band|caja|cajon|bolsa|ristra|sobre|atado|docena|maple)\b", re.I)


def leer_open_fruit(path: Path) -> list[dict]:
    """Planilla 'rubro | Descripción | Precio neto'. Sin aclaración el precio es por
    kilo ('KIWI', 'ZUCCHINI', 'X KG'); si el nombre dice PAQ/BJA/CAJON/BOLSA/... el
    precio es por ese envase (unidad)."""
    raw = pd.read_excel(path, header=None)
    hdr = next((i for i in range(min(10, len(raw)))
                if any("precio" in _txt(c).lower() for c in raw.iloc[i])), None)
    if hdr is None:
        raise ValueError("Open Fruit: no encuentro la fila de encabezado ('Precio neto')")
    items = []
    for _, r in raw.iloc[hdr + 1:].iterrows():
        nombre, precio = _txt(r[1]), parse_price(r[2])
        if not nombre or not precio:
            continue
        por_envase = bool(_OPENFRUIT_ENVASE.search(nombre))
        items.append(_item(nombre, precio, unidad_lista="UN" if por_envase else "KG"))
    return items


# palabra clave en el nombre del archivo -> (nombre proveedor, lector, rubros que cubre)
# rubros None = todo excepto Frutas y Verduras. El orden es el de las columnas del reporte.
PROVEEDORES: dict[str, tuple[str, Callable[[Path], list[dict]], set[str] | None]] = {
    "criollo": ("El Criollo", leer_criollo, None),
    "posta": ("Posta", leer_posta, None),
    "emporio": ("El Nuevo Emporio", leer_emporio, None),
    "alyser": ("Alyser", leer_alyser, None),
    "villares": ("Villares", leer_villares, None),
    "nogada": ("Nogada", leer_nogada, None),
    "lameson": ("La Mesón", leer_lameson, {"almacen", "lacteos", "quesos"}),
    "raices": ("Raices", leer_raices, {"frutas y verduras"}),
    "openfruit": ("Open Fruit", leer_open_fruit, {"frutas y verduras"}),
}


# =========================================================================== #
# Google Drive
# =========================================================================== #
def descargar_listas(dest: Path, folder_id: str, creds_path: Path) -> list[Path]:
    creds = Credentials.from_service_account_file(str(creds_path), scopes=DRIVE_SCOPES)
    session = AuthorizedSession(creds)
    resp = session.get(
        "https://www.googleapis.com/drive/v3/files",
        params={"q": f"'{folder_id}' in parents and trashed=false", "pageSize": 200,
                "fields": "files(id,name,mimeType,modifiedTime)",
                "supportsAllDrives": True, "includeItemsFromAllDrives": True},
    )
    resp.raise_for_status()

    paths = []
    for f in resp.json().get("files", []):
        mime, name = f["mimeType"], f["name"]
        if mime == "application/vnd.google-apps.spreadsheet":
            url, params, fname = f"https://www.googleapis.com/drive/v3/files/{f['id']}/export", {"mimeType": XLSX_MIME}, f"{name}.xlsx"
        elif mime.startswith("application/vnd.google-apps"):
            LOG.warning("Drive: '%s' es un %s, no se procesa", name, mime)
            continue
        else:
            url, params, fname = f"https://www.googleapis.com/drive/v3/files/{f['id']}", {"alt": "media"}, name
        r = session.get(url, params=params)
        r.raise_for_status()
        out = dest / re.sub(r'[\\/:*?"<>|]', "_", fname)
        out.write_bytes(r.content)
        LOG.info("Drive: descargado '%s' (modif. %s)", fname, f["modifiedTime"][:10])
        paths.append(out)
    return paths


def cargar_listas(paths: list[Path]) -> pd.DataFrame:
    frames = []
    for path in paths:
        key = next((k for k in PROVEEDORES if k in _sin_simbolos(path.name)), None)
        if key is None:
            LOG.warning("Sin lector para '%s' (agregalo en PROVEEDORES)", path.name)
            continue
        nombre, lector, rubros = PROVEEDORES[key]
        try:
            items = lector(path)
        except Exception:
            LOG.exception("Error leyendo %s", path.name)
            continue
        df = pd.DataFrame(items)
        if df.empty:
            LOG.warning("%s: 0 productos leídos", nombre)
            continue
        df["proveedor"] = nombre
        df["rubros"] = [rubros] * len(df)
        df["_key"] = df["producto"].map(match_key)
        df["_tokens"] = df["_key"].str.split()
        df["_skey"] = df["corto"].map(match_key)   # lo que se puntúa (sin origen/marca del productor)
        df = df[df["_key"].ne("")].drop_duplicates(["_key", "precio_lista"])
        LOG.info("%s: %d productos", nombre, len(df))
        frames.append(df)
    if not frames:
        raise RuntimeError("No se pudo leer ninguna lista de proveedor.")
    orden = [n for n, _, _ in PROVEEDORES.values()]
    frames.sort(key=lambda f: orden.index(f["proveedor"].iat[0]))
    return pd.concat(frames, ignore_index=True)


# =========================================================================== #
# Comparación
# =========================================================================== #
def _norm_cat(cat: Any) -> str:
    return strip_accents(cat).strip().lower()


def _sin_simbolos(texto: Any) -> str:
    """'La Mesón.pdf' -> 'lamesonpdf': para reconocer al proveedor por el nombre del archivo."""
    return re.sub(r"[^a-z0-9]+", "", strip_accents(texto).lower())


def aplicar_exclusiones_dossier(dossier: pd.DataFrame) -> pd.DataFrame:
    """Saca las filas de DOSSIER_EXCLUIR (errores de origen: bulto/caja cargado como unidad)."""
    fuera = pd.Series(False, index=dossier.index)
    for r in DOSSIER_EXCLUIR:
        m = (dossier["Código SKU"].astype(str).str.strip().eq(r["sku"])
             & pd.to_numeric(dossier["Precio Ponderado"], errors="coerce").lt(r["precio_max"]))
        for _, f in dossier[m].iterrows():
            LOG.info("Fila del Dossier excluida: %s '%s' $%s x %s -> %s", f["Código SKU"], f["Productos"],
                     f"{f['Precio Ponderado']:,.2f}", f["Cantidad Inventario"], r["motivo"])
        fuera |= m
    return dossier[~fuera].reset_index(drop=True)


def aplicar_lista_negra(dossier: pd.DataFrame) -> pd.DataFrame:
    negra = {_sin_simbolos(n) for n in LISTA_NEGRA}
    fuera = dossier["Productos"].map(_sin_simbolos).isin(negra)
    if fuera.any():
        LOG.info("Lista negra: %d filas excluidas (%s)", int(fuera.sum()),
                 ", ".join(sorted(set(dossier.loc[fuera, "Productos"].astype(str).str.strip()))))
    return dossier[~fuera].reset_index(drop=True)


def _crema_gastronomica(c: pd.Series) -> bool:
    """Crema de leche: solo balde/sachet gastronómico de 5 L al 44% (no potes chicos de 520 cc)."""
    texto = strip_accents(c["producto"])
    m = _QTY_RE.search(texto)
    if "44" not in c["_tokens"] or not m:
        return False
    cantidad, base = parse_presentacion(texto)
    return base in {"LT", "KG"} and 4.5 <= cantidad <= 5.5


_PRIORITARIAS: dict[str, str] = {}   # línea normalizada -> clave de MARCAS_PRIORITARIAS
_PRIORITARIAS.update({_sin_simbolos(k): k for k in MARCAS_PRIORITARIAS})

# (palabras del insumo que activan la regla, condición que debe cumplir cada candidato)
RESTRICCIONES = [({"crema", "leche"}, _crema_gastronomica)]


_POR_KILO_RE = re.compile(r"\bx\s*(kg|kilo|lt|litro)s?\b", re.I)


def precio_alineado(precio_dossier: float, cand: pd.Series, producto_dossier: str) -> tuple[float, str] | None:
    """El Dossier no dice en qué unidad está 'Cantidad Inventario' (kg, litro,
    frasco...). Se arman las lecturas posibles del precio de lista y se queda la
    que, dentro de la tolerancia, más se parece al Precio Ponderado:
    - por unidad base ($/KG, $/LT);
    - si el nombre del Dossier trae tamaño ('Atun x 1.8kg'), por ese envase;
    - si no trae tamaño ni dice 'x KG', por el envase de la lista."""
    def ratio(precio: float) -> float:
        return abs(math.log(precio_dossier / precio))

    opciones = [(cand["precio_base"], f"$/{cand['unidad_base']}")]
    cant_d, base_d = parse_presentacion(str(producto_dossier))
    if (base_d != "UN" and ENVASE_FIJO_RE.search(str(producto_dossier)) and cant_d
            and cand["unidad_base"] == base_d and cand["cantidad_base"]
            and abs(cand["cantidad_base"] / cant_d - 1) > TOLERANCIA_TAMANO_ENVASE):
        return None          # otro tamaño de frasco/pote: no es el mismo producto
    if base_d != "UN":
        if cant_d != 1 and base_d == cand["unidad_base"]:
            opciones.append((cand["precio_base"] * cant_d, f"$/envase {cant_d:g} {base_d}"))
    elif cand["cantidad_base"] != 1 and not _POR_KILO_RE.search(str(producto_dossier)):
        opciones.append((cand["precio_lista"], "$/envase"))

    validas = [o for o in opciones if o[0] > 0 and ratio(o[0]) <= math.log(TOLERANCIA_UNIDAD)]
    return min(validas, key=lambda o: ratio(o[0])) if validas else None


def proveedor_aplica(rubros: set[str] | None, cat: str) -> bool:
    return cat != "frutas y verduras" if rubros is None else cat in rubros


SIN_MATCH = "-"            # el proveedor no tiene el producto
NO_COMPARABLE = "≠ unidad"  # lo tiene, pero en una presentación que no se puede llevar a la misma unidad


def _grupo_miel(c: pd.Series) -> tuple | None:
    """Miel: tres variantes fijas en vez de una por marca/tamaño -- baldes mayoristas
    (>= 4 kg), frasco de hasta 2 kg y orgánica. El tamaño sale del nombre (La Mesón
    cotiza 'balde x 10kgs' por kilo). Individuales/sachets y otros productos con miel quedan afuera."""
    texto = strip_accents(c["producto"])
    if not _QTY_RE.search(texto):
        return None
    cantidad, base = parse_presentacion(texto)
    if base != "KG":
        return None
    if cantidad >= 4:
        clase = "balde 4+ kg"
    elif 0.1 <= cantidad <= 2:
        clase = "frasco hasta 2 kg"
    else:
        return None
    organica = _stem("organico") in {_stem(t) for t in c["_tokens"]}
    return (frozenset(), frozenset({"organico"}) if organica else frozenset(), clase)


def _orden_miel(clave: tuple) -> int:
    return 2 if clave[1] else (0 if clave[2] == "balde 4+ kg" else 1)


def _nombre_miel(clave, n: int, total: int, dossier: Any) -> str:
    return f"Miel (Opción {n})" if total > 1 else "Miel"


def _grupo_sardo(c: pd.Series) -> tuple:
    """Sardo estacionado (cualquier marca prioritaria en una sola fila) vs semi (por marca)."""
    toks = {_stem(t) for t in c["_tokens"]}
    if _stem("semi") in toks:
        return (frozenset(toks & _BRAND_STEMS), frozenset({"semi"}), "")
    return (frozenset(), frozenset(), "")


def _orden_sardo(clave: tuple) -> int:
    return 1 if clave[1] else 0


def _nombre_sardo(clave, n: int, total: int, dossier: Any) -> str:
    if not clave or not clave[1]:
        return "Queso Sardo Estacionado"
    marcas = " ".join(_ORIGINAL.get(m, m).title() for m in sorted(clave[0]))
    return f"Queso Sardo Semi {marcas}".strip()


def _grupo_arroz(c: pd.Series) -> tuple | None:
    """Arroz de consumo interno: sin marca ni presentación; solo importa la variedad
    (largo fino / parboil). Basmati, integral, yamaní, para risotto, etc. quedan afuera."""
    toks = {_stem(t) for t in c["_tokens"]}
    if toks & {_stem(t) for t in ("basmati", "integral", "yamani", "carnaroli", "arborio", "risotto",
                                  "salvaje", "negro", "glutinoso", "jazmin", "sushi", "precocido",
                                  "instantaneo", "aromatico")}:
        return None
    if _stem("parboil") in toks or _stem("parboiled") in toks:
        return (frozenset(), frozenset({"parboil"}), "")
    if _stem("largo") in toks and _stem("fino") in toks:
        return (frozenset(), frozenset({"largo fino"}), "")
    return None


def _orden_arroz(clave: tuple) -> int:
    return 0 if "largo fino" in clave[1] else 1


def _nombre_arroz(clave, n: int, total: int, dossier: Any) -> str:
    return f"Arroz personal (Opción {n})" if total > 1 else "Arroz personal"


# Insumos con variantes armadas a mano: 'activa' recibe las palabras de la búsqueda;
# 'query' (opcional) reemplaza la búsqueda; 'grupo' devuelve la variante de cada
# candidato (None = descartarlo); 'orden' fija el orden de las opciones; 'nombre' el
# texto de la fila; 'score_set' puntúa solo por contener la palabra; 'dossier_por_fraccion'
# lleva el Dossier a $/KG cuando figura por 1/N del envase ('Miel (x 10kg)'); 'ventana_score'
# es la diferencia de score que cuenta como empate (ahí gana el más barato).
AGRUPADORES = [
    {"activa": lambda q: bool(q) and q[0] == "miel", "query": "miel", "grupo": _grupo_miel,
     "orden": _orden_miel, "nombre": _nombre_miel, "score_set": True, "dossier_por_fraccion": True,
     "ventana_score": 2},
    {"activa": lambda q: bool(q) and q[0] == "arroz" and "personal" in q, "query": "arroz",
     "grupo": _grupo_arroz, "orden": _orden_arroz, "nombre": _nombre_arroz, "score_set": True,
     "dossier_por_fraccion": False, "ventana_score": 10},   # sin marca: gana el más barato por $/KG
    {"activa": lambda q: "sardo" in q and ("estacionado" in q or "semi" in q), "query": None,
     "grupo": _grupo_sardo, "orden": _orden_sardo, "nombre": _nombre_sardo, "score_set": False,
     "dossier_por_fraccion": False, "ventana_score": 10},   # marcas prioritarias: decide el precio
]

_VARIANT_BRAND_STEMS = {_stem(t) for t in VARIANT_BRANDS}
_VARIANT_WORD_STEMS = {_stem(t) for t in CRITICAL_QUALIFIERS | VARIANT_WORDS}


_ORIGINAL: dict[str, str] = {}   # stem -> palabra como figura en las listas (para nombrar filas)


def _clave_variante(c: pd.Series, qstems: set[str]) -> tuple:
    """(marcas, variedades que el Dossier no pidió, tamaño de envase): dos
    candidatos con la misma clave son 'el mismo producto' para cotizarlo en una fila."""
    for t in c["_tokens"]:
        _ORIGINAL.setdefault(_stem(t), t)
    toks = {_stem(t) for t in c["_tokens"]}
    envase = None if pd.isna(c["envase"]) else (round(float(c["envase"]), 3), c["unidad_base"])
    marcas = frozenset(toks & _VARIANT_BRAND_STEMS)
    return (marcas, frozenset((toks & _VARIANT_WORD_STEMS) - qstems - marcas), envase)


def comparar(dossier: pd.DataFrame, listas: pd.DataFrame, lineas_txt: list[dict] | None = None) -> pd.DataFrame:
    """Matriz multi-proveedor: una columna de precio por proveedor (precio llevado
    a la unidad del Dossier, '-' sin match, '≠ unidad' si no es comparable) + su
    columna 'Match <proveedor>' con el producto de la lista, y la Mejor Opción
    (la de menor precio comparable). `lineas_txt` son las líneas de
    'productos a buscar.txt': su marca es obligatoria en el match.

    Si el insumo tiene varias variantes válidas (marca / variedad / tamaño de
    envase) se emiten hasta MAX_OPCIONES filas '<insumo> (Opción n)': cada
    proveedor cotiza esa variante o lleva '-'. Se muestran las variantes que más
    proveedores cotizan; la Opción 1 es la más cercana al Precio Ponderado actual."""
    proveedores = list(dict.fromkeys(listas["proveedor"]))
    por_prov = {p: listas[listas["proveedor"] == p] for p in proveedores}
    claves = {nombre: k for k, (nombre, _, _) in PROVEEDORES.items()}
    filas = []

    for _, d in dossier.iterrows():
        pond = d["Precio Ponderado"]
        marcas, consulta = marcas_requeridas(d["Productos"], lineas_txt or [])
        query = consulta or _consulta(d["Productos"])
        stems_q = {_stem(t) for t in query.split()}
        esp = next((e for e in PRODUCTOS_ESPECIALES if {_stem(t) for t in e["producto"]} <= stems_q), None)
        if esp:
            marcas = [frozenset({_stem(esp["marca"])})]            # solo la marca fijada
        # proveedor real de compra: se compara, pero contra sí mismo no genera ahorro
        prov_real = (esp["proveedor_real"] if esp and _sin_simbolos(esp["proveedor_real"])
                     in _sin_simbolos(d["Proveedor"]) else None)
        # Las reglas se evalúan sobre el nombre completo; para buscar, 'crema de leche'
        # es 'crema' (las listas dicen 'CREMA 44% MILKAUT BALDE', sin 'leche')
        reglas = [fn for palabras, fn in RESTRICCIONES if palabras <= {_stem(t) for t in query.split()}]
        query = re.sub(r"\bcrema leche\b", "crema", query)
        qtokens = query.split()
        cfg = next((a for a in AGRUPADORES if a["activa"](qtokens)), None)
        if cfg and cfg["query"]:
            # 'Miel Organica (x 10kg)' no se limita a orgánicas: se cotiza toda la miel y
            # la orgánica es una de las opciones
            query, qtokens = cfg["query"], cfg["query"].split()
        qstems = {_stem(t) for t in qtokens}
        cant_inv = d["Cantidad Inventario"]
        if cfg and cfg["dossier_por_fraccion"] and pd.notna(pond):
            # El Dossier guarda 'Miel Organica (x 10kg)' a $566 con 100 unidades (= $56.600 por un
            # balde de 10 kg): precio por 1/10 de balde. Se lleva a $/KG real ($5.660) y 10 kg.
            tam, unidad = parse_presentacion(str(d["Productos"]))
            if unidad == "KG" and tam >= 5:
                LOG.info("Dossier por 1/%g de envase: '%s' $%s x %s -> $%s/KG x %s kg", tam, d["Productos"],
                         f"{pond:,.0f}", f"{cant_inv:g}", f"{pond * tam:,.0f}", f"{cant_inv / tam:g}")
                pond, cant_inv = pond * tam, cant_inv / tam
        cat = _norm_cat(d["Categoría 2"])
        if marcas:
            LOG.info("Marca obligatoria: %-45s -> %s", str(d["Productos"])[:45],
                     " | ".join("+".join(sorted(m)) for m in marcas))
        base = {
            "Código SKU": d["Código SKU"], "Producto (Dossier)": d["Productos"], "Opción": 1,
            "Categoría 2": d["Categoría 2"], "Proveedor actual": d["Proveedor"],
            "Cantidad Inventario": cant_inv, "Monto Total": d["Monto Total"],
            "Precio Ponderado": pond,
        }

        # variante -> {proveedor: mejor candidato de ese proveedor para esa variante}
        ventana = cfg["ventana_score"] if cfg else 2
        variantes: dict[tuple, dict[str, dict]] = {}
        candidatos: list[tuple[tuple, str, dict]] = []
        if pd.notna(pond) and pond > 0 and qtokens:
            for prov, df in por_prov.items():
                if not proveedor_aplica(df["rubros"].iat[0], cat):
                    continue
                for _, c in df.iterrows():
                    if not candidato_valido(qtokens, c["_tokens"], marcas, solo_marca=bool(consulta)):
                        continue
                    if reglas and not all(fn(c) for fn in reglas):
                        continue
                    sc = (fuzz.token_set_ratio(query, c["_key"]) if cfg and cfg["score_set"]
                          else score_match(query, c["_skey"]))
                    # Con marca exigida, marca + sustantivo ya están garantizados: umbral más bajo
                    # ('SARDO SEMI PINT GAUCHITO' es corto y no dice 'estacionado')
                    if sc < (SCORE_MARCA if marcas else SCORE_MINIMO):
                        continue
                    grupo = cfg["grupo"](c) if cfg else None
                    if cfg and grupo is None:
                        continue
                    alineado = precio_alineado(pond, c, d["Productos"])
                    cand = {"prov": prov, "score": sc, "producto": c["producto"],
                            "precio": alineado[0] if alineado else None,
                            "base": alineado[1] if alineado else "unidad no comparable"}
                    candidatos.append((grupo or _clave_variante(c, qstems), prov, cand))

            def ofrecer(slot: dict, prov: str, cand: dict) -> None:
                best = slot.get(prov)
                # Mejor score; a igual score (±ventana), el más barato comparable
                if best is None or cand["score"] > best["score"] + ventana or (
                    abs(cand["score"] - best["score"]) <= ventana and cand["precio"] is not None
                    and (best["precio"] is None or cand["precio"] < best["precio"])
                ):
                    slot[prov] = cand

            # Primero lo envasado; después lo que se cotiza a granel (por kg/lt, sin tamaño
            # de envase), que es el mismo producto: completa las filas de igual marca/variedad
            # donde ese proveedor no tiene su propio envase. Si ya lo tiene en todas, el granel
            # es su propia variante (balde de 4,5 kg vs frasco de 360 g).
            for clave, prov, cand in candidatos:
                if clave[2] is not None:
                    ofrecer(variantes.setdefault(clave, {}), prov, cand)
            for clave, prov, cand in sorted((x for x in candidatos if x[0][2] is None),
                                            key=lambda x: -x[2]["score"]):
                libres = [k for k in variantes if k[:2] == clave[:2] and k[2] is not None
                          and prov not in variantes[k]]
                for k in libres:
                    ofrecer(variantes[k], prov, cand)
                if not libres:
                    ofrecer(variantes.setdefault(clave, {}), prov, cand)

        # Variantes con al menos un precio comparable, la más parecida al precio actual primero
        ranking = []
        tam_dossier, base_dossier = parse_presentacion(str(d["Productos"]))
        for clave, slot in variantes.items():
            comp = [m for m in slot.values() if m["precio"] is not None]
            # Una variedad distinta (Lechoso, con piel...) se muestra aunque su precio no sea
            # comparable con el del Dossier (sin ganador): que el usuario vea que existe.
            if comp or clave[1]:
                cerca = min(abs(math.log(pond / m["precio"])) for m in comp) if comp else 9.9
                # Envase de otra escala que el que nombra el Dossier (lata de 340 g vs '1kg'): al final
                env = clave[2]
                otra_escala = bool(env and not isinstance(env, str) and base_dossier != "UN"
                                   and env[1] == base_dossier and not 0.5 <= env[0] / tam_dossier <= 2)
                ranking.append((-len(slot), int(otra_escala), 0 if clave[0] or clave[1] else 1,
                                cfg["orden"](clave) if cfg else round(cerca / CERCANIA_BUCKET),
                                -max(m["score"] for m in slot.values()),
                                cerca, len(ranking), clave))
        # Se quedan las variantes que más proveedores cotizan y, a igual cobertura,
        # las que difieren por marca/variedad antes que las que solo cambian de tamaño
        # ('Sauco' vs 'Lechoso' antes que un sobre de 500 g). Opción 1 = la más parecida
        # al precio que se paga hoy (de ahí sale el ahorro del total).
        # Un tamaño más de una marca/variedad ya elegida espera a que entren las demás.
        ordenadas = sorted(ranking)
        elegidas, grupos = [], set()
        for r in ordenadas:
            if r[-1][:2] not in grupos and len(elegidas) < MAX_OPCIONES:
                elegidas.append(r)
                grupos.add(r[-1][:2])
        elegidas += [r for r in ordenadas if r not in elegidas][:MAX_OPCIONES - len(elegidas)]
        elegidas.sort(key=lambda r: (r[3], r[0], r[4], r[5], r[6]))
        opciones: list[tuple[tuple | None, dict[str, dict]]] = [(r[-1], variantes[r[-1]]) for r in elegidas]
        if not any(m["precio"] is not None for _, slot in opciones for m in slot.values()) and variantes:
            # nada comparable: una sola fila 'UNIDAD NO COMPARABLE'
            mejor_por_prov: dict[str, dict] = {}
            for slot in variantes.values():
                for prov, m in slot.items():
                    if prov not in mejor_por_prov or m["score"] > mejor_por_prov[prov]["score"]:
                        mejor_por_prov[prov] = m
            opciones = [(None, mejor_por_prov)]
        if not opciones:
            opciones = [(None, {})]

        for n, (clave, slot) in enumerate(opciones, start=1):
            fila = dict(base)
            fila.update({p: SIN_MATCH for p in proveedores})
            fila.update({f"Match {p}": SIN_MATCH for p in proveedores})
            if cfg:
                fila["Producto (Dossier)"] = cfg["nombre"](clave, n, len(opciones), d["Productos"])
            elif len(opciones) > 1:
                fila["Producto (Dossier)"] = f"{d['Productos']} (Opción {n})"
            if len(opciones) > 1:
                fila["Opción"] = n
            mejores = list(slot.values())
            for m in mejores:
                fila[m["prov"]] = round(m["precio"], 2) if m["precio"] is not None else NO_COMPARABLE
                fila[f"Match {m['prov']}"] = f"{m['producto']}  [{m['base']}, score {m['score']:.0f}]"

            comparables = [m for m in mejores if m["precio"] is not None]
            if comparables:
                g = min(comparables, key=lambda m: m["precio"])
                es_real = bool(prov_real and _sin_simbolos(g["prov"]) == _sin_simbolos(prov_real))
                dif = 0.0 if es_real else pond - g["precio"]     # el proveedor real no compite contra sí mismo
                pct = dif / pond * 100
                mismo = _sin_simbolos(claves.get(g["prov"], g["prov"])) in _sin_simbolos(d["Proveedor"])
                if pct < AHORRO_MINIMO_PCT:
                    accion = "Precio actual competitivo"
                elif mismo:
                    accion = "Renegociar (mismo proveedor, lista más barata)"
                else:
                    accion = f"Evaluar cambio a {g['prov']}"
                fila.update({
                    "Mejor Opción": g["prov"], "Producto Mejor Opción": g["producto"],
                    "Base comparada": g["base"], "Precio Mejor Opción": round(g["precio"], 2),
                    "Ahorro Unitario": round(dif, 2),
                    "Ahorro %": round(pct, 1),
                    "Ahorro estimado mes": round(dif * cant_inv, 2),
                    "Confianza match": ("Revisar (dif. > 40%)" if pct > SOBREPRECIO_SOSPECHOSO
                                        else "Alta" if g["score"] >= SCORE_ALTA else "Media"),
                    "Score match": round(g["score"], 1),
                    "Acción": accion,
                })
            elif mejores:
                fila["Mejor Opción"] = "UNIDAD NO COMPARABLE"
            else:
                fila["Mejor Opción"] = "SIN MATCH"
            if prov_real and fila.get("Mejor Opción") == prov_real:
                fila["Acción"] = f"Proveedor actual ({prov_real}): sin ahorro vs {prov_real}; marca {esp['nombre']}"
            filas.append(fila)

    return pd.DataFrame(filas)


# =========================================================================== #
# Excel
# =========================================================================== #
COLS_INICIO = ["Código SKU", "Producto (Dossier)", "Opción", "Proveedor actual",
               "Cantidad Inventario", "Precio Ponderado"]
COLS_RESULTADO = ["Mejor Opción", "Precio Mejor Opción", "Ahorro Unitario", "Ahorro %", "Ahorro estimado mes",
                  "Confianza match", "Acción", "Producto Mejor Opción", "Base comparada"]
MONEY_COLS = {"Precio Ponderado", "Precio Mejor Opción", "Ahorro Unitario", "Ahorro estimado mes", "Monto Total",
              "Ahorro a revisar"}
WINNER_FILL = PatternFill("solid", fgColor="C8E6C9")


def columnas_matriz(proveedores: list[str]) -> list[str]:
    """Producto | precio de cada proveedor | Mejor Opción/Ahorro | qué producto matcheó en cada lista."""
    return COLS_INICIO + proveedores + COLS_RESULTADO + [f"Match {p}" for p in proveedores]


def _ordenar(df: pd.DataFrame) -> pd.DataFrame:
    """Estrictamente alfabético (A-Z) por 'Producto (Dossier)', sin distinguir
    mayúsculas ni tildes; a igual nombre, mantiene el orden del Dossier."""
    clave = df["Producto (Dossier)"].map(lambda v: re.sub(r"\s+", " ", strip_accents(v)).strip().lower())
    return df.assign(_k=clave).sort_values("_k", kind="stable").drop(columns="_k")


def _formatear(path: Path, proveedores: list[str]) -> None:
    money = MONEY_COLS | set(proveedores)
    wb = load_workbook(path)
    for ws in wb.worksheets:
        ws.freeze_panes = "C2"
        if ws.max_row > 1:
            ws.auto_filter.ref = ws.dimensions
        headers = [c.value for c in ws[1]]
        for cell in ws[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="1F4E78")
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        for idx, h in enumerate(headers, start=1):
            letter = ws.cell(row=1, column=idx).column_letter
            largo = str(h).startswith(("Producto", "Match", "Acción"))
            ws.column_dimensions[letter].width = 45 if largo else (9 if h == "Opción" else 16)
            if h in money:
                for row in ws.iter_rows(min_row=2, min_col=idx, max_col=idx):
                    row[0].number_format = '"$"#,##0.00'
                    row[0].alignment = Alignment(horizontal="right")
        # Resalta en verde el precio del proveedor ganador de cada fila
        if "Mejor Opción" in headers:
            col_mejor = headers.index("Mejor Opción")
            col_prov = {h: i for i, h in enumerate(headers) if h in proveedores}
            for row in ws.iter_rows(min_row=2):
                i = col_prov.get(row[col_mejor].value)
                if i is not None:
                    row[i].fill = WINNER_FILL
                    row[i].font = Font(bold=True)
    wb.save(path)


def _slug(cat: str) -> str:
    return re.sub(r"[^A-Za-z0-9]+", "_", strip_accents(cat).title()).strip("_")


def exportar(res: pd.DataFrame, out_dir: Path, hoja: str, rubros: dict[str, set[str] | None]) -> None:
    """Un Excel por categoría con la matriz multi-proveedor ('Comparativa'), las
    'Oportunidades' (ahorro >= AHORRO_MINIMO_PCT) y los 'Sin match'; más el Resumen.
    rubros: {proveedor: rubros que cubre}, en el orden de las columnas."""
    out_dir.mkdir(exist_ok=True)
    for viejo in out_dir.glob("Reporte_Ahorro_*.xlsx"):  # que no queden categorías de una corrida anterior
        try:
            viejo.unlink()
        except PermissionError:
            raise RuntimeError(f"No se puede reemplazar {viejo.name}: cerralo en Excel y volvé a correr.") from None
    proveedores = list(rubros)
    res = res.reindex(columns=list(dict.fromkeys([*res.columns, *columnas_matriz(proveedores)])))
    # Unifica variantes de la categoría ('Almacen'/'Almacén', 'frutas y Verduras'...)
    res["_cat"] = res["Categoría 2"].map(_norm_cat)
    nombre_cat = res.groupby("_cat")["Categoría 2"].agg(lambda s: s.value_counts().index[0])

    con_match = res["Mejor Opción"].ne("SIN MATCH")
    es_oportunidad = res["Ahorro %"].fillna(0) >= AHORRO_MINIMO_PCT
    # Los totales cuentan solo la Opción 1 de cada insumo: las otras filas son
    # alternativas del mismo consumo y sumarlas duplicaría el ahorro.
    principal = res["Opción"].fillna(1).eq(1)
    resumen = []
    for cat_key, grupo in res.groupby("_cat"):
        cat = nombre_cat[cat_key]
        provs_cat = [p for p in proveedores if proveedor_aplica(rubros[p], cat_key)]
        cols = columnas_matriz(provs_cat)
        matriz = _ordenar(grupo[con_match.loc[grupo.index]])
        opp = matriz[es_oportunidad.loc[matriz.index]]
        m1 = matriz[principal.loc[matriz.index]]
        opp1 = opp[principal.loc[opp.index]]
        a_revisar = opp1["Confianza match"].str.startswith("Revisar")
        resumen.append({"Categoría 2": cat, "Productos en Dossier": int(principal.loc[grupo.index].sum()),
                        "Con match": len(m1),
                        "Con precio comparable": int(m1["Precio Mejor Opción"].notna().sum()),
                        "Oportunidades": int((~a_revisar).sum()),
                        "Ahorro estimado mes": round(opp1.loc[~a_revisar, "Ahorro estimado mes"].sum(), 2),
                        "Oportunidades a revisar": int(a_revisar.sum()),
                        "Ahorro a revisar": round(opp1.loc[a_revisar, "Ahorro estimado mes"].sum(), 2)})
        if matriz.empty:
            LOG.info("%s: ningún producto encontrado en las listas", cat)
            continue
        path = out_dir / f"Reporte_Ahorro_{_slug(cat)}.xlsx"
        with pd.ExcelWriter(path, engine="openpyxl") as w:
            matriz[cols].to_excel(w, sheet_name="Comparativa", index=False)
            opp[cols].to_excel(w, sheet_name="Oportunidades", index=False)
            _ordenar(grupo.loc[~con_match.loc[grupo.index]])[COLS_INICIO] \
                .to_excel(w, sheet_name="Sin match", index=False)
        _formatear(path, provs_cat)
        LOG.info("%s: %d/%d con match (%d filas con opciones), %d oportunidades (+%d a revisar), ahorro estimado $%s -> %s",
                 cat, len(m1), resumen[-1]["Productos en Dossier"], len(matriz), resumen[-1]["Oportunidades"],
                 resumen[-1]["Oportunidades a revisar"],
                 f"{resumen[-1]['Ahorro estimado mes']:,.0f}", path.name)

    path = out_dir / "Reporte_Ahorro_Resumen.xlsx"
    df_res = pd.DataFrame(resumen).sort_values("Ahorro estimado mes", ascending=False)
    cols = ["Categoría 2"] + columnas_matriz(proveedores)
    with pd.ExcelWriter(path, engine="openpyxl") as w:
        df_res.to_excel(w, sheet_name=f"Resumen {hoja}"[:31], index=False)
        _ordenar(res[con_match])[cols].to_excel(w, sheet_name="Comparativa general", index=False)
        _ordenar(res[con_match & es_oportunidad])[cols].to_excel(w, sheet_name="Todas las oportunidades", index=False)
    _formatear(path, proveedores)
    LOG.info("Resumen: %d/%d productos con match, ahorro estimado $%s (+ $%s a revisar) -> %s",
             int((con_match & principal).sum()), int(principal.sum()), f"{df_res['Ahorro estimado mes'].sum():,.0f}",
             f"{df_res['Ahorro a revisar'].sum():,.0f}", path.name)


# =========================================================================== #
# CLI
# =========================================================================== #
def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--hoja", default="Cadena Octubre", help="Pestaña del Dossier a comparar (default: Cadena Octubre)")
    ap.add_argument("--dossier", default=str(DOSSIER))
    ap.add_argument("--carpeta-id", default=DRIVE_FOLDER_ID, help="ID de la carpeta de listas en Drive")
    ap.add_argument("--creds", default=str(DEFAULT_CREDS))
    ap.add_argument("--listas-dir", help="Usar listas ya descargadas en esta carpeta en vez de Drive")
    ap.add_argument("--criollo-csv", default=str(CRIOLLO_CSV),
                    help="Catálogo scrapeado de El Criollo (se usa si existe y no hay uno en Drive); '' para omitirlo")
    ap.add_argument("--nogada-csv", default=str(NOGADA_CSV),
                    help="Dónde guardar/leer el catálogo scrapeado de Nogada")
    ap.add_argument("--sin-nogada", action="store_true", help="No scrapear ni usar Nogada")
    ap.add_argument("--productos-txt", default=str(PRODUCTOS_TXT),
                    help="Lista de productos a buscar; sus marcas son obligatorias en el match ('' para desactivar)")
    ap.add_argument("--out-dir", default=str(OUTPUT_DIR))
    args = ap.parse_args()

    logging.basicConfig(level=logging.INFO, format="%(asctime)s | %(levelname)s | %(message)s", datefmt="%H:%M:%S")

    dossier = pd.read_excel(args.dossier, sheet_name=args.hoja)
    LOG.info("Dossier '%s': %d filas", args.hoja, len(dossier))
    dossier = aplicar_lista_negra(dossier)
    dossier = aplicar_exclusiones_dossier(dossier)

    def con_criollo(paths: list[Path]) -> list[Path]:
        criollo = Path(args.criollo_csv) if args.criollo_csv else None
        if criollo and criollo.is_file() and not any("criollo" in p.name.lower() for p in paths):
            LOG.info("El Criollo: usando %s (scrapeado el %s)", criollo.name,
                     pd.Timestamp(criollo.stat().st_mtime, unit="s").strftime("%d/%m/%Y"))
            paths = [*paths, criollo]
        return paths

    def con_nogada(paths: list[Path]) -> list[Path]:
        if args.sin_nogada or any("nogada" in p.name.lower() for p in paths):
            return paths
        cache = Path(args.nogada_csv)
        try:
            LOG.info("Nogada: %d productos scrapeados de nogada.com.ar", scrapear_nogada(cache))
        except Exception as e:
            if not cache.is_file():
                LOG.warning("Nogada: no se pudo scrapear (%s) y no hay copia previa; se omite", e)
                return paths
            LOG.warning("Nogada: no se pudo scrapear (%s); uso la copia del %s", e,
                        pd.Timestamp(cache.stat().st_mtime, unit="s").strftime("%d/%m/%Y"))
        return [*paths, cache]

    if args.listas_dir:
        listas = cargar_listas(con_nogada(con_criollo(
            sorted(p for p in Path(args.listas_dir).iterdir() if p.is_file()))))
    else:
        with tempfile.TemporaryDirectory(prefix="listas_proveedores_") as tmp:
            listas = cargar_listas(con_nogada(con_criollo(
                descargar_listas(Path(tmp), args.carpeta_id, Path(args.creds)))))

    lineas_txt = leer_productos_txt(Path(args.productos_txt)) if args.productos_txt else []
    LOG.info("productos a buscar: %d líneas, %d con marca", len(lineas_txt), sum(bool(l["marcas"]) for l in lineas_txt))
    res = comparar(dossier, listas, lineas_txt)
    n_comp = int(res["Precio Mejor Opción"].notna().sum()) if "Precio Mejor Opción" in res else 0
    prin = res["Opción"].eq(1)
    LOG.info("Con match: %d / %d insumos (con precio comparable: %d; %d filas incl. opciones)",
             int((res["Mejor Opción"].ne("SIN MATCH") & prin).sum()), int(prin.sum()), n_comp, len(res))
    if not n_comp:
        LOG.warning("Ningún producto del Dossier tuvo match comparable.")
        return 1
    rubros = listas.drop_duplicates("proveedor").set_index("proveedor")["rubros"].to_dict()
    try:
        exportar(res, Path(args.out_dir), args.hoja, rubros)
    except RuntimeError as e:
        LOG.error("%s", e)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
