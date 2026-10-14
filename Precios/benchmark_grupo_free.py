"""Grupo Free (Free Vegetales) — proveedor monopólico de papa/cebolla/zanahoria procesadas: PRECIO DE REFERENCIA.

No compite en el comparador (no tiene competidor directo en esos procesados): sirve para
  1. auditar que las facturas de Free cumplan la lista (`auditar_precio`),
  2. comparar a futuro contra elaboración propia en el Obrador o contra proveedores nuevos (`precio_referencia`).

    python Precios/benchmark_grupo_free.py "Grupo Free - Naturales Septiembre v2.pdf"
genera Precios/benchmark_grupo_free.json (precios SIN impuestos y SIN envío).
"""

from __future__ import annotations

import json
import re
import sys
import unicodedata
from pathlib import Path

import pdfplumber

AQUI = Path(__file__).resolve().parent
JSON_BENCH = AQUI / "benchmark_grupo_free.json"
_PROD_RE = re.compile(r"^(?P<p>[A-ZÁÉÍÓÚÑ ]+?)\s+PRESENTACION\s+\$\s*x\s*(?P<u>KG|Litro)\s*$", re.I)
_ITEM_RE = re.compile(r"^(?P<d>.+?)\s+\$\s*(?P<v>[\d,]+\.\d{2})\s*$")


def _n(t: str) -> str:
    return re.sub(r"[^a-z0-9]+", " ", unicodedata.normalize("NFKD", t).encode("ascii", "ignore").decode().lower()).strip()


def parsear_lista(pdf_path: Path) -> dict:
    with pdfplumber.open(pdf_path) as pdf:
        texto = "\n".join(p.extract_text() or "" for p in pdf.pages)
    items, prod, unidad, cond = [], None, "KG", []
    for ln in (l.strip() for l in texto.splitlines()):
        m = _PROD_RE.match(ln)
        if m:
            prod, unidad = m["p"].strip(), "LT" if m["u"].lower().startswith("l") else "KG"
            continue
        if re.match(r"^(PRECIOS SIN|ENVIOS|MINIMO)", ln, re.I):
            cond.append(ln)
            continue
        if re.match(r"^(ACEITES GASTRONOMICOS)", ln, re.I):
            prod, unidad = "ACEITE", "LT"
            continue
        m = _ITEM_RE.match(ln)
        if m and prod:
            items.append({"producto": prod, "presentacion": m["d"].strip(), "precio": float(m["v"].replace(",", "")),
                          "unidad": unidad, "clave": _n(f"{prod} {m['d']}")})
    return {"proveedor": "GRUPO FREE (Free Vegetales)", "lista": pdf_path.stem, "tipo": "PRECIO_DE_REFERENCIA / BENCHMARK",
            "unidad_precio": "$/KG o $/LT, SIN impuestos y SIN costo de envío",
            "condiciones": cond, "items": items}


def guardar(pdf_path: Path) -> dict:
    d = parsear_lista(pdf_path)
    JSON_BENCH.write_text(json.dumps(d, ensure_ascii=False, indent=1), encoding="utf-8")
    return d


def cargar() -> dict:
    return json.loads(JSON_BENCH.read_text(encoding="utf-8"))


def precio_referencia(descripcion: str, bench: dict | None = None) -> dict | None:
    """Mejor ítem de la lista para una descripción ('PAPA ENTERA COMÚN x KG (BOLSA 5 KGS)') por solapamiento de palabras."""
    bench = bench or cargar()
    q = set(_n(descripcion).split())
    mejor, mejor_s = None, 0.0
    for it in bench["items"]:
        toks = set(it["clave"].split())
        clave = set(_n(it["producto"]).split())
        if not clave <= q:                       # el producto (papa, cebolla…) tiene que estar
            continue
        s = len(q & toks) / len(toks)
        if s > mejor_s:
            mejor, mejor_s = it, s
    return dict(mejor, coincidencia=round(mejor_s, 2)) if mejor and mejor_s >= 0.5 else None


def auditar_precio(descripcion: str, precio_neto_unitario: float, bench: dict | None = None, tol: float = 0.03) -> dict:
    """Compara el precio de una factura/NC contra la lista. `precio_neto_unitario` ya con la bonificación aplicada."""
    ref = precio_referencia(descripcion, bench)
    if not ref:
        return {"estado": "SIN_REFERENCIA"}
    dif = precio_neto_unitario / ref["precio"] - 1
    return {"estado": "OK" if abs(dif) <= tol else ("BAJO_LISTA" if dif < 0 else "SOBRE_LISTA"),
            "lista": ref["precio"], "item_lista": f"{ref['producto']} {ref['presentacion']}", "dif_pct": round(dif * 100, 1)}


if __name__ == "__main__":
    ruta = Path(sys.argv[1]) if len(sys.argv) > 1 else AQUI / "Grupo Free - Naturales Septiembre v2.pdf"
    r = guardar(ruta)
    print(f"{len(r['items'])} precios de referencia -> {JSON_BENCH.name}")
