"""Precios de transferencia del Obrador (Congreso) hacia Chimidoha y los demás locales ("pasamano").

Congreso (Obrador) COMPRA la mercadería a los proveedores —factura a nombre de Lafinur 3368 SRL / sello Cerviño— y luego se la
"vende" internamente a Chimidoha y a los locales. La matriz de recargos sobre el costo de compra es:

    Panificados y Descartables (insumos) ..... costo x 1.20  (+20 %)
    Proteínas (Carnes y Pescados) ............ costo x 1.10  (+10 %)
    Aceite e insumos normales de almacén ..... costo x 1.00  (al costo)
    Productos elaborados por el Obrador ...... explosión completa por unidad (GR, UN, PORC, KG) contra las recetas (BOM)

`aprender_de_remitos_chimidoha()` detecta la hoja/archivo "Remitos Chimidoha" (cuando exista) y aprende de sus cruces el
recargo realmente aplicado por rubro; el resultado queda en aprendizaje_transferencias.json y se compara con la matriz.

Uso:
    python transferencias_obrador.py            # busca "Remitos Chimidoha" y resume lo aprendido
"""

from __future__ import annotations

import json
import re
import statistics
import sys
import unicodedata
from pathlib import Path
from typing import Any

BASE_DIR = Path(__file__).resolve().parent
APRENDIZAJE_TRANSFERENCIAS = BASE_DIR / "aprendizaje_transferencias.json"

RECARGO_PANIFICADOS_DESCARTABLES = 1.20
RECARGO_PROTEINAS = 1.10
RECARGO_AL_COSTO = 1.00

# Categoría 2 (col E de Productos) -> tipo de transferencia
_PANIF = {"panificados y elaborados", "descartables", "panificados", "compras varias por ml y caja", "imprenta"}
_PROTEINAS = {"carnes", "pescados"}
UNIDADES_ELABORADOS = {"GR": 1.0, "G": 1.0, "KG": 1000.0, "UN": None, "PORC": None}   # None = unidades de receta sin convertir


def _n(t: Any) -> str:
    return re.sub(r"\s+", " ", unicodedata.normalize("NFKD", str(t or "")).encode("ascii", "ignore").decode().lower()).strip()


def tipo_transferencia(categoria2: str | None, elaborado_por_obrador: bool = False) -> str:
    if elaborado_por_obrador:
        return "ELABORADO_BOM"
    c = _n(categoria2)
    if c in _PANIF:
        return "PANIFICADO_DESCARTABLE"
    if c in _PROTEINAS:
        return "PROTEINA"
    return "AL_COSTO"


def factor_transferencia(categoria2: str | None) -> float:
    return {"PANIFICADO_DESCARTABLE": RECARGO_PANIFICADOS_DESCARTABLES, "PROTEINA": RECARGO_PROTEINAS,
            "AL_COSTO": RECARGO_AL_COSTO}.get(tipo_transferencia(categoria2), 1.0)


def precio_transferencia(costo: float, categoria2: str | None) -> float:
    """Precio al que el Obrador transfiere un insumo comprado (no elaborado)."""
    return round(float(costo) * factor_transferencia(categoria2), 4)


def costo_elaborado(motor, codigo: str, cantidad: float, unidad: str = "UN") -> dict | None:
    """Producto elaborado por Obrador: explosión completa de la receta (BOM) por la unidad vendida.
    `motor` = motor_recetas.MotorRecetas. GR/KG se llevan a gramos; UN/PORC usan la unidad de receta."""
    u = unidad.upper()
    mult = UNIDADES_ELABORADOS.get(u, None)
    q = cantidad * mult if mult else cantidad
    r = motor.explotar(codigo, q)
    if r is None:
        return None
    return {"codigo": codigo, "cantidad": cantidad, "unidad": u, "cantidad_receta": q,
            "costo_total": round(r.costo_total, 2), "costo_unitario": round(r.costo_total / cantidad, 4) if cantidad else None,
            "tipo": "ELABORADO_BOM"}


# --------------------------------------------------------------------------- #
# "Remitos Chimidoha": detección y aprendizaje
# --------------------------------------------------------------------------- #
_PATRON_ARCHIVO = re.compile(r"remitos?\s*[_ -]?\s*chimi", re.I)


def buscar_remitos_chimidoha(extra_dirs: list[Path] | None = None) -> list[Path]:
    dirs = [BASE_DIR, BASE_DIR / "Arcos_data", BASE_DIR / "Facturas Mishiguene", BASE_DIR / "Facturas_WPP",
            BASE_DIR / "Copias_Gestion", BASE_DIR / "Carpeta Revision"] + list(extra_dirs or [])
    out: list[Path] = []
    for d in dirs:
        if d.is_dir():
            out += [p for p in d.iterdir() if p.suffix.lower() in (".xlsx", ".xls", ".csv") and _PATRON_ARCHIVO.search(p.name)]
    return sorted(set(out))


def _leer_tablas(path: Path):
    import pandas as pd
    if path.suffix.lower() == ".csv":
        return {"csv": pd.read_csv(path, header=None, sep=None, engine="python", dtype=str)}
    return pd.read_excel(path, sheet_name=None, header=None)


def _columnas(df) -> dict | None:
    """Encabezado de remitos internos: producto / cantidad / precio (unitario o de transferencia) / total."""
    for i in range(min(15, len(df))):
        celdas = [_n(c) for c in df.iloc[i]]
        def col(*claves):
            return next((j for j, c in enumerate(celdas) if c and any(k in c for k in claves)), None)
        prod, cant = col("producto", "descripcion", "articulo", "insumo"), col("cantidad", "cant")
        prec = col("precio", "p. unit", "p.unit", "costo", "valor")
        if prod is not None and cant is not None and prec is not None:
            return {"hdr": i, "prod": prod, "cant": cant, "prec": prec,
                    "rubro": col("rubro", "categoria"), "costo": col("costo obrador", "costo compra", "costo base")}
    return None


def aprender_de_remitos_chimidoha(paths: list[Path] | None = None, costo_fn=None) -> dict:
    """Lee los remitos Obrador→Chimidoha y calcula el recargo real (precio transferido / costo de compra) por rubro.
    `costo_fn(producto) -> (costo, categoria2) | None` da el costo vigente en Congreso (por defecto lo toma de las
    planillas de Facturas)."""
    import pandas as pd
    paths = paths if paths is not None else buscar_remitos_chimidoha()
    res: dict[str, Any] = {"archivos": [str(p.name) for p in paths], "por_tipo": {}, "filas": 0, "sin_costo": 0}
    if not paths:
        res["estado"] = "Sin 'Remitos Chimidoha' todavía: se aplica la matriz de recargos por defecto"
        return res
    if costo_fn is None:
        sys.path.insert(0, str(BASE_DIR))
        import imputacion_facturas as imp
        loc = imp.Local("Congreso")

        def costo_fn(producto):
            m = imp.mapear_producto(str(producto), None, loc, {"mapeos": {}})
            if not m["presentacion"]:
                return None
            vig, _ = loc.costo_vigente(m["presentacion"])
            cat = loc.catalogo[loc.catalogo["PRESENTACION"] == m["presentacion"]].iloc[0]
            return (vig, cat["CAT2"]) if vig else None
    ratios: dict[str, list[float]] = {}
    for p in paths:
        for nombre, df in _leer_tablas(p).items():
            c = _columnas(df)
            if not c:
                continue
            for _, r in df.iloc[c["hdr"] + 1:].iterrows():
                prod = r[c["prod"]]
                precio = pd.to_numeric(r[c["prec"]], errors="coerce")
                if pd.isna(precio) or precio <= 0 or not str(prod).strip() or str(prod) == "nan":
                    continue
                res["filas"] += 1
                if c["costo"] is not None and pd.notna(pd.to_numeric(r[c["costo"]], errors="coerce")):
                    costo, cat = float(pd.to_numeric(r[c["costo"]])), (r[c["rubro"]] if c["rubro"] is not None else None)
                else:
                    got = costo_fn(prod)
                    if not got:
                        res["sin_costo"] += 1
                        continue
                    costo, cat = got
                tipo = tipo_transferencia(cat)
                ratios.setdefault(tipo, []).append(float(precio) / costo)
    for tipo, v in ratios.items():
        esperado = {"PANIFICADO_DESCARTABLE": 1.2, "PROTEINA": 1.1, "AL_COSTO": 1.0}.get(tipo)
        med = statistics.median(v)
        res["por_tipo"][tipo] = {"n": len(v), "recargo_mediano": round(med, 3), "matriz": esperado,
                                 "coincide": esperado is not None and abs(med - esperado) <= 0.03}
    res["estado"] = "aprendido" if ratios else "archivo detectado pero sin cruces utilizables"
    APRENDIZAJE_TRANSFERENCIAS.write_text(json.dumps(res, ensure_ascii=False, indent=1), encoding="utf-8")
    return res


if __name__ == "__main__":
    print(json.dumps(aprender_de_remitos_chimidoha(), ensure_ascii=False, indent=1))
