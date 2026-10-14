"""Indexa los informes operativos que llegan junto a las facturas de WhatsApp (Maxirest) para el cierre mensual.

Archivos reconocidos (se copian a Datos_Operativos_Historicos/ y se resumen en indice_operativo.json):
  * VENTAS RESUMIDAS.xlsx           -> ventas por forma de pago / comprobante / canal, cubiertos y promedio por cubierto
  * INFORME DE MOV DE STOCK.xlsx    -> egresos de stock del mes (Maxirest): consumo interno, errores de comanda,
    INFORME MOV.DE STOCK JULIO.xlsx    marketing, desperdicio, invitaciones por mozo, etc. (valor a precio de venta)

Uso:
    python datos_operativos.py [carpeta_origen]   # default Facturas_WPP/
"""

from __future__ import annotations

import json
import logging
import re
import shutil
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd

BASE_DIR = Path(__file__).resolve().parent
DESTINO = BASE_DIR / "Datos_Operativos_Historicos"
LOG = logging.getLogger("datos_operativos")

# Conceptos de egreso de stock que NO son venta: se separan por naturaleza para el cierre
CONCEPTO_TIPO = {
    "ERROR COMANDA": "error_comanda", "DESPERDICIO": "desperdicio", "MARKETING": "marketing",
    "CONSUMO INTERNO": "consumo_interno", "CONSUMO ADMINISTRACI": "consumo_administracion", "BIG BOX": "promocion_big_box",
    "EGRESO DE STOCK": "egreso_manual",
}


def _txt(s) -> str:
    return re.sub(r"\s+", " ", str(s)).strip()


def indexar_ventas_resumidas(path: Path) -> dict:
    """Maxirest 'Ventas resumidas': la columna ID viene corrupta (fechas) pero DETALLE/IMPORTE/CNTOPER son confiables."""
    v = pd.read_excel(path)
    v["DETALLE"] = v["DETALLE"].map(_txt)
    v = v[v["DETALLE"] != ""]
    d = {r.DETALLE.lower(): r for r in v.itertuples()}

    def imp(k):
        return float(d[k].IMPORTE) if k in d else None

    def ops(k):
        return int(d[k].CNTOPER) if k in d else None

    def tot(k):
        return float(d[k].TOTAL) if k in d else None
    total = tot("total")
    canales = {k: {"importe": imp(k), "operaciones": ops(k), "porcentaje": float(d[k].PORCENTAJE)}
               for k in ("salon", "mostrador", "vereda", "rappi") if k in d}
    res = {
        "archivo": path.name,
        "periodo": "no figura en el archivo (la columna ID quedó con fechas corruptas); cierre mensual de un local",
        "total_ventas": total, "operaciones": ops("total"),
        "ticket_promedio": round(total / ops("total"), 2) if total and ops("total") else None,
        "forma_de_pago": {"efectivo": imp("efectivo"), "tarjetas": imp("tarjetas"), "otros_rappi": imp("otros")},
        "comprobantes": {"factura_a": imp("factura a fiscal"), "ticket": imp("ticket"), "factura_b": imp("factura b fiscal"),
                         "nota_credito_b": imp("n.crédito b fiscal")},
        "anulaciones": imp("anulaciones"),
        "descuentos": {"socios": imp("5.socios"), "empleados": imp("12.descuento empleados"),
                       "rappi": imp("28.dto restaurante- rappi")},
        "cubiertos": {"total": tot("total cubiertos"), "salon": tot("cubiertos de salon"),
                      "mostrador": tot("cubiertos de mostrador"), "vereda": tot("cubiertos de vereda")},
        "promedio_por_cubierto": imp("promedio por cubierto"),
        "canales": canales,
        "ventas_mostrador": imp("mostrador"), "ventas_rappi": imp("rappi"),
        "neto_21": tot("neto acf 21.00%") or imp("neto acf 21.00%"), "iva_21": imp("iva  acf 21.00%"),
    }
    res["cubierto_promedio_calculado"] = (round(total / res["cubiertos"]["total"], 2)
                                          if total and res["cubiertos"]["total"] else None)
    return res


def indexar_mov_stock(path: Path) -> dict:
    sheets = pd.read_excel(path, sheet_name=None)
    nombre = "Hoja1" if "Hoja1" in sheets else next(iter(sheets))
    d = sheets[nombre]
    d.columns = [_txt(c) for c in d.columns]
    for c in ("TURNO", "CONCEPTO", "NOMBRE", "SUCURSAL", "EMP_NOM"):
        if c in d:
            d[c] = d[c].map(_txt)
    d["FECHA"] = pd.to_datetime(d["FECHA"], errors="coerce")
    d["VALOR"] = pd.to_numeric(d["PRECIO"], errors="coerce").fillna(0).abs()
    por_concepto = d.groupby("CONCEPTO").agg(lineas=("VALOR", "size"), valor=("VALOR", "sum")).sort_values("valor", ascending=False)
    clasif: dict[str, dict] = {}
    for conc, r in por_concepto.iterrows():
        tipo = CONCEPTO_TIPO.get(conc.upper(), "invitacion_o_consumo_personal")
        c = clasif.setdefault(tipo, {"valor": 0.0, "lineas": 0, "conceptos": []})
        c["valor"] += float(r.valor)
        c["lineas"] += int(r.lineas)
        c["conceptos"].append(conc)
    top = (d[d["VALOR"] > 0].groupby("NOMBRE")["VALOR"].sum().sort_values(ascending=False).head(10)).round(0).to_dict()
    cubiertos = d[d["NOMBRE"].str.contains("cubierto", case=False, na=False)]["CANTIDAD"].abs().sum()
    meses = d["FECHA"].dt.to_period("M").dropna().unique()
    return {
        "archivo": path.name, "periodo": str(meses[0]) if len(meses) == 1 else [str(m) for m in meses],
        "desde": str(d["FECHA"].min().date()), "hasta": str(d["FECHA"].max().date()),
        "lineas": int(len(d)), "remitos_distintos": int(d["REMITO"].nunique()),
        "sucursal": sorted(set(d["SUCURSAL"].dropna())) if "SUCURSAL" in d else None,
        "valor_total_egresos": float(d["VALOR"].sum()),
        "por_naturaleza": {k: {"valor": round(v["valor"], 0), "lineas": v["lineas"], "conceptos": v["conceptos"]}
                           for k, v in sorted(clasif.items(), key=lambda kv: -kv[1]["valor"])},
        "top_productos_por_valor": top, "cubiertos_registrados": float(cubiertos),
        "uso": "auditoría de movimientos de stock del cierre (error de comanda, marketing, desperdicio, consumo interno)",
    }


def indexar(origen: Path = BASE_DIR / "Facturas_WPP", destino: Path = DESTINO) -> dict:
    destino.mkdir(exist_ok=True)
    idx: dict = {"generado": datetime.now().isoformat(timespec="seconds"), "origen": str(origen.name), "informes": []}
    for p in sorted(origen.iterdir()):
        n = p.name.lower()
        if p.suffix.lower() not in (".xlsx",):
            continue
        try:
            if "ventas resumidas" in n:
                info = indexar_ventas_resumidas(p)
                info["tipo"] = "ventas_resumidas"
            elif "mov" in n and "stock" in n:
                info = indexar_mov_stock(p)
                info["tipo"] = "movimientos_de_stock"
            else:
                continue
        except Exception as e:
            LOG.error("%s: %s", p.name, e)
            continue
        shutil.copy2(p, destino / p.name)
        info["copiado_a"] = str((destino / p.name).relative_to(BASE_DIR))
        idx["informes"].append(info)
    (destino / "indice_operativo.json").write_text(json.dumps(idx, ensure_ascii=False, indent=1, default=str),
                                                   encoding="utf-8")
    return idx


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, format="%(message)s")
    carpeta = Path(sys.argv[1]) if len(sys.argv) > 1 else BASE_DIR / "Facturas_WPP"
    res = indexar(carpeta)
    print(json.dumps(res, ensure_ascii=False, indent=1, default=str)[:6000])
