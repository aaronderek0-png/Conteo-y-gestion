"""Desperdicios y Traspasos de TODOS los locales, explotados a insumo base con el motor de recetas.

Regla global (misma para Arcos, Cabello, Lafinur y Congreso):
  * Insumo directo / comprado (PITA, BAGEL, CROISSANT..., materia prima con precio en 'Insumos'): se computa 1:1
    al insumo, llevando la cantidad a su unidad base (g / ml / un).
  * Elaboración o semielaborado (código de receta, Madre, OBR/M/MR/MO, o un nombre que la identifica): se explota
    recursivamente (BOM) con `MotorRecetas.explotar` y se imputan los insumos base proporcionales a la cantidad.
  * Lo que no se puede resolver (código sin receta ni insumo, unidad incompatible, cantidad fuera de rango) se
    informa aparte y NO se imputa a ningún insumo.

Salida: Copias_Gestion/../Movimientos_Explotados_<MM>_<AAAA>.xlsx (en la raíz) y, si el archivo no está bloqueado,
la hoja 'Desperdicios y Traspasos (BOM)' dentro de cada Analisis_Gestion_Comercial_<Local>_<MM>_<AAAA>.xlsx.

Uso:  python movimientos_locales.py [--anio 2026] [--meses 9 10] [--sin-hojas]
"""
from __future__ import annotations

import argparse
import logging
import sys
from pathlib import Path

import pandas as pd

import analisis_gestion_comercial as agc
from defensivo import es_panificado_comprado
from motor_recetas import MotorRecetas, _base, es_elaboracion

LOG = logging.getLogger("movimientos_locales")
BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "Arcos_data"

# local -> (nombre en la planilla de traspasos, libro con la hoja 'Desperdicios' o None si no existe)
LOCALES_MOV = {
    "Arcos": ("ARCOS", "arcos.xlsx"),
    "Congreso": ("CONGRESO", "congreso_facturas.xlsx"),
    "Lafinur": ("LAFINUR", "lafinur.xlsx"),
    "Cabello": ("CABELLO", "cabello.xlsx"),
}
COLS_DETALLE = ["Local", "Mes", "Tipo", "Fecha", "Código", "Producto", "Cantidad", "Unidad", "Importe fichado $",
                "Vía", "Estado", "Insumo base", "Nombre insumo base", "Cantidad base", "Unidad base", "Costo base $"]


def _clase_insumo(motor: MotorRecetas, sku: str) -> tuple[str, str]:
    ins = motor.insumos[sku]
    return _base(ins["unidad"])[0], (ins["unidad"] or "un")


def explotar_fila(motor: MotorRecetas, codigo: str, producto: str, qty: float, unidad: str):
    """-> (vía, estado, [(sku_base, cantidad_base, unidad_base, costo)])."""
    cod = str(codigo or "").strip().upper()
    clase_mov, mult = _base(unidad)
    ins = motor.insumos.get(cod)
    directo = es_panificado_comprado(cod) or (
        ins is not None and ins["precio"] is not None and not es_elaboracion(cod, producto)
        and motor.resolver_receta(cod) is None)
    if directo and ins is not None:
        clase_ins, uni = _clase_insumo(motor, cod)
        if clase_ins != clase_mov:
            return "Directo 1:1", f"unidad incompatible ({unidad} -> {uni}): sin imputar", []
        q = qty * mult
        return "Directo 1:1", "ok", [(cod, q, uni, q * (ins["precio"] or 0.0))]
    cod_rec = motor.resolver_receta(cod) if cod else None
    if cod_rec is None and producto:
        cod_rec, sim, _ = motor.buscar_por_nombre(producto)
    if cod_rec is not None:
        u_rec = motor.unidad_receta.get(cod_rec) or "gr"
        clase_rec = _base(u_rec)[0]
        if clase_rec != clase_mov:
            return "BOM", f"unidad incompatible ({unidad} vs receta en {u_rec}): sin imputar", []
        r = motor.explotar(cod_rec, qty * mult)
        if r is None or (not r.primarios and r.faltantes):
            falt = "; ".join(r.faltantes[:2]) if r else "receta inexistente"
            return "BOM", f"receta no calculable ({falt})", []
        filas = []
        for sku, q in r.primarios.items():
            i2 = motor.insumos.get(sku)
            uni = (i2["unidad"] if i2 else "") or ""
            filas.append((sku, q, uni, q * ((i2 or {}).get("precio") or 0.0)))
        estado = "ok" if r.completo else f"ok parcial (sin precio: {len(r.faltantes)} insumo/s)"
        return "BOM", estado, filas
    if ins is not None:                      # insumo sin precio y sin receta: igual se imputa 1:1 en cantidad
        clase_ins, uni = _clase_insumo(motor, cod)
        if clase_ins == clase_mov:
            q = qty * mult
            return "Directo 1:1", "ok (insumo sin precio)", [(cod, q, uni, 0.0)]
    return "Sin resolver", "no es insumo ni receta conocida", []


def movimientos_local(motor: MotorRecetas, local: str, anio: int, mes: int) -> pd.DataFrame:
    nombre_trasp, libro = LOCALES_MOV[local]
    filas: list[dict] = []

    def agregar(tipo, fecha, cod, prod, qty, uni, importe, error=False):
        if error:
            via, estado, base = "Excluido", "cantidad fuera de rango (¿error de unidad al cargar?)", []
        else:
            via, estado, base = explotar_fila(motor, cod, prod, qty, uni)
        comunes = dict(Local=local, Mes=mes, Tipo=tipo, Fecha=fecha, Código=cod, Producto=prod, Cantidad=qty,
                       Unidad=uni, **{"Importe fichado $": importe}, Vía=via, Estado=estado)
        if not base:
            filas.append({**comunes, "Insumo base": "", "Nombre insumo base": "", "Cantidad base": None,
                          "Unidad base": "", "Costo base $": None})
        for sku, q, u, costo in base:
            ins = motor.insumos.get(sku, {})
            filas.append({**comunes, "Insumo base": sku, "Nombre insumo base": ins.get("nombre") or "",
                          "Cantidad base": q, "Unidad base": u, "Costo base $": costo})

    if libro and (DATA_DIR / libro).exists():
        d = agc.cargar_desperdicios(DATA_DIR / libro, anio, mes)
        for r in d.itertuples(index=False):
            agregar("Desperdicio fichado", r.fecha, r.codigo, r.producto, r.qty, r.unidad, r.total, r.error_carga)
    t = agc.cargar_traspasos(DATA_DIR / "traspasos.xlsx", anio, mes, nombre_trasp)
    for r in t.itertuples(index=False):
        agregar("Traspaso entrante" if r.sentido == "IN" else "Traspaso saliente", r.fecha, r.codigo, r.producto,
                r.qty, r.unidad, r.total)
    return pd.DataFrame(filas, columns=COLS_DETALLE)


def resumen_insumos(det: pd.DataFrame) -> pd.DataFrame:
    """Por local, tipo e insumo base: cantidad y costo imputados."""
    d = det[det["Insumo base"] != ""]
    if d.empty:
        return pd.DataFrame(columns=["Local", "Tipo", "Insumo base", "Nombre insumo base", "Unidad base",
                                     "Cantidad base", "Costo base $", "Movimientos"])
    g = (d.groupby(["Local", "Tipo", "Insumo base", "Nombre insumo base", "Unidad base"], as_index=False)
         .agg(**{"Cantidad base": ("Cantidad base", "sum"), "Costo base $": ("Costo base $", "sum"),
                 "Movimientos": ("Código", "count")}))
    return g.sort_values(["Local", "Tipo", "Costo base $"], ascending=[True, True, False])


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--anio", type=int, default=2026)
    ap.add_argument("--meses", type=int, nargs="+", default=[9, 10])
    ap.add_argument("--sin-hojas", action="store_true", help="no agregar la hoja a los Analisis_*.xlsx")
    args = ap.parse_args()
    logging.basicConfig(level=logging.INFO, format="%(asctime)s | %(levelname)s | %(message)s", datefmt="%H:%M:%S")
    motor = MotorRecetas(DATA_DIR / "recetas.xlsx")
    for av in motor.avisos[:5]:
        LOG.warning("Recetas: %s", av)
    for mes in args.meses:
        partes = [movimientos_local(motor, loc, args.anio, mes) for loc in LOCALES_MOV]
        det = pd.concat(partes, ignore_index=True)
        res_ins = resumen_insumos(det)
        por_local = (det.groupby(["Local", "Tipo", "Vía"], as_index=False)
                     .agg(Filas=("Código", "count"), **{"Importe fichado $": ("Importe fichado $", "sum")}))
        sin = det[det["Vía"].isin(["Sin resolver", "Excluido"]) | det["Estado"].str.startswith("unidad incompatible")
                  | det["Estado"].str.startswith("receta no calculable")]
        (BASE_DIR / "Reportes_Generados").mkdir(exist_ok=True)
        destino = BASE_DIR / "Reportes_Generados" / f"Movimientos_Explotados_{mes:02d}_{args.anio}.xlsx"
        try:
            with pd.ExcelWriter(destino) as w:
                por_local.to_excel(w, sheet_name="Resumen por local", index=False)
                res_ins.to_excel(w, sheet_name="Por insumo base", index=False)
                det.to_excel(w, sheet_name="Detalle", index=False)
                sin.to_excel(w, sheet_name="Sin imputar", index=False)
        except PermissionError:
            destino = destino.with_name(destino.stem + "_nuevo.xlsx")
            with pd.ExcelWriter(destino) as w:
                por_local.to_excel(w, sheet_name="Resumen por local", index=False)
                res_ins.to_excel(w, sheet_name="Por insumo base", index=False)
                det.to_excel(w, sheet_name="Detalle", index=False)
                sin.to_excel(w, sheet_name="Sin imputar", index=False)
            LOG.warning("Archivo bloqueado: se guardó como %s", destino.name)
        LOG.info("Mes %02d: %d filas de detalle -> %s", mes, len(det), destino.name)
        for loc in LOCALES_MOV:
            x = det[det["Local"] == loc]
            LOG.info("  %-9s desperdicio $%12s | traspasos IN/OUT %3d/%3d | sin imputar %d", loc,
                     f"{x.loc[x['Tipo'] == 'Desperdicio fichado', 'Importe fichado $'].drop_duplicates().sum():,.0f}",
                     x[x['Tipo'] == 'Traspaso entrante'].drop_duplicates(['Fecha', 'Código']).shape[0],
                     x[x['Tipo'] == 'Traspaso saliente'].drop_duplicates(['Fecha', 'Código']).shape[0],
                     int(x["Vía"].isin(["Sin resolver", "Excluido"]).sum()))
        if not args.sin_hojas:
            sys.path.insert(0, str(BASE_DIR))
            from consolidador import agregar_hoja_df
            for loc in LOCALES_MOV:
                cand = sorted([*(BASE_DIR / "Reportes_Generados").glob(f"Analisis_Gestion_Comercial_{loc}_{mes:02d}_{args.anio}*.xlsx"),
                               *BASE_DIR.glob(f"Analisis_Gestion_Comercial_{loc}_{mes:02d}_{args.anio}*.xlsx")],
                              key=lambda p: p.stat().st_mtime, reverse=True)
                if not cand:
                    continue
                try:
                    agregar_hoja_df(cand[0], "Desperdicios y Traspasos (BOM)", det[det["Local"] == loc]
                                    .drop(columns=["Local", "Mes"]))
                    LOG.info("  hoja agregada en %s", cand[0].name)
                except Exception as exc:  # noqa: BLE001 - archivo abierto/bloqueado
                    LOG.warning("  no se pudo agregar la hoja en %s (%s)", cand[0].name, exc)
    return 0


if __name__ == "__main__":
    sys.exit(main())
