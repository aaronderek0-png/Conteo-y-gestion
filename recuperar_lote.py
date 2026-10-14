"""Recuperación acotada de los 6 adjuntos que fallaron con APIError (14/10/2026).

Dos modos, nada más:
  python recuperar_lote.py plan   -> SOLO LECTURA. Lee las planillas vivas (col A/D de Facturas, fila 2 de fórmulas),
                                     arma el plan exacto y lo guarda en recuperacion_plan.json (con huella). No escribe.
  python recuperar_lote.py apply  -> exige que el plan actual coincida con recuperacion_plan.json; respalda A:I de cada
                                     destino en backups_recuperacion/, escribe SOLO filas INYECTAR (las REVISAR quedan
                                     fuera y van a Carpeta Revision), re-lee cada celda y registra en logs/.

Lista de documentos fija (no recorre carpetas): ver ADJUNTOS. No hace push de git ni toca GAS ni credenciales.
"""
from __future__ import annotations

import hashlib
import json
import shutil
import sys
import tempfile
from datetime import datetime
from pathlib import Path
from types import SimpleNamespace

BASE = Path(__file__).resolve().parent
sys.path.insert(0, str(BASE))

import imputacion_facturas as imp          # noqa: E402
import procesar_lote_wpp as lote           # noqa: E402
import registro_fallos as rf               # noqa: E402

ENTRADA = BASE / "Facturas Mishiguene"
ADJUNTOS = ["IMG20260828153210.jpg", "IMG20260819152602.jpg", "IMG20260808100551.jpg",
            "IMG20260804133616.jpg", "IMG20260707124118.jpg", "IMG20260725120055.jpg"]
PLAN = BASE / "recuperacion_plan.json"
BACKUPS = BASE / "backups_recuperacion"


def construir() -> tuple[dict, dict]:
    """(reporte, filas_por_local). Lee SOLO de la caché de lecturas; sin lectura previa no invoca visión (dry)."""
    with tempfile.TemporaryDirectory() as tmp:
        for n in ADJUNTOS:
            shutil.copy2(ENTRADA / n, Path(tmp) / n)
        ns = SimpleNamespace(push_drive=False, confirmo_escribir_produccion=False, dry_run=True, json=None)
        reporte, _loc, _ap, cl = lote.ejecutar(Path(tmp), False, ns, indexar_informes=False)
    return reporte, cl.get("_plan_por_local", {})


def _hoja(local: str):
    gc, _ = imp._gspread()
    sh = gc.open_by_key(imp.IDS_PLANILLAS[local])
    return sh, imp._hoja_por_clave(sh, "factura")


def previsto(local: str, filas: list[dict]) -> dict:
    """Lectura viva (solo lectura): duplicados y primera fila libre, sin escribir."""
    sh, ws = _hoja(local)
    col_a, col_d = ws.col_values(1), ws.col_values(4)
    existentes = {(imp.ultimo_entero(a), imp.norm(d)) for a, d in zip(col_a[1:], (col_d + [""] * len(col_a))[1:]) if a}
    cols = imp._columnas_de_carga(ws)
    ok, revisar, dup = [], [], []
    for f in filas:
        fila = imp._fila_push(f)
        if (imp.ultimo_entero(fila[0]), imp.norm(fila[3])) in existentes:
            dup.append(fila)
        elif f.get("accion") == "INYECTAR":
            ok.append(fila)
        else:
            revisar.append(fila + [f.get("motivos")])
    return {"planilla": sh.title, "hoja": ws.title, "ultima_fila_actual": len(col_a), "primera_libre": len(col_a) + 1,
            "columnas_manuales": [imp._col_letra(i) for i in cols["manuales"]],
            "columnas_con_formula_respetadas": [imp._col_letra(i) for i in cols["con_formula"]],
            "sku_col": imp._col_letra(cols["sku_col"]) if cols["sku_col"] is not None else None,
            "escribir": ok, "fuera_por_revisar": revisar, "ya_cargadas": dup}


def huella(plan: dict) -> str:
    base = {l: [p["planilla"], p["primera_libre"], p["escribir"]] for l, p in plan.items()}
    return hashlib.sha256(json.dumps(base, sort_keys=True, default=str).encode()).hexdigest()[:16]


def hacer_plan() -> dict:
    reporte, por_local = construir()
    plan = {l: previsto(l, fs) for l, fs in por_local.items() if l in imp.IDS_PLANILLAS and fs}
    sin_destino = [a for a in reporte["auditoria"] if a["archivo"].split(" [")[0] in ADJUNTOS and not a.get("filas_imputadas")]
    salida = {"generado": datetime.now().isoformat(timespec="seconds"), "huella": huella(plan), "plan": plan,
              "sin_filas_a_imputar": [{k: a.get(k) for k in ("archivo", "tipo", "numero", "proveedor", "local", "estado")}
                                      for a in sin_destino]}
    PLAN.write_text(json.dumps(salida, ensure_ascii=False, indent=1, default=str), encoding="utf-8")
    print(f"Plan guardado en {PLAN.name} (huella {salida['huella']}). NO se escribió nada.")
    for l, p in plan.items():
        print(f"\n== {l}: {p['planilla']} / pestaña {p['hoja']} · última fila {p['ultima_fila_actual']} · "
              f"escribiría desde la fila {p['primera_libre']} en {', '.join(p['columnas_manuales'])}"
              f" (respeta fórmulas en {', '.join(p['columnas_con_formula_respetadas']) or '—'})")
        for i, fila in enumerate(p["escribir"]):
            print(f"  fila {p['primera_libre'] + i}: " + " | ".join("" if x is None else str(x) for x in fila))
        for fila in p["fuera_por_revisar"]:
            print(f"  [NO se escribe: REVISAR] N° {fila[0]} {fila[3]} · {fila[5][:90]}")
        for fila in p["ya_cargadas"]:
            print(f"  [ya estaba cargada] N° {fila[0]} {fila[3]}")
    for s in salida["sin_filas_a_imputar"]:
        print(f"  [sin filas] {s['archivo']} · {s['tipo']} · {s['estado']}")
    return salida


def aplicar() -> None:
    if not PLAN.exists():
        sys.exit("Falta recuperacion_plan.json: corré primero `plan`.")
    guardado = json.loads(PLAN.read_text(encoding="utf-8"))
    reporte, por_local = construir()
    plan_vivo = {l: previsto(l, fs) for l, fs in por_local.items() if l in imp.IDS_PLANILLAS and fs}
    if huella(plan_vivo) != guardado["huella"]:
        sys.exit("La planilla o la lectura cambió desde el plan aprobado: volvé a correr `plan` y aprobar de nuevo.")
    BACKUPS.mkdir(exist_ok=True)
    sello = datetime.now().strftime("%Y%m%d_%H%M%S")
    for local, p in plan_vivo.items():
        filas = [f for f in por_local[local] if f.get("accion") == "INYECTAR"]
        if not p["escribir"]:
            print(f"{local}: nada para escribir")
            continue
        sh, ws = _hoja(local)
        (BACKUPS / f"{local}_{sello}_Facturas_AI.json").write_text(
            json.dumps(ws.get("A1:I" + str(p["ultima_fila_actual"])), ensure_ascii=False), encoding="utf-8")
        try:
            r = imp.push_drive(local, filas, confirmado=True)
        except Exception as e:                                   # noqa: BLE001
            rf.evento(f"recuperar:{local}", "push_facturas", "ERROR", error=e, local=local)
            print(f"{local}: ERROR {rf.detalle_excepcion(e)}")
            continue
        if not r.get("escrito"):
            print(f"{local}: no escribió ({r.get('motivo')})")
            continue
        leido = ws.get(f"A{r['desde_fila']}:I{r['hasta_fila']}", value_render_option="UNFORMATTED_VALUE")
        buenos = sum(1 for d, fila in zip(r["detalle"], leido)
                     if str(fila[0]) == str(d["N"]) and imp.norm(fila[3] if len(fila) > 3 else "") == imp.norm(d["proveedor"]))
        rf.evento(f"recuperar:{local}", "push_facturas", "OK", local=local, filas=r["filas"], verificadas=buenos)
        print(f"{local}: escribí filas {r['desde_fila']}-{r['hasta_fila']} en «{r['planilla']}»/{r['hoja']}; "
              f"read-back {buenos}/{r['filas']} · backup {local}_{sello}")
        (BACKUPS / f"{local}_{sello}_resultado.json").write_text(
            json.dumps(r, ensure_ascii=False, indent=1, default=str), encoding="utf-8")


if __name__ == "__main__":
    modo = sys.argv[1] if len(sys.argv) > 1 else "plan"
    if modo == "plan":
        hacer_plan()
    elif modo == "apply":
        aplicar()
    else:
        sys.exit(__doc__)
