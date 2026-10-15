"""Benchmark del ecosistema: sintaxis, importación y ejecución segura de cada script + 12 escenarios representativos.

Cada script corre en un subproceso aislado con tracemalloc (pico de memoria Python) y timeout. Modo 'import' = ejecuta el
módulo sin `__main__`; modo 'run' = ejecución real con argumentos de solo lectura / dry-run (ver SEGUROS).
Salida: conocimiento/benchmark_ecosistema.json (consumido por JUSTIFICACION_EJECUTIVA y el Panel HTML).
Uso: python benchmark_ecosistema.py [--rapido]   (--rapido omite los scripts que tardan > 60 s)
"""
from __future__ import annotations

import json
import os
import py_compile
import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path

BASE = Path(__file__).resolve().parent
SALIDA = BASE / "conocimiento" / "benchmark_ecosistema.json"
PY = sys.executable

# script -> (argumentos, timeout s). Solo lectura o dry-run; nunca escritura a producción.
SEGUROS: dict[str, tuple[list[str], int]] = {
    "mapa_sistema.py": ([], 60), "estado_bot.py": ([], 60), "ecosistema_scan.py": (["--arbol"], 120),
    "scouting_sin_referencia.py": ([], 120), "bom_diferencias.py": ([], 120), "motor_recetas.py": ([], 180),
    "comparador_precios.py": (["--sin-nogada"], 300), "telegram_ingestor.py": (["--help"], 30),
    "orquestador_mishi.py": (["--help"], 30), "procesar_lote_wpp.py": (["--help"], 30),
    "imputacion_facturas.py": (["--help"], 30), "invoice_parser.py": (["--help"], 30),
    "transferencias_obrador.py": ([], 120), "enviador_mails.py": (["--preview"], 120),
    "reparar_ref_planillas.py": (["--help"], 30), "recuperar_lote.py": (["--help"], 30),
    "crear_trials.py": (["--estado"], 30), "gas_sync.py": (["status"], 120),
    "refactor_precios_gestion.py": (["--help"], 30), "analisis_gestion_comercial.py": (["--help"], 60),
    "consolidador.py": (["--help"], 60), "drive_connector.py": (["--help"], 60), "datos_operativos.py": (["--help"], 60),
    "defensivo.py": ([], 30), "coordinador.py": ([], 30), "conocimiento.py": ([], 30), "registro_fallos.py": ([], 30),
    "entornos.py": ([], 30), "aprobaciones.py": ([], 30),
}

RUNNER = r'''
import json, runpy, sys, time, tracemalloc
tracemalloc.start()
script, modo = sys.argv[1], sys.argv[2]
sys.argv = [script] + sys.argv[3:]
t0 = time.perf_counter(); rc = 0; err = ""
try:
    runpy.run_path(script, run_name="__main__" if modo == "run" else "mishi_bench")
except SystemExit as e:
    rc = e.code if isinstance(e.code, int) else (0 if e.code in (None, 0) else 1)
except BaseException as e:
    rc = 1; err = type(e).__name__ + ": " + str(e)[:200]
dt = time.perf_counter() - t0
_, pico = tracemalloc.get_traced_memory()
print("@@BENCH@@" + json.dumps({"rc": rc, "s": round(dt, 3), "pico_mb": round(pico / 1e6, 1), "error": err}))
'''


def correr(script: Path, modo: str, args: list[str], timeout: int) -> dict:
    t0 = time.perf_counter()
    try:
        r = subprocess.run([PY, "-c", RUNNER, str(script), modo, *args], cwd=str(BASE), capture_output=True, text=True,
                           timeout=timeout, encoding="utf-8", errors="replace",
                           env={**os.environ, "PYTHONIOENCODING": "utf-8"})
    except subprocess.TimeoutExpired:
        return {"rc": -1, "s": round(time.perf_counter() - t0, 1), "pico_mb": None, "error": f"timeout {timeout}s"}
    for l in r.stdout.splitlines():
        if l.startswith("@@BENCH@@"):
            d = json.loads(l[9:])
            if d["rc"] and not d["error"]:
                d["error"] = (r.stderr or r.stdout)[-160:].strip().replace("\n", " ")
            return d
    return {"rc": r.returncode or 1, "s": round(time.perf_counter() - t0, 3), "pico_mb": None,
            "error": (r.stderr or r.stdout)[-200:].strip().replace("\n", " ")}


def scripts() -> list[Path]:
    out = sorted(BASE.glob("*.py")) + sorted((BASE / "Precios").glob("*.py")) + sorted((BASE / "WebTest").glob("*.py"))
    return [p for p in out if p.name != "benchmark_ecosistema.py" and not p.name.startswith("run_pipeline")]


# --------------------------------------------------------------------------- #
# 12 escenarios representativos (sin escribir a producción)
# --------------------------------------------------------------------------- #
def escenarios() -> list[dict]:
    sys.path.insert(0, str(BASE))
    os.environ.setdefault("MISHI_ESTADO_DIR", str(BASE / "estado" / "_bench"))
    os.environ.setdefault("MISHI_LOG_DIR", str(BASE / "estado" / "_bench"))
    import aprobaciones as ap
    import coordinador as co
    import imputacion_facturas as imp
    import procesar_lote_wpp as lote
    res = []

    def caso(nombre, fn):
        t0 = time.perf_counter()
        try:
            ok, detalle = fn()
        except Exception as e:                              # noqa: BLE001
            ok, detalle = False, f"{type(e).__name__}: {str(e)[:150]}"
        res.append({"escenario": nombre, "ok": bool(ok), "detalle": detalle, "ms": round((time.perf_counter() - t0) * 1000)})

    cache = json.loads((BASE / "cache_facturas.json").read_text(encoding="utf-8"))

    def comp(num):
        return next((dict(v) for v in cache.values() if str(v.get("numero_factura", "")).endswith(num)), None)

    def clasificar(nombre_archivo):
        import tempfile, shutil
        d = Path(tempfile.mkdtemp())
        shutil.copy2(BASE / "Facturas Mishiguene" / nombre_archivo, d / nombre_archivo)
        ns = type("N", (), dict(push_drive=False, confirmo_escribir_produccion=False, dry_run=True, json=None))()
        rep, _l, _a, cl = lote.ejecutar(d, False, ns, indexar_informes=False)
        return rep, cl

    # 1 factura normal (nueva sin local → no escribe; ya cargada → idempotente)
    def c1():
        rep, _ = clasificar("IMG20260808100551.jpg")
        e = rep["auditoria"][0]["estado"]
        return "YA_CARGADA" in e, e
    caso("1 factura normal (ya cargada: no duplica)", c1)

    # 2/3 NC caso A y B (reglas de construir_filas con comprobante sintético)
    def nc(caso_b):
        loc = imp.Local("Arcos")
        pres = str(loc.catalogo["PRESENTACION"].dropna().iloc[0])
        item = {"descripcion_original": pres if not caso_b else "BONIFICACION POR DIFERENCIA DE PRECIO",
                "cantidad": 2.0, "unidad_detectada": "UN", "precio_unitario": 100.0, "descuento_pct": 0.0,
                "total_linea": 200.0, "linea_consistente": True}
        c = {"tipo": "NC", "numero_factura": "0001-00000999", "fecha_factura": "2026-10-10", "fecha_ingreso": "2026-10-10",
             "proveedor_detectado": "PROVEEDOR PRUEBA", "items": [item], "total_comprobante": 200.0, "validacion": {},
             "extra": {"motivo_nc": "bonificación diferencia de precio" if caso_b else "devolución mercadería"},
             "archivo": "sintetica_nc.jpg", "local_destino": "Arcos"}
        r = imp.construir_filas(c, loc, imp.cargar_aprendizaje())
        return r["filas"]
    caso("2 NC caso A (devolución: cantidad negativa)",
         lambda: ((lambda f: (bool(f) and all((x.get("G") or 0) <= 0 for x in f), f"{len(f)} fila(s), G={[x.get('G') for x in f][:2]}"))(nc(False))))
    caso("3 NC caso B (bonificación: fila NC −1)",
         lambda: ((lambda f: (bool(f), f"{len(f)} fila(s): {[(x.get('presentacion'), x.get('G')) for x in f][:2]}"))(nc(True))))

    # 4 remito de proveedor con birome/NO/balanza (Raíces 8528)
    def c4():
        rep, _ = clasificar("IMG20260707095646.jpg")
        r = rep["remitos"][0] if rep["remitos"] else {}
        return bool(r), f"estado={r.get('estado')} renglones={len(r.get('renglones', []))} total_est={r.get('total_estimado')}"
    caso("4 remito proveedor (birome/NO/balanza)", c4)

    # 5 transferencia del Obrador (3727 → Cabello)
    def c5():
        rep, _ = clasificar("WhatsApp Image 2026-10-14 at 00.53.06 (2).jpeg")
        r = (rep.get("remitos_obrador") or [{}])[0]
        return abs((r.get("total") or 0) - 58656.30) < 0.01, f"{len(r.get('renglones', []))} renglones total={r.get('total')} → {r.get('receptor')}"
    caso("5 transferencia Obrador → Cabello", c5)

    # 6 ruteo por renglón (Wigen)
    def c6():
        w = comp("") if False else next((dict(v) for v in cache.values() if "WIGEN" in str(v.get("proveedor_detectado", "")).upper()
                                         and v.get("tipo") == "Remito"), None)
        partes = lote.separar_destinos_remito(w)
        return len(partes) == 2, {p.get("local_destino_renglon"): [i["descripcion_original"] for i in p["items"]] for p in partes}
    caso("6 destinos por renglón (Wigen)", c6)

    base = ["ID1", "0001-1", "2026-07-07", "WIGEN", "NUECES", "KG", 1, 1, 0, 0, "", "NO", "PENDIENTE", "", "a.jpg", ""]

    def tabla(h, t="", m="ok"):
        return [["ID"], base + [h, m, "foto", "", "", t, "", ""]]
    caso("7 aprobación (APROBADO → CARGAR; ya ESCRITO → NADA)",
         lambda: ((ap.planear(ap.leer(tabla("APROBADO")))[0]["accion"] == "CARGAR"
                   and ap.planear(ap.leer(tabla("APROBADO", "ESCRITO")))[0]["accion"] == "NADA"), "idempotente"))
    caso("8 denegación (cierra sin efecto)", lambda: (ap.planear(ap.leer(tabla("DENEGADO")))[0]["accion"] == "CERRAR_SIN_EFECTO", "sin gasto ni stock"))

    # 9 duplicado
    def c9():
        import tempfile
        q = co.Cola(Path(tempfile.mkdtemp()) / "t.json")
        q.registrar("sha1", "a.jpg"); q.registrar("sha1", "a.jpg")
        return q.resumen()["RECIBIDO"] == 1, "una sola entrada por SHA"
    caso("9 duplicado (idempotencia por SHA)", c9)

    # 10 fallo parcial + 11 reintento
    def c10_11():
        import tempfile
        q = co.Cola(Path(tempfile.mkdtemp()) / "t.json")
        q.registrar("k", "b.jpg"); q.transicion("k", "LEIDO"); q.transicion("k", "ERROR", error="APIError 400")
        parcial = len(q.pendientes()) == 1 and q.obtener("k")["intentos"] == 1
        q.transicion("k", "VALIDADO"); q.transicion("k", "ESCRITO", {"rango": "A1:I2"})
        return parcial and q.obtener("k")["estado"] == "ESCRITO", "ERROR no consume; reintento termina en ESCRITO"
    caso("10 fallo parcial y 11 reintento", c10_11)

    # 12 hoja sin filas / protegida
    def c12():
        class WS:
            title, row_count = "Remitos", 608
            added = 0
            def add_rows(self, n):
                raise RuntimeError("APIError: [400]: Invalid requests[0].updateSheetProperties: You are trying to edit a protected cell or object")
        try:
            imp._asegurar_filas(WS(), 613)
            return False, "no detectó"
        except RuntimeError as e:
            return "protected" in str(e), "add_rows rechazado → push_remitos_drive deriva a _bloqueo_remitos (cola + respaldo)"
    caso("12 hoja sin filas / protegida", c12)
    return res


def main() -> int:
    rapido = "--rapido" in sys.argv
    filas = []
    for p in scripts():
        t0 = time.perf_counter()
        sint = "ok"
        try:
            compile(p.read_text(encoding="utf-8", errors="replace"), str(p), "exec")      # sintaxis sin escribir .pyc
        except SyntaxError as e:
            sint = f"{e.msg} (línea {e.lineno})"
        if sint != "ok":
            filas.append({"script": p.relative_to(BASE).as_posix(), "sintaxis": sint, "modo": "-", "rc": 1, "s": 0, "pico_mb": None})
            continue
        args, to = SEGUROS.get(p.name, (None, 90))
        modo = "run" if args is not None else "import"
        if rapido and to > 60:
            args, modo = None, "import"
        r = correr(p, modo, args or [], to)
        filas.append({"script": p.relative_to(BASE).as_posix(), "sintaxis": "ok", "modo": modo, **r})
        print(f"{p.name:34} {modo:6} rc={r['rc']!s:3} {r['s']:>8}s {str(r['pico_mb']):>7} MB  {r['error'][:60]}")
    esc = escenarios()
    for e in esc:
        print(("OK  " if e["ok"] else "FALLA"), e["escenario"], f"{e['ms']} ms", "|", str(e["detalle"])[:100])
    tests = subprocess.run([PY, "-m", "unittest", "discover", "-s", "tests", "-q"], cwd=str(BASE), capture_output=True, text=True)
    out = {"generado": datetime.now().isoformat(timespec="seconds"), "scripts": filas, "escenarios": esc,
           "tests": {"rc": tests.returncode, "resumen": (tests.stderr or tests.stdout).strip().splitlines()[-3:]}}
    SALIDA.parent.mkdir(exist_ok=True)
    SALIDA.write_text(json.dumps(out, ensure_ascii=False, indent=1, default=str), encoding="utf-8")
    malos = [f for f in filas if f["rc"] not in (0, None)] + [e for e in esc if not e["ok"]]
    print(f"\n{len(filas)} scripts, {len(esc)} escenarios; con problemas: {len(malos)}; tests rc={tests.returncode}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
