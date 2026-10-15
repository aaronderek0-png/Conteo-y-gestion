"""Auto-consciencia del repositorio: grafo vivo en conocimiento/arbol_sistema.json.

Por cada `.py` (excepto node_modules, backup_previo, Scripts Auxiliares, __pycache__): imports locales, archivos de datos
que menciona (literales .xlsx/.json/.csv/.txt/.pdf), IDs de Sheets fijos, uso de gspread/escritura, CLI (argparse),
sintaxis (ast) y marca de modificación. Carpetas: conteo y tamaño. GAS: proyectos y sus openById/IDs.
Uso: python mapa_sistema.py   (también: python ecosistema_scan.py --arbol)
"""
from __future__ import annotations

import ast
import json
import re
from datetime import datetime
from pathlib import Path

BASE = Path(__file__).resolve().parent
SALIDA = BASE / "conocimiento" / "arbol_sistema.json"
EXCLUIR = {"node_modules", "backup_previo", "Scripts Auxiliares", "__pycache__", ".git", "originales_export"}
CARPETAS = [".", "GAS", "Precios", "Verduleria", "Arcos_data", "Copias_Gestion", "Carpeta Revision",
            "Datos_Operativos_Historicos", "Reportes_Ahorro", "WebTest", "tests", "estado", "conocimiento", "logs",
            "Facturas Mishiguene", "backups_recuperacion"]
RX_DATO = re.compile(r"""["']([\w .()/\\-]+\.(?:xlsx|xls|json|csv|txt|pdf|html|jsonl|bat))["']""", re.I)
RX_ID = re.compile(r"\b(1[A-Za-z0-9_-]{43})\b")


def analizar_py(p: Path) -> dict:
    txt = p.read_text(encoding="utf-8", errors="replace")
    info = {"ruta": p.relative_to(BASE).as_posix(), "lineas": txt.count("\n") + 1,
            "modificado": datetime.fromtimestamp(p.stat().st_mtime).isoformat(timespec="seconds")}
    try:
        tree = ast.parse(txt)
        info["sintaxis_ok"] = True
    except SyntaxError as e:
        info.update({"sintaxis_ok": False, "error": f"{e.msg} (línea {e.lineno})"})
        return info
    mods = set()
    for n in ast.walk(tree):
        if isinstance(n, ast.Import):
            mods.update(a.name.split(".")[0] for a in n.names)
        elif isinstance(n, ast.ImportFrom) and n.module and n.level == 0:
            mods.add(n.module.split(".")[0])
    locales = {q.stem for q in BASE.glob("*.py")} | {q.stem for q in (BASE / "Precios").glob("*.py")}
    info["depende_de"] = sorted(m for m in mods if m in locales and m != p.stem)
    info["lee_o_escribe_archivos"] = sorted(set(RX_DATO.findall(txt)))[:25]
    info["ids_sheets_fijos"] = len(set(RX_ID.findall(txt)))
    info["usa_gspread"] = "gspread" in txt
    info["escribe_en_sheets"] = bool(re.search(r"\.(update|batch_update|append_rows|values_batch_update|add_rows|add_worksheet|delete_rows)\(", txt))
    info["cli"] = "argparse" in mods or "sys.argv" in txt
    info["invocacion"] = f"python {p.name}" + (" --help" if info["cli"] else "")
    info["tiene_main"] = "__main__" in txt
    return info


def construir() -> dict:
    scripts = []
    for p in sorted(list(BASE.glob("*.py")) + list((BASE / "Precios").glob("*.py")) + list((BASE / "WebTest").glob("*.py"))
                    + list((BASE / "tests").glob("*.py"))):
        if EXCLUIR & set(p.parts):
            continue
        scripts.append(analizar_py(p))
    por = {s["ruta"]: s for s in scripts}
    usados_por: dict[str, list[str]] = {}
    for s in scripts:
        for d in s.get("depende_de", []):
            usados_por.setdefault(d, []).append(Path(s["ruta"]).stem)
    for r, s in por.items():
        s["usado_por"] = sorted(usados_por.get(Path(r).stem, []))
    carpetas = {}
    for c in CARPETAS:
        d = BASE / c if c != "." else BASE
        if not d.is_dir():
            carpetas[c] = {"existe": False}
            continue
        archivos = [f for f in (d.iterdir() if c == "." else d.rglob("*")) if f.is_file() and not (EXCLUIR & set(f.parts))]
        carpetas[c] = {"existe": True, "archivos": len(archivos), "mb": round(sum(f.stat().st_size for f in archivos) / 1e6, 2)}
    gas = {}
    for d in sorted((BASE / "GAS").glob("*")):
        if d.is_dir():
            t = "\n".join(f.read_text(encoding="utf-8", errors="replace") for f in d.glob("*.js"))
            gas[d.name] = {"openById": len(re.findall(r"openById\(", t)), "ids_fijos": len(set(RX_ID.findall(t))),
                           "importrange": "IMPORTRANGE" in t}
    return {"generado": datetime.now().isoformat(timespec="seconds"), "version": 1,
            "scripts": scripts, "carpetas": carpetas, "gas": gas,
            "datos_clave": {"planillas": "entornos.json / imputacion_facturas.IDS_PLANILLAS", "cola": "estado/trabajos.json",
                            "eventos": "logs/pipeline_eventos.jsonl", "conocimiento": "conocimiento/fuentes.json"}}


def main() -> dict:
    d = construir()
    SALIDA.parent.mkdir(exist_ok=True)
    tmp = SALIDA.with_suffix(".tmp")
    tmp.write_text(json.dumps(d, ensure_ascii=False, indent=1), encoding="utf-8")
    tmp.replace(SALIDA)
    malos = [s["ruta"] for s in d["scripts"] if not s.get("sintaxis_ok")]
    print(f"{len(d['scripts'])} scripts, {len(d['gas'])} proyectos GAS, sintaxis con error: {malos or 'ninguno'} -> {SALIDA.name}")
    return d


if __name__ == "__main__":
    main()
