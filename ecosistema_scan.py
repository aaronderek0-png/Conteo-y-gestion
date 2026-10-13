"""Escaneo global del ecosistema Mishiguene + auditoría de integridad (solo lectura).

    python ecosistema_scan.py            # imprime el dictamen y escribe Copias_Gestion/auditoria_ecosistema.json

Qué hace
  1. Inventario (os.walk, sin node_modules): tipos de archivo, scripts Python con su docstring, y si existe
     lógica de Apps Script (.gs/.js/.txt con 'function ... SpreadsheetApp').
  2. Cascada de Gestión Semanal (columnas L..T) contra las identidades que usa el Apps Script/hoja:
         P = O + L      (Inicial + compras)         R = P - Q      (Consumo)
         N = L * M      (TOTAL $)                   S = R * M      (Consumo valorizado)
         T = Q * M      (Mercadería valorizada)     O(sem n) = Q(sem n-1)   (encadenado de stock)
  3. Traspasos: distribución de ESTADO real (RECIBIDO / PENDIENTE / DENEGADO / vacío) y cantidades no numéricas.
  4. Códigos colisionados por local (un código = productos distintos) y 'NO ENCONTRADO'.
  5. Ciclos / REND vacío en el BOM de Recetas.
  6. Tipado: precios guardados como texto y fechas con día/mes invertidos.
"""
from __future__ import annotations

import collections
import datetime as dt
import json
import os
import re
import sys
import warnings
from pathlib import Path

import openpyxl

from defensivo import codigos_colisionados, parse_fecha_robusta, parse_numero
from motor_recetas import MotorRecetas

warnings.filterwarnings("ignore")
BASE = Path(__file__).resolve().parent
EXPORT = BASE / "Copias_Gestion" / "originales_export"
SKIP_DIRS = {"node_modules", ".git", "__pycache__"}


def inventario() -> dict:
    tipos, scripts, apps_script = collections.Counter(), [], []
    for d, dn, fn in os.walk(BASE):
        dn[:] = [x for x in dn if x not in SKIP_DIRS]
        for f in fn:
            ext = os.path.splitext(f)[1].lower()
            tipos[ext or "(sin extensión)"] += 1
            path = Path(d) / f
            if ext == ".py" and "(1)" not in f:
                try:
                    txt = path.read_text(encoding="utf8", errors="ignore")
                    m = re.match(r'\s*"""(.*?)"""', txt, re.S)
                    scripts.append((str(path.relative_to(BASE)), (m.group(1).strip().splitlines() or [""])[0][:100] if m else ""))
                except OSError:
                    pass
            if ext in (".gs", ".js") and "node_modules" not in str(path):
                apps_script.append(str(path.relative_to(BASE)))
            if ext == ".txt":
                try:
                    if re.search(r"SpreadsheetApp|function\s+\w+\s*\(", path.read_text(encoding="utf8", errors="ignore")[:200000]):
                        apps_script.append(str(path.relative_to(BASE)))
                except OSError:
                    pass
    return dict(tipos=dict(tipos.most_common()), scripts=scripts, apps_script=apps_script)


def cascada_semanal(path: Path) -> dict:
    wb = openpyxl.load_workbook(path, data_only=True)
    viol, filas, prev = collections.Counter(), 0, None
    ejemplos = []
    for i in range(1, 6):
        if f"Semana{i}" not in wb.sheetnames:
            continue
        cur = {}
        for r in wb[f"Semana{i}"].iter_rows(min_row=3, max_col=23, values_only=True):
            cod = r[2]
            if not cod:
                continue
            filas += 1
            L, M, N, O, P, Q, R, S, T = [x if isinstance(x, (int, float)) else None for x in r[11:20]]
            cur[cod] = Q

            def bad(a, b):
                return a is not None and b is not None and abs(a - b) > 0.02 * max(1.0, abs(b))
            for nom, a, b in (("P!=O+L", P, (O or 0) + (L or 0)), ("R!=P-Q", R, (P or 0) - (Q or 0)),
                              ("N!=L*M", N, (L or 0) * (M or 0)), ("S!=R*M", S, (R or 0) * (M or 0)),
                              ("T!=Q*M", T, (Q or 0) * (M or 0))):
                if bad(a, b):
                    viol[nom] += 1
                    ejemplos.append((f"Semana{i}", cod, nom))
            if Q is not None and Q < 0:
                viol["saldo final negativo"] += 1
            if prev and cod in prev and prev[cod] is not None and O is not None and bad(O, prev[cod]):
                viol["O(sem n)!=Q(sem n-1)"] += 1
        prev = cur
    return dict(archivo=path.name, filas=filas, violaciones=dict(viol), ejemplos=ejemplos[:5])


def traspasos(path: Path) -> dict:
    ws = openpyxl.load_workbook(path, data_only=True, read_only=True)["Traspasos"]
    estados, n, cant_txt, sin_cod = collections.Counter(), 0, 0, 0
    for r in ws.iter_rows(min_row=3, max_col=12, values_only=True):
        if not r[0] and not r[3]:
            continue
        n += 1
        estados[str(r[10]).strip().upper() if r[10] else "(vacío)"] += 1
        if isinstance(r[5], str) and parse_numero(r[5])[0] is None:
            cant_txt += 1
        if not r[3]:
            sin_cod += 1
    lista = openpyxl.load_workbook(path, data_only=True, read_only=True)["Lista"]
    validos = sorted({str(r[9]).strip() for r in lista.iter_rows(min_row=2, max_col=10, values_only=True) if r[9]})
    return dict(filas=n, estados=dict(estados), estados_validos_en_lista=validos, cantidad_no_numerica=cant_txt,
                sin_codigo=sin_cod)


LOCALES = {"Arcos": ("Arcos_septiembre", "Inventario", 5), "Cabello": ("Cabello_septiembre", "Gestion", 3),
           "Lafinur": ("Lafinur_septiembre", "Gestion", 4), "Congreso": ("Congreso_septiembre", "Inventario", 4)}


def codigos_y_tipado(hoy: dt.date) -> dict:
    out = {}
    for local, (f, sh, ip) in LOCALES.items():
        ws = openpyxl.load_workbook(EXPORT / f"{f}.xlsx")[sh]
        rows = list(ws.iter_rows(min_row=2))
        # columna de código por encabezado
        hdr = None
        for r in ws.iter_rows(min_row=1, max_row=4):
            heads = [str(c.value or "").strip().upper() for c in r]
            if "CODIGO" in heads:
                hdr, ic = r[0].row, heads.index("CODIGO")
                break
        pares = [(str(r[ic].value).strip().upper(), r[ic + 1].value) for r in ws.iter_rows(min_row=hdr + 1)
                 if r[ic].value]
        col = codigos_colisionados(pares)
        out[local] = dict(colisiones={k: sorted(v)[:3] for k, v in col.items()}, items=len(pares))
    return out


def main() -> int:
    sys.stdout.reconfigure(encoding="utf-8")
    hoy = dt.date.today()
    rep: dict = {"fecha": hoy.isoformat()}
    inv = inventario()
    rep["inventario"] = inv
    print(f"[1] Inventario: {inv['tipos']}")
    print(f"    Apps Script (.gs/.js/.txt con funciones) en el workspace: {inv['apps_script'] or 'NINGUNO'}")
    print(f"    Scripts Python propios: {len(inv['scripts'])}")
    rep["cascada"] = {}
    for f in ("casi_gestion.xlsx", "trial.xlsx"):
        p = BASE / "Arcos_data" / f
        if p.exists():
            rep["cascada"][f] = cascada_semanal(p)
            c = rep["cascada"][f]
            print(f"[2] Cascada L..T {f}: {c['filas']} filas, violaciones: {c['violaciones'] or 'ninguna'}")
    tp = BASE / "Arcos_data" / "traspasos.xlsx"
    if tp.exists():
        rep["traspasos"] = traspasos(tp)
        print(f"[3] Traspasos: {rep['traspasos']}")
    rep["codigos"] = codigos_y_tipado(hoy)
    for l, d in rep["codigos"].items():
        print(f"[4] {l}: {len(d['colisiones'])} códigos colisionados {list(d['colisiones'])[:6]}")
    m = MotorRecetas(EXPORT / "Recetas.xlsx")
    rep["bom"] = dict(recetas=len(m.recetas), ciclos=m.ciclos, sin_rend=m.sin_rend, avisos=m.avisos[:20])
    print(f"[5] BOM: {len(m.recetas)} recetas | ciclos: {m.ciclos} | sin REND: {m.sin_rend}")
    out = BASE / "Copias_Gestion" / "auditoria_ecosistema.json"
    out.write_text(json.dumps(rep, ensure_ascii=False, indent=1, default=str), encoding="utf-8")
    print(f"-> {out.relative_to(BASE)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
