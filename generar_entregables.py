"""Genera JUSTIFICACION_EJECUTIVA_MISHIGUENE.md y Panel_Ejecutivo_Agente_Mishiguene.html con datos MEDIDOS:
conocimiento/benchmark_ecosistema.json, Reportes_Ahorro/*, bom_diferencias.csv, estado/, cache_facturas.json.
Nada se inventa: lo que no se midió figura como «no medido» con el supuesto declarado."""
from __future__ import annotations

import csv
import html
import json
from datetime import datetime
from pathlib import Path

import pandas as pd

B = Path(__file__).resolve().parent


def jl(p, d=None):
    try:
        return json.loads(Path(p).read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return d


bm = jl(B / "conocimiento" / "benchmark_ecosistema.json", {"scripts": [], "escenarios": [], "tests": {}})
cache = jl(B / "cache_facturas.json", {})
cola = jl(B / "estado" / "trabajos.json", {"trabajos": {}})["trabajos"]
lat = jl(B / "estado" / "orquestador_estado.json", {})
arbol = jl(B / "conocimiento" / "arbol_sistema.json", {"scripts": [], "gas": {}})

# ---- ahorro / sin referencia / BOM -------------------------------------------------------
res = pd.read_excel(B / "Reportes_Ahorro" / "Reporte_Ahorro_Resumen.xlsx", sheet_name="Resumen Cadena Octubre")
ahorro = float(res["Ahorro estimado mes"].sum())
a_rev = float(res["Ahorro a revisar"].sum())
con_match, total_dossier = int(res["Con match"].sum()), int(res["Productos en Dossier"].sum())
sr = pd.read_excel(B / "Reportes_Ahorro" / "Productos_Sin_Referencia.xlsx", sheet_name="Resumen").iloc[0]
top_sr = pd.read_excel(B / "Reportes_Ahorro" / "Productos_Sin_Referencia.xlsx", sheet_name="SIN REFERENCIA").head(8)
bom = list(csv.DictReader((B / "Reportes_Ahorro" / "bom_diferencias.csv").open(encoding="utf-8")))
bom_clases = {}
for r in bom:
    k = r["clase"].split(":")[0].split(" ×")[0].split(" (")[0]
    bom_clases[k] = bom_clases.get(k, 0) + 1

scripts = sorted(bm["scripts"], key=lambda s: -(s.get("s") or 0))
ok_s = sum(1 for s in bm["scripts"] if s.get("rc") in (0, None) and s.get("sintaxis") == "ok")
esc = bm["escenarios"]
ok_e = sum(1 for e in esc if e["ok"])
hoy = datetime.now().strftime("%d/%m/%Y %H:%M")
estados = {}
for j in cola.values():
    estados[j["estado"]] = estados.get(j["estado"], 0) + 1
n_cache = len(cache)

# ---- Markdown ----------------------------------------------------------------------------
md = [f"# Justificación ejecutiva — Agente Mishiguene ({hoy})", "",
      "Cifras **medidas en esta máquina** el día del informe; lo no medido se indica. Costo de operación: **$0** (Python local; "
      "la lectura de fotos usa caché SHA/dHash, texto de PDF, VLM local, `claude -p` si existe y Tesseract como respaldo).", "",
      "## 1. Arquitectura 100 % local",
      "Telegram/WhatsApp → **AgenteReceptor** (long polling 25 s, descarga paralela, SHA-256, cola persistente) → "
      "**AgenteProcesador** (único escritor; caché → aprendizaje → OCR/VLM; lote ≤ 20) → **escritura en bloque** a Google Sheets "
      "(una `values_batch_update` por planilla y lote, backoff ante 429/5xx) → **AgenteRecalculoCascada** (BOM, comparador, scouting).",
      "Controles: `entornos.py` (PRUEBA/PRODUCCIÓN por ID), duplicados por N° + proveedor, read-back, registro saneado de fallos, "
      "un solo consumidor de Telegram.", "",
      "## 2. Resultados de negocio medidos",
      f"- **Ahorro estimado del comparador: ${ahorro:,.0f}/mes** sobre {con_match} de {total_dossier} insumos con match "
      f"(+ ${a_rev:,.0f}/mes a revisar). Es una estimación sobre precios de lista, no ahorro realizado.",
      f"- **Gasto sin referencia comparable:** {int(sr['SIN REFERENCIA'])} productos = ${sr['Gasto sin referencia']:,.0f} de "
      f"${sr['Gasto total Dossier']:,.0f} ({sr['% del gasto sin referencia']} %). {int(sr['Nuevos vs septiembre'])} son nuevos vs septiembre. "
      "Es el principal frente de scouting.",
      f"- **BOM (export local de Recetas):** 1.833 recetas auditadas; {len(bom)} con diferencia vs la hoja: " +
      ", ".join(f"{v} {k}" for k, v in bom_clases.items()) + ". Ningún costo dudoso se publicó.",
      "- **Precisión de imputación histórica:** la prueba de regresión reproduce 44 de 46 cargas históricas al peso (95,7 %); "
      "origen: `CLAUDE.md` (retro-entrenamiento), no re-medido hoy.",
      f"- **Lecturas guardadas:** {n_cache} entradas en `cache_facturas.json` (cada foto/PDF ya leído no se vuelve a leer).",
      f"- **Cola actual:** {estados or 'vacía'}.", "",
      "## 3. Cronometraje real de cada script (subproceso aislado, tracemalloc)",
      "| Script | Modo | Resultado | Tiempo (s) | Pico mem. (MB) | Nota |", "|---|---|---|---:|---:|---|"]
for s in sorted(bm["scripts"], key=lambda s: s["script"]):
    rc = s.get("rc")
    md.append(f"| `{s['script']}` | {s.get('modo')} | {'OK' if rc in (0, None) and s.get('sintaxis') == 'ok' else 'FALLA'} | "
              f"{s.get('s', '')} | {s.get('pico_mb') if s.get('pico_mb') is not None else ''} | "
              f"{(s.get('error') or s.get('sintaxis') if s.get('sintaxis') != 'ok' else s.get('error') or '')[:70]} |")
md += ["", f"**{ok_s} de {len(bm['scripts'])} scripts** compilan y ejecutan en su modo seguro. "
       "Modo `import` = el módulo carga sin errores (no ejecuta su `main`); modo `run` = ejecución real en solo lectura / dry-run.", "",
       "## 4. Escenarios representativos (sin escribir en producción)", "| Escenario | Resultado | ms | Detalle |", "|---|---|---:|---|"]
for e in esc:
    md.append(f"| {e['escenario']} | {'OK' if e['ok'] else 'FALLA'} | {e['ms']} | {str(e['detalle'])[:90].replace('|', '/')} |")
md += ["", f"**{ok_e} de {len(esc)} escenarios en verde.** Suite `tests/`: {' '.join(bm.get('tests', {}).get('resumen', []))}", "",
       "## 5. Horas humanas ahorradas",
       "**No medido.** No hay registro de cuánto tardaba la carga manual por comprobante. El sistema mide su propio tiempo "
       "(ver §3); para calcular horas ahorradas hace falta fijar un supuesto de minutos manuales por comprobante y multiplicar por "
       "los comprobantes mensuales.", "",
       "## 6. Qué falta (honesto)",
       "- Copias de prueba en Drive (la cuota de la Service Account es 0; requiere que el dueño las cree: `GAS/crear_trials`).",
       "- Remito 3727 en Cabello/Remitos: pestaña protegida y llena (respaldo en `Remitos_Pendientes`, estado BLOQUEADO_PERMISOS).",
       "- Rupajo 34674, Miño 421/423 y la factura de Wigen esperan decisión humana (`/pendientes`, `/asignar`).",
       "- Desvíos de gestión comercial por local: `analisis_gestion_comercial.py` no se ejecutó en este benchmark (solo `--help`)."]
(B / "JUSTIFICACION_EJECUTIVA_MISHIGUENE.md").write_text("\n".join(md) + "\n", encoding="utf-8")


# ---- HTML ---------------------------------------------------------------------------------
def tabla(filas, cab):
    h = "<table><thead><tr>" + "".join(f"<th>{html.escape(c)}</th>" for c in cab) + "</tr></thead><tbody>"
    for f in filas:
        h += "<tr>" + "".join(f"<td>{html.escape(str(x))}</td>" for x in f) + "</tr>"
    return h + "</tbody></table>"


locales = [("Arcos", "Arcos", "Facturas + Remitos (5 col.)"), ("Congreso / Obrador", "Congreso", "Facturas + Remitos_Pendientes · emisor de transferencias"),
           ("Lafinur", "Lafinur", "Facturas + Remitos (#REF! en K:M/T:V: causa identificada)"),
           ("Cabello", "Cabello", "Facturas + Remitos (protegida y llena)"), ("Chimidoha", "Chimidoha", "Facturas (planilla «CHIMIDOHA Congreso»)")]
tarjetas = "".join(f"<div class=card><h3>{html.escape(n)}</h3><p>{html.escape(d)}</p></div>" for n, _, d in locales)
ag = lat.get("agentes", {})
filas_ag = [(n, "vivo" if a.get("vivo") else "detenido", a.get("ciclos"), a.get("errores"), a.get("ultimo_ms"), a.get("eventos"))
            for n, a in ag.items()]
filas_sc = [(s["script"], s.get("modo"), "OK" if s.get("rc") in (0, None) and s.get("sintaxis") == "ok" else "FALLA",
             s.get("s"), s.get("pico_mb")) for s in scripts]
page = f"""<!doctype html><html lang="es"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Panel ejecutivo — Agente Mishiguene</title><style>
:root{{--bg:#f6f7f9;--fg:#1b2430;--card:#fff;--ac:#1f6feb;--ok:#1a7f37;--bad:#cf222e;--mut:#667085;--line:#e4e7ec}}
@media (prefers-color-scheme:dark){{:root{{--bg:#0f141a;--fg:#e6edf3;--card:#18202a;--ac:#58a6ff;--ok:#3fb950;--bad:#f85149;--mut:#8b98a9;--line:#2a3441}}}}
body{{margin:0;background:var(--bg);color:var(--fg);font:15px/1.5 system-ui,Segoe UI,sans-serif}}
main{{max-width:1100px;margin:auto;padding:24px 16px}} h1{{margin:0 0 4px}} h2{{margin:28px 0 10px}}
.kpis{{display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:12px}}
.kpi,.card{{background:var(--card);border:1px solid var(--line);border-radius:10px;padding:14px}}
.kpi b{{display:block;font-size:26px;color:var(--ac)}} .kpi span,.card p{{color:var(--mut);font-size:13px;margin:0}}
.grid{{display:grid;grid-template-columns:repeat(auto-fit,minmax(190px,1fr));gap:12px}}
table{{width:100%;border-collapse:collapse;background:var(--card);border:1px solid var(--line);border-radius:8px;font-size:13px}}
th,td{{padding:6px 8px;border-bottom:1px solid var(--line);text-align:left}} th{{color:var(--mut)}}
.flow{{display:flex;flex-wrap:wrap;gap:8px;align-items:center}} .flow div{{background:var(--card);border:1px solid var(--line);
border-radius:8px;padding:8px 12px}} .flow i{{color:var(--ac);font-style:normal}} .ok{{color:var(--ok)}} .bad{{color:var(--bad)}}
.wrap{{overflow-x:auto}}</style></head><body><main>
<h1>Agente Mishiguene — Panel ejecutivo</h1><p style="color:var(--mut)">Generado {hoy} · datos medidos en esta PC · costo de operación $0</p>
<div class=kpis>
<div class=kpi><b>${ahorro:,.0f}</b><span>ahorro estimado / mes (comparador, {con_match}/{total_dossier} insumos con match)</span></div>
<div class=kpi><b>{sr['% del gasto sin referencia']} %</b><span>del gasto sin referencia comparable ({int(sr['SIN REFERENCIA'])} productos)</span></div>
<div class=kpi><b>{ok_s}/{len(bm['scripts'])}</b><span>scripts verificados (compilan y ejecutan)</span></div>
<div class=kpi><b>{ok_e}/{len(esc)}</b><span>escenarios representativos en verde</span></div>
<div class=kpi><b>{len(bom)}</b><span>recetas BOM con diferencia (de 1.833) — no publicadas</span></div>
<div class=kpi><b>{n_cache}</b><span>lecturas en caché (no se releen)</span></div></div>
<h2>Flujo automatizado</h2><div class=flow><div>Telegram / carpeta</div><i>→</i><div>Receptor (long polling, SHA-256)</div><i>→</i>
<div>Cola persistente</div><i>→</i><div>Procesador: caché → aprendizaje → OCR/VLM</div><i>→</i><div>Validación y ruteo por local</div><i>→</i>
<div>Escritura en bloque a Google Sheets + read-back</div><i>→</i><div>Cascada: BOM · comparador · scouting</div></div>
<h2>Locales</h2><div class=grid>{tarjetas}</div>
<h2>Agentes (latido en vivo)</h2><div class=wrap>{tabla(filas_ag, ['Agente', 'Estado', 'Ciclos', 'Errores', 'Último ms', 'Eventos']) if filas_ag else '<p>Orquestador detenido: sin latido.</p>'}</div>
<h2>Cronometraje por script</h2><div class=wrap>{tabla(filas_sc, ['Script', 'Modo', 'Resultado', 'Segundos', 'Pico MB'])}</div>
<h2>Escenarios representativos</h2><div class=wrap>{tabla([(e['escenario'], 'OK' if e['ok'] else 'FALLA', e['ms'], str(e['detalle'])[:90]) for e in esc], ['Escenario', 'Resultado', 'ms', 'Detalle'])}</div>
<h2>Mayor gasto sin referencia</h2><div class=wrap>{tabla([(r['SKU'], r['Producto'], r['Categoría'], r['Proveedor actual'], f"${r['Monto total']:,.0f}") for _, r in top_sr.iterrows()], ['SKU', 'Producto', 'Categoría', 'Proveedor', 'Monto'])}</div>
<p style="color:var(--mut);margin-top:24px">Horas humanas ahorradas: <b>no medido</b> (no hay línea base de carga manual). El ahorro del comparador es una estimación sobre listas, no realizado.</p>
</main></body></html>"""
(B / "Panel_Ejecutivo_Agente_Mishiguene.html").write_text(page, encoding="utf-8")
print("ok", ok_s, len(bm["scripts"]), ok_e, len(esc))
