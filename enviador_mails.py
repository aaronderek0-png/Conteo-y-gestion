"""Envía por mail los reportes de Reportes_Ahorro/ (generados por comparador_precios.py).

Reutiliza el envío SMTP (Gmail, SSL 465) y el diseño HTML del 'Reporte Pricing
Tri-Proveedor' de Precios/run_pipeline_pricing.py: una tabla matricial por
'Categoría 2' con el precio de CADA proveedor (El Criollo, Posta, El Nuevo
Emporio, Alyser, Villares, Nogada...), la Mejor Opción y el Ahorro Unitario contra
el Precio Ponderado del Dossier, con los insumos en orden alfabético (A-Z) y, si hay
variantes (marca/tamaño), una fila por Opción (el nombre del insumo va limpio; la
marca y presentación se leen en la celda de cada proveedor).
Adjunta todos los Excels.

Credenciales: variables de entorno, como el pipeline anterior. GMAIL_USER y
DESTINATARIO_MAIL tienen default; GMAIL_APP_PASSWORD es obligatoria (no se
guarda en el código):
    setx GMAIL_APP_PASSWORD "xxxx xxxx xxxx xxxx"   (una sola vez, luego abrir terminal nueva)

Uso:
    python enviador_mails.py                 # arma y envía
    python enviador_mails.py --preview       # solo guarda el HTML para revisarlo
    python enviador_mails.py --to otro@mail.com
"""

from __future__ import annotations

import argparse
import email.message
import html
import logging
import os
import re
import smtplib
import sys
import unicodedata
from datetime import datetime
from pathlib import Path

import pandas as pd

BASE_DIR = Path(__file__).resolve().parent
REPORTES_DIR = BASE_DIR / "Reportes_Ahorro"
RESUMEN = "Reporte_Ahorro_Resumen.xlsx"
PREVIEW_HTML = REPORTES_DIR / "preview_mail.html"
AHORRO_MINIMO_PCT = 3.0  # mismo umbral que comparador_precios.py

LOG = logging.getLogger("enviador_mails")


def _money(v: float) -> str:
    return f"${v:,.2f}" if pd.notna(v) else "-"


def _es_revisar(r: pd.Series) -> bool:
    return str(r.get("Confianza match", "")).startswith("Revisar")


def _es_oportunidad(r: pd.Series) -> bool:
    return pd.notna(r.get("Ahorro %")) and r["Ahorro %"] >= AHORRO_MINIMO_PCT and not _es_revisar(r)


# --------------------------------------------------------------------------- #
# Lectura de reportes
# --------------------------------------------------------------------------- #
def leer_reportes(carpeta: Path) -> dict[str, tuple[pd.DataFrame, list[str]]]:
    """{categoría: (matriz, proveedores)} de la hoja 'Comparativa' de cada
    Reporte_Ahorro_<Categoria>.xlsx. Los proveedores son las columnas que
    tienen su 'Match <proveedor>' al lado."""
    reportes = {}
    for path in sorted(carpeta.glob("Reporte_Ahorro_*.xlsx")):
        if path.name == RESUMEN or path.name.startswith("~$"):
            continue
        cat = path.stem.removeprefix("Reporte_Ahorro_").replace("_", " ")
        try:
            df = pd.read_excel(path, sheet_name="Comparativa")
        except ValueError:
            LOG.warning("%s no tiene hoja 'Comparativa' (formato viejo): volvé a correr comparador_precios.py", path.name)
            continue
        proveedores = [c.removeprefix("Match ") for c in df.columns
                       if c.startswith("Match ") and c.removeprefix("Match ") in df.columns]
        reportes[cat] = (df, proveedores)
    return reportes


# --------------------------------------------------------------------------- #
# HTML (diseño de run_pipeline_pricing.send_email_report)
# --------------------------------------------------------------------------- #
_BASE_RE = re.compile(r"\s*\[([^\]]*?)(?:, score \d+)?\]\s*$")


def _celda_proveedor(precio, match: str, ganador: bool, zebra: bool) -> str:
    """Precio del proveedor ('$1,234.00/KG') con el producto de la lista abajo."""
    match = "" if pd.isna(match) else str(match)
    base = _BASE_RE.search(match)
    desc = html.escape(_BASE_RE.sub("", match))
    bg = "background-color: #c8e6c9;" if ganador else ("background-color: #fafafa;" if zebra else "")
    if isinstance(precio, (int, float)) and pd.notna(precio):
        unidad = base.group(1).removeprefix("$") if base else ""
        unidad = unidad.replace("/envase ", "/")  # '$/envase 5 KG' -> '/5 KG'
        valor = f"<b>${precio:,.2f}{html.escape(unidad)}</b>"
    elif str(precio).strip() == "-" or pd.isna(precio) or not desc:
        return f'<td style="padding: 10px 8px; text-align: center; color: #9e9e9e; {bg}">-</td>'
    else:  # '≠ unidad': lo tiene, pero en otra presentación
        valor = f'<span style="color: #9e9e9e;">{html.escape(str(precio))}</span>'
    color = "#1b5e20" if ganador else "#333"
    return (f'<td style="padding: 10px 8px; text-align: right; color: {color}; {bg}">{valor}'
            f'<br><span style="font-size: 10px; color: #616161; font-weight: normal;">{desc}</span></td>')


def _clave_alfabetica(nombre) -> str:
    """'Álamo' ordena junto a 'alamo': sin tildes ni mayúsculas."""
    texto = unicodedata.normalize("NFKD", str(nombre)).encode("ascii", "ignore").decode()
    return re.sub(r"\s+", " ", texto).strip().lower()


def _filas_html(df: pd.DataFrame, proveedores: list[str]) -> str:
    rows = ""
    # Estrictamente A-Z por 'Producto (Dossier)' (la columna 'Insumo (Dossier)' de la tabla)
    df = df.sort_values("Producto (Dossier)", key=lambda col: col.map(_clave_alfabetica), kind="stable")
    for _, r in df.iterrows():
        mejor = str(r["Mejor Opción"])
        celdas = "".join(_celda_proveedor(r[p], r.get(f"Match {p}"), p == mejor, i % 2 == 0)
                         for i, p in enumerate(proveedores))

        if _es_revisar(r):
            nota = '<br><span style="font-size: 10px; color: #e65100; font-weight: normal;">revisar match (dif. &gt; 40%)</span>'
        elif pd.notna(r.get("Ahorro %")) and r["Ahorro %"] < AHORRO_MINIMO_PCT:
            nota = '<br><span style="font-size: 10px; color: #616161; font-weight: normal;">precio actual competitivo</span>'
        else:
            nota = ""
        if mejor == "UNIDAD NO COMPARABLE":
            mejor_html = '<span style="color: #9e9e9e;">Unidad no comparable</span>'
        else:
            mejor_html = html.escape(mejor) + nota

        ahorro = r.get("Ahorro Unitario")
        if pd.notna(ahorro) and ahorro > 0:
            ahorro_html = (f'{_money(ahorro)} ({r["Ahorro %"]}%)<br><span style="font-size: 10px; color: #616161; '
                           f'font-weight: normal;">mes: {_money(r["Ahorro estimado mes"])}</span>')
        else:
            ahorro_html = "-"

        rows += f"""
        <tr style="border-bottom: 1px solid #e0e0e0;">
            <td style="padding: 10px 8px; font-weight: bold; color: #1a237e;">{html.escape(str(r['Producto (Dossier)']))}</td>
            <td style="padding: 10px 8px; text-align: right;">
                <b>{_money(r['Precio Ponderado'])}</b><br><span style="font-size: 10px; color: #616161;">{html.escape(str(r['Proveedor actual']))}</span>
            </td>
            {celdas}
            <td style="padding: 10px 8px; font-weight: bold; color: #00695c;">{mejor_html}</td>
            <td style="padding: 10px 8px; text-align: right; color: #2e7d32; font-weight: bold;">{ahorro_html}</td>
        </tr>
        """
    return rows


def armar_html(reportes: dict[str, tuple[pd.DataFrame, list[str]]]) -> tuple[str, float]:
    # Categorías ordenadas por ahorro confirmado
    datos = []
    for cat, (df, proveedores) in reportes.items():
        # Solo la Opción 1 de cada insumo suma: las otras son alternativas del mismo consumo
        principal = df[df["Opción"].fillna(1).eq(1)] if "Opción" in df.columns else df
        subtotal = (principal.loc[principal.apply(_es_oportunidad, axis=1), "Ahorro estimado mes"].sum()
                    if not principal.empty else 0.0)
        datos.append((cat, df, proveedores, subtotal))
    datos.sort(key=lambda x: (x[3], len(x[1])), reverse=True)

    secciones, total = "", 0.0
    for cat, df, proveedores, subtotal in datos:
        if df.empty:
            continue
        total += subtotal
        th_prov = "".join(f'<th style="padding: 10px 8px; text-align: right;">{html.escape(p)}</th>' for p in proveedores)
        secciones += f"""
        <h3 style="color: #1F4E78; margin: 28px 0 8px; font-size: 15px;">{html.escape(cat)}
            <span style="color: #2e7d32; font-size: 13px;">&nbsp;·&nbsp;Ahorro estimado: {_money(subtotal)}</span></h3>
        <table style="border-collapse: collapse; width: 100%; max-width: 1400px; font-size: 12px;">
            <thead>
                <tr style="background-color: #1F4E78; color: white;">
                    <th style="padding: 10px 8px; text-align: left;">Insumo (Dossier)</th>
                    <th style="padding: 10px 8px; text-align: right;">Precio Actual</th>
                    {th_prov}
                    <th style="padding: 10px 8px; text-align: left;">Mejor Opción</th>
                    <th style="padding: 10px 8px; text-align: right;">Ahorro Unitario</th>
                </tr>
            </thead>
            <tbody>
                {_filas_html(df, proveedores)}
            </tbody>
        </table>
        """

    html_content = f"""
    <html>
    <body style="font-family: Arial, sans-serif; color: #333; margin: 20px;">
        <h2 style="color: #1F4E78; border-bottom: 2px solid #1F4E78; padding-bottom: 8px;">
            Comparativa de Precios Multi-Proveedor: Dossier Mensual vs Listas de Proveedores
        </h2>
        <p style="font-size: 13px; color: #555;">
            Precio Ponderado del Dossier contra el precio de cada proveedor, llevado a la misma unidad
            (<b>-</b> = el proveedor no tiene el producto; <b>≠ unidad</b> = lo tiene en una presentación no comparable).
            En verde, la mejor opción; si un insumo tiene variantes (marca, variedad o tamaño) cada una va en su
            fila (Opción 1, 2, 3; la marca y presentación se leen en cada proveedor) y el ahorro total
            cuenta solo la Opción 1. Ahorro estimado del mes confirmado: <b style="color: #2e7d32;">{_money(total)}</b>.
            Detalle completo, productos sin match y matches a revisar en los Excels adjuntos.
        </p>
        {secciones}
    </body>
    </html>
    """
    return html_content, total


# --------------------------------------------------------------------------- #
# Envío (run_pipeline_pricing.send_email_report)
# --------------------------------------------------------------------------- #
def enviar(user: str, password: str, recipient: str, html_content: str, adjuntos: list[Path]) -> None:
    msg = email.message.EmailMessage()
    msg["Subject"] = f"Reporte Pricing Multi-Proveedor - {datetime.now().strftime('%d/%m/%Y')}"
    msg["From"] = user
    msg["To"] = recipient
    msg.set_content("Adjunto encontrarás los reportes de precios por categoría.")
    msg.add_alternative(html_content, subtype="html")

    for path in adjuntos:
        msg.add_attachment(path.read_bytes(), maintype="application",
                           subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename=path.name)

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(user, password)
        server.send_message(msg)
    LOG.info("Mail enviado con éxito a %s (%d adjuntos).", recipient, len(adjuntos))


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--carpeta", default=str(REPORTES_DIR))
    ap.add_argument("--to", help="Destinatario (default: DESTINATARIO_MAIL o el propio GMAIL_USER)")
    ap.add_argument("--preview", action="store_true", help="No envía: guarda el HTML en <carpeta>/preview_mail.html")
    args = ap.parse_args()

    logging.basicConfig(level=logging.INFO, format="%(asctime)s | %(levelname)s | %(message)s", datefmt="%H:%M:%S")

    carpeta = Path(args.carpeta)
    reportes = leer_reportes(carpeta)
    if not reportes:
        LOG.error("No hay reportes en %s. Corré primero comparador_precios.py.", carpeta)
        return 1

    html_content, total = armar_html(reportes)
    LOG.info("%d categorías, %d productos, ahorro confirmado %s", len(reportes),
             sum(len(df) for df, _ in reportes.values()), _money(total))

    if args.preview:
        preview = carpeta / PREVIEW_HTML.name
        preview.write_text(html_content, encoding="utf-8")
        LOG.info("Preview guardado en %s", preview)
        return 0

    user = os.getenv("GMAIL_USER", "aaron.derek0@gmail.com")
    pw = os.getenv("GMAIL_APP_PASSWORD")
    if not pw:
        LOG.error("Falta la variable de entorno GMAIL_APP_PASSWORD (app password de Gmail). "
                  "Configurala una vez con: setx GMAIL_APP_PASSWORD \"xxxx xxxx xxxx xxxx\" y abrí una terminal nueva.")
        return 1
    dest = args.to or os.getenv("DESTINATARIO_MAIL", user)
    adjuntos = sorted(p for p in carpeta.glob("Reporte_Ahorro_*.xlsx") if not p.name.startswith("~$"))
    enviar(user, pw.replace(" ", ""), dest, html_content, adjuntos)
    return 0


if __name__ == "__main__":
    sys.exit(main())
