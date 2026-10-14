import os
import shutil
from pathlib import Path

os.chdir(r"C:\Users\EXO G12\Desktop\Agente Mishiguene")
SP = Path(os.environ["TEMP"]) / "claude" / "C--Users-EXO-G12-Desktop-Agente-Mishiguene" / "9189c61d-c6bb-43c1-bef9-9781623b448f" / "scratchpad"

# ---- rescate del scratchpad
dest = Path("Scripts Auxiliares")
dest.mkdir(exist_ok=True)
n = 0
for p in SP.iterdir():
    if p.is_file() and p.suffix in (".py", ".md"):
        shutil.copy2(p, dest / p.name)
        n += 1
tmp_wpp = Path(os.environ["TEMP"]) / "wpp"
if tmp_wpp.is_dir():
    for p in tmp_wpp.glob("*.py"):
        if not (dest / p.name).exists():
            shutil.copy2(p, dest / p.name)
(dest / "LEEME.txt").write_text(
    "Scripts auxiliares de la sesion de carga de facturas (rescatados del scratchpad temporal).\n"
    "lecturas.py      -> lectura visual de las fotos de WhatsApp (fuente de cache_facturas.json)\n"
    "bloque_wpp.py    -> bloque de reglas de local y lectores de PDF que se incorporo a invoice_parser.py\n"
    "bloque_imp.py    -> bloque de retro-entrenamiento/remitos/NC/prorrateo incorporado a imputacion_facturas.py\n"
    "multimodal_block.py -> motor multimodal inicial de invoice_parser.py\n"
    "claude_add.md    -> texto de CLAUDE.md del lote\n"
    "aplicar_cambios*.py, cierre.py -> parches aplicados (referencia historica)\n", encoding="utf-8")
print("scripts rescatados:", n)

# ---- CLAUDE.md
c = open("CLAUDE.md", encoding="utf-8").read()
c = c.replace("## Lote de WhatsApp (`Facturas_WPP/`): geografía, reglas y pipeline",
              "## Lote de WhatsApp (`Facturas Mishiguene/`): geografía, reglas y pipeline")
c = c.replace("Entrada: PDFs digitales", "Carpeta de entrada: **`Facturas Mishiguene/`** (sin guiones; `Facturas_WPP/` se lee solo si la nueva no existe). "
              "Lo que requiere revisión se refleja en **`Carpeta Revision/`** (`Revision Facturas IA.xlsx`).\nEntrada: PDFs digitales")
add = """
## Compra centralizada y precios de transferencia (Obrador → Chimidoha y locales)
**Ruteo.** Congreso (Obrador) **compra** la mercadería y luego se la "vende" internamente a Chimidoha y a los locales:
- Razón social `Lafinur 3368 SRL` o sello físico `Cerviño` → se imputa en **Congreso**, aunque el cliente diga Chimi Doha SRL
  (casos Kosta 1440 y Ecobaires 1610: quedan en Congreso con la alerta INFO de compra centralizada).
- Si el proveedor ya tiene historial cargado directamente en Chimidoha con razón social de Chimi → Chimidoha. Fallback: sello
  `CAFE MISHIGUENE NUÑEZ` o firma/anotación manuscrita `Núñez` / `Chimi`.
- `Sorrentino 58` (flete `O;A;L;C`): 100 % en Congreso con etiqueta `MULTILOCAL`.

**Matriz de recargos** (`transferencias_obrador.py`; la usan `imputacion_facturas.py` —cada fila inyectada en Congreso trae su precio
de transferencia— y `analisis_gestion_comercial.py`):
| Rubro | Precio de transferencia |
|---|---|
| Panificados y Descartables (insumos) | costo × 1.20 |
| Proteínas (Carnes y Pescados) | costo × 1.10 |
| Aceite e insumos normales de almacén | costo × 1.00 (al costo) |
| Elaborados por el Obrador (`Ventas Obrador`) | explosión completa por unidad (GR, UN, PORC, KG) contra las recetas (BOM, `motor_recetas.py`) |

**"Remitos Chimidoha"** (hoja/archivo que se carga a mano): `analisis_gestion_comercial.py` y `python transferencias_obrador.py` lo buscan
(nombres tipo `Remitos Chimidoha*.xlsx` en la raíz, `Arcos_data/`, `Facturas Mishiguene/`, `Copias_Gestion/`), calculan el recargo real
por rubro (precio remito / costo de compra en Congreso) y lo comparan con la matriz en `aprendizaje_transferencias.json`.

## Inyección pre-aprobada, color y aprobación
Toda fila que el agente inyecta (copias o `--push-drive`) va pintada en A:I: **`#FFF8E1`** amarillo suave = OK pendiente de aprobar,
**`#FFEBEE`** rojo suave = requiere revisión. `GAS/aprobador_facturas_ia.gs` agrega al menú de la planilla
`Facturas IA ▸ ✅ Aprobar Facturas IA (Quitar color)` (también "Aprobar TODAS las amarillas"). Hay que pegarlo en cada planilla real
(Extensiones ▸ Apps Script) y compartir la planilla con la Service Account como **Editor** para que `--push-drive` pueda escribir.

## Telegram: conexión
`python telegram_ingestor.py --conectar <TOKEN>`: guarda `TELEGRAM_BOT_TOKEN` en `.env`, detecta el grupo `Facturas Mishiguene`
(o el último grupo con actividad), guarda `TELEGRAM_CHAT_ID`, manda el mensaje de prueba y procesa lo pendiente en `Facturas Mishiguene/`.
El bot debe estar en el grupo y escribir algo antes (Group Privacy desactivada en @BotFather o bot admin). `.env` está en `.gitignore`.
"""
if "## Compra centralizada" not in c:
    c = c.rstrip("\n") + "\n" + add
open("CLAUDE.md", "w", encoding="utf-8").write(c)

g = open(".gitignore", encoding="utf-8").read()
for l in ["Facturas Mishiguene/", "Carpeta Revision/", "aprendizaje_transferencias.json", ".env"]:
    if l not in g:
        g = g.rstrip("\n") + "\n" + l
open(".gitignore", "w", encoding="utf-8").write(g + "\n")
print("CLAUDE.md y .gitignore actualizados")
