# Agente Mishiguene

Ecosistema de gestión de costos/inventario de Mishiguene (locales: Arcos, Congreso/Obrador, Lafinur, Cabello, Chimidoha).

## Piezas
- `GAS/` — 15 proyectos de Apps Script clonados con clasp. **Leer `GAS/ECOSISTEMA.md` primero.** Registro en `GAS/gas_registry.json`.
- `*.py` (raíz) — pipeline Python: `drive_connector`, `consolidador`, `comparador_precios`, `analisis_gestion_comercial`, `motor_recetas`, `defensivo`, `ecosistema_scan`.
- `Copias_Gestion/` — copias de las Gestiones y logs de auditoría; se trabaja sobre copias, no sobre originales.
- `WebTest/app.py` — dashboard Streamlit (a conectar con datos reales).
- `credentials.json` — Service Account (no subir a repos).

## Comandos
- `python gas_sync.py pull` — actualiza todo el código GAS desde Google.
- `python gas_sync.py push <carpeta>` — sube UN proyecto, con confirmación.
- clasp: `node_modules\.bin\clasp.cmd` (login ya hecho con aaron.derek0@gmail.com).

## Reglas
- Nunca `push` a GAS ni escribir en Sheets de producción sin que el usuario lo pida explícitamente; primero probar en copia.
- No imprimir ni commitear la API key/teléfono de CallMeBot ni `credentials.json`.
- El Maestro de Productos (`1lOabw9U…`) es la fuente única de SKU/rubros.

## Modo autónomo (cómo trabajar sin preguntar todo)
Autorizado sin consultar: leer todo, `gas_sync.py pull/status`, editar `GAS/**` y `*.py` localmente, correr `refactor_precios_gestion.py --dry-run/--apply` (escribe solo en `Copias_Gestion/*_COPIA.xlsx`), commits locales.
Siempre con confirmación explícita del usuario (bloqueado en `.claude/settings.json`): `clasp push/deploy`, `gas_sync.py push`, `git push`, escribir en Sheets de producción, tocar `Copias_Gestion/originales_export/`.

Flujo para cualquier cambio con efecto en datos:
1. `--dry-run` y mostrar el resumen.
2. Aplicar solo sobre `Copias_Gestion/<Local>_COPIA.xlsx` (read-back verifica cada celda escrita).
3. Revisar logs `Copias_Gestion/log_<Local>.csv` (filas REVISAR = no inyectadas, a decisión humana).
4. Recién entonces proponer al usuario el push a producción, listando exactamente qué cambia.
Variaciones de precio > ±50 % nunca se inyectan solas.

## Circuito de datos (cómo gobierna el Maestro a los locales)
Todo lo de esta sección está verificado contra los exports de `Arcos_data/` y `GAS/`.

**Maestro (`Arcos_data/maestro.xlsx`, ID `1lOabw9U…`).** Pestañas reales: `Productos` (SKU, Producto, Tipo, Rubro, Precio, UM, P.U),
`Presentacion` (SKU, Presentación/Artículo, Contenido, UM, Factor), `CONFIGURACION` (listas: Tipo, Rubro, Unidad, PROVEEDORES + ESTADO),
`Solicitudes` (cola de altas: Fecha, Local, ¿ = PRODUCTO|PROVEEDOR, Producto Base/Propuesto, Proveedor, Estado PENDIENTE/APROBADO/RECHAZADO).
Los nombres que se usan en conversación (`00_MAESTRO`, `PRODUCTOS`, `PRESENTACIONES`, `SOLICITUDES_PRODUCTOS`) corresponden a esas pestañas.
- Cada planilla de local (Arcos, Cabello, Lafinur, Congreso) trae por `IMPORTRANGE` el catálogo en su pestaña `Productos`
  (A=SKU, B=Producto base, **C=Presentación/"Producto Neto"**, D/E=Categoría 1/2, F=UN.PRESENT, G=UN.MEDIDA, H=CANTIDAD PRESENT) y los
  proveedores en `Listas!B4:B`.
- Esas dos listas son los desplegables de `Facturas`: validación de la col **D** = `Listas!B`, de la col **E** = `Productos!C`.
- Un producto/proveedor nuevo entra por `Solicitudes`; al pasar `Estado` (col L) a APROBADO, `GAS/maestro_productos/Aprobar SKU.js`
  genera el SKU siguiente y lo replica en las 4 planillas.

**Carga de una factura (pestaña `Facturas` de cada local).** Una fila por renglón:
`A`=N° factura · `B`=Fecha factura · `C`=Fecha ingreso · `D`=Proveedor (desplegable) · `E`=Producto/Presentación (desplegable) ·
`F`=libre/anotación · `G`=Cantidad compra · `H`=Precio · `I`=Descuento (**fracción**: 0.10 = 10 %).
Todo lo demás son fórmulas ya arrastradas (no se escriben): `J`=Total=G·H·(1−I) · `K:M`=UN.PRESENT/UN.MEDIDA/CANT.PRESENT (XLOOKUP a `Productos`) ·
`N`=Cantidad inventario=G·K·M · **`O`=Precio un. inventario (P_UN)=J/N** · `P:S`=mes/semana · `T:U`=categorías · `V`=alerta de cambio de precio ·
**`X`=SKU** (XLOOKUP de E en `Productos!C`, "SIN SKU" si no existe). La primera fila libre es la siguiente a la última con N°.
Los GAS `*_precios` leen `Facturas` y suben el último P_UN a `Inventario`/`Gestion`: **una fila mal cargada contamina el costo**.

**Sinonimia y recetas.** `Obrador == Congreso`: el Obrador es la planta de Congreso (`consolidador.LOCAL_PLANTA`), por eso un comprobante
"Obrador" se imputa en la planilla de Congreso. Recetas: `motor_recetas.py` explota recursivamente las elaboraciones — código `M\d`,
`MR\d`, `MO\d`, `OBR\d`, `P0\d` o nombre con `MADRE` — hasta insumos con precio (costo, factor y precio por kg a Gestión).
**Regla Verdulería:** solo Limón, Pepino, Berenjena, Tomate Perita y Cebolla se compran por bulto/cajón; el resto va por kilo.
En el catálogo todas están en KG, así que un bulto se convierte a kilos con el peso que diga el comprobante (`imputacion_facturas.BULTO_KG_DEFAULT`
está vacío: sin peso → REVISAR). Una verdura fuera de esa lista que llega por bulto también va a REVISAR.

## Motores de pricing e imputación (nuevos)
- `comparador_precios.py` — compara el Dossier contra las listas de proveedores. Descarta ofertas con `VENCIDA(S)`, `SOLO QUEDAN VENCIDAS`,
  `SIN STOCK`, `AGOTADO` (registro en `Precios/ofertas_descartadas.csv`); guardas de especie (`cabra/oveja/bufala` bidireccional;
  `cerdo/cordero/pollo/pavo/gallina/conejo` si el insumo lo nombra). `bufalo` NO es especie: es marca de fécula.
  Listas sin lector propio entran por `Precios/ingesta_listas.py` (xlsx/xls/csv/pdf/txt, normaliza a $/KG, $/LT, $/UN con el factor de
  empaque, caché SHA-256 en `Precios/cache_listas.json`). Carpeta de entrada `Precios/Listas_Mail/`; `--gmail` baja adjuntos por IMAP
  (`GMAIL_USER`/`GMAIL_APP_PASSWORD`). Reporte: `python comparador_precios.py` → `python enviador_mails.py --preview`.
- `invoice_parser.py` — `procesar_comprobante(path)`: PDF digital, PDF escaneado y fotos (`.jpg/.jpeg/.png/.webp`) → JSON
  `{local_destino, numero_factura, fecha_factura, fecha_ingreso, proveedor_detectado, items[...], total_comprobante, validacion}`.
  Pre-proceso de imagen (aplanado de luz, ruido, CLAHE, rotación/enderezado) + OCR (Tesseract → RapidOCR → EasyOCR; **hoy ninguno instalado**,
  ver abajo). Caché `cache_facturas.json`. Validación: suma de renglones = total ±1 % (acepta total con IVA 21/10,5/27).
- `imputacion_facturas.py` — JSON → filas de `Facturas`. Proveedor contra `Listas!B`; producto contra `Productos!C` por
  `aprendizaje_proveedores.json` → nombre idéntico → fuzzy con los filtros duros del comparador. **Regla:** confianza ≥ 0,90 y precio neto
  dentro de ±25 % del costo vigente → inyecta; si no → `REVISAR / NO CATALOGADO` en rojo (original y sugerencia en la col F) y registro
  para `Solicitudes`. Variación > ±50 % nunca se inyecta. Dry-run por defecto; `--apply` escribe solo en
  `Copias_Gestion/<Local>_FACTURAS_COPIA.xlsx` y `Copias_Gestion/Maestro_COPIA.xlsx` y verifica cada celda.
  Prueba sin comprobantes reales: `python imputacion_facturas.py --demo-historico 6 --local Arcos`.

## Pendientes conocidos
- Las `*_COPIA.xlsx` de `Copias_Gestion/` son Gestiones (Resumen/Inventario/Mov de stock): **no tienen pestaña `Facturas`**. Esa vive en
  `Arcos_data/{arcos,cabello,lafinur,congreso_facturas}.xlsx`; por eso la imputación usa `*_FACTURAS_COPIA.xlsx`. Chimidoha no tiene planilla de
  Facturas propia (su costo sale de la planilla de Lafinur/Chimi): sus comprobantes quedan a REVISAR hasta definirlo.
- OCR: instalar un motor (`choco install tesseract` + idioma `spa` + `pip install pytesseract`, o `easyocr`). `rapidocr-onnxruntime` no tiene
  wheel para Python 3.15.
- `aprendizaje_proveedores.json` se siembra con las solicitudes APROBADAS del Maestro y el historial de Facturas; crece con
  `imputacion_facturas.py --apply --aprender`.

