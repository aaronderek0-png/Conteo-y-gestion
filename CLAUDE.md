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
  `Arcos_data/{arcos,cabello,lafinur,congreso_facturas}.xlsx`; por eso la imputación usa `*_FACTURAS_COPIA.xlsx`. Chimidoha no tiene planilla real de
  Facturas: su copia se crea desde la de Congreso (ver sección del lote de WhatsApp).
- OCR: Tesseract ya instalado (ver sección del lote). `rapidocr-onnxruntime` no tiene wheel para Python 3.15.
- `aprendizaje_proveedores.json` se siembra con las solicitudes APROBADAS del Maestro y el historial de Facturas; crece con
  `imputacion_facturas.py --apply --aprender`.

## Lote de WhatsApp (`Facturas_WPP/`): geografía, reglas y pipeline
Ejecutar: `python procesar_lote_wpp.py` (escribe SOLO en copias; `--dry-run` no escribe; es idempotente).
Entrada: PDFs digitales (se usa solo la copia ORIGINAL; ignora DUPLICADO/TRIPLICADO), fotos `.jpeg` de WhatsApp (leídas con visión
multimodal y guardadas en `cache_facturas.json` por SHA-256; sin lectura previa se usa Tesseract `spa` y todo queda a REVISAR),
informes Maxirest (`Datos_Operativos_Historicos/`) y listas de precios.

**Local de destino (`invoice_parser.detectar_local`).** Prioridad: sello/RECIBIDO manuscrito → anotación manuscrita del local →
dirección de entrega → razón social propia → historial (N° ya cargado en una planilla).
- **Trampa fiscal:** `LAFINUR 3368 SRL` (CUIT 30-71456509-1) es la razón social de los 4 locales (Arcos, Congreso, Cabello, Lafinur).
  Por sí sola **no** asigna Lafinur: sin sello ni dirección de entrega el local queda INDETERMINADO.
- **Edificio Av. Congreso esq. Montañeses 1602 (Núñez):** Planta Baja = café **Chimidoha** (Café Mishiguene Núñez, razón social
  Chimi Doha SRL, CUIT 30-71567213-4); Planta Alta = Centro de Producción **Congreso = Obrador**.
  Sello `CAFE MISHIGUENE NUÑEZ - RECIBIDO` → Chimidoha. Sello `MISHIGUENE CERVIÑO - RECIBIDO` (o "CERVIÑO RECIBIDO" manuscrito) → Congreso.
  Dirección `CONGRESO 1602` + Lafinur 3368 SRL sin sello de Núñez → Congreso. Dirección Cerviño 4417 → Congreso.
- Sello `CAFE MISHIGUENE CABELLO` / entrega `Cabello 3181` → Cabello · `ARCOS` / calle Arcos → Arcos · `LAFINUR`/`LAFI` → Lafinur.
- Alerta si la razón social es Chimi Doha pero el sello es Cerviño (verificar de quién es el gasto) y si se cargó en una planilla
  distinta de la que indica el sello (`REASIGNAR A …`: se marca en la col F de la copia, no se borra).
- **Chimidoha no tiene planilla real**: `Copias_Gestion/Chimidoha_FACTURAS_COPIA.xlsx` se crea desde la de Congreso (misma estructura,
  Facturas vacía). Sus precios vigentes se toman de Congreso.

**Modo de carga por proveedor** (`aprendizaje_proveedores.json` → `modos_proveedor`). `AGRUPADO_TOTAL` = una sola fila por factura con
el neto total y un concepto genérico: Raíces y Open Fruit → `Verduleria` (171/173 y 154/159 facturas), Aguilera → `LAVADERO`, Sorrentino
→ `Flete`. El resto va renglón a renglón. Un único renglón ya aprendido por nombre (p. ej. `Huevo cajon`) va por su nombre.
El retro-entrenamiento cruza cada comprobante con la col A de las 4 planillas (N° + proveedor/fecha ±4 días + monto) y aprende
`descripción del proveedor → Presentación (col E)`; la prueba de regresión reproduce 44 de 46 cargas históricas al peso.

**Remitos con birome** (`imputacion_facturas.procesar_remito`). `NO` escrito a mano o tachado → `cantidad_real = 0`; kilos de balanza
anotados al pie se vinculan al renglón impreso (`peso_real_recibido_kg`, `costo_real_por_kg`). Se guardan en `remitos_pendientes.json`
y en la pestaña `Remitos_Pendientes` de la copia. Al llegar la factura que cita el N° de remito se concilia (renglones cobrados que
estaban anulados, kilos facturados de más) y el remito pasa a `CONCILIADO`. Si la factura ya está en la planilla como suma agrupada, no
se insertan filas de detalle (se evita duplicar el gasto). Bultos (`BULTO_KG_DEFAULT`): tomate perita 16 kg, limón 16, berenjena 9,
cebolla 17, pepino 13.

**Notas de crédito.** En las planillas van con **cantidad negativa y precio positivo**. Caso A (devolución física / faltante / error
de carga / "rechaza cliente"): producto específico con cantidad negativa → resta stock y gasto. Caso B (bonificación / diferencia de
precio / importe parcial): una fila `Nota de credito Materia Prima` (SKU LV0218, categoría NC) con cantidad −1, no mueve kilos.
**Prorrateo** (Quilmes `5548-00008853` en Cabello): impuestos internos y servicio logístico se distribuyen en el precio del producto:
`H = (neto con descuento + cargos) / cantidad / (1 − desc)` y el % de descuento queda en la col I; `envio` e `IMPUESTOS INTERNOS`
quedan en $0 solo para trazabilidad. Sin impuestos internos, un envío va como renglón propio `Envio` (convención de Congreso).
Percepciones de IVA/IIBB no se imputan.

**Informes operativos** (`datos_operativos.py` → `Datos_Operativos_Historicos/indice_operativo.json`): `VENTAS RESUMIDAS.xlsx`
(forma de pago, comprobantes, canales salón/mostrador/vereda/Rappi, cubiertos, promedio por cubierto, descuentos) e `INFORME DE MOV DE
STOCK` de Julio y Agosto (egresos Maxirest: error de comanda, marketing, desperdicio, consumo interno, invitaciones, Big Box).
**Benchmark Grupo Free** (`Precios/benchmark_grupo_free.py` → `benchmark_grupo_free.json`, 54 precios $/kg sin impuestos ni envío):
precio de referencia del proveedor monopólico de papa/cebolla/zanahoria procesadas; audita facturas y NC de Free y sirve para
comparar contra elaboración propia en el Obrador. No compite en el comparador.

**Publicar en producción:** `procesar_lote_wpp.py --push-drive` solo muestra el plan; escribir exige además
`--confirmo-escribir-produccion` (`imputacion_facturas.push_drive`, usa `credentials.json`; agrega A:I a la primera fila libre de
`Facturas`). No usar sin OK explícito del usuario.

**Telegram** (`telegram_ingestor.py`): `--once` (Tarea Programada), `--watch --interval 60`, `--setup-task`. Usa `TELEGRAM_BOT_TOKEN`
(y opcional `TELEGRAM_CHAT_ID`); `telegram_state.json` guarda el `offset` y los SHA-256 ya procesados; sin novedades cuesta una llamada HTTP.

**OCR:** Tesseract 5.4 (winget `UB-Mannheim.TesseractOCR`) + `tessdata/spa.traineddata` local (TESSDATA_PREFIX se fija solo).
Tesseract lee bien el texto impreso pero **no** la birome (ej. "APIO 1,090 KG" → "MIO 4,080"): lo manuscrito se lee con visión.
Las copias `*_FACTURAS_COPIA.xlsx` se guardan con openpyxl (sin valores cacheados de fórmulas): el agente siempre lee catálogo e
historial del export original y de la copia solo las filas que agregó.
