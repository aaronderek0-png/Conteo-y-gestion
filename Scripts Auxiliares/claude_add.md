
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
