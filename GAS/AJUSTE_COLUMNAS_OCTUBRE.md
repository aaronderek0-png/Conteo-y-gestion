# Ajuste de columnas de los GAS para los libros de octubre

Los libros `Gestion_<Local>_octubre.xlsx` (carpeta `Copias_Gestion/octubre_VF/`) usan el layout de Arcos **sin** la columna "precio anterior"
y con encabezado en la **fila 1** (datos desde la **fila 2**). Los GAS escriben por número de columna, así que hay que corrimiento -1 (o -2) en los
números de abajo. Todos los números son de columna (A=1) / de fila. **Cambiar también el ID de la planilla de gestión** (`ID_GESTION_MENSUAL` / `ID_GESTION`) por el del
libro nuevo una vez subido a Drive.

## Layout nuevo (todos los locales)
`A RUBRO · B UN. · C MEDIDA · D CODIGO · E PRODUCTO · F $ unitario actual · G… sectores · FINAL · VALORIZADO · % · Origen precio · …`

| Local | Sectores | FINAL | VALORIZADO | % valorizado | Origen del precio |
|---|---|---|---|---|---|
| Arcos | G:N (7–14) | O (15) | P (16) | Q (17) | R (18) |
| Cabello | G:M (7–13) | N (14) | O (15) | P (16) | Q (17) |
| Lafinur | G (7) | H (8) | I (9) | J (10) | K (11) |
| Congreso/Obrador | G:I (7–9) | J (10) | K (11) | L (12) | M (13) |
| Chimidoha | G:M (7–13) | N (14) | O (15) | P (16) | Q (17) |

## GAS de CONTEO (los que hay que tocar a mano)

### `arcos_conteo_gestion/Code.js`
| Línea | Hoy | Poner |
|---|---|---|
| 19 | `GESTION_FILA_DATOS: 3,` | `GESTION_FILA_DATOS: 2,` |
| 20 | `GESTION_COL_DESTINO_INICIO: 8,` | `GESTION_COL_DESTINO_INICIO: 7,` (G) |
| 622 | `` `=SUM(H${filaReal}:O${filaReal})` `` | `` `=SUM(G${filaReal}:N${filaReal})` `` |
| 623 | `` `=P${filaReal}*G${filaReal}` `` | `` `=O${filaReal}*F${filaReal}` `` |
| 624 | `` `=IFERROR(Q${filaReal}/SUM(Q$3:Q$${uFilaRemota}),0)` `` | `` `=IFERROR(P${filaReal}/SUM(P$2:P$${uFilaRemota}),0)` `` |
| 629 | `getRange(CONFIG.GESTION_FILA_DATOS, 8, cantFilasRemotas, 8)` | `getRange(CONFIG.GESTION_FILA_DATOS, 7, cantFilasRemotas, 8)` |
| 630 | `getRange(CONFIG.GESTION_FILA_DATOS, 16, cantFilasRemotas, 3)` | `getRange(CONFIG.GESTION_FILA_DATOS, 15, cantFilasRemotas, 3)` |
| 637 | `item.prod, "", "", ...item.conteos, "", "", "", "Nuevo item sin precio"` | `item.prod, "", ...item.conteos, "", "", "", "Nuevo item sin precio"` (un solo `""` antes de los conteos) |
| 650 | `` `=SUM(H${f}:O${f})` `` | `` `=SUM(G${f}:N${f})` `` |
| 651 | `` `=P${f}*G${f}` `` | `` `=O${f}*F${f}` `` |
| 652 | `` `=IFERROR(Q${f}/SUM(Q$3:Q$${filaAlta + filasNuevas.length - 1}),0)` `` | `` `=IFERROR(P${f}/SUM(P$2:P$${filaAlta + filasNuevas.length - 1}),0)` `` |
| 655 | `getRange(filaAlta, 16, filasNuevas.length, 3)` | `getRange(filaAlta, 15, filasNuevas.length, 3)` |
| 599 (mensaje) | `"…desde la fila 3."` | `"…desde la fila 2."` (cosmético) |

Nota: con `GESTION_FILA_DATOS: 3` sobre un libro con datos desde la fila 2, el primer ítem (AAO) **nunca recibía conteo**: ese bug ya existía.
Sin cambios: línea 602 (`getRange(…, 1, n, 18)` lee 18 columnas, alcanza) y línea 610 (SKU en `[3]` = columna D, igual que antes).

### `cabello_conteo_gestion/Code.js`
| Línea | Hoy | Poner |
|---|---|---|
| 31 | `GESTION_COL_DESTINO_INICIO: 8,` | `GESTION_COL_DESTINO_INICIO: 7,` (G) |
| 510 | `item.prod, "", "", ...item.conteos, item.final` | `item.prod, "", ...item.conteos, item.final` (un solo `""`) |
| 536 (texto) | `(H:O actualizadas)` | `(G:N actualizadas)` (cosmético) |
Sin cambios: `GESTION_FILA_DATOS: 2`, línea 486 (SKU en columna 4 = D, el libro nuevo ya trae MEDIDA en C) y línea 503 (ancho 8 = 7 sectores + FINAL).
Ojo: el FINAL de Cabello se escribe como valor (pisa la fórmula `=SUM`). Si querés conservar la fórmula, escribí solo 7 columnas: `getRange(…, CONFIG.GESTION_COL_DESTINO_INICIO, cantFilasRemotas, 7)` y `filaConteos = [...item.conteos]`.

### `congreso_obrador_conteo/Code.js`
| Línea | Hoy | Poner |
|---|---|---|
| 26 | `GESTION_FILA_DATOS: 3,` | `GESTION_FILA_DATOS: 2,` |
| 27 | `GESTION_COL_DESTINO_INICIO: 9,` | `GESTION_COL_DESTINO_INICIO: 7,` (G) |
| 572 | `rubro, un, med, skuRaw, prodRaw, "", "", "", ...conteos4` | `rubro, un, med, skuRaw, prodRaw, "", ...conteos4` |
| 578 | `getRange(nFila, CONFIG.GESTION_COL_DESTINO_INICIO, 1, 4).setValues([valores4])` | `getRange(nFila, CONFIG.GESTION_COL_DESTINO_INICIO, 1, 3).setValues([valores4.slice(0, 3)])` (el FINAL del libro nuevo es fórmula `=SUM(G:I)`; no pisarlo) |
| 587 | `getRange(filaDestinoInsert, 1, filasParaInsertarAbajo.length, 12)` | `…, 10)` |
Sin cambios: línea 505 (SKU en D, nombre en E).

## GAS de PRECIOS (mismo corrimiento; el parche "no pisar G con F" ya está en el código)
Con una sola columna de precio (F), `PRECIO_VIEJO` y `PRECIO_NUEVO` apuntan a la misma columna 6.

| Script | Constante | Hoy → Poner |
|---|---|---|
| `arcos_precios` | `FILA_INICIO_INV` | 3 → 2 |
| | `COL_PRECIO_NUEVO_INV` | 7 → 6 |
| | `COL_FINAL_INV` / `COL_VALORIZADO_INV` / `COL_PORC_VALORIZADO_INV` / `COL_ORIGEN_PRECIO_INV` | 16/17/18/19 → 15/16/17/18 |
| `cabello_precios` | `COLS.PRECIO_ACTUAL` | 7 → 6 |
| | `COLS.VALORIZADO` / `COLS.FECHA_ULT_FACT` | 16 / 18 → 15 / 17 |
| `lafinur_precios` | `COL_PRECIO_DESTINO` | 7 → 6 |
| `congreso_precios` | `FILA_INICIO_INV` | 3 → 2 |
| | `COL_PRECIO_NUEVO_INV` / `COL_ORIGEN_PRECIO_INV` | 7 / 8 → 6 / 13 |
| | `COL_FINAL_INV` / `COL_VALORIZADO_INV` / `COL_PORCENTAJE_INV` | 12 / 13 / 14 → 10 / 11 / 12 |
(`lafinur_precios`: `COL_CODIGO` 4 y `COL_FECHA_DESTINO` 11 no cambian.)

## Aviso sobre los libros "vivos" de septiembre
Los GAS apuntan hoy al libro **sin** `_vf` (p. ej. Arcos `1mzZE5ZZ…`), que **conserva** la columna "precio anterior" (sectores en H:O). Agus borró esa columna en las copias
`_vf`. Estas correcciones son para los libros nuevos de octubre; no las apliques al libro viejo.
