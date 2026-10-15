# Trials aislados: estado, auditoría de copias y diagnóstico #REF! de Lafinur (14/10/2026)

## 1. Copias «NO CARGAR - PRUEBA - <Local>» — NO creadas (bloqueo observado)
- Operación: `POST drive/v3/files/{id}/copy` con la Service Account (`crear_trials.py Cabello`). Resultado: **HTTP 403
  `storageQuotaExceeded` — "The user's Drive storage quota has been exceeded."** La SA no puede ser propietaria de archivos.
- Alternativa sin permisos nuevos: el dueño hace «Archivo ▸ Hacer una copia» de las 5 planillas, la nombra
  `NO CARGAR - PRUEBA - <Local>`, la comparte con `agente-mishi@black-obelisk-460914-d9.iam.gserviceaccount.com` como Editor y pasa
  los 5 IDs. Se registran con `entornos.registrar_prueba(...)` (verifica ID≠producción, prefijo del título y editable).
- Mientras tanto, las pruebas corren sobre copias **locales** (`Copias_Gestion/*_FACTURAS_COPIA.xlsx`).

## 2. Controles que impiden escribir en producción desde PRUEBA
- `MISHI_ENTORNO=PRUEBA` → `imputacion_facturas.IDS_PLANILLAS` = IDs de `entornos.json["prueba"]` (si no hay, falla al importar).
- `_gspread()` envuelve `open_by_key`: cualquier ID de producción o no registrado como prueba lanza `EscrituraProhibida`.
- Registro de copia: rechaza ID de producción, título sin prefijo o copia no editable (tests en `tests/test_entornos_aprobaciones.py`).
- Límite honesto: el control cubre el código Python del proyecto. `consolidador.py` y `analisis_gestion_comercial.py` tienen IDs
  fijos pero solo LEEN; los GAS no se ejecutan desde Python.

## 3. Auditoría de IDs fijos en scripts GAS (riesgo de modificar originales desde una copia)
Una copia de planilla copia su script ligado, y ese script usa `openById` con IDs de producción. NO ejecutar menús/triggers en las copias:
- Escriben o leen originales: `*_conteo_gestion` (ID_GESTION / HISTORIAL_EXTERNO / MAESTRO), `congreso_consumo_teorico`
  (FACTURAS, CONTEO_INICIAL, LOGS_RECETAS, TRASPASOS, RECETAS, VENTAS_MAXI), `*_precios` (ID_FACTURAS_*), `*_remitos_alertas`
  (MAESTRO_ID: Solicitudes), `recetas_costos` (TARGET_LOGS_ID), `maestro_productos` (replica SKU en las 4 planillas).
- Los triggers no se copian con la planilla, pero un menú manual sí puede ejecutar el código.
- Fórmulas `IMPORTRANGE` de las copias siguen leyendo el Maestro real (lectura; no modifica).

## 4. #REF! de Lafinur — causa en el export (reproducción en copia pendiente de los IDs de prueba)
- Export `Arcos_data/lafinur.xlsx`, pestaña `Facturas`: 23.077 filas con `#REF!` en K, T y V (UN. PRESENT, CATEGORÍA GASTO, ALERTA).
  `Control precios`: 290 celdas con `#REF!` (FILTER sobre `Facturas!…2:…1001`).
- Mecanismo (fórmulas de K111/T111): `=IF(OR(A111="",E111=""), {"";"";""}, XLOOKUP(...))`. Cuando A tiene dato y E (PRODUCTOS) está
  vacía —o la fila está vacía—, el comodín `{"";"";""}` es una columna de 3 filas (`;` separa filas), no 3 columnas.
  Cada celda intenta derramar 3 filas hacia abajo y choca con la fórmula de la fila siguiente → «El resultado de la matriz no se
  expandió porque sobrescribiría datos» = `#REF!`. Lo mismo `{"";""}` en T:U y `HSTACK("","",)` en V.
- Filas con A sin E detectadas: 111 (factura 121, DAICH MARCOS R) y 854. La fila 111 es la que dispara el error visible.
- Arreglo propuesto (NO aplicado en producción): comodines horizontales `{"","",""}` (en es-AR: `{""\\""\\""}`) en K:M, `{"",""}` en T:U,
  y completar la columna E de las filas 111 y 854. Verificar en la copia de prueba antes de llevarlo a producción.
