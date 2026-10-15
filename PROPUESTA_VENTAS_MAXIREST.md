# Ventas Maxirest y pronóstico de panes — estudio (SIN cambios aplicados)

Evidencia: exports locales en `Arcos_data/` (no es lectura viva de Drive) + código en `GAS/`. Fecha: 14/10/2026.

## Estructura real
- **Entrada Maxirest**: pestaña `Historial_Productos` (A FECHA, B CODIGO, C PRODUCTO, D UNIDADES, E PRECIO, F TOTAL, G DIA_SEM). 4.926 filas
  desde 01/07/2026 en el export de octubre; una sola fórmula (G, array).
- **Consumidores por ID**: `GAS/congreso_consumo_teorico` lee `ID_VENTAS_MAXI = 1IRRh0wf…`, hoja `Historial_Productos`, columnas A:D
  (fecha, código, cantidad), mapea código Maxirest → SKU de receta y calcula consumo teórico vs real. `ventas_bom_rentabilidad`
  usa `ID_RECETAS 1u7QSi7J…` (recetas, no ventas).
- **Archivos por período**: `ventas_obrador_octubre.xlsx` (doc 1ZoHhFnn…) y `ventas_obrador_septiembre` (doc 1lDrAOdA…) tienen IDs distintos
  entre sí y distintos del `1IRRh0wf…` de los scripts. Cada período nuevo = ID nuevo = IDs hardcodeados desactualizados
  (ya señalado como deuda 5 en `GAS/ECOSISTEMA.md`).
- **Pestañas del archivo mensual**: Resumen (14 fórmulas → `Por semana`), Por día (62), Por semana (49), Historial_Productos (entrada),
  Planificacion_Bloques (336 fórmulas; 962 referencias a Historial_Productos, 156 a Stock_Conteo, 24 a Incidencias_Produccion, 11 #REF),
  Stock_Conteo, Pronostico (46 filas × 19 col.; 438 fórmulas, 2.418 referencias), Mapeo_Productos (42 filas, sin fórmulas).
- **Pronóstico de panes**: `=MAX(SUMIFS(Historial!D, Historial!C, producto, Historial!A, día-7), SUMIFS(..., día-14))` por producto y día.
  Necesita ≥14 días de historia previos al día pronosticado: **los primeros 14 días de un período nuevo dependen del período anterior**.

## Variante A — un archivo vivo, separando capas por pestaña + cierres
Entradas (`Historial_Productos`, `Mapeo_Productos`) / cálculos (`Por día`, `Por semana`, `Resumen`) / pronóstico, y un cierre por período
(snapshot de valores en pestaña `Cierre_AAAA-MM`, o export a archivo de solo lectura).
- Ventajas: conserva el ID que usan los scripts; el pronóstico mantiene la ventana de 14 días sin tocar fórmulas; refresco = lo de hoy.
- Riesgos: el archivo crece (SUMIFS de columna completa se enlentece); una edición accidental de entradas altera el histórico;
  hay que cerrar períodos de forma disciplinada.
- Efecto sobre lo que funciona: ninguno inmediato (solo agrega pestañas de cierre).

## Variante B — entradas aisladas + archivos derivados
Archivo de entradas (solo Maxirest, permisos restringidos) y archivos derivados (resumen, gestión, pronóstico) que leen por IMPORTRANGE.
- Ventajas: entradas inmutables y auditables; derivados regenerables; separa responsabilidades.
- Riesgos: IMPORTRANGE necesita autorización y tiene latencia/límites; cada derivado depende de un ID más; hay que reapuntar
  `congreso_consumo_teorico` y re-armar 2.418 referencias del pronóstico; más puntos de rotura si cambia un ID.
- Efecto: cambios en scripts y fórmulas que hoy funcionan.

## Recomendación: A (con cierres), evolucionable a B
Con la estructura real: el ID vivo ya lo fija un consumidor (`1IRRh0wf…`), el pronóstico exige historia continua y 962+2.418 referencias
internas a `Historial_Productos`. A no rompe nada y resuelve el problema de períodos (cierres por pestaña/archivo). B se puede
introducir después aislando solo `Historial_Productos` como fuente única cuando el volumen lo justifique.
Pendiente de confirmar con lectura viva: que `1IRRh0wf…` sea el archivo actual y quién más lo consume (no solo GAS del repo).
