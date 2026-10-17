# Gestión de octubre 2026 — paquete listo para usar (copias locales)

Generado con `gestion_mensual.py` + `conteos_octubre.py`; verificado con `verificar_gestion.py` (0 errores de fórmula).
**Nada se subió a Drive.** El único cambio en Drive es el conteo nuevo de Chimidoha (columna B = SKU + hoja `Comentarios`), que autorizaste.

## Qué es y qué NO es
- Es la gestión de octubre **armada y valorizada**: ítems depurados con las correcciones de Agus, SKU unificados, precios de octubre con la
  lista de precios del Obrador, fórmulas en todo, hojas `Resumen`, `Controles` y `Comentarios`.
- **No** tiene el inventario final de octubre: hoy es 17/10 y los conteos son del 30/09 (Arcos `Control_Semanal` termina el 30/09; solo el Obrador
  tiene snapshots semanales 05/10 y 12/10). Los sectores quedan en blanco. Al cierre: cargar el conteo (GAS o a mano), volver a correr
  `python gestion_mensual.py` para refrescar precios con las facturas del mes y completar compras/ventas en `Resumen`.
- `_prueba/` = mismos libros con los conteos de septiembre, solo para validar fórmulas y precios.

## Archivos
| Archivo | Contenido |
|---|---|
| `Gestion_<Local>_octubre.xlsx` (Arcos, Cabello, Lafinur, Congreso, Chimidoha) | Hojas: Resumen, Inventario (formato Arcos), Controles, Comentarios |
| `Conteo_<Local>_octubre.xlsx` (Arcos, Cabello, Obrador, Chimidoha) | Conteos en blanco, SKU unificado, `Total2` (Obrador: `Total`) con el layout de Arcos, hoja `Cambios octubre` |

## Cómo se fija cada precio (columna F de Inventario; origen en la columna "origen del precio")
1. Ítem de transferencia del Obrador (markup 1,1–1,4) → precio de la lista: Lafinur = `PRECIO_LAFINUR`; Arcos/Cabello/Chimi = `PRECIO_CHIMIDOHA`.
   Conversión g→kg ×1000 según el prefijo de la descripción (GR./UN./PORC.).
2. Factura del local de octubre → 3. lista Obrador (insumo de paso) → 4. factura de los 3 meses previos → 5. Maestro (P.U) → 6. se mantiene el cierre de septiembre.
3. **Congreso valoriza su producción al COSTO del Obrador** (`Recetas!Obrador`, sin margen). Chimidoha: factura de Congreso × recargo por rubro
   (1,2 panificados/descartables, 1,1 proteínas, 1,0 resto) o el mismo ítem de Arcos/Cabello.
4. Regla del proyecto: una variación > ±50 % de factura/Maestro **no se inyecta** (queda el precio de septiembre y la fila va a `Controles` como REVISAR).
   **Excepción que tomé yo:** la lista oficial del Obrador se aplica hasta +300 % / −70 % (los precios de septiembre estaban desactualizados, dijo Agus);
   esas filas quedan marcadas "Aplicado (lista Obrador)" en `Controles`.

## Resultado por local
| Local | Ítems | OK | Aplicado Obrador >±50 % | Sin factura (se mantuvo sept) | REVISAR | Sin precio | Sin SKU |
|---|---|---|---|---|---|---|---|
| Arcos | 363 | 231 | 15 | 75 | 24 | 18 | 3 |
| Cabello | 364 | 210 | 12 | 97 | 30 | 15 | 5 |
| Lafinur | 402 | 187 | 1 | 162 | 36 | 16 | 10 |
| Congreso | 269 | 211 | 1 | 18 | 14 | 25 | 0 |
| Chimidoha | 359 | 233 | 0 | 0 | 0 | 126 | 99 |
Arcos y Cabello comparten 359 ítems; solo difiere TSAL (trucha fresca, Cabello no la trabaja).

## Correcciones de Agus aplicadas (detalle en cada hoja `Comentarios`)
Códigos fusionados (I0021→I0025; TBC5 e I0561→I0560; bagels de Cabello→códigos de Arcos; BURR→PHBURR; FAR→FER; OBR20→OBR3031) · ítems de producción propia,
librería, posnet y velas fuera del conteo · Harina 0000 fuera de Arcos/Cabello · rubro OBRADOR para fritas, shakshuka verde, zhoug rojo y almíbar; vinagre de vino = Almacén ·
crema de almendras con su propio SKU (M54), separada de cremoso de tahina · Vino Bikurim con SKU VBIK · Arcos y Cabello espejados. Lo que Agus pidió *verificar*
(guefilte, salsa inglesa, farfalaj, vinos Serie A, Nicassia, quesos, etc.) está en `Controles` con su nota.

## Pendientes que requieren decisión o datos
- **Resumen de Arcos**: el STOCK INICIAL sale del Inventario de septiembre ($10,6 M); el Resumen viejo decía $14,9 M. No pude conciliar la diferencia.
- **Chimidoha**: 99 ítems sin SKU y 126 sin precio (sobre todo limpieza y descartables sin facturas); revisar `Comentarios` y `Controles`.
- **GAS**: se corrigió en el código local (sin `push`) que `arcos/cabello/lafinur/congreso_precios` pisaran G con F. Siguen sin tocar los GAS de conteo: Agus borró la
  columna "precio anterior" y esos scripts escriben por número de columna e ID de planilla fijos, así que para el libro nuevo hay que ajustar columnas e IDs antes de usarlos.
- Los cambios de GAS y los libros nuevos no se publicaron; para pasar a Drive hace falta tu OK (la cuenta de servicio no puede crear archivos: cuota).
