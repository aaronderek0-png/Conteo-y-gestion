# Ecosistema Apps Script (GAS) — Mishiguene

Cuenta: aaron.derek0@gmail.com · 15 proyectos clonados con clasp · IDs en `gas_registry.json`.
Sincronizar: `python gas_sync.py pull` (desde la raíz del proyecto).

## Flujo de datos (cómo encaja todo)

```
Facturas (Sheet por local)  ──► *_precios  ──► hoja Gestion/Inventario (precio viejo/nuevo, valorizado)
Remitos (Gmail, INFORME*)   ──► *_remitos_alertas ──► alerta diaria 11 AM por mail (variación > 6 %)
Conteos por sector (Total)  ──► *_conteo_gestion ──► Gestión mensual + snapshot Control_Semanal
Maxirest (ventas) + Recetas ──► congreso_consumo_teorico / ventas_bom_rentabilidad ──► consumo teórico vs real, rentabilidad
MAESTRO DE PRODUCTOS (1lOabw9U…) ◄── fuente única de SKU y rubro, la leen casi todos
```

## Proyectos

| Carpeta | Local | Qué hace |
|---|---|---|
| `maestro_productos` | Central | Maestro de SKU: concilia facturas + verdulería, aprobación de SKU nuevos, memoria de resoluciones. Triggers: facturas c/24 h, verdulería c/2 días. |
| `recetas_costos` | Central | Pipeline de costos de recetas (Insumos → Recetas → Costos → Maestro), snapshot al histórico, alertas WhatsApp (umbral insumos ±15 %, recetas ±10 %). |
| `ventas_bom_rentabilidad` | Central | Explosión de recetas (BOM) sobre ventas, reporte teórico y rentabilidad por plato. |
| `congreso_consumo_teorico` | Congreso | Sincronización total: stock inicial + facturas + traspasos + ventas Maxirest − consumo teórico = desvíos. |
| `arcos_conteo_gestion` | Arcos | Conteos por sector (Total → Total2) → Gestión mensual de la rotisería; snapshot semanal. |
| `arcos_precios` | Arcos | Último precio mensual desde hoja Facturas hacia Inventario. |
| `arcos_remitos_alertas` | Arcos | Importa remitos de Gmail (Maxirest "Ventas por artículo"), prorrateo de gastos/impuestos, auditoría diaria de precios. |
| `cabello_conteo_gestion` | Cabello | Idem conteo → Gestión para Cabello. |
| `cabello_precios` | Cabello | Precios y rubros desde Facturas (2 variantes: Gestion y Sheet6). |
| `cabello_remitos_alertas` | Cabello | Mismo script de remitos/alertas (LOCAL = CABELLO). |
| `congreso_obrador_conteo` | Congreso/Obrador | Conteo de producción/pastelería → consolidación. |
| `congreso_precios` | Congreso | Último precio mensual + valorizado y % sobre total. |
| `congreso_alertas` | Congreso | Auditoría diaria de variación de precios (sin importar remitos). |
| `lafinur_precios` | Lafinur | Precios y rubros desde Facturas + Maestro. |
| `lafinur_remitos_alertas` | Lafinur | Mismo script de remitos/alertas (LOCAL = LAFINUR). |
| `_ajeno_quini6` / `_vacio` | — | No son de Mishiguene. Ignorar. |

## IDs compartidos
- Maestro de productos: `1lOabw9U4ilRwE5r6QXQ2Jc1nqkKUK3Hi096xga0XFnI` (hoja `Productos`)
- Facturas Arcos `1frh93qB…` · Cabello `1zAxYyNB…` · Lafinur `1tv0uI4K…` · Congreso `1udLEQpQ…`
- Recetas `1u7QSi7J…` · Logs maestro `1SXFt1Kd…` · Ventas Maxirest `1IRRh0wf…` · Conteo inicial `1i0LPYWI…` · Traspasos `1FbRoJwx…`

## Hallazgos / deuda técnica
1. **Código copiado ×4**: los 4 scripts de remitos/alertas son casi idénticos (diferencias de 64 a 644 líneas). Cualquier arreglo hay que hacerlo 4 veces → candidato a unificar.
2. **Lógica de conteo ×3** (arcos/cabello/congreso): misma estructura con CONFIG distinto.
3. `norm`, `parseNum`, `parseFecha` están reescritos en casi todos.
4. **Secretos en texto plano**: API key y teléfono de CallMeBot (`recetas_costos`). Mover a `PropertiesService`.
5. Los IDs de hojas están hardcodeados: si se duplica un Sheet (p. ej. por mes) hay que editar cada script.
6. Los scripts atados a un Sheet **no aparecen en `clasp list-scripts`**; se clonan solo con su Script ID.
