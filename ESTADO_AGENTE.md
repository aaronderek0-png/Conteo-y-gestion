# Estado del agente (para Codex y revisión) — actualizado 14/10/2026 03:2x

| Fase | Comprobante / ID | Estado | Último fallo | Recurso | Acción pendiente | Evidencia |
|---|---|---|---|---|---|---|
| Recepción | 6 adjuntos IMG2026{0828,0819,0808,0804,0707,0725} | RECIBIDO, guardados en `Facturas Mishiguene/` | APIError (gspread, causa exacta NO observada) | Sheets vivos | `python recuperar_lote.py plan` (solo lectura) | `recuperar_lote.py`, lectura offline de caché |
| 0828153210 | Resta 0024-00097543 | YA_CARGADA (export local) | — | Congreso/Facturas | verificar en vivo con `plan` | auditoría dry-run |
| 0819152602 | Rupajo 00002-00034674 | REVISAR: sin local | — | — | decisión humana de local | auditoría dry-run |
| 0808100551 | MIX-57 0001-00035844 | YA_CARGADA (export local) | — | Congreso/Facturas | verificar en vivo | auditoría dry-run |
| 0804133616 | NC Buenos Mares 3-19303 | NC_APLICADA | — | Congreso | verificar en vivo | auditoría dry-run |
| 0725120055 | Remito Zuccardi 0029-00006172 | REMITO_PENDIENTE | — | Remitos_Pendientes | ninguna (Codex lo publicó) | Resultado_Mishiguene.md |
| 0707124118 | Remito Wigen (2 destinos) | REMITO_PENDIENTE | — | Remitos_Pendientes Congreso/Lafinur | falta factura de julio con referencia al remito | no hay en PDFs ni exports: la 552 (30/07) coincide solo en 10 kg nuez |

Escrituras previstas a `Facturas` por la recuperación: **0 filas** (ninguna INYECTAR).
Bot: PID 23916, commit 9866f59. Eventos: `logs/pipeline_eventos.jsonl`.

## Actualización 14/10/2026 ~03:30
- Bot: PID 16980 (arrancó 03:25), commit a531601, único consumidor; offset 223419362 sin cambios (sin mensajes nuevos).
- Lectura viva (verificacion_vivo.json): Resta 97543 (Congreso fila 391), MIX-57 35844 (filas 93-94) y NC Buenos Mares 19303 (fila 34) están cargadas.
  Rupajo 34674 no está en Congreso ni Lafinur (sin local: decisión humana). Miño 421/423 no están en Facturas (4 renglones REVISAR, no se publican).
  Raíces 8528: sus 5 renglones ya están en Congreso/Remitos_Pendientes. Remito Obrador 3727 NO está en Cabello/Remitos.
- PENDIENTE de aprobación (bloqueado por clasificador al ejecutar `python recuperar_lote.py apply`): escribir Cabello/Remitos filas 609-613
  (03/06 Flan 4×461,75; mammoul 10×192,43; Lock food 230 100×269; Lock food 390 100×278; Caja pizza 1×185; total $58.656,30).
- Comparador: $346.639/mes estimado, 124/324 insumos con match (46,9 s). Sin referencia: 206 productos = 81,4 % del gasto Dossier (scouting_sin_referencia.py).
- BOM (motor_recetas): 1833 recetas; 1660 coinciden con C_UN ±2 %, 117 difieren, 25 incompletas.

## 14/10/2026 ~10:50 — causa observada del APIError y bloqueo real
- Reproducido con el remito 3727: `APIError [400] Range (Remitos!A609:E613) exceeds grid limits. Max rows: 608`. La pestaña Remitos de
  Cabello tiene la grilla llena (608/608) y está PROTEGIDA: `add_rows` devuelve `400 You are trying to edit a protected cell or object`.
  No se escribió nada (backup en backups_recuperacion/Cabello_*_Remitos.json). El código ahora agrega filas cuando la hoja no está protegida.
- Acción pendiente (dueño de la planilla): ampliar ≥25 filas en CABELLO/Remitos o quitar la protección / dejar a la Service Account
  como editora de ese rango. Luego `python recuperar_lote.py apply` (vigencia del plan se revalida sola).
- BOM: `bom_diferencias.py` → 142 recetas con diferencia (25 incompletas, 24 con factor entero = hoja sin dividir por REND,
  93 sin patrón). Cálculo sobre export local Recetas.xlsx, no Drive en vivo. No se publica ningún costo.

## 14/10/2026 — trials (misión de ambientes)
- Copias de prueba: NO creadas. Drive `files.copy` con la SA → 403 storageQuotaExceeded. Acción: dueño copia las 5 planillas, comparte con la SA
  como Editor y pasa los IDs (ver TRIALS_Y_REF.md). Controles por ID implementados (entornos.py) y probados.
- Registro de aprobación (aprobaciones.py): modelo Q:X + reglas probadas en memoria; sin aplicar a planillas reales.
- Ventas Maxirest: solo estudio (PROPUESTA_VENTAS_MAXIREST.md); recomendación A. Nada modificado.
- #REF! Lafinur: causa en TRIALS_Y_REF.md (comodín vertical {"";"";""} derramando sobre la fila siguiente). Sin tocar producción.
