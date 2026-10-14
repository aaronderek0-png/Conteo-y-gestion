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
