/**
 * Aprobador de Facturas IA — Mishiguene
 *
 * El agente (procesar_lote_wpp.py / telegram_ingestor.py) inyecta filas en la pestaña "Facturas" y las pinta:
 *   #FFF8E1  amarillo suave = cargada OK, pendiente de aprobación
 *   #FFEBEE  rojo suave     = requiere revisión (E dice "REVISAR / NO CATALOGADO"; el original y el motivo van en la col F)
 *
 * Este script agrega al menú de Google Sheets:  Facturas IA  >  ✅ Aprobar Facturas IA (Quitar color)
 * Instalación: Extensiones > Apps Script > pegar este archivo > guardar > recargar la planilla (una vez autorizar).
 * No cambia ningún dato: solo quita el color de fondo de A:I de las filas aprobadas.
 */
var HOJA_FACTURAS = 'Facturas';
var COLOR_OK = '#fff8e1';
var COLOR_REVISAR = '#ffebee';
var COLS = 9;            // A:I

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Facturas IA')
    .addItem('✅ Aprobar Facturas IA (Quitar color)', 'aprobarFacturasIA')
    .addItem('Aprobar TODAS las amarillas (OK)', 'aprobarTodasLasOK')
    .addSeparator()
    .addItem('Ir a la próxima fila a revisar (rojo)', 'irAProximaRevision')
    .addToUi();
}

/** Quita el color de las filas seleccionadas (amarillas y rojas). Pide confirmación si hay rojas. */
function aprobarFacturasIA() {
  var ui = SpreadsheetApp.getUi();
  var hoja = SpreadsheetApp.getActiveSheet();
  if (hoja.getName() !== HOJA_FACTURAS) {
    ui.alert('Abrí la pestaña "' + HOJA_FACTURAS + '" y seleccioná las filas a aprobar.');
    return;
  }
  var rango = hoja.getActiveRange();
  var filas = [];
  for (var r = rango.getRow(); r < rango.getRow() + rango.getNumRows(); r++) {
    var c = hoja.getRange(r, 1).getBackground();
    if (c === COLOR_OK || c === COLOR_REVISAR) filas.push({fila: r, rojo: c === COLOR_REVISAR});
  }
  if (!filas.length) {
    ui.alert('En la selección no hay filas pintadas por el agente (amarillo o rojo suave).');
    return;
  }
  var rojas = filas.filter(function (f) { return f.rojo; }).length;
  if (rojas) {
    var resp = ui.alert('Hay ' + rojas + ' fila(s) en rojo (a revisar). ¿Aprobarlas igual?', ui.ButtonSet.YES_NO);
    if (resp !== ui.Button.YES) filas = filas.filter(function (f) { return !f.rojo; });
  }
  filas.forEach(function (f) { hoja.getRange(f.fila, 1, 1, COLS).setBackground(null); });
  SpreadsheetApp.getActive().toast(filas.length + ' fila(s) aprobada(s).', 'Facturas IA', 5);
}

/** Quita el color de TODAS las filas amarillas (las rojas no se tocan). */
function aprobarTodasLasOK() {
  var hoja = SpreadsheetApp.getActive().getSheetByName(HOJA_FACTURAS);
  var ultima = hoja.getLastRow();
  if (ultima < 2) return;
  var fondos = hoja.getRange(2, 1, ultima - 1, 1).getBackgrounds();
  var n = 0;
  for (var i = 0; i < fondos.length; i++) {
    if (fondos[i][0] === COLOR_OK) {
      hoja.getRange(i + 2, 1, 1, COLS).setBackground(null);
      n++;
    }
  }
  SpreadsheetApp.getActive().toast(n + ' fila(s) amarilla(s) aprobada(s).', 'Facturas IA', 5);
}

function irAProximaRevision() {
  var hoja = SpreadsheetApp.getActive().getSheetByName(HOJA_FACTURAS);
  var ultima = hoja.getLastRow();
  var fondos = hoja.getRange(2, 1, Math.max(ultima - 1, 1), 1).getBackgrounds();
  for (var i = 0; i < fondos.length; i++) {
    if (fondos[i][0] === COLOR_REVISAR) {
      hoja.setActiveRange(hoja.getRange(i + 2, 1, 1, COLS));
      return;
    }
  }
  SpreadsheetApp.getActive().toast('No quedan filas en rojo.', 'Facturas IA', 4);
}
