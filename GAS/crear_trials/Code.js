/**
 * Crea las 5 copias «NO CARGAR - PRUEBA - <Local>» en el Drive del usuario y las comparte con la Service Account.
 * USO MANUAL: pegar en https://script.google.com (cuenta aaron.derek0@gmail.com) y ejecutar crearTrials().
 * NO se publica con clasp. No toca los originales: solo lee y copia. Devuelve el JSON para entornos.json (prueba).
 * Aviso: la copia trae su script ligado con IDs de producción; NO ejecutar sus menús.
 */
function crearTrials() {
  var SA = 'agente-mishi@black-obelisk-460914-d9.iam.gserviceaccount.com';
  var ORIGENES = {
    Congreso: '1udLEQpQ29avK4c6fNnQig-w02xIV82IgLe85yVJXaOY',
    Arcos: '1frh93qB8Od5mrtAfMuObKwRW93p7kev6DNgmERqz6ho',
    Lafinur: '1tv0uI4KsSEPRUecYfnwcDqpGbUDelGqWwe3Hp-fikOc',
    Cabello: '1zAxYyNB98Sb1-D0HtfokESXc6FrDS96EbFYyMmsseIE',
    Chimidoha: '1gSHbDSgICeTy6lZpyIPRskFbEs0JB5mxw1e71oOkIzU'
  };
  var out = {};
  Object.keys(ORIGENES).forEach(function (local) {
    var copia = DriveApp.getFileById(ORIGENES[local]).makeCopy('NO CARGAR - PRUEBA - ' + local);
    copia.addEditor(SA);
    out[local] = copia.getId();
  });
  Logger.log(JSON.stringify({ prueba: out }, null, 1));
  return out;
}

/**
 * Quita las protecciones de hoja/rango SOLO en las 5 copias de PRUEBA (IDs fijos de abajo; nunca los originales), para que
 * la Service Account pueda escribir en Facturas/Remitos del sandbox. Ejecutar a mano como dueño. No toca producción.
 */
function desprotegerCopiasDePrueba() {
  var PRUEBA = {
    Chimidoha: '1wkf1bu1m4K3ZOgCznJ3oDtN_nnjtaWCGxhONVSfTQxs',
    Cabello: '1QdHOCClPQU0tOSMzBq56wCPYLhsleZbr7gFIJrlZRFM',
    Lafinur: '1fqBlOxDKMGzYz7lG0-BNe1gJPVZGH01k5GxSgmz-vi4',
    Arcos: '1bkSFcrUG0pAnO0VrxHcQKWM6h8M3pf8YLUfKKVH4XTg',
    Congreso: '1YGNpT3zhnfGDdIzrDalNo8BRBBvvX51qCT5WuQsyME4'
  };
  var quitadas = {};
  Object.keys(PRUEBA).forEach(function (local) {
    var ss = SpreadsheetApp.openById(PRUEBA[local]);
    if (ss.getName().indexOf('NO CARGAR - PRUEBA - ') !== 0) throw new Error('No es una copia de prueba: ' + local);
    var n = 0;
    ss.getSheets().forEach(function (sh) {
      sh.getProtections(SpreadsheetApp.ProtectionType.SHEET).concat(
        sh.getProtections(SpreadsheetApp.ProtectionType.RANGE)).forEach(function (p) { p.remove(); n++; });
    });
    quitadas[local] = n;
  });
  Logger.log(JSON.stringify(quitadas));
  return quitadas;
}
