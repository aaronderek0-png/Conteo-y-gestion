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
