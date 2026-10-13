function quitarTildesMultiplesHojas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Lista de configuraciones: [Nombre de la hoja, Rango a procesar]
  var tareas = [
    { nombreHoja: "Presentacion", rangoStr: "B2:C20000" },
    { nombreHoja: "Productos", rangoStr: "B2:B20000" }
  ];
  
  var conTildes = "áéíóúÁÉÍÓÚäëïöüÄËÏÖÜàèìòùÀÈÌÒÙ";
  var sinTildes = "aeiouAEIOUaeiouAEIOUaeiouAEIOU";
  
  for (var t = 0; t < tareas.length; t++) {
    var hoja = ss.getSheetByName(tareas[t].nombreHoja);
    if (!hoja) continue; // Si no encuentra la hoja, pasa a la siguiente
    
    var rango = hoja.getRange(tareas[t].rangoStr);
    var valores = rango.getValues();
    
    for (var i = 0; i < valores.length; i++) {
      for (var col = 0; col < valores[i].length; col++) {
        var texto = valores[i][col];
        if (typeof texto === "string" && texto !== "") {
          for (var j = 0; j < conTildes.length; j++) {
            // Reemplazamos todas las ocurrencias de cada vocal acentuada
            texto = texto.split(conTildes[j]).join(sinTildes[j]);
          }
          valores[i][col] = texto;
        }
      }
    }
    
    rango.setValues(valores);
  }
}