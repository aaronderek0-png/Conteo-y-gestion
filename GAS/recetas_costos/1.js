function arreglarInsumosRapido() {
  var hojaRecetas = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Recetas");
  var rango = hojaRecetas.getRange("E2:E20000");
  var valores = rango.getValues();
  var validaciones = rango.getDataValidations();
  
  var masterSet = new Set();
  
  // 1. Buscamos la lista maestra de la validación UNA SOLA VEZ para no saturar
  for (var i = 0; i < validaciones.length; i++) {
    var regla = validaciones[i][0];
    if (regla != null && regla.getCriteriaType() === SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE) {
      var rangoMaestro = regla.getCriteriaValues()[0];
      var valoresMaestros = rangoMaestro.getValues().flat();
      for (var v = 0; v < valoresMaestros.length; v++) {
        if (valoresMaestros[v] !== "") masterSet.add(String(valoresMaestros[v]));
      }
      break; // Ya memorizó toda la lista de insumos permitidos, cortamos acá
    }
  }

  var conTildes = "áéíóúÁÉÍÓÚäëïöüÄËÏÖÜàèìòùÀÈÌÒÙ";
  var sinTildes = "aeiouAEIOUaeiouAEIOUaeiouAEIOU";

  // 2. Procesamos las 20k filas directo en la memoria (rapidísimo)
  for (var i = 0; i < valores.length; i++) {
    var texto = valores[i][0];
    
    if (typeof texto === "string" && texto !== "") {
      // Si el texto NO está en la lista permitida (tiene el cartel rojo)
      if (!masterSet.has(texto)) {
        var textoLimpio = texto;
        
        for (var j = 0; j < conTildes.length; j++) {
          textoLimpio = textoLimpio.split(conTildes[j]).join(sinTildes[j]);
        }
        
        // Si al limpiarle las tildes ahora SÍ coincide con la lista, lo pisamos en la memoria
        if (masterSet.has(textoLimpio)) {
          valores[i][0] = textoLimpio;
        }
      }
    }
  }

  // 3. Pegamos las 20k filas actualizadas de un solo golpe
  rango.setValues(valores);
}