function marcarCodigosEnCostos() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Obtener la hoja actual (Recetas) y la hoja de destino (Costos)
  var hojaRecetas = ss.getActiveSheet(); 
  var hojaCostos = ss.getSheetByName("Costos");
  
  // Verificar que la hoja 'Costos' exista
  if (!hojaCostos) {
    SpreadsheetApp.getUi().alert("Error: No se encontró la hoja llamada 'Costos'. Verifica el nombre.");
    return;
  }
  
  // 2. Lista completa de códigos a buscar
  var codigosBuscar = [
    // Nuevos códigos agregados
    "OBR11", "OBR28", "OBR48", "OBR14", "M20", "OBR20", "OBR24", "OBR25", 
    "OBR53", "OBR59", "OBR64", "P202", "OBR63", "OBR149", "P170", "OBR103",
    // Códigos anteriores
    "P2", "A10", "A12", "D6", "P10", "P103", "P107", "P11", "P116", "P123", 
    "P127", "P129", "P13", "P137", "P14", "P140", "P15", "P154", "P156", "P158", 
    "P159", "P16", "P160", "P169", "P17", "P176", "P179", "P18", "P193", "P201", 
    "P205", "P210", "P23", "P232", "P238", "P24", "P244", "P245", "P247", "P25", 
    "P252", "P254", "P261", "P264", "P266", "P286", "P290", "P292", "P293", "P295", 
    "P296", "P298", "P299", "P30", "P300", "P301", "P302", "P303", "P304", "P305", 
    "P37", "P39", "P4", "P40b", "P42", "P43", "P44", "P47", "P53", "P6", 
    "P62", "P70", "P72", "P74", "P76", "P77", "P80", "P81", "P90", "P97"
  ];
  
  // Convertir a Conjunto (Set) para búsquedas rápidas
  var setCodigos = new Set(codigosBuscar);
  
  // 3. Obtener los datos de la columna A en la hoja Recetas
  var ultimaFilaRecetas = hojaRecetas.getLastRow();
  if (ultimaFilaRecetas < 2) return; 
  
  var rangoCodigos = hojaRecetas.getRange(2, 1, ultimaFilaRecetas - 1, 1); 
  var valoresCodigos = rangoCodigos.getValues();
  
  // 4. Preparar la lectura y escritura en la hoja 'Costos'
  var ultimaFilaCostos = hojaCostos.getLastRow();
  if (ultimaFilaCostos < 2) {
    SpreadsheetApp.getUi().alert("La hoja 'Costos' no tiene filas de datos.");
    return;
  }
  
  // Códigos en columna A de Costos
  var rangoCodigosCostos = hojaCostos.getRange(2, 1, ultimaFilaCostos - 1, 1);
  var valoresCostosA = rangoCodigosCostos.getValues();
  
  // Valores actuales en columna G (Columna 7) de Costos
  var rangoColumnag = hojaCostos.getRange(2, 7, ultimaFilaCostos - 1, 1); 
  var valoresCostosG = rangoColumnag.getValues();
  
  // 5. Procesar coincidencias
  for (var i = 0; i < valoresCodigos.length; i++) {
    var codigoActual = valoresCodigos[i].toString().trim();
    
    if (codigoActual === "") continue;
    
    if (setCodigos.has(codigoActual)) {
      for (var j = 0; j < valoresCostosA.length; j++) {
        if (valoresCostosA[j].toString().trim() === codigoActual) {
          valoresCostosG[j][0] = "SI"; // CORREGIDO: Asigna a la celda específica dentro de la fila
        }
      }
    }
  }
  
  // 6. Escribir los resultados corregidos en la columna G de la hoja Costos
  rangoColumnag.setValues(valoresCostosG);
  
  // Notificación de éxito
  SpreadsheetApp.getUi().alert("¡Proceso completado con éxito! Se marcaron con 'SI' en la hoja Costos.");
}
