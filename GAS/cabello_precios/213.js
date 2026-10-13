/**
 * Cruza los códigos de la columna D (D2:D) contra el Maestro de Insumos (Hoja Productos, A:Código, D:Rubro)
 * y asigna el rubro resultante en la columna A (A2:A). Si no lo encuentra, asigna "Obrador".
 */
function actualizarRubrosPorCodigo() {
  const ID_MAESTRO = "1lOabw9U4ilRwE5r6QXQ2Jc1nqkKUK3Hi096xga0XFnI";
  const HOJA_MAESTRO_PRODUCTOS = "Productos";
  
  // 1. Abrir Maestro y armar diccionario Código -> Rubro
  let mapaRubros = new Map();
  try {
    const ssMaestro = SpreadsheetApp.openById(ID_MAESTRO);
    const hojaMaestro = ssMaestro.getSheetByName(HOJA_MAESTRO_PRODUCTOS);
    
    if (!hojaMaestro) {
      throw new Error(`No se encontró la pestaña "${HOJA_MAESTRO_PRODUCTOS}" en el archivo Maestro.`);
    }
    
    const ultimaFilaMaestro = hojaMaestro.getLastRow();
    if (ultimaFilaMaestro >= 2) {
      // Leemos de A2 a D (Col A = Código, Col D = Rubro)
      const datosMaestro = hojaMaestro.getRange(2, 1, ultimaFilaMaestro - 1, 4).getValues();
      
      datosMaestro.forEach(fila => {
        const codigo = String(fila[0] || "").trim().toUpperCase();
        const rubro = fila[3] ? String(fila[3]).trim() : "";
        if (codigo) {
          mapaRubros.set(codigo, rubro || "Obrador");
        }
      });
    }
  } catch (err) {
    SpreadsheetApp.getUi().alert("Error al conectar con el Maestro: " + err.message);
    return;
  }

  // 2. Leer la hoja activa local (o podés forzar getSheetByName("Facturas"))
  const ssLocal = SpreadsheetApp.getActiveSpreadsheet();
  const hojaLocal = ssLocal.getActiveSheet();
  const ultimaFilaLocal = hojaLocal.getLastRow();

  if (ultimaFilaLocal < 2) return;

  // Leemos los códigos de la columna D (columna 4)
  const codigosLocales = hojaLocal.getRange(2, 4, ultimaFilaLocal - 1, 1).getValues();
  
  // 3. Procesar y mapear resultados para la Columna A
  const resultadoRubros = codigosLocales.map(fila => {
    const cod = String(fila[0] || "").trim().toUpperCase();
    
    // Si la celda del código está vacía, dejamos la celda de Rubro vacía
    if (!cod) return [""]; 
    
    // Si existe en el mapa pone el rubro (o si viene vacío, Obrador), si no existe pone Obrador
    const rubroEncontrado = mapaRubros.get(cod);
    return [rubroEncontrado ? rubroEncontrado : "Obrador"];
  });

  // 4. Volcar todo en la columna A (columna 1) de una sola vez
  hojaLocal.getRange(2, 1, resultadoRubros.length, 1).setValues(resultadoRubros);
}