// ==========================================================================
// CONFIGURACIÓN DE ORIGEN (MAESTRO Y FACTURAS) Y GESTIÓN LOCAL
// ==========================================================================

const ID_MAESTRO_PRODUCTOS = "1lOabw9U4ilRwE5r6QXQ2Jc1nqkKUK3Hi096xga0XFnI";
const ID_FACTURAS_ORIGEN = "1tv0uI4KsSEPRUecYfnwcDqpGbUDelGqWwe3Hp-fikOc";

const CONFIG_LOCAL = {
  HOJA_GESTION: "Gestion",
  FILA_INICIO: 2,          // Arranca en la fila 2
  COL_RUBRO_DESTINO: 1,    // Columna A (RUBRO)
  COL_CODIGO: 4,           // Columna D (CODIGO)
  COL_PRECIO_VIEJO: 6,     // Columna F ($ unitario anterior)
  COL_PRECIO_DESTINO: 7,   // Columna G ($ unitario actual)
  COL_FECHA_DESTINO: 11    // Columna K (Fecha / Auditoría)
};

const MAESTRO_COLS = {
  HOJA: "Productos",
  COL_SKU: 0,              // Columna A (SKU)
  COL_RUBRO: 3             // Columna D (Rubro / Categoría)
};

const FACTURAS_COLS = {
  FECHA_INGRESO: 2,        // Columna C (Fecha ingreso remota)
  PRECIO_UNITARIO: 14,     // Columna O (Precio unitario remoto)
  CODIGO_SKU: 23           // Columna X (Código remoto)
};

// ==========================================================================
// MENÚ DE EJECUCIÓN
// ==========================================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("⚡ LAFINUR")
    .addItem("🏷️ Actualizar Rubros desde Maestro", "actualizarRubrosDesdeMaestro")
    .addItem("💵 Actualizar Precios desde Facturas", "actualizarPreciosDesdeFacturas")
    .addToUi();
}

// ==========================================================================
// 1. CARGA DE RUBROS (MAESTRO -> GESTION)
// ==========================================================================
function actualizarRubrosDesdeMaestro() {
  const ssLocal = SpreadsheetApp.getActiveSpreadsheet();
  const hojaLocal = ssLocal.getSheetByName(CONFIG_LOCAL.HOJA_GESTION) || ssLocal.getActiveSheet();
  
  // 1. Cargar catálogo Maestro en memoria
  let ssMaestro;
  try {
    ssMaestro = SpreadsheetApp.openById(ID_MAESTRO_PRODUCTOS);
  } catch (err) {
    SpreadsheetApp.getUi().alert("❌ Error abriendo el Maestro de Productos: " + err.message);
    return;
  }
  
  const hojaMaestro = ssMaestro.getSheetByName(MAESTRO_COLS.HOJA);
  if (!hojaMaestro) {
    SpreadsheetApp.getUi().alert(`❌ No se encontró la hoja '${MAESTRO_COLS.HOJA}' en el Maestro.`);
    return;
  }
  
  const ultFilaMaestro = hojaMaestro.getLastRow();
  if (ultFilaMaestro < 2) {
    SpreadsheetApp.getUi().alert("❌ La hoja Productos del Maestro está vacía.");
    return;
  }
  
  // Leemos de A hasta D (4 columnas)
  const datosMaestro = hojaMaestro.getRange(2, 1, ultFilaMaestro - 1, 4).getValues();
  const mapaRubros = new Map();
  
  for (let i = 0; i < datosMaestro.length; i++) {
    const sku = String(datosMaestro[i][MAESTRO_COLS.COL_SKU] || '').trim().toUpperCase();
    const rubro = String(datosMaestro[i][MAESTRO_COLS.COL_RUBRO] || '').trim();
    if (sku && rubro) {
      mapaRubros.set(sku, rubro);
    }
  }
  
  // 2. Procesar hoja local Gestion
  const filaInicio = CONFIG_LOCAL.FILA_INICIO;
  const ultFilaLocal = hojaLocal.getLastRow();
  
  if (ultFilaLocal < filaInicio) {
    SpreadsheetApp.getUi().alert("No hay filas para procesar en Gestion.");
    return;
  }
  
  const numFilas = ultFilaLocal - filaInicio + 1;
  const codigosLocales = hojaLocal.getRange(filaInicio, CONFIG_LOCAL.COL_CODIGO, numFilas, 1).getValues();
  const rubrosActuales = hojaLocal.getRange(filaInicio, CONFIG_LOCAL.COL_RUBRO_DESTINO, numFilas, 1).getValues();
  
  let rubrosActualizados = 0;
  
  for (let i = 0; i < numFilas; i++) {
    const codLocal = String(codigosLocales[i][0] || '').trim().toUpperCase();
    
    if (!codLocal) continue;
    
    if (mapaRubros.has(codLocal)) {
      rubrosActuales[i][0] = mapaRubros.get(codLocal);
      rubrosActualizados++;
    }
  }
  
  // 3. Volcado en Columna A
  hojaLocal.getRange(filaInicio, CONFIG_LOCAL.COL_RUBRO_DESTINO, numFilas, 1).setValues(rubrosActuales);
  SpreadsheetApp.flush();
  
  SpreadsheetApp.getActiveSpreadsheet().toast(
    `Se actualizaron ${rubrosActualizados} rubros desde el Maestro.`,
    "🏷️ Rubros Completados",
    5
  );
}

// ==========================================================================
// 2. CARGA DE PRECIOS DESDE FACTURAS REMOTAS
// ==========================================================================
function actualizarPreciosDesdeFacturas() {
  const ssLocal = SpreadsheetApp.getActiveSpreadsheet();
  const hojaLocal = ssLocal.getSheetByName(CONFIG_LOCAL.HOJA_GESTION) || ssLocal.getActiveSheet();
  
  // 1. Obtener período dinámico (mes y año del archivo o pestaña)
  const textoReferencia = ssLocal.getName() + " " + hojaLocal.getName();
  const periodo = detectarPeriodo(textoReferencia);
  
  // Ventana: mes actual + 3 meses hacia atrás (4 meses total)
  const mesesValidos = [];
  for (let i = 0; i < 4; i++) {
    let m = periodo.mes - i;
    let a = periodo.ano;
    if (m <= 0) {
      m += 12;
      a -= 1;
    }
    mesesValidos.push({ mes: m, ano: a });
  }

  // 2. Cargar datos remotos
  let ssRemoto;
  try {
    ssRemoto = SpreadsheetApp.openById(ID_FACTURAS_ORIGEN);
  } catch (err) {
    SpreadsheetApp.getUi().alert("❌ Error abriendo el archivo de facturas: " + err.message);
    return;
  }

  const hojaFacturas = ssRemoto.getSheetByName("Facturas");
  
  if (!hojaFacturas) {
    SpreadsheetApp.getUi().alert("❌ No se encontró la pestaña 'Facturas' en el archivo remoto.");
    return;
  }
  
  const uFilaRemota = hojaFacturas.getLastRow();
  if (uFilaRemota < 2) {
    SpreadsheetApp.getUi().alert("❌ La hoja Facturas remota no tiene datos.");
    return;
  }
  
  const dataFacturas = hojaFacturas.getRange(2, 1, uFilaRemota - 1, 24).getValues();
  const mapaUltimaFactura = new Map();
  
  for (let i = 0; i < dataFacturas.length; i++) {
    const fila = dataFacturas[i];
    const codigoRaw = String(fila[FACTURAS_COLS.CODIGO_SKU] || '').trim().toUpperCase();
    const fechaRaw = fila[FACTURAS_COLS.FECHA_INGRESO];
    const precio = parseNumero(fila[FACTURAS_COLS.PRECIO_UNITARIO]);
    
    if (!codigoRaw || !fechaRaw || precio <= 0) continue;
    
    let fechaObj = null;
    if (fechaRaw instanceof Date) {
      fechaObj = fechaRaw;
    } else {
      const partes = String(fechaRaw).split(/[\/\-]/);
      if (partes.length >= 2) {
        const dia = parseInt(partes[0], 10);
        const mes = parseInt(partes[1], 10) - 1;
        const anio = partes[2] ? parseInt(partes[2], 10) : periodo.ano;
        fechaObj = new Date(anio, mes, dia);
      }
    }
    
    if (!fechaObj || isNaN(fechaObj.getTime())) continue;
    
    const mFactura = fechaObj.getMonth() + 1;
    const aFactura = fechaObj.getFullYear();
    
    const entraEnVentana = mesesValidos.some(p => p.mes === mFactura && p.ano === aFactura);
    if (!entraEnVentana) continue;
    
    const diaTxt = String(fechaObj.getDate()).padStart(2, '0');
    const mesTxt = String(mFactura).padStart(2, '0');
    const fechaTexto = `${diaTxt}/${mesTxt}/${fechaObj.getFullYear()}`;
    
    if (!mapaUltimaFactura.has(codigoRaw)) {
      mapaUltimaFactura.set(codigoRaw, { fecha: fechaObj, fechaTexto: fechaTexto, precio: precio, idx: i });
    } else {
      const prev = mapaUltimaFactura.get(codigoRaw);
      if (fechaObj > prev.fecha || (fechaObj.getTime() === prev.fecha.getTime() && i >= prev.idx)) {
        mapaUltimaFactura.set(codigoRaw, { fecha: fechaObj, fechaTexto: fechaTexto, precio: precio, idx: i });
      }
    }
  }

  // 3. Procesar hoja local
  const filaInicio = CONFIG_LOCAL.FILA_INICIO;
  const uFilaLocal = hojaLocal.getLastRow();
  
  if (uFilaLocal < filaInicio) {
    SpreadsheetApp.getUi().alert("No hay filas para procesar.");
    return;
  }
  
  const totalFilas = uFilaLocal - filaInicio + 1;
  
  const codigosLocales = hojaLocal.getRange(filaInicio, CONFIG_LOCAL.COL_CODIGO, totalFilas, 1).getValues();
  const preciosViejosF = hojaLocal.getRange(filaInicio, CONFIG_LOCAL.COL_PRECIO_VIEJO, totalFilas, 1).getValues();
  const matrizG = hojaLocal.getRange(filaInicio, CONFIG_LOCAL.COL_PRECIO_DESTINO, totalFilas, 1).getValues(); 
  const matrizK = hojaLocal.getRange(filaInicio, CONFIG_LOCAL.COL_FECHA_DESTINO, totalFilas, 1).getValues();  
  
  let encontrados = 0;
  let noEncontrados = 0;
  
  for (let i = 0; i < totalFilas; i++) {
    const cod = String(codigosLocales[i][0] || '').trim().toUpperCase();
    const precioViejo = preciosViejosF[i][0];
    
    if (!cod) {
      matrizK[i][0] = "";
      continue;
    }
    
    if (mapaUltimaFactura.has(cod)) {
      const match = mapaUltimaFactura.get(cod);
      matrizG[i][0] = match.precio;
      matrizK[i][0] = match.fechaTexto;
      encontrados++;
    } else {
      matrizG[i][0] = precioViejo;
      const textoPrecio = (precioViejo !== "" && precioViejo !== null && precioViejo !== undefined) ? precioViejo : "-";
      matrizK[i][0] = `No encontrado: ${textoPrecio}`;
      noEncontrados++;
    }
  }

  // 4. Escritura masiva en G y K
  hojaLocal.getRange(filaInicio, CONFIG_LOCAL.COL_PRECIO_DESTINO, totalFilas, 1).setValues(matrizG);
  hojaLocal.getRange(filaInicio, CONFIG_LOCAL.COL_FECHA_DESTINO, totalFilas, 1).setValues(matrizK);
  
  SpreadsheetApp.flush();
  SpreadsheetApp.getActiveSpreadsheet().toast(`Listo: ${encontrados} encontrados, ${noEncontrados} mantenidos de Col F.`, "Éxito", 5);
}

// ==========================================================================
// UTILITARIOS
// ==========================================================================
function detectarPeriodo(nombre) {
  const str = String(nombre || '').toLowerCase();
  const meses = {
    "enero": 1, "febrero": 2, "marzo": 3, "abril": 4, "mayo": 5, "junio": 6,
    "julio": 7, "agosto": 8, "septiembre": 9, "octubre": 10, "noviembre": 11, "diciembre": 12
  };
  
  let mes = null;
  for (const m in meses) {
    if (str.includes(m)) {
      mes = meses[m];
      break;
    }
  }
  
  let ano = 2026;
  const matchAno = str.match(/(20\d{2})/);
  if (matchAno) ano = parseInt(matchAno[1], 10);
  
  return { mes: mes || 9, ano: ano };
}

function parseNumero(val) {
  if (val === null || val === undefined || val === "") return 0;
  if (typeof val === "number") return isNaN(val) ? 0 : val;
  const clean = val.toString().trim().replace("$", "").replace(/\./g, "").replace(",", ".");
  const num = parseFloat(clean);
  return isNaN(num) ? 0 : num;
}