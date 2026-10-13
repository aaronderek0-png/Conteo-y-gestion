// ==========================================================================
// CABELLO GESTIÓN: CRUCE SKU Y RUBROS PARA SHEET6
// ==========================================================================

const ARCHIVOS_CONFIG_SHEET6 = {
  ID_FACTURAS_CABELLO: "1zAxYyNB98Sb1-D0HtfokESXc6FrDS96EbFYyMmsseIE",
  NOMBRE_HOJA_FACTURAS: "Facturas"
};

const COLS_SHEET6 = {
  FILA_INICIO: 2,
  HOJA_DESTINO: "Sheet6",
  
  CODIGO: 3,             // Columna C (Código SKU en Sheet6)
  PRECIO_ANTERIOR: 5,    // Columna E ($ unitario anterior en Sheet6)
  PRECIO_ACTUAL: 6,      // Columna F ($ unitario actual en Sheet6)
  VALORIZADO: 15,        // Columna O (Valorizado cantidad * precio)
  FECHA_ULT_FACT: 17     // Columna Q/R (Fecha última factura)
};

const FACTURAS_COLS_SHEET6 = {
  COL_FECHA_FACTURA: 1,  // Columna B (Fecha Factura)
  COL_FECHA_INGRESO: 2,  // Columna C (Fecha Entrega/Ingreso)
  COL_PRECIO_UN_INV: 14, // Columna O (Precio unitario inventario)
  COL_CODIGO_W: 22,      // Columna W (Código SKU en Cabello)
  COL_CODIGO_X: 23       // Columna X (Código SKU alternativo)
};

function obtenerVentanaMesesDesdeTitulo_Sheet6_() {
  const nombreLibro = SpreadsheetApp.getActiveSpreadsheet().getName().toLowerCase();
  
  const meses = {
    "enero": 1, "febrero": 2, "marzo": 3, "abril": 4, "mayo": 5, "junio": 6,
    "julio": 7, "agosto": 8, "septiembre": 9, "setiembre": 9, 
    "octubre": 10, "noviembre": 11, "diciembre": 12
  };

  let mesTarget = null;
  let nombreMesStr = "";

  for (const [mTexto, mNum] of Object.entries(meses)) {
    if (nombreLibro.includes(mTexto)) {
      mesTarget = mNum;
      nombreMesStr = mTexto.toUpperCase();
      break;
    }
  }

  let anoTarget = 2026;
  const matchAno = nombreLibro.match(/(20\d{2})/);
  if (matchAno) anoTarget = parseInt(matchAno[1], 10);

  if (!mesTarget) {
    mesTarget = 9;
    nombreMesStr = "SEPTIEMBRE";
  }

  const fechaTopeSuperior = new Date(anoTarget, mesTarget, 0, 23, 59, 59, 999);
  const fechaTopeInferior = new Date(anoTarget, mesTarget - 4, 1, 0, 0, 0, 0);

  return {
    mesNum: mesTarget,
    nombreMes: nombreMesStr,
    ano: anoTarget,
    tsTopeSuperior: fechaTopeSuperior.getTime(),
    tsTopeInferior: fechaTopeInferior.getTime()
  };
}

function norm_Sheet6_(texto) {
  if (texto === null || texto === undefined) return "";
  return texto.toString().trim().toUpperCase()
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
    .replace(/[^A-Z0-9 ]/g, "")
    .replace(/\s+/g, " ");
}

function parseNum_Sheet6_(val) {
  if (val === null || val === undefined || val === "") return 0;
  if (typeof val === "number") return isNaN(val) ? 0 : val;
  const limpio = val.toString().trim().replace(/[^0-9.,]/g, "");
  
  let numStr = limpio;
  if (numStr.includes(".") && numStr.includes(",")) {
    numStr = numStr.replace(/\./g, "").replace(",", ".");
  } else if (numStr.includes(".")) {
    const partes = numStr.split(".");
    if (partes.length === 2 && partes[1].length === 3) {
      numStr = numStr.replace(".", "");
    }
  } else if (numStr.includes(",")) {
    numStr = numStr.replace(",", ".");
  }
  const num = parseFloat(numStr);
  return isNaN(num) ? 0 : num;
}

function parseFecha_Sheet6_(fechaIngresoRaw, fechaFacturaRaw, anoReferencia) {
  let val = (fechaIngresoRaw !== "" && fechaIngresoRaw !== null && fechaIngresoRaw !== undefined)
            ? fechaIngresoRaw
            : fechaFacturaRaw;

  if (!val) return null;

  if (val instanceof Date && !isNaN(val.getTime())) {
    const d = val.getDate().toString().padStart(2, "0");
    const m = (val.getMonth() + 1).toString().padStart(2, "0");
    const a = val.getFullYear();
    return {
      timestamp: val.getTime(),
      fechaTexto: `${d}/${m}/${a}`,
      mes: val.getMonth() + 1
    };
  }

  const s = val.toString().trim();
  if (s.includes("/")) {
    const partes = s.split("/");
    const d = parseInt(partes[0], 10);
    const m = parseInt(partes[1], 10);
    let a = partes.length >= 3 ? parseInt(partes[2], 10) : anoReferencia;
    if (a < 100) a += 2000;
    
    if (!isNaN(d) && !isNaN(m)) {
      const fObj = new Date(a, m - 1, d);
      return {
        timestamp: fObj.getTime(),
        fechaTexto: `${d.toString().padStart(2, "0")}/${m.toString().padStart(2, "0")}/${a}`,
        mes: m
      };
    }
  }
  return null;
}

function buscarHojaFacturas_Sheet6_(ss) {
  const hoja = ss.getSheetByName(ARCHIVOS_CONFIG_SHEET6.NOMBRE_HOJA_FACTURAS);
  if (hoja) return hoja;

  for (const h of ss.getSheets()) {
    const n = norm_Sheet6_(h.getName());
    if (n === "FACTURAS" || n.includes("FACTURA") || n.includes("CABELLO")) {
      return h;
    }
  }
  return ss.getSheets()[0];
}

// --------------------------------------------------------------------------
// PROCESO 1: ACTUALIZAR PRECIOS EN SHEET6
// --------------------------------------------------------------------------
function actualizarPreciosCabelloSheet6() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let ventana;

  try {
    ventana = obtenerVentanaMesesDesdeTitulo_Sheet6_();
  } catch(e) {
    SpreadsheetApp.getUi().alert(e.message);
    return;
  }

  const mapaUltimoPrecio = new Map();

  try {
    const ssFacturas = SpreadsheetApp.openById(ARCHIVOS_CONFIG_SHEET6.ID_FACTURAS_CABELLO);
    const hojaFacturas = buscarHojaFacturas_Sheet6_(ssFacturas);

    const uFilaF = hojaFacturas.getLastRow();
    if (uFilaF >= 2) {
      const datosFacturas = hojaFacturas.getRange(2, 1, uFilaF - 1, 24).getValues();

      for (let i = 0; i < datosFacturas.length; i++) {
        const fila = datosFacturas[i];
        
        let rawSku = fila[FACTURAS_COLS_SHEET6.COL_CODIGO_W] || fila[FACTURAS_COLS_SHEET6.COL_CODIGO_X] || "";
        let codigoSku = String(rawSku).trim().toUpperCase();

        if (!codigoSku || codigoSku === "SIN SKU" || codigoSku.includes("#N/A") || codigoSku.includes("ERROR")) {
          continue;
        }

        const precioUnitario = parseNum_Sheet6_(fila[FACTURAS_COLS_SHEET6.COL_PRECIO_UN_INV]);
        if (precioUnitario <= 0) continue;

        const infoFecha = parseFecha_Sheet6_(fila[FACTURAS_COLS_SHEET6.COL_FECHA_INGRESO], fila[FACTURAS_COLS_SHEET6.COL_FECHA_FACTURA], ventana.ano);
        if (!infoFecha) continue;

        const t = infoFecha.timestamp;

        if (t <= ventana.tsTopeSuperior && t >= ventana.tsTopeInferior) {
          if (!mapaUltimoPrecio.has(codigoSku) || t > mapaUltimoPrecio.get(codigoSku).timestamp || (t === mapaUltimoPrecio.get(codigoSku).timestamp && i >= mapaUltimoPrecio.get(codigoSku).idx)) {
            mapaUltimoPrecio.set(codigoSku, {
              precio: precioUnitario,
              fechaTexto: infoFecha.fechaTexto,
              timestamp: t,
              esDelMesActual: (infoFecha.mes === ventana.mesNum),
              idx: i
            });
          }
        }
      }
    }
  } catch (err) {
    SpreadsheetApp.getUi().alert("❌ Error en Facturas Cabello:\n" + err.message);
    return;
  }

  let hojaLocal = ss.getSheetByName(COLS_SHEET6.HOJA_DESTINO);
  if (!hojaLocal) {
    SpreadsheetApp.getUi().alert(`❌ No se encontró la hoja "${COLS_SHEET6.HOJA_DESTINO}".`);
    return;
  }

  const uFilaLocal = hojaLocal.getLastRow();
  if (uFilaLocal < COLS_SHEET6.FILA_INICIO) {
    SpreadsheetApp.getUi().alert(`No hay filas en la hoja "${COLS_SHEET6.HOJA_DESTINO}".`);
    return;
  }

  const cantFilas = uFilaLocal - COLS_SHEET6.FILA_INICIO + 1;
  const codigos = hojaLocal.getRange(COLS_SHEET6.FILA_INICIO, COLS_SHEET6.CODIGO, cantFilas, 1).getValues();
  const preciosAnteriores = hojaLocal.getRange(COLS_SHEET6.FILA_INICIO, COLS_SHEET6.PRECIO_ANTERIOR, cantFilas, 1).getValues();

  const salidaPrecios = [];
  const salidaAudit = [];

  let countMesActual = 0;
  let countMesesPrevios = 0;
  let countColF = 0;
  let countFaltantes = 0;

  for (let i = 0; i < cantFilas; i++) {
    const cod = String(codigos[i][0] || "").trim().toUpperCase();
    const precioViejo = parseNum_Sheet6_(preciosAnteriores[i][0]);

    if (cod && mapaUltimoPrecio.has(cod)) {
      const info = mapaUltimoPrecio.get(cod);
      salidaPrecios.push([info.precio]);
      
      if (info.esDelMesActual) {
        salidaAudit.push([`Factura: ${info.fechaTexto}`]);
        countMesActual++;
      } else {
        salidaAudit.push([`Histórico factura: ${info.fechaTexto}`]);
        countMesesPrevios++;
      }
    } else if (precioViejo > 0) {
      salidaPrecios.push([precioViejo]);
      salidaAudit.push(["Precio anterior (Col E)"]);
      countColF++;
    } else {
      salidaPrecios.push([preciosAnteriores[i][0] || "No encontrado"]);
      salidaAudit.push(["Sin precio asignable"]);
      countFaltantes++;
    }
  }

  // Escribe en Columna F ($ unitario actual) y Columna R (Auditoría)
  hojaLocal.getRange(COLS_SHEET6.FILA_INICIO, COLS_SHEET6.PRECIO_ACTUAL, cantFilas, 1).setValues(salidaPrecios);
  hojaLocal.getRange(COLS_SHEET6.FILA_INICIO, COLS_SHEET6.FECHA_ULT_FACT, cantFilas, 1).setValues(salidaAudit);

  SpreadsheetApp.flush();

  SpreadsheetApp.getUi().alert(
    `✅ PRECIOS CABELLO ACTUALIZADOS EN ${COLS_SHEET6.HOJA_DESTINO}`,
    `Cruce: SKU Col C vs Facturas Col W/X\n\n` +
    `• Mes actual: ${countMesActual}\n` +
    `• Históricos (4 meses): ${countMesesPrevios}\n` +
    `• Heredados Col E: ${countColF}\n` +
    `• Sin compras: ${countFaltantes}`,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// --------------------------------------------------------------------------
// PROCESO 2: ACTUALIZAR RUBROS EN SHEET6 (C2:C -> A2:A)
// --------------------------------------------------------------------------
function actualizarRubrosPorCodigoSheet6() {
  const ID_MAESTRO = "1lOabw9U4ilRwE5r6QXQ2Jc1nqkKUK3Hi096xga0XFnI";
  const HOJA_MAESTRO_PRODUCTOS = "Productos";
  
  let mapaRubros = new Map();
  try {
    const ssMaestro = SpreadsheetApp.openById(ID_MAESTRO);
    const hojaMaestro = ssMaestro.getSheetByName(HOJA_MAESTRO_PRODUCTOS);
    
    if (!hojaMaestro) {
      throw new Error(`No se encontró la pestaña "${HOJA_MAESTRO_PRODUCTOS}" en el Maestro.`);
    }
    
    const ultimaFilaMaestro = hojaMaestro.getLastRow();
    if (ultimaFilaMaestro >= 2) {
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

  const ssLocal = SpreadsheetApp.getActiveSpreadsheet();
  const hojaLocal = ssLocal.getSheetByName(COLS_SHEET6.HOJA_DESTINO);
  if (!hojaLocal) {
    SpreadsheetApp.getUi().alert(`❌ No se encontró la hoja "${COLS_SHEET6.HOJA_DESTINO}".`);
    return;
  }

  const ultimaFilaLocal = hojaLocal.getLastRow();
  if (ultimaFilaLocal < 2) return;

  // En Sheet6 los códigos están en Columna C (columna 3)
  const codigosLocales = hojaLocal.getRange(2, COLS_SHEET6.CODIGO, ultimaFilaLocal - 1, 1).getValues();
  
  const resultadoRubros = codigosLocales.map(fila => {
    const cod = String(fila[0] || "").trim().toUpperCase();
    if (!cod) return [""]; 
    const rubroEncontrado = mapaRubros.get(cod);
    return [rubroEncontrado ? rubroEncontrado : "Obrador"];
  });

  // Escribe en Columna A (RUBRO)
  hojaLocal.getRange(2, 1, resultadoRubros.length, 1).setValues(resultadoRubros);
  SpreadsheetApp.getUi().alert(`✅ Rubros actualizados con éxito en "${COLS_SHEET6.HOJA_DESTINO}".`);
}