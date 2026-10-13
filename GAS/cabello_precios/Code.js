// ==========================================================================
// CABELLO GESTIÓN: ÚLTIMO PRECIO CON CRUCE ESTRICTO SKU (COL W / X)
// ==========================================================================

const ARCHIVOS_CONFIG = {
  ID_FACTURAS_CABELLO: "1zAxYyNB98Sb1-D0HtfokESXc6FrDS96EbFYyMmsseIE",
  NOMBRE_HOJA_FACTURAS: "Facturas"
};

const COLS = {
  FILA_INICIO: 2,
  HOJA_DESTINO: "Gestion", // Tu solapa de Gestión
  
  CODIGO: 4,             // Columna D (Código SKU local)
  PRECIO_ANTERIOR: 6,    // Columna F ($ unitario anterior)
  PRECIO_ACTUAL: 7,      // Columna G ($ unitario actual)
  VALORIZADO: 16,        // Columna P (Calculado por tu fórmula MAP / cantidad * precio)
  FECHA_ULT_FACT: 18     // Columna R (Auditoría / Fecha última factura)
};

// Índices base 0 para la hoja Facturas remota
const FACTURAS_COLS = {
  COL_FECHA_FACTURA: 1,  // Columna B (Fecha Factura)
  COL_FECHA_INGRESO: 2,  // Columna C (Fecha Entrega/Ingreso - Prioridad)
  COL_PRECIO_UN_INV: 14, // Columna O (Precio unitario inventario)
  COL_CODIGO_W: 22,      // Columna W (Código SKU en Cabello)
  COL_CODIGO_X: 23       // Columna X (Código SKU alternativo)
};

function obtenerVentanaMesesDesdeTitulo() {
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
    mesTarget = 9; // Fallback Septiembre
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

function norm(texto) {
  if (texto === null || texto === undefined) return "";
  return texto.toString().trim().toUpperCase()
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
    .replace(/[^A-Z0-9 ]/g, "")
    .replace(/\s+/g, " ");
}

function parseNum(val) {
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

function parseFecha(fechaIngresoRaw, fechaFacturaRaw, anoReferencia) {
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

function buscarHojaFacturas(ss) {
  const hoja = ss.getSheetByName(ARCHIVOS_CONFIG.NOMBRE_HOJA_FACTURAS);
  if (hoja) return hoja;

  for (const h of ss.getSheets()) {
    const n = norm(h.getName());
    if (n === "FACTURAS" || n.includes("FACTURA") || n.includes("CABELLO")) {
      return h;
    }
  }
  return ss.getSheets()[0];
}

// --------------------------------------------------------------------------
// PROCESO: ACTUALIZA G Y R (MATQUEO ESTRICTO SKU W o X -> D)
// --------------------------------------------------------------------------
function actualizarPreciosCabello() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let ventana;

  try {
    ventana = obtenerVentanaMesesDesdeTitulo();
  } catch(e) {
    SpreadsheetApp.getUi().alert(e.message);
    return;
  }

  const mapaUltimoPrecio = new Map();

  try {
    const ssFacturas = SpreadsheetApp.openById(ARCHIVOS_CONFIG.ID_FACTURAS_CABELLO);
    const hojaFacturas = buscarHojaFacturas(ssFacturas);

    const uFilaF = hojaFacturas.getLastRow();
    if (uFilaF >= 2) {
      // Leemos 24 columnas (A a X) para cubrir tanto W como X
      const datosFacturas = hojaFacturas.getRange(2, 1, uFilaF - 1, 24).getValues();

      for (let i = 0; i < datosFacturas.length; i++) {
        const fila = datosFacturas[i];
        
        // 1. Prioriza W (índice 22); si está vacío usa X (índice 23)
        let rawSku = fila[FACTURAS_COLS.COL_CODIGO_W] || fila[FACTURAS_COLS.COL_CODIGO_X] || "";
        let codigoSku = String(rawSku).trim().toUpperCase();

        if (!codigoSku || codigoSku === "SIN SKU" || codigoSku.includes("#N/A") || codigoSku.includes("ERROR")) {
          continue;
        }

        // 2. Precio unitario inventario desde Columna O (índice 14)
        const precioUnitario = parseNum(fila[FACTURAS_COLS.COL_PRECIO_UN_INV]);
        if (precioUnitario <= 0) continue;

        // 3. Fecha (prioridad ingreso en Col C; fallback factura en Col B)
        const infoFecha = parseFecha(fila[FACTURAS_COLS.COL_FECHA_INGRESO], fila[FACTURAS_COLS.COL_FECHA_FACTURA], ventana.ano);
        if (!infoFecha) continue;

        const t = infoFecha.timestamp;

        // Filtro de ventana temporal (4 meses hacia atrás)
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

  let hojaLocal = ss.getSheetByName(COLS.HOJA_DESTINO) || ss.getSheets()[0];
  const uFilaLocal = hojaLocal.getLastRow();

  if (uFilaLocal < COLS.FILA_INICIO) {
    SpreadsheetApp.getUi().alert("No hay filas en la hoja de Gestión.");
    return;
  }

  const cantFilas = uFilaLocal - COLS.FILA_INICIO + 1;
  const codigos = hojaLocal.getRange(COLS.FILA_INICIO, COLS.CODIGO, cantFilas, 1).getValues();
  const preciosAnteriores = hojaLocal.getRange(COLS.FILA_INICIO, COLS.PRECIO_ANTERIOR, cantFilas, 1).getValues();

  const salidaG = [];
  const salidaR = [];

  let countMesActual = 0;
  let countMesesPrevios = 0;
  let countColF = 0;
  let countFaltantes = 0;

  for (let i = 0; i < cantFilas; i++) {
    const cod = String(codigos[i][0] || "").trim().toUpperCase();
    const precioViejo = parseNum(preciosAnteriores[i][0]);

    if (cod && mapaUltimoPrecio.has(cod)) {
      const info = mapaUltimoPrecio.get(cod);
      salidaG.push([info.precio]);
      
      if (info.esDelMesActual) {
        salidaR.push([`Factura: ${info.fechaTexto}`]);
        countMesActual++;
      } else {
        salidaR.push([`Histórico factura: ${info.fechaTexto}`]);
        countMesesPrevios++;
      }
    } else if (precioViejo > 0) {
      salidaG.push([precioViejo]);
      salidaR.push(["Precio anterior (Col F)"]);
      countColF++;
    } else {
      salidaG.push([preciosAnteriores[i][0] || "No encontrado"]);
      salidaR.push(["Sin precio asignable"]);
      countFaltantes++;
    }
  }

  // Escribe G (Precio unitario actual) y R (Auditoría)
  hojaLocal.getRange(COLS.FILA_INICIO, COLS.PRECIO_ACTUAL, cantFilas, 1).setValues(salidaG);
  hojaLocal.getRange(COLS.FILA_INICIO, COLS.FECHA_ULT_FACT, cantFilas, 1).setValues(salidaR);

  // Recalculo inmediato
  SpreadsheetApp.flush();

  // Lee el total valorizado de la Columna P
  const valoresP = hojaLocal.getRange(COLS.FILA_INICIO, COLS.VALORIZADO, cantFilas, 1).getValues();
  let granTotalValorizado = 0;
  for (let i = 0; i < valoresP.length; i++) {
    granTotalValorizado += parseNum(valoresP[i][0]);
  }

  SpreadsheetApp.getUi().alert(
    `✅ PRECIOS CABELLO ACTUALIZADOS (${ventana.nombreMes} ${ventana.ano})`,
    `Tope máximo: Fin de ${ventana.nombreMes}\n` +
    `Cruce: SKU Col D (Gestión) vs SKU Col W/X (Facturas)\n\n` +
    `• Compras del mes actual: ${countMesActual}\n` +
    `• Compras históricas (hasta 4 meses): ${countMesesPrevios}\n` +
    `• Heredados de Columna F: ${countColF}\n` +
    `• Sin compras registradas: ${countFaltantes}\n\n` +
    `Total Valorizado (Col P): $ ${granTotalValorizado.toLocaleString("es-AR", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Gestión Cabello")
    .addItem("Actualizar Precios por Código SKU (Col W/X)", "actualizarPreciosCabello")
    .addToUi();
}