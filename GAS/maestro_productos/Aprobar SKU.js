/**
 * SISTEMA INTEGRAL DE CATALOGACIÓN AUTOMÁTICA (DOBLE IMPACTO)
 * Impacta en pestañas: 'Presentacion' y 'Productos'
 */

const CONFIG_MAESTRO = {
  HOJA_SOLICITUDES: "Solicitudes",
  HOJA_PRESENTACION: "Presentacion",
  HOJA_PRODUCTOS: "Productos",
  
  // Índices en la hoja Solicitudes (1 = A, 2 = B...)
  COL_TIPO_SOLICITUD: 3,  // C: PRODUCTO / PROVEEDOR
  COL_TIPO_INSUMO: 4,     // D: Tipo (Materia prima, Limpieza...)
  COL_RUBRO: 5,           // E: Rubro Propuesto
  COL_PROD_BASE: 6,       // F: Producto Base
  COL_PROD_PROP: 7,       // G: Producto Propuesto (Presentación)
  COL_CONTENIDO: 8,       // H: Contenido
  COL_UM: 9,              // I: UM Contenido
  COL_FACTOR: 10,         // J: Factor Conversión
  COL_PROVEEDOR: 11,      // K: Proveedor
  COL_ESTADO: 12,         // L: Estado (PENDIENTE / APROBADO / RECHAZADO)
  COL_SKU: 13             // M: SKU Asignado
};

/**
 * Evento al editar el desplegable de Estado (Columna L)
 */
function onEdit(e) {
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();
  const row = e.range.getRow();
  const col = e.range.getColumn();

  // Solo ejecuta en la columna L (12) a partir de la fila 2 en Solicitudes
  if (row < 2 || col !== CONFIG_MAESTRO.COL_ESTADO || sheet.getName() !== CONFIG_MAESTRO.HOJA_SOLICITUDES) return;

  const nuevoEstado = String(e.value || e.range.getValue() || "").trim().toUpperCase();

  if (nuevoEstado === "APROBADO") {
    ejecutarAprobacionFila(sheet, row);
  }
}

/**
 * Procesa la aprobación de una fila individual
 */
function ejecutarAprobacionFila(hojaSol, row) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaPres = ss.getSheetByName(CONFIG_MAESTRO.HOJA_PRESENTACION);
  const hojaProd = ss.getSheetByName(CONFIG_MAESTRO.HOJA_PRODUCTOS);

  if (!hojaPres || !hojaProd) {
    SpreadsheetApp.getUi().alert("❌ Error: No se encontraron las pestañas 'Presentacion' o 'Productos'.");
    return;
  }

  // Leer la fila completa (Columnas A a M -> 13 columnas)
  const fila = hojaSol.getRange(row, 1, 1, 13).getValues()[0];

  const tipoSolicitud = String(fila[CONFIG_MAESTRO.COL_TIPO_SOLICITUD - 1] || "").trim().toUpperCase();
  const tipoInsumo    = String(fila[CONFIG_MAESTRO.COL_TIPO_INSUMO - 1] || "Materia prima").trim();
  const rubro         = String(fila[CONFIG_MAESTRO.COL_RUBRO - 1] || "General").trim();
  const prodPropuesto = String(fila[CONFIG_MAESTRO.COL_PROD_PROP - 1] || "").trim();
  const prodBase      = String(fila[CONFIG_MAESTRO.COL_PROD_BASE - 1] || prodPropuesto).trim();
  const contenido     = fila[CONFIG_MAESTRO.COL_CONTENIDO - 1] !== "" ? fila[CONFIG_MAESTRO.COL_CONTENIDO - 1] : 1;
  const umContenido   = String(fila[CONFIG_MAESTRO.COL_UM - 1] || "UN").trim().toUpperCase();
  const factorConv    = fila[CONFIG_MAESTRO.COL_FACTOR - 1] !== "" ? fila[CONFIG_MAESTRO.COL_FACTOR - 1] : 1;
  const skuActual     = String(fila[CONFIG_MAESTRO.COL_SKU - 1] || "").trim();

  // 1. Caso PRODUCTO
  if (tipoSolicitud === "PRODUCTO") {
    if (!prodBase && !prodPropuesto) {
      SpreadsheetApp.getUi().alert("⚠️ Debe completar al menos el Producto Base o el Producto Propuesto.");
      return;
    }

    // Obtener o generar SKU correlativo
    let skuFinal = skuActual;
    if (!skuFinal) {
      const ultimoSku = obtenerMayorSku([hojaPres, hojaProd]);
      skuFinal = generarSiguienteSku(ultimoSku);
      hojaSol.getRange(row, CONFIG_MAESTRO.COL_SKU).setValue(skuFinal);
    }

    // A. Insertar en pestaña 'Presentacion':
    // [SKU, Producto Base, Presentación/Artículo, Tipo, Rubro/Categoría, Contenido, UM Contenido, Factor Conversión, Activa]
    hojaPres.appendRow([
      skuFinal,
      prodBase,
      prodPropuesto || prodBase,
      tipoInsumo,
      rubro,
      contenido,
      umContenido,
      factorConv,
      "Sí"
    ]);

    // B. Insertar en pestaña 'Productos' (solo si no existe ya el SKU):
    const yaExisteEnProductos = verificarExisteSku(hojaProd, skuFinal);
    if (!yaExisteEnProductos) {
      // [SKU, Producto, Tipo, Rubro / Categoría, Precio, Activo (UM), Inventario (Factor)]
      hojaProd.appendRow([
        skuFinal,
        prodBase,
        tipoInsumo,
        rubro,
        "",          // Precio vacío (lo calcula el algoritmo de compras nocturno)
        umContenido, // UM
        factorConv   // Cantidad base inventario
      ]);
    }

    // Formato visual de confirmación en la fila de Solicitudes
    hojaSol.getRange(row, 1, 1, 13).setBackground("#d9ead3");
  } 
  // 2. Caso PROVEEDOR
  else if (tipoSolicitud === "PROVEEDOR") {
    hojaSol.getRange(row, CONFIG_MAESTRO.COL_SKU).setValue("N/A");
    hojaSol.getRange(row, 1, 1, 13).setBackground("#d9ead3");
  }
}

/**
 * Busca el SKU correlativo más alto registrado en el catálogo
 */
function obtenerMayorSku(hojas) {
  let mayorSku = "AA0000";

  hojas.forEach(hoja => {
    const ultFila = hoja.getLastRow();
    if (ultFila >= 2) {
      const skus = hoja.getRange(2, 1, ultFila - 1, 1).getValues().flat();
      for (let s of skus) {
        const str = String(s || "").trim().toUpperCase();
        if (/^[A-Z]{2}\d{4}$/.test(str) && str > mayorSku) {
          mayorSku = str;
        }
      }
    }
  });

  return mayorSku;
}

/**
 * Verifica si un SKU ya existe en la hoja Productos
 */
function verificarExisteSku(hojaProd, sku) {
  const ultFila = hojaProd.getLastRow();
  if (ultFila < 2) return false;
  const skus = hojaProd.getRange(2, 1, ultFila - 1, 1).getValues().flat().map(String);
  return skus.includes(sku);
}

/**
 * Generador alfanumérico correlativo: AA0001 -> AA9999 -> AB0001...
 */
function generarSiguienteSku(sku) {
  let prefijo = sku.substring(0, 2);
  let num = parseInt(sku.substring(2), 10) + 1;

  if (num > 9999) {
    num = 1;
    let c1 = prefijo.charAt(0);
    let c2 = prefijo.charAt(1);

    if (c2 === 'Z') {
      c1 = String.fromCharCode(c1.charCodeAt(0) + 1);
      c2 = 'A';
    } else {
      c2 = String.fromCharCode(c2.charCodeAt(0) + 1);
    }
    prefijo = `${c1}${c2}`;
  }

  return `${prefijo}${num.toString().padStart(4, "0")}`;
}