/**
 * Tablero de Control de Precios - ETL Conciliador de Precios (Versión 00_MAESTRO)
 * Congreso, Arcos, Lafinur, Cabello y Verdulería Raíces
 * Lógica: Control de desfasaje (±40%) -> Hoja Observaciones -> Data Entry automático a Productos con APROBADO
 * Memoria persistente anti-choclo con ScriptProperties (Clave: SKU + Fecha + Origen + Precio)
 */

const CONGRESO_SPREADSHEET_ID = '1udLEQpQ29avK4c6fNnQig-w02xIV82IgLe85yVJXaOY';
const ARCOS_SPREADSHEET_ID = '1frh93qB8Od5mrtAfMuObKwRW93p7kev6DNgmERqz6ho';
const LAFINUR_SPREADSHEET_ID = '1tv0uI4KsSEPRUecYfnwcDqpGbUDelGqWwe3Hp-fikOc';
const CABELLO_SPREADSHEET_ID = '1zAxYyNB98Sb1-D0HtfokESXc6FrDS96EbFYyMmsseIE';

const HOJA_MAESTRO_NAME = 'Productos';
const HOJA_EQUIVALENCIAS_NAME = 'Equivalencias';
const HOJA_OBSERVACIONES_NAME = 'Observaciones';
const HOJA_CONFIG_NAME = 'CONFIGURACION';

// Parámetro de desfasaje: 40%
const UMBRAL_DESFASE = 0.40;

// Carpeta compartida Raíces
const CARPETA_RAICES_ID = '18ZrxfnLYsYhKq4L1y9xPsbLv1GEJgj49';
const ARCHIVO_VERDULERIA_NAME = 'precios_verduleria_limpios';

// Fallback de índices base 0 para las hojas Facturas de sucursales
const FALLBACK_COL_FECHA_FACTURA = 1;  // B
const FALLBACK_COL_FECHA_INGRESO = 2;  // C (Prioridad)
const FALLBACK_COL_PRODUCTOS = 4;      // E
const FALLBACK_COL_PRECIO_UN_INV = 14; // O (PRECIO UN. INVENTARI)

// Índices base 0 fijos de la hoja Productos (Maestro)
const IDX_MAESTRO_SKU = 0;      // Col A (1)
const IDX_MAESTRO_NOMBRE = 1;   // Col B (2)
const IDX_MAESTRO_PRECIO = 4;   // Col E (5)
const IDX_MAESTRO_INVENT = 6;   // Col G (7)
const IDX_MAESTRO_PU = 7;       // Col H (8)
const IDX_MAESTRO_FECHA = 8;    // Col I (9)
const IDX_MAESTRO_ORIGEN = 9;   // Col J (10)

// Color institucional para feedback visual de data entry completado
const COLOR_VERDE_OK = '#d9ead3';

// =======================================================
// 🛠️ UTILITARIOS DE LIMPIEZA, NÚMEROS Y FECHAS
// =======================================================
function normalizar(texto) {
  return String(texto || '')
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}

function parsearNumero(val) {
  if (val === null || val === undefined || val === '') return 0;
  if (typeof val === 'number') return isNaN(val) ? 0 : val;
  
  let str = String(val).trim().replace(/[^0-9.,-]/g, '');
  if (!str) return 0;
  
  if (str.includes('.') && str.includes(',')) {
    str = str.replace(/\./g, '').replace(',', '.');
  } else if (str.includes(',')) {
    str = str.replace(',', '.');
  }
  let num = parseFloat(str);
  return isNaN(num) ? 0 : num;
}

function parsearFecha(fechaIngresoRaw, fechaFacturaRaw) {
  let raw = (fechaIngresoRaw !== '' && fechaIngresoRaw !== null && fechaIngresoRaw !== undefined) 
            ? fechaIngresoRaw 
            : fechaFacturaRaw;

  if (!raw) return null;

  if (raw instanceof Date && !isNaN(raw.getTime())) {
    return new Date(raw.getFullYear(), raw.getMonth(), raw.getDate());
  }

  if (typeof raw === 'string' && raw.includes('/')) {
    let partes = raw.trim().split('/');
    let dia = parseInt(partes[0], 10);
    let mes = parseInt(partes[1], 10) - 1;
    let anio = partes.length >= 3 ? parseInt(partes[2], 10) : 2026;
    if (anio < 100) anio += 2000;
    let f = new Date(anio, mes, dia);
    if (!isNaN(f.getTime())) return f;
  }

  let f = new Date(raw);
  return !isNaN(f.getTime()) ? new Date(f.getFullYear(), f.getMonth(), f.getDate()) : null;
}

function formatearFechaSheet(fecha) {
  if (!fecha) return '';
  if (typeof fecha === 'string') return fecha;
  if (!(fecha instanceof Date) || isNaN(fecha.getTime())) return '';
  let dia = fecha.getDate();
  let mes = fecha.getMonth() + 1;
  let anio = fecha.getFullYear();
  return `${dia}/${mes}/${anio}`;
}

// =======================================================
// 🧠 MEMORIA PERSISTENTE DE RESOLUCIONES (ANTI-CHOCLO)
// =======================================================
function armarClaveMemoria_(sku, fecha, origen, precio) {
  const s = normalizar(sku);
  const f = normalizar(fecha);
  const o = normalizar(origen);
  const p = parsearNumero(precio);
  return `OBS_${s}_${f}_${o}_${p}`;
}

function registrarEnMemoriaPersistente_(clave, estado) {
  const props = PropertiesService.getScriptProperties();
  props.setProperty(clave, estado);
}

function existeEnMemoriaPersistente_(clave) {
  const props = PropertiesService.getScriptProperties();
  return props.getProperty(clave) !== null;
}

function vaciarMemoriaResoluciones() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.alert(
    '⚠️ RESETEAR MEMORIA DE OBSERVACIONES',
    '¿Seguro que querés vaciar la caché de ítems aprobados y denegados? Las observaciones históricas volverán a evaluarse en la próxima conciliación.',
    ui.ButtonSet.YES_NO
  );
  if (resp !== ui.Button.YES) return;

  const props = PropertiesService.getScriptProperties();
  const keys = props.getKeys();
  let borradas = 0;
  for (let k of keys) {
    if (k.startsWith('OBS_')) {
      props.deleteProperty(k);
      borradas++;
    }
  }
  ui.alert(`✅ Se borraron ${borradas} registros de la memoria persistente.`);
}

// =======================================================
// ⚙️ GESTIÓN DE HOJA OBSERVACIONES Y DESPLEGABLE
// =======================================================
function obtenerOCrearHojaObservaciones(ss) {
  let hojaObs = ss.getSheetByName(HOJA_OBSERVACIONES_NAME);
  if (!hojaObs) {
    hojaObs = ss.insertSheet(HOJA_OBSERVACIONES_NAME);
    const encabezados = [
      ['SKU', 'Producto', 'Precio Anterior', 'Precio Nuevo', 'Variación %', 'Fecha Detección', 'Origen', 'ESTADO']
    ];
    hojaObs.getRange(1, 1, 1, encabezados[0].length).setValues(encabezados).setFontWeight('bold');
    hojaObs.setFrozenRows(1);
  }
  return hojaObs;
}

function aplicarReglaDesplegableEstado(hojaObs, fila) {
  const ss = hojaObs.getParent();
  const hojaConfig = ss.getSheetByName(HOJA_CONFIG_NAME);
  if (!hojaConfig) return;

  const rangoValores = hojaConfig.getRange('E2:E4');
  const regla = SpreadsheetApp.newDataValidation()
    .requireValueInRange(rangoValores, true)
    .setAllowInvalid(false)
    .build();

  hojaObs.getRange(fila, 8).setDataValidation(regla);
}

// =======================================================
// 🚀 LÓGICA PRINCIPAL DE CONCILIACIÓN
// =======================================================
function conciliarProcesos(opciones) {
  const ssActual = SpreadsheetApp.getActiveSpreadsheet();
  const hojaMaestro = ssActual.getSheetByName(HOJA_MAESTRO_NAME);
  const hojaEquiv = ssActual.getSheetByName(HOJA_EQUIVALENCIAS_NAME);

  if (!hojaMaestro) {
    Logger.log('No se encontró la hoja ' + HOJA_MAESTRO_NAME);
    return;
  }

  // 1. CARGAR EQUIVALENCIAS (Pestaña Equivalencias)
  let mapaEquivalenciasRaices = {};
  if (opciones.verduleria && hojaEquiv) {
    const datosEquiv = hojaEquiv.getDataRange().getValues();
    const cabecerasEquiv = datosEquiv[0].map(c => normalizar(c));

    let idxSkuEquiv = cabecerasEquiv.findIndex(c => c.includes('sku') || c.includes('cod'));
    let idxFacturaEquiv = cabecerasEquiv.findIndex(c => c.includes('factur') || c.includes('raic') || c.includes('art') || c.includes('nombre'));

    if (idxSkuEquiv === -1) idxSkuEquiv = 0;
    if (idxFacturaEquiv === -1) idxFacturaEquiv = 1;

    for (let i = 1; i < datosEquiv.length; i++) {
      let sku = String(datosEquiv[i][idxSkuEquiv]).trim().toUpperCase();
      let nomFactura = normalizar(datosEquiv[i][idxFacturaEquiv]);
      if (sku && nomFactura) {
        mapaEquivalenciasRaices[sku] = nomFactura;
      }
    }
  }

  // 2. PROCESAR FACTURAS (Congreso, Arcos, Lafinur, Cabello)
  let mapaNombreACodigo = {};
  let mapaPreciosPorCodigo = {};
  let mapaPreciosPorNombre = {};

  if (opciones.facturas) {
    const origenesFacturas = [
      { id: CONGRESO_SPREADSHEET_ID, nombre: 'Congreso' },
      { id: ARCOS_SPREADSHEET_ID, nombre: 'Arcos' },
      { id: LAFINUR_SPREADSHEET_ID, nombre: 'Lafinur' },
      { id: CABELLO_SPREADSHEET_ID, nombre: 'Cabello' }
    ];

    for (let orig of origenesFacturas) {
      try {
        const ssOrigen = SpreadsheetApp.openById(orig.id);
        const todasLasHojas = ssOrigen.getSheets();
        const hojaProductos = todasLasHojas.find(h => normalizar(h.getName()).includes('producto'));
        if (hojaProductos) {
          const datosProductos = hojaProductos.getDataRange().getValues();
          for (let i = 1; i < datosProductos.length; i++) {
            let cod = String(datosProductos[i][0] || '').trim().toUpperCase();
            let nomFactura = normalizar(datosProductos[i][1]);
            if (cod && nomFactura) {
              mapaNombreACodigo[nomFactura] = cod;
            }
          }
        }
      } catch (err) {
        Logger.log(`Error leyendo Productos de ${orig.nombre}: ${err.message}`);
      }
    }

    for (let orig of origenesFacturas) {
      try {
        const ssFacturas = SpreadsheetApp.openById(orig.id);
        const todasLasHojas = ssFacturas.getSheets();
        let hojaFacturas = todasLasHojas.find(h => normalizar(h.getName()).includes('factura'));
        if (!hojaFacturas) continue;

        const rows = hojaFacturas.getDataRange().getValues();
        const cabeceraFacturas = rows[0].map(c => normalizar(c));

        let colCodigoFactura = cabeceraFacturas.findIndex(c => c === 'codigo' || c === 'sku' || c.includes('cod'));
        if (colCodigoFactura === -1) colCodigoFactura = 22;

        let colProdFactura = cabeceraFacturas.findIndex(c => c.includes('producto') || c.includes('descripcion') || c.includes('articulo'));
        if (colProdFactura === -1) colProdFactura = FALLBACK_COL_PRODUCTOS;

        let colPrecioFactura = cabeceraFacturas.findIndex(c => c.includes('precio un. inventari') || c.includes('precio unitario') || c.includes('precio un') || c.includes('p.u'));
        if (colPrecioFactura === -1) colPrecioFactura = FALLBACK_COL_PRECIO_UN_INV;

        let colFecIngreso = cabeceraFacturas.findIndex(c => c.includes('ingreso') || c.includes('recepcion'));
        if (colFecIngreso === -1) colFecIngreso = FALLBACK_COL_FECHA_INGRESO;

        let colFecFactura = cabeceraFacturas.findIndex(c => c.includes('factura') && c.includes('fecha'));
        if (colFecFactura === -1) colFecFactura = FALLBACK_COL_FECHA_FACTURA;

        for (let i = 1; i < rows.length; i++) {
          let row = rows[i];
          if (!row || row.length <= colPrecioFactura) continue;

          let nombreEnFactura = normalizar(row[colProdFactura]);
          let precioUnitario = parsearNumero(row[colPrecioFactura]);
          if (precioUnitario <= 0) continue;

          let fecha = parsearFecha(row[colFecIngreso], row[colFecFactura]);
          let fechaComparar = fecha ? fecha.getTime() : 0;

          let codigoDirecto = '';
          if (row.length > colCodigoFactura && row[colCodigoFactura]) {
            codigoDirecto = String(row[colCodigoFactura]).trim().toUpperCase();
            if (codigoDirecto === 'SIN SKU' || codigoDirecto.includes('#N/A') || codigoDirecto.includes('VALUE')) {
              codigoDirecto = '';
            }
          }

          let codigoMatch = codigoDirecto || mapaNombreACodigo[nombreEnFactura] || null;

          if (codigoMatch) {
            if (!mapaPreciosPorCodigo[codigoMatch] || fechaComparar >= mapaPreciosPorCodigo[codigoMatch].fechaTime) {
              mapaPreciosPorCodigo[codigoMatch] = { 
                precio: precioUnitario, 
                fecha: fecha, 
                fechaTime: fechaComparar, 
                origen: orig.nombre 
              };
            }
          }

          if (nombreEnFactura) {
            if (!mapaPreciosPorNombre[nombreEnFactura] || fechaComparar >= mapaPreciosPorNombre[nombreEnFactura].fechaTime) {
              mapaPreciosPorNombre[nombreEnFactura] = { 
                precio: precioUnitario, 
                fecha: fecha, 
                fechaTime: fechaComparar, 
                origen: orig.nombre 
              };
            }
          }
        }
      } catch (err) {
        Logger.log(`Error en facturas de ${orig.nombre}: ${err.message}`);
      }
    }
  }

  // 3. PROCESAR VERDULERÍA RAÍCES
  let mapaPreciosVerduleria = {};
  if (opciones.verduleria) {
    try {
      const carpetaRaices = DriveApp.getFolderById(CARPETA_RAICES_ID);
      let archivoVerduleria = null;

      const archivosPorNombre = carpetaRaices.getFilesByName(ARCHIVO_VERDULERIA_NAME);
      if (archivosPorNombre.hasNext()) {
        archivoVerduleria = archivosPorNombre.next();
      } else {
        const todosLosArchivos = carpetaRaices.getFiles();
        while (todosLosArchivos.hasNext()) {
          let f = todosLosArchivos.next();
          if (f.getMimeType() === MimeType.GOOGLE_SHEETS || f.getName().toLowerCase().includes('verduleria') || f.getName().toLowerCase().includes('precio')) {
            archivoVerduleria = f;
            break;
          }
        }
      }

      if (archivoVerduleria) {
        const ssVerduleria = SpreadsheetApp.open(archivoVerduleria);
        const hojaPrecios = ssVerduleria.getSheets()[0];
        const datosPrecios = hojaPrecios.getDataRange().getValues();
        const cabecerasPrecios = datosPrecios[0].map(c => normalizar(c));

        let idxArticulo = cabecerasPrecios.findIndex(c => c.includes('articulo') || c.includes('art') || c.includes('prod') || c.includes('desc'));
        let idxPrecioKilo = cabecerasPrecios.findIndex(c => c.includes('precio_kilo') || c.includes('kilo') || c.includes('kg'));
        let idxPrecioBulto = cabecerasPrecios.findIndex(c => c.includes('precio_bulto') || c.includes('bulto') || c.includes('precio'));

        if (idxArticulo === -1) idxArticulo = 0;

        for (let i = 1; i < datosPrecios.length; i++) {
          let art = normalizar(datosPrecios[i][idxArticulo]);
          let precioKilo = idxPrecioKilo !== -1 ? parsearNumero(datosPrecios[i][idxPrecioKilo]) : 0;
          let precioBulto = idxPrecioBulto !== -1 ? parsearNumero(datosPrecios[i][idxPrecioBulto]) : 0;
          let precioFinal = precioKilo > 0 ? precioKilo : precioBulto;

          if (art && precioFinal > 0) {
            mapaPreciosVerduleria[art] = precioFinal;
          }
        }
        Logger.log(`Verdulería cargada: ${Object.keys(mapaPreciosVerduleria).length} productos.`);
      }
    } catch (err) {
      Logger.log(`Error leyendo archivo de verdulería: ${err.message}`);
    }
  }

  // 4. LEER OBSERVACIONES VISIBLES EN HOJA PARA EVITAR DUPLICADOS EN PENDIENTE
  const hojaObs = obtenerOCrearHojaObservaciones(ssActual);
  const ultFilaObs = hojaObs.getLastRow();
  const setObsEnHoja = new Set();
  
  if (ultFilaObs > 1) {
    const datosObsExistentes = hojaObs.getRange(2, 1, ultFilaObs - 1, 8).getValues();
    datosObsExistentes.forEach(f => {
      const skuEx = String(f[0] || '').trim().toUpperCase();
      const precioEx = parsearNumero(f[3]);
      const fechaEx = f[5];
      const origenEx = f[6];
      if (skuEx) {
        setObsEnHoja.add(armarClaveMemoria_(skuEx, fechaEx, origenEx, precioEx));
      }
    });
  }

  // 5. COMPARAR PRECIOS: IMPACTAR EN PRODUCTOS O ENVIAR A OBSERVACIONES
  const ultimaFila = hojaMaestro.getLastRow();
  if (ultimaFila < 2) return;

  const datosMaestro = hojaMaestro.getRange(1, 1, ultimaFila, 10).getValues();

  let columnaPrecio = [];
  let columnaPU = [];
  let columnaFecha = [];
  let columnaOrigen = [];
  let hoy = new Date();

  let filasObservadas = [];
  let actualizadosDirecto = 0;

  for (let i = 1; i < ultimaFila; i++) {
    let codigoMaestro = String(datosMaestro[i][IDX_MAESTRO_SKU]).trim().toUpperCase();
    let nombreMaestro = datosMaestro[i][IDX_MAESTRO_NOMBRE];
    let nombreNorm = normalizar(nombreMaestro);
    
    let factorInvent = parsearNumero(datosMaestro[i][IDX_MAESTRO_INVENT]);
    if (factorInvent <= 0) factorInvent = 1;

    let precioActual = parsearNumero(datosMaestro[i][IDX_MAESTRO_PRECIO]);
    let fechaActual = datosMaestro[i][IDX_MAESTRO_FECHA];
    let origenActual = datosMaestro[i][IDX_MAESTRO_ORIGEN];

    let precioEncontrado = null;
    let fechaEncontrada = null;
    let origenEncontrado = null;

    // A. Verdulería Raíces
    let nombreEnPdfRaices = mapaEquivalenciasRaices[codigoMaestro];
    if (opciones.verduleria && nombreEnPdfRaices && mapaPreciosVerduleria.hasOwnProperty(nombreEnPdfRaices)) {
      precioEncontrado = mapaPreciosVerduleria[nombreEnPdfRaices];
      fechaEncontrada = formatearFechaSheet(hoy);
      origenEncontrado = 'Raíces';
    }

    // B. Facturas
    if (!precioEncontrado && opciones.facturas) {
      let matchData = null;
      if (mapaPreciosPorCodigo.hasOwnProperty(codigoMaestro)) {
        matchData = mapaPreciosPorCodigo[codigoMaestro];
      } else if (nombreNorm && mapaPreciosPorNombre.hasOwnProperty(nombreNorm)) {
        matchData = mapaPreciosPorNombre[nombreNorm];
      }

      if (matchData) {
        precioEncontrado = matchData.precio;
        fechaEncontrada = formatearFechaSheet(matchData.fecha);
        origenEncontrado = matchData.origen;
      }
    }

    // C. Control del Desfasaje (±40%)
    let esDesfasado = false;
    if (precioEncontrado !== null && precioEncontrado > 0 && precioActual > 0) {
      let variacion = Math.abs((precioEncontrado - precioActual) / precioActual);
      if (variacion >= UMBRAL_DESFASE) {
        esDesfasado = true;
        
        // Clave única exacta: SKU (A) + Fecha (F) + Origen (G) + Precio Nuevo (D)
        const claveUnica = armarClaveMemoria_(codigoMaestro, fechaEncontrada, origenEncontrado, precioEncontrado);

        // BLINDAJE ANTI-PLAGA:
        // Solo ingresa si: NO está resuelto en memoria persistente Y NO está en la hoja física
        if (!existeEnMemoriaPersistente_(claveUnica) && !setObsEnHoja.has(claveUnica)) {
          let porcentajeFormat = ((precioEncontrado - precioActual) / precioActual * 100).toFixed(1) + '%';
          
          filasObservadas.push({
            sku: codigoMaestro,
            nombre: nombreMaestro,
            precioViejo: precioActual,
            precioNuevo: precioEncontrado,
            variacion: porcentajeFormat,
            fecha: fechaEncontrada,
            origen: origenEncontrado
          });
          setObsEnHoja.add(claveUnica);
        }
      }
    }

    if (precioEncontrado !== null && precioEncontrado > 0 && !esDesfasado) {
      let puCalculado = factorInvent > 0 ? (precioEncontrado / factorInvent) : precioEncontrado;
      columnaPrecio.push([precioEncontrado]);
      columnaPU.push([puCalculado]);
      columnaFecha.push([fechaEncontrada]);
      columnaOrigen.push([origenEncontrado]);
      actualizadosDirecto++;
    } else {
      let puRecalculado = factorInvent > 0 ? (precioActual / factorInvent) : precioActual;
      columnaPrecio.push([precioActual > 0 ? precioActual : (datosMaestro[i][IDX_MAESTRO_PRECIO] || '')]);
      columnaPU.push([puRecalculado]);
      columnaFecha.push([fechaActual || '']);
      columnaOrigen.push([origenActual || '']);
    }
  }

  let numFilas = ultimaFila - 1;

  hojaMaestro.getRange(2, IDX_MAESTRO_PRECIO + 1, numFilas, 1).setValues(columnaPrecio);
  hojaMaestro.getRange(2, IDX_MAESTRO_PU + 1, numFilas, 1).setValues(columnaPU);
  hojaMaestro.getRange(2, IDX_MAESTRO_FECHA + 1, numFilas, 1).setValues(columnaFecha);
  hojaMaestro.getRange(2, IDX_MAESTRO_ORIGEN + 1, numFilas, 1).setValues(columnaOrigen);

  Logger.log(`✅ Conciliación terminada: ${actualizadosDirecto} productos actualizados directamente.`);

  if (filasObservadas.length > 0) {
    let filaInicio = hojaObs.getLastRow() + 1;
    let filasAInsertar = filasObservadas.map(f => [
      f.sku,
      f.nombre,
      f.precioViejo,
      f.precioNuevo,
      f.variacion,
      f.fecha,
      f.origen,
      'PENDIENTE'
    ]);

    hojaObs.getRange(filaInicio, 1, filasAInsertar.length, 8).setValues(filasAInsertar);

    for (let r = filaInicio; r < filaInicio + filasAInsertar.length; r++) {
      aplicarReglaDesplegableEstado(hojaObs, r);
    }
    Logger.log(`⚠️ Se enviaron ${filasObservadas.length} artículos nuevos a Observaciones.`);
  }
}

// =======================================================
// 🔁 ONEDIT Y DATA ENTRY AUTOMÁTICO
// =======================================================
function onEdit(e) {
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();
  if (sheet.getName() !== HOJA_OBSERVACIONES_NAME) return;

  const col = e.range.getColumn();
  const row = e.range.getRow();

  if (col === 8 && row > 1) {
    const valorEstado = String(e.range.getValue() || '').trim().toUpperCase();
    if (valorEstado === 'APROBADO') {
      procesarFilaAprobada(sheet, row);
    }
  }
}

function procesarFilaAprobada(hojaObs, fila) {
  const ss = hojaObs.getParent();
  const hojaProductos = ss.getSheetByName(HOJA_MAESTRO_NAME);
  if (!hojaProductos) return;

  const datosFila = hojaObs.getRange(fila, 1, 1, 8).getValues()[0];
  const skuObs = String(datosFila[0] || '').trim().toUpperCase();
  const precioNuevo = parsearNumero(datosFila[3]);
  const fechaDetec = datosFila[5];
  const origenDetec = datosFila[6];

  if (!skuObs || precioNuevo <= 0) return;

  const ultFilaProductos = hojaProductos.getLastRow();
  if (ultFilaProductos < 2) return;

  const skus = hojaProductos.getRange(2, 1, ultFilaProductos - 1, 1).getValues();
  let filaDestino = -1;

  for (let i = 0; i < skus.length; i++) {
    if (String(skus[i][0] || '').trim().toUpperCase() === skuObs) {
      filaDestino = i + 2;
      break;
    }
  }

  if (filaDestino === -1) return;

  let factorInvent = parsearNumero(hojaProductos.getRange(filaDestino, 7).getValue());
  if (factorInvent <= 0) factorInvent = 1;
  let nuevoPU = precioNuevo / factorInvent;

  hojaProductos.getRange(filaDestino, 5).setValue(precioNuevo);
  hojaProductos.getRange(filaDestino, 8).setValue(nuevoPU);
  hojaProductos.getRange(filaDestino, 9).setValue(fechaDetec);
  hojaProductos.getRange(filaDestino, 10).setValue(origenDetec);

  hojaProductos.getRange(filaDestino, 1, 1, 10).setBackground(COLOR_VERDE_OK);
  hojaObs.getRange(fila, 1, 1, 8).setBackground(COLOR_VERDE_OK);
}

// =======================================================
// 🧹 PROCESAR RESOLUCIONES, GUARDAR MEMORIA Y LIMPIAR
// =======================================================
function procesarTodosLosAprobadosObservaciones() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaObs = ss.getSheetByName(HOJA_OBSERVACIONES_NAME);
  if (!hojaObs) return;

  const ultFila = hojaObs.getLastRow();
  if (ultFila < 2) return;

  const datos = hojaObs.getRange(2, 1, ultFila - 1, 8).getValues();
  let aprobadosCount = 0;
  let denegadosCount = 0;

  // 1. Data entry de aprobados y registro en ScriptProperties (Memoria Persistente)
  for (let i = 0; i < datos.length; i++) {
    const fila = datos[i];
    const sku = String(fila[0] || '').trim().toUpperCase();       // Col A: SKU
    const precioNuevo = parsearNumero(fila[3]);                    // Col D: Precio Nuevo
    const fecha = fila[5];                                        // Col F: Fecha
    const origen = fila[6];                                       // Col G: Origen
    const estado = String(fila[7] || '').trim().toUpperCase();    // Col H: Estado

    const filaReal = i + 2;
    const clave = armarClaveMemoria_(sku, fecha, origen, precioNuevo);

    if (estado === 'APROBADO') {
      procesarFilaAprobada(hojaObs, filaReal);
      registrarEnMemoriaPersistente_(clave, 'APROBADO');
      aprobadosCount++;
    } else if (estado === 'DENEGADO' || estado === 'RECHAZADO') {
      registrarEnMemoriaPersistente_(clave, 'DENEGADO');
      denegadosCount++;
    }
  }

  SpreadsheetApp.flush();

  // 2. Eliminación física de filas resueltas (de abajo hacia arriba)
  // Solo se quedan las que sigan en PENDIENTE
  for (let i = datos.length - 1; i >= 0; i--) {
    let estado = String(datos[i][7] || '').trim().toUpperCase();
    let filaReal = i + 2;
    if (estado === 'APROBADO' || estado === 'DENEGADO' || estado === 'RECHAZADO') {
      hojaObs.deleteRow(filaReal);
    }
  }

  SpreadsheetApp.getUi().alert(
    '🧹 RESOLUCIONES PROCESADAS Y GUARDADAS EN MEMORIA',
    `• Aprobados actualizados en Productos: ${aprobadosCount}\n` +
    `• Denegados archivados en memoria: ${denegadosCount}\n\n` +
    `Se eliminaron las filas resueltas de la hoja. Los ítems denegados no volverán a ingresar en futuras conciliaciones.`,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// =======================================================
// 🟢 MENÚ SUPERIOR Y DISPARADORES PROGRAMADOS
// =======================================================
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🟢 Actualizar Precios')
    .addItem('🔄 Ejecutar Conciliación Completa', 'ejecutarConciliacionUnificada')
    .addItem('📄 Actualizar solo Facturas', 'ejecutarSoloFacturas')
    .addItem('🥬 Actualizar solo Verdulería', 'ejecutarSoloVerduleria')
    .addSeparator()
    .addItem('✅ Procesar y Limpiar Observaciones', 'procesarTodosLosAprobadosObservaciones')
    .addSeparator()
    .addItem('🗑️ Vaciar Memoria de Denegados (Reset)', 'vaciarMemoriaResoluciones')
    .addItem('⚙️ Reinstalar Triggers Automáticos', 'crearTriggersAutomaticos')
    .addToUi();
}

function crearTriggersAutomaticos() {
  const triggersExistentes = ScriptApp.getProjectTriggers();
  for (let t of triggersExistentes) {
    ScriptApp.deleteTrigger(t);
  }

  ScriptApp.newTrigger('ejecutarSoloFacturas')
    .timeBased()
    .everyHours(24)
    .create();

  ScriptApp.newTrigger('ejecutarSoloVerduleria')
    .timeBased()
    .everyDays(2)
    .create();

  Logger.log('✅ Triggers programados instalados.');
}

function ejecutarSoloFacturas() {
  conciliarProcesos({ facturas: true, verduleria: false });
}

function ejecutarSoloVerduleria() {
  conciliarProcesos({ facturas: false, verduleria: true });
}

function ejecutarConciliacionUnificada() {
  conciliarProcesos({ facturas: true, verduleria: true });
}