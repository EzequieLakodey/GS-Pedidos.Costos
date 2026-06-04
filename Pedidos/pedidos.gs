// --- CONFIGURACIÓN Y CONEXIONES ---
const CONFIG = {
  ID_COSTOS: "1JHKK7tu30UyPeeq9hgwF3U2P8_84NtzJfINVRfa0WPA",
  HOJA_MATERIALES: "Materiales",
  HOJA_COSTOS: "Productos",
  HOJA_PEDIDOS: "Pedidos",
  HOJA_TERMINADOS: "Terminados",
  HOJA_ENTREGADOS: "Entregados",
  HOJA_LOG: "LogMovimientos",
  HOJA_CLIENTES: "Clientes",
  FILA_HEADERS: 3,
  FILA_INICIO_DATOS: 4
};

// Función para obtener la conexión a Costos de forma segura
function obtenerSSCostos() {
  return SpreadsheetApp.openById(CONFIG.ID_COSTOS);
}

/**
 * Busca dinámicamente en qué fila están los encabezados de una hoja externa,
 * buscando palabras clave como "ID" y "MODELO" en las primeras 5 filas.
 * Esto evita errores cuando las hojas externas tienen filas de título o encabezados en posiciones distintas.
 */
function detectarFilaHeaders(hoja) {
  const maxRows = Math.min(5, hoja.getLastRow());
  if (maxRows === 0) return 1;
  const data = hoja.getRange(1, 1, maxRows, hoja.getLastColumn()).getValues();
  for (let i = 0; i < data.length; i++) {
    const row = data[i].map(v => v.toString().trim().toUpperCase());
    if (row.includes("ID") && (row.includes("MODELO") || row.includes("CATEGORIA") || row.includes("STOCK"))) {
      return i + 1; // 1-indexed
    }
  }
  return 1; // fallback a fila 1
}

/**
 * Obtiene la hoja de Productos del archivo Costos,
 * con fallback al nombre "Costos" si no existe "Productos".
 */
function obtenerHojaProductos(ssCostos) {
  return ssCostos.getSheetByName(CONFIG.HOJA_COSTOS) || ssCostos.getSheetByName("Costos");
}

// Headers exactos para mapeo
const HEADERS_COSTOS = {
  ID: "ID", CATEGORIA: "CATEGORIA", MODELO: "MODELO", MATERIALES: "MATERIALES",
  COSTO_MAT: "COSTO MAT", HORAS_MOBRA: "HORAS M.OBRA", MOBRA_USD: "M.OBRA USD",
  SUBTOTAL: "SUBTOTAL", DESPERDICIOS: "DESPERDICIOS", IMPUESTOS: "IMPUESTOS",
  IMPORTES_IMP: "IMPORTES IMPUESTOS", TOTAL: "TOTAL", STOCK: "STOCK"
};
const HEADERS_MATERIALES = {
  ID: "ID", CATEGORIA: "Categoria", DETALLES: "Detalles", UNIDAD: "Unidad de Medida",
  PROVEEDOR: "Proveedor", PU_USD: "P/U USD", PU_ARS: "P/U ARS",
  IMP: "%IMP", DTO: "%DTO", TOTAL: "Total", STOCK: "STOCK"
};

// --- INICIO DE LA APLICACIÓN ---
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('⚙️ Gestión Integral')
    .addItem('➕ Cargar Nuevo Pedido', 'abrirModalPedidos')
    .addSeparator()
    .addItem('💱 Modificar Tipo de Cambio', 'modificarTipoCambio')
    .addItem('🌾 Aumentar Stock Materiales', 'abrirModalStocksMateriales')
    .addItem('📦 Aumentar Stock Productos', 'abrirModalStocksProductos')
    .addItem('📉 Gestionar Alertas de Stock', 'abrirModalAlertasStock')
    .addItem('🔎 Revisar Stocks Mínimos', 'revisarAlertasStock')
    .addSeparator()
    .addItem('🔠 Convertir Clientes a Mayúsculas', 'normalizarRazonesSociales')
    .addToUi();

  try {
    revisarAlertasStock(true);
  } catch (e) {
    console.error("Error revisando alertas al abrir: " + e.message);
  }
}

// Renombramos a onEditSimple para que Google no lo bloquee por seguridad
function onEditSimple(e) {
  if (!e?.range) return;
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('BLOQUEAR_ONEDIT') === '1') return;
  const sheet = e.range.getSheet();
  const name = sheet.getName();
  const col = e.range.getColumn();
  const row = e.range.getRow();
  const val = e.value || "";

  // Solo procesar filas de datos (ignorar encabezados)
  if (row <= CONFIG.FILA_HEADERS) return;

  // 2. Lógica de Flujo de Pedidos: detectar columna ESTADO dinámicamente
  if (name === CONFIG.HOJA_PEDIDOS || name === CONFIG.HOJA_TERMINADOS) {
    // Leer los encabezados de la fila 3 para encontrar la columna ESTADO
    const headers = sheet.getRange(CONFIG.FILA_HEADERS, 1, 1, sheet.getLastColumn()).getValues()[0];
    const colEstado = headers.findIndex(h => h.toString().trim().toUpperCase() === "ESTADO") + 1;

    if (colEstado > 0 && col === colEstado) {
      const valNorm = val.toString().trim().toUpperCase();
      if (name === CONFIG.HOJA_PEDIDOS) {
        if (valNorm === "PRODUCIR" || valNorm === "Terminado") {
          iniciarDecisionStockPorEstado(sheet, row, colEstado, valNorm, e.oldValue);
          return;
        }
        procesarFlujoPedidos(sheet, row, col, valNorm);
      } else if (name === CONFIG.HOJA_TERMINADOS) {
        // En la hoja TERMINADOS, también manejamos ENTREGA PARCIAL
        if (valNorm === "ENTREGA PARCIAL") {
          PropertiesService.getScriptProperties().setProperty('FILA_ENTREGA_PARCIAL', row.toString());
          PropertiesService.getScriptProperties().setProperty('HOJA_ENTREGA_PARCIAL', CONFIG.HOJA_TERMINADOS);
          abrirModalEntregaParcial();
        } else {
          procesarFlujoTerminados(sheet, row, col, valNorm);
        }
      }
    }
  }
}

// --- INSTALADOR DE TRIGGERS (FUERA DE ONEDIT) ---
function instalarTodoElSistema() {
  const ss = SpreadsheetApp.getActive();
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('onOpen').forSpreadsheet(ss).onOpen().create();
  ScriptApp.newTrigger('onEditInstalable').forSpreadsheet(ss).onEdit().create();

  SpreadsheetApp.getUi().alert('🚀 ¡Sistema activado! El menú y los cambios de estado funcionarán automáticamente.');
}

function onEditInstalable(e) {
  if (!e || !e.range) return;
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(5000)) {
    console.warn("onEditInstalable omitido por lock ocupado.");
    return;
  }
  try {
    onEditSimple(e);
    // 2. Manejar conversión de moneda
    manejarConversionMoneda(e);
  } finally {
    lock.releaseLock();
  }
}
// --- LÓGICA DE PEDIDOS (FORMULARIO Y GUARDADO) ---
function abrirModalPedidos() {
  const html = HtmlService.createTemplateFromFile('FormularioPedidos');
  SpreadsheetApp.getUi().showModalDialog(
    html.evaluate().setWidth(780).setHeight(880),
    '📝 Cargar Nuevo Pedido'
  );
}


function obtenerDataInicialPedidos() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Clientes de la hoja local - con manejo seguro si no existe
    let clientes = [];
    try {
      const hojaClientes = ss.getSheetByName(CONFIG.HOJA_CLIENTES);
      if (hojaClientes && hojaClientes.getLastRow() >= 2) {
        // Obtenemos Razon, Alias, Contacto para el filtro
        clientes = hojaClientes.getRange("A2:C" + hojaClientes.getLastRow()).getValues()
          .map(r => ({ razon: r[0], alias: r[1] || "", contacto: r[2] || "" }))
          .filter(c => c.razon && c.razon.toString().trim() !== "");
      }
    } catch (eClientes) {
      clientes = [];
    }

    // Conexión a Costos
    const ssCostos = obtenerSSCostos();
    if (!ssCostos) throw new Error("No se pudo conectar al archivo de Costos. Verifica el ID.");

    // Obtener hoja de productos con fallback
    const hojaProds = obtenerHojaProductos(ssCostos);
    if (!hojaProds) throw new Error("No se encontró la hoja 'Productos' ni 'Costos' en el archivo de Costos. Verifica el nombre de la hoja.");

    // Detectar fila de headers dinámicamente
    const filaHeadersProds = detectarFilaHeaders(hojaProds);
    const filaInicioProds = filaHeadersProds + 1;

    let productos = [];
    let tc = 1;

    const data = hojaProds.getDataRange().getValues();
    if (data.length < filaHeadersProds) throw new Error("La hoja de Productos no tiene suficientes filas.");

    const headers = data[filaHeadersProds - 1]; // 0-indexed
    const colId = headers.findIndex(h => h && h.toString().trim().toUpperCase() === "ID");
    const colCat = headers.findIndex(h => h && h.toString().trim().toUpperCase() === "CATEGORIA");
    const colMod = headers.findIndex(h => h && h.toString().trim().toUpperCase() === "MODELO");
    const colStock = headers.findIndex(h => h && h.toString().trim().toUpperCase() === "STOCK");
    const colMat = headers.findIndex(h => h && h.toString().trim().toUpperCase() === "MATERIALES");
    let colPrecio = headers.findIndex(h => h && h.toString().trim().toUpperCase() === "PRECIO");
    if (colPrecio === -1 && headers.length >= 14) {
      colPrecio = 13; // Columna N es el índice 13
    }
    let colUltimoPrecio = headers.findIndex(h => h && h.toString().trim().toUpperCase() === "ULTIMO PRECIO");

    for (let i = filaInicioProds - 1; i < data.length; i++) {
      if (colId !== -1 && data[i][colId] && data[i][colId].toString().trim() !== "") {
        let recetaTexto = colMat !== -1 && data[i][colMat] ? data[i][colMat].toString() : "";
        let precioVal = colPrecio !== -1 && data[i][colPrecio] !== undefined ? parseFloat(data[i][colPrecio]) || 0 : 0;
        let ultimoPrecioVal = colUltimoPrecio !== -1 && data[i][colUltimoPrecio] !== undefined ? parseFloat(data[i][colUltimoPrecio]) || 0 : 0;
        productos.push({
          id: data[i][colId],
          categoria: colCat !== -1 && data[i][colCat] ? data[i][colCat].toString().trim() : "Sin Categoría",
          nombre: ((colCat !== -1 && data[i][colCat] ? data[i][colCat] : "") + " - " + (colMod !== -1 && data[i][colMod] ? data[i][colMod] : "")).trim(),
          stock: colStock !== -1 ? parseFloat(data[i][colStock]) || 0 : 0,
          receta: parsearReceta(recetaTexto),
          precio: precioVal,
          ultimoPrecio: ultimoPrecioVal
        });
      }
    }

    // Intentar leer el Tipo de Cambio desde G1
    try {
      tc = parseFloat(hojaProds.getRange("G1").getValue()) || 1;
    } catch (e) { tc = 1; }

    // Obtener Materiales para validación de stock
    let materiales = [];
    try {
      const hojaMat = ssCostos.getSheetByName(CONFIG.HOJA_MATERIALES);
      if (hojaMat) {
        const dataMat = hojaMat.getDataRange().getValues();
        const headersMat = dataMat[0]; // Fila 1 (index 0)
        const colMatDetalle = headersMat.findIndex(h => h && h.toString().trim().toUpperCase() === "DETALLES");
        const colMatStock = headersMat.findIndex(h => h && h.toString().trim().toUpperCase() === "STOCK");

        for (let i = 1; i < dataMat.length; i++) {
          if (colMatDetalle !== -1 && dataMat[i][colMatDetalle]) {
            materiales.push({
              nombre: dataMat[i][colMatDetalle].toString().trim(),
              stock: colMatStock !== -1 ? parseFloat(dataMat[i][colMatStock]) || 0 : 0
            });
          }
        }
      }
    } catch (eMat) {
      materiales = [];
    }

    return { clientes: (clientes || []).sort(), productos: productos || [], materiales: materiales || [], tc };
  } catch (e) {
    throw new Error("Error en obtenerDataInicialPedidos: " + e.message);
  }
}



function guardarNuevoPedido(datos) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.HOJA_PEDIDOS);

    const filaNum = sheet.getLastRow() + 1;
    const headers = sheet.getRange(CONFIG.FILA_HEADERS, 1, 1, sheet.getLastColumn()).getValues()[0]
      .map(h => h.toString().trim().toUpperCase());

    // Crear el texto del detalle del pedido: Cantidad | Categoria | Modelo | Precio Unitario
    let pedidoDetalle = datos.productos.map(p =>
      `${p.cantidad} | ${p.categoria} | ${p.modelo} | $${p.precio_unitario} ${datos.moneda === 'ARS' ? 'AR$' : 'U$D'}`
    ).join('\n');

    // Parsear fecha para evitar el error de formato
    let dateObj = new Date(datos.fecha + "T12:00:00");

    // Construir la forma de entrega
    let f_entrega_str = datos.f_entrega.toUpperCase() === "RETIRA"
      ? `RETIRA`
      : `ENVIO`;

    let estadoFinal = datos.estado || "No Encomendado";

    // Calcular descuento equivalente
    let dtoEquiv = datos.factorDesc === 1 ? 0 : (1 - datos.factorDesc);
    let impPctNum = parseFloat(datos.impuesto_pct) || 0;

    // Obtener la fila del cliente en la hoja "Clientes" para el hipervínculo
    const clienteLink = construirLinkCliente(datos.cliente) || datos.cliente;

    // Mapeamos los datos correspondientes a cada encabezado
    const setValByHeader = (headerName, value, numberFormat = null) => {
      const idx = headers.indexOf(headerName.toUpperCase());
      if (idx !== -1) {
        const range = sheet.getRange(filaNum, idx + 1);
        range.setValue(value);
        if (numberFormat) {
          range.setNumberFormat(numberFormat);
        }
      }
    };

    setValByHeader("FECHA", dateObj, 'dd/mm/yyyy');
    setValByHeader("CLIENTE", clienteLink);
    setValByHeader("PEDIDO", pedidoDetalle);
    setValByHeader("SUBTOTAL", datos.subtotalBase);
    setValByHeader("TOTAL", datos.totalGeneral);
    setValByHeader("CUENTA", datos.montoPago || 0);
    setValByHeader("T.C", datos.tc);
    setValByHeader("DTO", dtoEquiv, "0.00%");
    setValByHeader("IMP", impPctNum / 100, "0.00%");
    setValByHeader("PLAZO DE ENTREGA", datos.plazo);
    setValByHeader("PLAZO", datos.plazo);
    setValByHeader("FORMA ENTREGA", f_entrega_str);
    setValByHeader("ESTADO", estadoFinal);
    setValByHeader("MONEDA", datos.moneda === "ARS" ? "AR$" : "U$D");
    setValByHeader("MONEDA_ESTADO", datos.moneda === "ARS" ? "AR$" : "U$D"); // <--- AGREGADO AQUÍ


    // ACTUALIZAR PRECIOS EN LA HOJA PRODUCTOS SI FUERON MODIFICADOS
    actualizarPreciosProductos(datos.productos);

    // LÓGICA DE DESCUENTO DE STOCK
    if (estadoFinal === "Terminado" || estadoFinal === "Producir") {
      descontarStockDePedido(datos.productos);
    }

    // SI ES TERMINADO, MOVER DIRECTAMENTE A LA HOJA TERMINADOS
    if (estadoFinal === "Terminado") {
      const hojaTerminados = ss.getSheetByName(CONFIG.HOJA_TERMINADOS);
      if (hojaTerminados) {
        transferirFilaDinamica(sheet, filaNum, hojaTerminados);
        registrarMovimientoLog({
          accion: "ALTA_DIRECTA_TERMINADO",
          hojaOrigen: CONFIG.HOJA_PEDIDOS,
          hojaDestino: CONFIG.HOJA_TERMINADOS,
          fila: filaNum,
          estado: "Terminado",
          detalle: `Pedido creado y movido a ${CONFIG.HOJA_TERMINADOS}`
        });
        sheet.deleteRow(filaNum);
      }
    }

    return { success: true };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function descontarStockDePedido(productosPedido) {
  const ssCostos = obtenerSSCostos();
  const hojaProds = obtenerHojaProductos(ssCostos);
  const hojaMat = ssCostos.getSheetByName("Materiales");

  if (!hojaProds || !hojaMat) throw new Error("No se encontraron las hojas de Productos o Materiales.");

  const filaHeadersProds = detectarFilaHeaders(hojaProds);
  const colsP = getColMap(hojaProds, HEADERS_COSTOS, filaHeadersProds);
  const colsM = getColMap(hojaMat, HEADERS_MATERIALES, 1);

  const datosP = hojaProds.getDataRange().getValues();
  const datosM = hojaMat.getDataRange().getValues();

  for (const prod of productosPedido) {
    if (!prod.id) continue;

    let filaProdIdx = datosP.findIndex(r => r[colsP.ID - 1] === prod.id);
    if (filaProdIdx === -1) continue;

    let filaProdSheet = filaProdIdx + 1;
    let stockActual = parseFloat(datosP[filaProdIdx][colsP.STOCK - 1]) || 0;

    if (stockActual >= prod.cantidad) {
      hojaProds.getRange(filaProdSheet, colsP.STOCK).setValue(stockActual - prod.cantidad);
    } else {
      let faltante = prod.cantidad - stockActual;
      hojaProds.getRange(filaProdSheet, colsP.STOCK).setValue(0);

      let recetaTexto = datosP[filaProdIdx][colsP.MATERIALES - 1] || "";
      let materialesReceta = parsearReceta(recetaTexto);

      for (const mat of materialesReceta) {
        let cantTotalRequerida = mat.cantidad * faltante;
        let filaMatIdx = datosM.findIndex(r => r[colsM.DETALLES - 1] === mat.nombre);
        if (filaMatIdx !== -1) {
          let filaMatSheet = filaMatIdx + 1;
          let stockMatActual = parseFloat(datosM[filaMatIdx][colsM.STOCK - 1]) || 0;
          hojaMat.getRange(filaMatSheet, colsM.STOCK).setValue(stockMatActual - cantTotalRequerida);
          datosM[filaMatIdx][colsM.STOCK - 1] = stockMatActual - cantTotalRequerida; // update memory
        }
      }
    }
  }
}

function parsearReceta(texto) {
  if (!texto) return [];
  const regex = /(.+?)\s*-\s*(.+?)\s*-\s*(.+?)\s*\(([\d.]+)\s*(.+?)\)/;
  return texto.split('\n').map(linea => {
    const m = linea.match(regex);
    if (!m) return null;
    return {
      nombre: m[3].trim(),
      cantidad: parseFloat(m[4])
    };
  }).filter(Boolean);
}

/**
 * Función Maestra: Transfiere una fila entre dos hojas emparejando por nombres de encabezado.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} hojaOrigen
 * @param {number} filaOrigen
 * @param {GoogleAppsScript.Spreadsheet.Sheet} hojaDestino
 * @param {Object} valoresExtras Objeto con { "HEADER": valor } para columnas que no están en el origen o queremos sobreescribir.
 */
function transferirFilaDinamica(hojaOrigen, filaOrigen, hojaDestino, valoresExtras = {}) {
  const headersOrigen = hojaOrigen.getRange(CONFIG.FILA_HEADERS, 1, 1, hojaOrigen.getLastColumn()).getValues()[0]
    .map(h => h.toString().trim().toUpperCase());
  const headersDestino = hojaDestino.getRange(CONFIG.FILA_HEADERS, 1, 1, hojaDestino.getLastColumn()).getValues()[0]
    .map(h => h.toString().trim().toUpperCase());

  const formulasOrigen = hojaOrigen.getRange(filaOrigen, 1, 1, hojaOrigen.getLastColumn()).getFormulas()[0];
  const valoresOrigen = hojaOrigen.getRange(filaOrigen, 1, 1, hojaOrigen.getLastColumn()).getValues()[0];

  // Construir la nueva fila para el destino
  let nuevaFila = headersDestino.map(header => {
    // 1. Prioridad: Valores extras pasados por parámetro (ej: FECHA ENTREGA)
    if (valoresExtras[header] !== undefined) return valoresExtras[header];

    // 2. Buscar por nombre en el origen
    const idxOrigen = headersOrigen.indexOf(header);
    if (idxOrigen !== -1) {
      // Si es fórmula (empieza por =), usamos la fórmula para preservar hipervínculos
      return formulasOrigen[idxOrigen].toString().startsWith('=') ? formulasOrigen[idxOrigen] : valoresOrigen[idxOrigen];
    }

    // 3. Si no existe en el origen ni en extras, devolver vacío
    return "";
  });

  // Forzar que CLIENTE se mantenga como hyperlink hacia hoja Clientes.
  const idxClienteDestino = headersDestino.indexOf("CLIENTE");
  if (idxClienteDestino !== -1) {
    let clienteNombre = null;
    const idxClienteOrigen = headersOrigen.indexOf("CLIENTE");
    if (idxClienteOrigen !== -1) {
      const formulaCliente = formulasOrigen[idxClienteOrigen];
      const valorCliente = valoresOrigen[idxClienteOrigen];
      clienteNombre = extraerNombreCliente(valorCliente, formulaCliente);
    }
    const linkCliente = construirLinkCliente(clienteNombre);
    if (linkCliente) nuevaFila[idxClienteDestino] = linkCliente;
  }

  hojaDestino.appendRow(nuevaFila);

  // Aplicar formatos básicos si es necesario (ej: Fechas)
  const lastRow = hojaDestino.getLastRow();
  headersDestino.forEach((h, i) => {
    if (h.includes("FECHA")) {
      hojaDestino.getRange(lastRow, i + 1).setNumberFormat('dd/mm/yyyy');
    }
  });
}

// --- LÓGICA DE FLUJO (MOVIMIENTO DE FILAS) ---
function procesarFlujoPedidos(hoja, fila, columna, valor) {
  if (!valor) return;
  const valNorm = valor.toString().trim().toUpperCase();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (valNorm === "Terminado") {
    const destino = ss.getSheetByName(CONFIG.HOJA_TERMINADOS);
    if (!destino) return;
    transferirFilaDinamica(hoja, fila, destino);
    registrarMovimientoLog({
      accion: "CAMBIO_ESTADO",
      hojaOrigen: hoja.getName(),
      hojaDestino: destino.getName(),
      fila,
      estado: "Terminado",
      detalle: "Movimiento automático por estado TERMINADO"
    });
    hoja.deleteRow(fila);
  }
  else if (valNorm === "ENTREGADO") {
    moverAEntregados(hoja, fila);
  }
  else if (valNorm === "ENTREGA PARCIAL") {
    const headers = hoja.getRange(CONFIG.FILA_HEADERS, 1, 1, hoja.getLastColumn()).getValues()[0]
      .map(h => h.toString().trim().toUpperCase());
    const colEstado = headers.indexOf("ESTADO") + 1;
    if (colEstado > 0) hoja.getRange(fila, colEstado).setValue("En Produccion");
    PropertiesService.getScriptProperties().setProperty('FILA_ENTREGA_PARCIAL', fila.toString());
    PropertiesService.getScriptProperties().setProperty('HOJA_ENTREGA_PARCIAL', CONFIG.HOJA_PEDIDOS);
    abrirModalEntregaParcial();
  }
}

function procesarFlujoTerminados(hoja, fila, columna, valor) {
  if (!valor) return;
  const valNorm = valor.toString().trim().toUpperCase();
  if (valNorm === "ENTREGADO") {
    moverAEntregados(hoja, fila);
  }
}

function moverAEntregados(hojaOrigen, fila) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const destino = ss.getSheetByName(CONFIG.HOJA_ENTREGADOS);
  if (!destino) return;

  // Transferir dinámicamente agregando la Fecha de Entrega automática
  transferirFilaDinamica(hojaOrigen, fila, destino, {
    "FECHA ENTREGA": new Date(),
    "ESTADO": "ENTREGADO"
  });

  registrarMovimientoLog({
    accion: "CAMBIO_ESTADO",
    hojaOrigen: hojaOrigen.getName(),
    hojaDestino: destino.getName(),
    fila,
    estado: "ENTREGADO",
    detalle: "Movimiento automático a ENTREGADOS"
  });

  hojaOrigen.deleteRow(fila);
}

// --- ENTREGA PARCIAL ---
function abrirModalEntregaParcial() {
  const html = HtmlService.createTemplateFromFile('FormularioEntregaParcial');
  SpreadsheetApp.getUi().showModalDialog(
    html.evaluate().setWidth(780).setHeight(750),
    '📦 Registrar Entrega Parcial'
  );
}

function obtenerDatosPedidoParcial() {
  const props = PropertiesService.getScriptProperties();
  const fila = parseInt(props.getProperty('FILA_ENTREGA_PARCIAL'));
  const hojaOrigen = props.getProperty('HOJA_ENTREGA_PARCIAL') || CONFIG.HOJA_PEDIDOS;
  if (!fila) return null;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(hojaOrigen);
  if (!sheet) return null;

  const headers = sheet.getRange(CONFIG.FILA_HEADERS, 1, 1, sheet.getLastColumn()).getValues()[0]
    .map(h => h.toString().trim().toUpperCase());
  const datos = sheet.getRange(fila, 1, 1, sheet.getLastColumn()).getValues()[0];

  const getValByHeader = (headerName) => {
    const idx = headers.indexOf(headerName.toUpperCase());
    return idx !== -1 ? datos[idx] : null;
  };

  let pedidoTexto = getValByHeader("PEDIDO") ? getValByHeader("PEDIDO").toString() : "";
  let productos = pedidoTexto.split('\n').filter(l => l.trim()).map(linea => {
    let partes = linea.split('|').map(s => s.trim());
    if (partes.length >= 4) {
      return {
        cantidad: parseFloat(partes[0]) || 0,
        categoria: partes[1] || '',
        modelo: partes[2] || '',
        precio_unitario: parseFloat(partes[3].replace('$', '')) || 0
      };
    }
    return null;
  }).filter(Boolean);

  let subtotalLimpio = parseFloat(getValByHeader("SUBTOTAL")) || 0;
  let totalLimpio = parseFloat(getValByHeader("TOTAL")) || 0;
  let tcOriginal = parseFloat(getValByHeader("T.C")) || 1;
  let cuentaVal = parseFloat(getValByHeader("CUENTA")) || 0;
  let descPct = parseFloat(getValByHeader("DTO")) || 0;
  let impPct = parseFloat(getValByHeader("IMP")) || 0;

  return {
    fila,
    hoja_origen: hojaOrigen,
    cliente: getValByHeader("CLIENTE") || "",
    productos,
    subtotal: subtotalLimpio,
    total: totalLimpio,
    tc: tcOriginal,
    descuento_pct: descPct,
    impuesto_pct: impPct,
    pagado_hasta_ahora: cuentaVal
  };
}

function guardarEntregaParcial(datos) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hojaOrigenNombre = datos.hoja_origen || CONFIG.HOJA_PEDIDOS;
    const sheet = ss.getSheetByName(hojaOrigenNombre);
    const hEntregados = ss.getSheetByName(CONFIG.HOJA_ENTREGADOS);
    const fila = datos.fila;
    if (!sheet || !hEntregados) throw new Error("No se encontraron las hojas requeridas para registrar la entrega parcial.");

    const headersOrigen = sheet.getRange(CONFIG.FILA_HEADERS, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => h.toString().trim().toUpperCase());
    const originalValues = sheet.getRange(fila, 1, 1, sheet.getLastColumn()).getValues()[0];

    let entregados = datos.productos_entregados.filter(p => p.cantidad_entregada > 0);
    if (!entregados.length) throw new Error('No se seleccionaron productos.');

    let subEntBruto = entregados.reduce((sum, p) => sum + (p.cantidad_entregada * p.precio_unitario), 0);
    let subtotalPendienteActual = parseFloat(originalValues[headersOrigen.indexOf("SUBTOTAL")]) || 0;
    let totalPendienteActual = parseFloat(originalValues[headersOrigen.indexOf("TOTAL")]) || 0;
    let pagadoAnterior = parseFloat(originalValues[headersOrigen.indexOf("CUENTA")]) || 0;
    let pagoRecibido = parseFloat(datos.pago_recibido) || 0;
    let pagadoTotalAcumulado = pagadoAnterior + pagoRecibido;

    let subBrutoPendiente = (datos.productos_entregados || []).reduce((sum, p) => sum + ((parseFloat(p.cantidad_original) || 0) * (parseFloat(p.precio_unitario) || 0)), 0);
    let ratioEntregado = subBrutoPendiente > 0 ? (subEntBruto / subBrutoPendiente) : 0;
    ratioEntregado = Math.max(0, Math.min(1, ratioEntregado));

    // Repartimos sobre el saldo pendiente actual para mantener coherencia con FormularioPedidos.
    let subEnt = subtotalPendienteActual * ratioEntregado;
    let totalEnt = totalPendienteActual * ratioEntregado;
    let cuentaEntregada = pagadoTotalAcumulado * ratioEntregado;
    let subEntPendiente = Math.max(0, subEnt - cuentaEntregada);
    let totalEntPendiente = Math.max(0, totalEnt - cuentaEntregada);

    // 1. ENVIAR LO ENTREGADO A LA HOJA "ENTREGADOS"
    const headersDest = hEntregados.getRange(CONFIG.FILA_HEADERS, 1, 1, hEntregados.getLastColumn()).getValues()[0].map(h => h.toString().trim().toUpperCase());
    let filaParaEntregados = headersDest.map(header => {
      if (header === "PEDIDO") return entregados.map(p => `${p.cantidad_entregada} | ${p.categoria} | ${p.modelo} | $${p.precio_unitario}`).join('\n');
      if (header === "SUBTOTAL") return subEntPendiente;
      if (header === "TOTAL") return totalEntPendiente;
      if (header === "T.C") return datos.tc;
      if (header === "ESTADO") return "ENTREGADO";
      if (header === "CUENTA") return cuentaEntregada;
      if (header === "FECHA ENTREGA") return new Date();
      if (header === "CLIENTE") {
        const idxCli = headersOrigen.indexOf("CLIENTE");
        const clienteNombre = idxCli !== -1 ? extraerNombreCliente(originalValues[idxCli], "") : "";
        return construirLinkCliente(clienteNombre) || (idxCli !== -1 ? originalValues[idxCli] : "");
      }

      const idx = headersOrigen.indexOf(header);
      return idx !== -1 ? originalValues[idx] : "";
    });

    hEntregados.appendRow(filaParaEntregados);
    const lastRowEnt = hEntregados.getLastRow();
    headersDest.forEach((h, i) => { if (h.includes("FECHA")) hEntregados.getRange(lastRowEnt, i + 1).setNumberFormat('dd/mm/yyyy'); });
    registrarMovimientoLog({
      accion: "ENTREGA_PARCIAL",
      hojaOrigen: hojaOrigenNombre,
      hojaDestino: CONFIG.HOJA_ENTREGADOS,
      fila,
      estado: "ENTREGA PARCIAL",
      detalle: `Se transfirió parte del pedido. Pago recibido: ${pagoRecibido}`
    });

    // 2. ACTUALIZAR LO QUE QUEDA EN LA HOJA "PEDIDOS"
    let restantes = datos.productos_entregados.filter(p => (p.cantidad_original - p.cantidad_entregada) > 0);

    if (!restantes.length) {
      sheet.deleteRow(fila);
      return { success: true };
    }

    let detalleRestante = restantes.map(p => `${p.cantidad_original - p.cantidad_entregada} | ${p.categoria} | ${p.modelo} | $${p.precio_unitario}`).join('\n');
    let subRestBruto = restantes.reduce((sum, p) => sum + ((p.cantidad_original - p.cantidad_entregada) * p.precio_unitario), 0);
    let ratioRestante = subBrutoPendiente > 0 ? (subRestBruto / subBrutoPendiente) : 0;
    ratioRestante = Math.max(0, Math.min(1, ratioRestante));

    let subRestSaldo = subtotalPendienteActual * ratioRestante;
    let totalRestanteSaldo = totalPendienteActual * ratioRestante;

    // El pago recibido en esta operación reduce el saldo que permanece en la hoja de trabajo.
    let nuevoSaldoSubtotal = Math.max(0, subRestSaldo - pagoRecibido);
    let nuevoSaldoTotal = Math.max(0, totalRestanteSaldo - pagoRecibido);
    let nuevoPagado = Math.max(0, pagadoTotalAcumulado - cuentaEntregada);

    const setVal = (header, val) => {
      const idx = headersOrigen.indexOf(header);
      if (idx !== -1) sheet.getRange(fila, idx + 1).setValue(val);
    };

    setVal("PEDIDO", detalleRestante);
    setVal("SUBTOTAL", nuevoSaldoSubtotal);
    setVal("TOTAL", nuevoSaldoTotal);
    setVal("ESTADO", "En Produccion");
    setVal("CUENTA", nuevoPagado);

    return { success: true };
  } catch (e) {
    console.error(e);
    return { success: false, message: e.message };
  }
}
// --- FUNCIONES DE GESTIÓN Y STOCK ---

function agregarCliente(razon, alias, contacto) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.HOJA_CLIENTES);
    if (!sheet) throw new Error('No se encontró la hoja de Clientes');

    // Determine if headers exist, but we assume row 1 is headers, data starts at row 2
    // We will append a row with values: RAZON SOCIAL, ALIAS, CONTACTO, and empty for other columns if needed? Actually the sheet may have more columns.
    // To keep it simple, we will append only the three columns; if the sheet has more columns, they will be left blank.
    // Better: get the number of columns in the sheet and fill missing with empty strings.
    const lastCol = sheet.getLastColumn();
    const valores = [razon.toUpperCase(), alias, contacto];
    // Pad with empty strings if there are more columns
    while (valores.length < lastCol) {
      valores.push('');
    }
    sheet.appendRow(valores);
    return { success: true };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function modificarTipoCambio() {
  const ssCostos = obtenerSSCostos();
  const hojaProds = obtenerHojaProductos(ssCostos);
  const tcActual = parseFloat(hojaProds.getRange("G1").getValue()) || 1;

  const response = SpreadsheetApp.getUi().prompt('💱 Tipo de Cambio', 'Ingrese el nuevo tipo de cambio (Ref):', SpreadsheetApp.getUi().ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() === SpreadsheetApp.getUi().Button.OK) {
    const nuevoTc = parseFloat(response.getResponseText());
    if (!isNaN(nuevoTc) && nuevoTc > 0) {
      hojaProds.getRange("G1").setValue(nuevoTc);
      SpreadsheetApp.getUi().alert('✅ Tipo de cambio actualizado a: ' + nuevoTc);
    } else {
      SpreadsheetApp.getUi().alert('❌ Valor inválido.');
    }
  }
}

function abrirModalStocksMateriales() { abrirModalStocks("materiales"); }
function abrirModalStocksProductos() { abrirModalStocks("productos"); }

function abrirModalStocks(modo) {
  const template = HtmlService.createTemplateFromFile('FormularioModificarStocks');
  template.modo = modo;
  SpreadsheetApp.getUi().showModalDialog(
    template.evaluate().setWidth(600).setHeight(650),
    modo === "materiales" ? "🌾 Aumentar Stock Materiales" : "📦 Aumentar Stock Productos"
  );
}

function obtenerDataParaStocks() {
  const ssCostos = obtenerSSCostos();
  const hojaProds = obtenerHojaProductos(ssCostos);
  const hojaMat = ssCostos.getSheetByName(CONFIG.HOJA_MATERIALES);

  // Materiales
  let materiales = [];
  if (hojaMat) {
    const data = hojaMat.getDataRange().getValues();
    const headers = data[0].map(h => h.toString().trim());
    const colDetalle = headers.indexOf("Detalles");
    const colStock = headers.indexOf("STOCK");
    const colProv = headers.indexOf("Proveedor");
    const colCat = headers.indexOf("Categoria"); // Nuevo

    for (let i = 1; i < data.length; i++) {
      if (colDetalle !== -1 && data[i][colDetalle]) {
        materiales.push({
          nombre: data[i][colDetalle].toString().trim(),
          stock: colStock !== -1 ? parseFloat(data[i][colStock]) || 0 : 0,
          proveedor: colProv !== -1 ? data[i][colProv].toString().trim() : "Sin Proveedor",
          categoria: colCat !== -1 ? data[i][colCat].toString().trim() : "Sin Categoría"
        });
      }
    }
  }

  // Productos
  let productos = [];
  if (hojaProds) {
    const filaHeaders = detectarFilaHeaders(hojaProds);
    const data = hojaProds.getDataRange().getValues();
    const headers = data[filaHeaders - 1].map(h => h.toString().trim());
    const colCat = headers.indexOf("CATEGORIA");
    const colMod = headers.indexOf("MODELO");
    const colStock = headers.indexOf("STOCK");

    for (let i = filaHeaders; i < data.length; i++) {
      if (colMod !== -1 && data[i][colMod]) {
        productos.push({
          nombre: ((colCat !== -1 ? data[i][colCat] : "") + " - " + data[i][colMod]).trim(),
          stock: colStock !== -1 ? parseFloat(data[i][colStock]) || 0 : 0,
          categoria: colCat !== -1 ? data[i][colCat].toString().trim() : "Sin Categoría"
        });
      }
    }
  }

  return {
    materiales,
    productos,
    proveedores: [...new Set(materiales.map(m => m.proveedor).filter(p => p))],
    categorias: [...new Set(productos.map(p => p.categoria).filter(c => c))]
  };
}

function guardarStockLote(tipo, updates) {
  try {
    const ssCostos = obtenerSSCostos();
    const hoja = tipo === "materiales" ? ssCostos.getSheetByName(CONFIG.HOJA_MATERIALES) : obtenerHojaProductos(ssCostos);
    if (!hoja) throw new Error("No se encontró la hoja de " + tipo);

    const filaHeaders = tipo === "materiales" ? 1 : detectarFilaHeaders(hoja);
    const data = hoja.getDataRange().getValues();
    const headers = data[filaHeaders - 1].map(h => h.toString().trim());

    const colStock = headers.indexOf("STOCK") + 1;
    const colClave = tipo === "materiales" ? headers.indexOf("Detalles") + 1 : headers.indexOf("MODELO") + 1;
    const colCat = tipo === "productos" ? headers.indexOf("CATEGORIA") + 1 : -1;

    updates.forEach(u => {
      let filaIdx = -1;
      for (let i = filaHeaders; i < data.length; i++) {
        let claveMatch = data[i][colClave - 1] && data[i][colClave - 1].toString().trim() === u.nombre.split(" - ").pop();
        let catMatch = tipo === "productos" ? (data[i][colCat - 1] && data[i][colCat - 1].toString().trim() === u.nombre.split(" - ").shift()) : true;
        if (claveMatch && catMatch) { filaIdx = i + 1; break; }
      }

      if (filaIdx !== -1) {
        const stockActual = parseFloat(hoja.getRange(filaIdx, colStock).getValue()) || 0;
        hoja.getRange(filaIdx, colStock).setValue(stockActual + u.incremento);
      }
    });

    return { success: true };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// --- HELPER FUNCTIONS ---

function getColIndex(sheet, headerName, headerRow) {
  const row = headerRow || 1;
  const vals = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
  const idx = vals.findIndex(v => v.toString().trim().toUpperCase() === headerName.trim().toUpperCase());
  return idx === -1 ? -1 : idx + 1;
}
function getColMap(sheet, headersObj, headerRow) {
  const map = {};
  const missing = [];
  Object.entries(headersObj).forEach(([key, label]) => {
    const col = getColIndex(sheet, label, headerRow);
    if (col === -1) missing.push(label);
    map[key] = col;
  });
  return map;
}

function actualizarPreciosProductos(productosPedido) {
  try {
    const ssCostos = obtenerSSCostos();
    const hojaProds = obtenerHojaProductos(ssCostos);
    if (!hojaProds) return;

    const filaHeadersProds = detectarFilaHeaders(hojaProds);
    const headers = hojaProds.getRange(filaHeadersProds, 1, 1, hojaProds.getLastColumn()).getValues()[0]
      .map(h => h.toString().trim().toUpperCase());

    const colId = headers.indexOf("ID") + 1;
    const colUltimoPrecio = headers.indexOf("ULTIMO PRECIO") + 1;

    if (colId === 0 || colUltimoPrecio === 0) return;

    const data = hojaProds.getDataRange().getValues();

    for (const prod of productosPedido) {
      if (!prod.id) continue;
      // Solo sobreescribimos en "ULTIMO PRECIO" si seleccionó "ULTIMO_PRECIO"
      if (prod.tipo_precio !== "ULTIMO_PRECIO") continue;

      const filaIdx = data.findIndex(r => r[colId - 1] === prod.id);
      if (filaIdx !== -1) {
        const ultimoPrecioActual = parseFloat(data[filaIdx][colUltimoPrecio - 1]) || 0;
        if (prod.precio_unitario !== ultimoPrecioActual) {
          hojaProds.getRange(filaIdx + 1, colUltimoPrecio).setValue(prod.precio_unitario);
        }
      }
    }
  } catch (e) {
    console.error("Error al actualizar precios de productos: " + e.message);
  }
}

function normalizarRazonesSociales() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // Cambia "Clientes" por el nombre exacto de tu hoja si es diferente
  const sheet = ss.getSheetByName("Clientes");
  if (!sheet) return;

  const filaInicio = 2; // Suponiendo que la fila 1 son los encabezados
  const ultimaFila = sheet.getLastRow();

  if (ultimaFila < filaInicio) return; // Si no hay datos, salir

  // Obtenemos el rango de la columna A (Razón Social)
  const rango = sheet.getRange(filaInicio, 1, ultimaFila - filaInicio + 1, 1);
  const valores = rango.getValues();
  let huboCambios = false;

  // Recorremos fila por fila
  for (let i = 0; i < valores.length; i++) {
    const valorOriginal = valores[i][0].toString();
    const valorMayuscula = valorOriginal.toUpperCase();

    // Si el texto original es distinto a su versión en mayúsculas, lo reemplazamos
    if (valorOriginal !== valorMayuscula) {
      valores[i][0] = valorMayuscula;
      huboCambios = true;
    }
  }

  // Solo escribimos en la hoja si realmente se modificó algo (para que sea más rápido)
  if (huboCambios) {
    rango.setValues(valores);
    SpreadsheetApp.getUi().alert("✅ Éxito", "Se revisaron los clientes y se convirtieron las Razones Sociales a mayúsculas.", SpreadsheetApp.getUi().ButtonSet.OK);
  } else {
    SpreadsheetApp.getUi().alert("ℹ️ Información", "Todas las Razones Sociales ya estaban en mayúsculas.", SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function manejarConversionMoneda(e) {
  const sheet = e.range.getSheet();
  const nombreHoja = sheet.getName();
  const hojasPermitidas = [CONFIG.HOJA_PEDIDOS, CONFIG.HOJA_TERMINADOS, CONFIG.HOJA_ENTREGADOS];

  if (!hojasPermitidas.includes(nombreHoja)) return;

  const fila = e.range.getRow();
  if (fila <= CONFIG.FILA_HEADERS) return;
  // Obtenemos encabezados dinámicamente de la hoja actual
  const headers = sheet.getRange(CONFIG.FILA_HEADERS, 1, 1, sheet.getLastColumn()).getValues()[0]
    .map(h => h.toString().trim().toUpperCase());

  const colMoneda = headers.indexOf("MONEDA") + 1;
  const colMonedaEstado = headers.indexOf("MONEDA_ESTADO") + 1;
  const colTC = headers.indexOf("T.C") + 1;
  const colSubtotal = headers.indexOf("SUBTOTAL") + 1;
  const colTotal = headers.indexOf("TOTAL") + 1;
  const colCuenta = headers.indexOf("CUENTA") + 1;
  const colSaldo = headers.indexOf("SALDO") + 1;

  if (colMoneda === 0 || colMonedaEstado === 0 || colTC === 0 || colSubtotal === 0 || colTotal === 0) return;
  // Si estamos editando la columna Moneda
  if (e.range.getColumn() === colMoneda) {
    const nuevaMoneda = e.value;
    const estadoActual = sheet.getRange(fila, colMonedaEstado).getValue() || "AR$";
    if (nuevaMoneda === estadoActual) return; // No hay cambios
    const tc = parseFloat(sheet.getRange(fila, colTC).getValue()) || 1;
    let subtotal = parseFloat(sheet.getRange(fila, colSubtotal).getValue()) || 0;
    let total = parseFloat(sheet.getRange(fila, colTotal).getValue()) || 0;
    let cuenta = colCuenta > 0 ? (parseFloat(sheet.getRange(fila, colCuenta).getValue()) || 0) : 0;
    let saldo = colSaldo > 0 ? (parseFloat(sheet.getRange(fila, colSaldo).getValue()) || 0) : 0;
    // Lógica de conversión
    if (nuevaMoneda === "U$D" && estadoActual === "AR$") {
      subtotal /= tc;
      total /= tc;
      cuenta /= tc;
      saldo /= tc;
    } else if (nuevaMoneda === "AR$" && estadoActual === "U$D") {
      subtotal *= tc;
      total *= tc;
      cuenta *= tc;
      saldo *= tc;
    }
    // Guardar valores y actualizar estado
    sheet.getRange(fila, colSubtotal).setValue(subtotal);
    sheet.getRange(fila, colTotal).setValue(total);
    if (colCuenta > 0) sheet.getRange(fila, colCuenta).setValue(cuenta);
    if (colSaldo > 0) sheet.getRange(fila, colSaldo).setValue(saldo);
    sheet.getRange(fila, colMonedaEstado).setValue(nuevaMoneda);
  }
}

function iniciarDecisionStockPorEstado(sheet, fila, colEstado, estadoObjetivo, oldValue) {
  const headers = sheet.getRange(CONFIG.FILA_HEADERS, 1, 1, sheet.getLastColumn()).getValues()[0]
    .map(h => h.toString().trim().toUpperCase());
  const colPedido = headers.indexOf("PEDIDO") + 1;
  if (colPedido === 0) {
    SpreadsheetApp.getUi().alert("No se encontró la columna PEDIDO.");
    return;
  }

  const detalle = sheet.getRange(fila, colPedido).getValue();
  const productosPedido = parsearPedidoDetalle(detalle);
  if (!productosPedido.length) {
    SpreadsheetApp.getUi().alert("No se pudieron interpretar los productos del pedido.");
    return;
  }

  const dataDecision = construirDataDecisionStock(productosPedido);
  const props = PropertiesService.getScriptProperties();
  const estadoAnterior = (oldValue || "No Encomendado").toString();

  props.setProperty("DECISION_STOCK_CTX", JSON.stringify({
    hoja: sheet.getName(),
    fila,
    colEstado,
    estadoObjetivo,
    estadoAnterior,
    productosPedido,
    items: dataDecision.items,
    ts: Date.now()
  }));

  props.setProperty('BLOQUEAR_ONEDIT', '1');
  try {
    sheet.getRange(fila, colEstado).setValue(estadoAnterior);
  } finally {
    props.deleteProperty('BLOQUEAR_ONEDIT');
  }

  const html = HtmlService.createTemplateFromFile('FormularioDecisionStockEstado');
  SpreadsheetApp.getUi().showModalDialog(
    html.evaluate().setWidth(860).setHeight(680),
    'Decidir fuente de stock'
  );
}

function obtenerDataDecisionStockEstado() {
  const raw = PropertiesService.getScriptProperties().getProperty("DECISION_STOCK_CTX");
  if (!raw) return null;
  const ctx = JSON.parse(raw);
  return {
    hoja: ctx.hoja,
    fila: ctx.fila,
    estadoObjetivo: ctx.estadoObjetivo,
    items: ctx.items || []
  };
}

function aplicarDecisionStockEstado(decisiones) {
  try {
    const props = PropertiesService.getScriptProperties();
    const raw = props.getProperty("DECISION_STOCK_CTX");
    if (!raw) throw new Error("No hay una decisión de stock pendiente.");
    const ctx = JSON.parse(raw);

    const decisionesMap = {};
    (decisiones || []).forEach(d => {
      if (d && d.idProducto && d.fuente) decisionesMap[d.idProducto] = d.fuente.toUpperCase();
    });

    const faltantes = (ctx.productosPedido || []).filter(p => !decisionesMap[p.id]);
    if (faltantes.length) {
      throw new Error("Debes elegir una fuente de stock para todos los productos.");
    }

    descontarStockSegunDecision(ctx.productosPedido || [], decisionesMap);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName(ctx.hoja);
    if (!hoja) throw new Error("No se encontró la hoja del pedido.");

    props.setProperty('BLOQUEAR_ONEDIT', '1');
    try {
      hoja.getRange(ctx.fila, ctx.colEstado).setValue(ctx.estadoObjetivo);
    } finally {
      props.deleteProperty('BLOQUEAR_ONEDIT');
    }

    procesarFlujoPedidos(hoja, ctx.fila, ctx.colEstado, ctx.estadoObjetivo);
    registrarMovimientoLog({
      accion: "DECISION_STOCK",
      hojaOrigen: ctx.hoja,
      fila: ctx.fila,
      estado: ctx.estadoObjetivo,
      detalle: `Fuentes elegidas: ${JSON.stringify(decisionesMap)}`
    });
    props.deleteProperty("DECISION_STOCK_CTX");

    return { success: true };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function parsearPedidoDetalle(textoPedido) {
  if (!textoPedido) return [];
  const ssCostos = obtenerSSCostos();
  const hojaProds = obtenerHojaProductos(ssCostos);
  if (!hojaProds) return [];

  const filaHeaders = detectarFilaHeaders(hojaProds);
  const data = hojaProds.getDataRange().getValues();
  const headers = data[filaHeaders - 1].map(h => h.toString().trim().toUpperCase());
  const colId = headers.indexOf("ID");
  const colCat = headers.indexOf("CATEGORIA");
  const colMod = headers.indexOf("MODELO");

  return textoPedido.toString().split('\n').map(linea => {
    const partes = linea.split('|').map(s => s.trim());
    if (partes.length < 3) return null;
    const cantidad = parseFloat(partes[0]) || 0;
    const categoria = partes[1] || "";
    const modelo = partes[2] || "";
    if (!cantidad || !modelo) return null;

    const idx = data.findIndex((row, i) =>
      i >= filaHeaders &&
      row[colCat] && row[colMod] &&
      row[colCat].toString().trim().toUpperCase() === categoria.toUpperCase() &&
      row[colMod].toString().trim().toUpperCase() === modelo.toUpperCase()
    );
    if (idx === -1) return null;
    return {
      id: data[idx][colId],
      categoria,
      modelo,
      cantidad
    };
  }).filter(Boolean);
}

function construirDataDecisionStock(productosPedido) {
  const ssCostos = obtenerSSCostos();
  const hojaProds = obtenerHojaProductos(ssCostos);
  const hojaMat = ssCostos.getSheetByName(CONFIG.HOJA_MATERIALES);
  if (!hojaProds || !hojaMat) throw new Error("No se encontraron hojas de Productos/Materiales.");

  const filaHeadersProds = detectarFilaHeaders(hojaProds);
  const colsP = getColMap(hojaProds, HEADERS_COSTOS, filaHeadersProds);
  const colsM = getColMap(hojaMat, HEADERS_MATERIALES, 1);
  const datosP = hojaProds.getDataRange().getValues();
  const datosM = hojaMat.getDataRange().getValues();

  const items = productosPedido.map(prod => {
    const filaProdIdx = datosP.findIndex(r => r[colsP.ID - 1] === prod.id);
    if (filaProdIdx === -1) {
      return {
        idProducto: prod.id,
        nombre: `${prod.categoria} - ${prod.modelo}`,
        cantidad: prod.cantidad,
        stockProducto: 0,
        productoDisponible: false,
        materiales: [],
        materialesSuficientes: false
      };
    }

    const stockActual = parseFloat(datosP[filaProdIdx][colsP.STOCK - 1]) || 0;
    const recetaTexto = datosP[filaProdIdx][colsP.MATERIALES - 1] || "";
    const receta = parsearReceta(recetaTexto);

    let materiales = [];
    let materialesSuficientes = true; // Si no tiene receta (longitud 0), se considera suficiente (bypass)
    receta.forEach(mat => {
      const filaMatIdx = datosM.findIndex(r => r[colsM.DETALLES - 1] === mat.nombre);
      const stockMat = filaMatIdx !== -1 ? (parseFloat(datosM[filaMatIdx][colsM.STOCK - 1]) || 0) : 0;
      const requerido = (parseFloat(mat.cantidad) || 0) * prod.cantidad;
      if (stockMat < requerido) materialesSuficientes = false;
      materiales.push({
        nombre: mat.nombre,
        requerido,
        stock: stockMat
      });
    });

    return {
      idProducto: prod.id,
      nombre: `${prod.categoria} - ${prod.modelo}`,
      cantidad: prod.cantidad,
      stockProducto: stockActual,
      productoDisponible: stockActual >= prod.cantidad,
      materiales,
      materialesSuficientes
    };
  });

  return { items };
}

function descontarStockSegunDecision(productosPedido, decisionesMap) {
  const ssCostos = obtenerSSCostos();
  const hojaProds = obtenerHojaProductos(ssCostos);
  const hojaMat = ssCostos.getSheetByName("Materiales");

  if (!hojaProds || !hojaMat) throw new Error("No se encontraron las hojas de Productos o Materiales.");

  const filaHeadersProds = detectarFilaHeaders(hojaProds);
  const colsP = getColMap(hojaProds, HEADERS_COSTOS, filaHeadersProds);
  const colsM = getColMap(hojaMat, HEADERS_MATERIALES, 1);

  const datosP = hojaProds.getDataRange().getValues();
  const datosM = hojaMat.getDataRange().getValues();

  for (const prod of productosPedido) {
    if (!prod.id) continue;
    const decision = (decisionesMap[prod.id] || "").toUpperCase();
    if (!decision) throw new Error(`Falta decisión para ${prod.categoria} - ${prod.modelo}.`);

    const filaProdIdx = datosP.findIndex(r => r[colsP.ID - 1] === prod.id);
    if (filaProdIdx === -1) throw new Error(`No se encontró el producto ${prod.categoria} - ${prod.modelo}.`);

    const filaProdSheet = filaProdIdx + 1;
    const stockActual = parseFloat(datosP[filaProdIdx][colsP.STOCK - 1]) || 0;

    if (decision === "PRODUCTO") {
      if (stockActual < prod.cantidad) {
        throw new Error(`Stock insuficiente de producto para ${prod.categoria} - ${prod.modelo}.`);
      }
      const nuevoStock = stockActual - prod.cantidad;
      hojaProds.getRange(filaProdSheet, colsP.STOCK).setValue(nuevoStock);
      datosP[filaProdIdx][colsP.STOCK - 1] = nuevoStock;
      continue;
    }

    if (decision === "MATERIALES") {
      const recetaTexto = datosP[filaProdIdx][colsP.MATERIALES - 1] || "";
      const materialesReceta = parsearReceta(recetaTexto);
      if (!materialesReceta.length) {
        // BYPASS: Si no tiene receta de materiales, no restamos nada y continuamos
        continue;
      }

      for (const mat of materialesReceta) {
        const requerida = mat.cantidad * prod.cantidad;
        const filaMatIdx = datosM.findIndex(r => r[colsM.DETALLES - 1] === mat.nombre);
        if (filaMatIdx === -1) throw new Error(`No se encontró el material ${mat.nombre}.`);
        const filaMatSheet = filaMatIdx + 1;
        const stockMatActual = parseFloat(datosM[filaMatIdx][colsM.STOCK - 1]) || 0;
        const nuevoStockMat = stockMatActual - requerida;
        hojaMat.getRange(filaMatSheet, colsM.STOCK).setValue(nuevoStockMat);
        datosM[filaMatIdx][colsM.STOCK - 1] = nuevoStockMat;
      }
      continue;
    }

    throw new Error(`Fuente inválida para ${prod.categoria} - ${prod.modelo}.`);
  }
}

function abrirModalAlertasStock() {
  const html = HtmlService.createTemplateFromFile('FormularioAlertasStock');
  SpreadsheetApp.getUi().showModalDialog(
    html.evaluate().setWidth(920).setHeight(700),
    'Gestión de alertas de stock mínimo'
  );
}

function obtenerDataAlertasStock() {
  const materiales = obtenerCatalogoMaterialesAlertas();
  const config = leerAlertasStockConfig();
  return {
    materiales,
    alertas: config.items || []
  };
}

function guardarAlertasStock(items) {
  const saneadas = (items || [])
    .filter(i => i && i.material)
    .map(i => ({
      material: i.material.toString().trim(),
      minimo: Math.max(0, parseFloat(i.minimo) || 0)
    }));
  const payload = {
    version: 1,
    updatedAt: new Date().toISOString(),
    items: saneadas
  };
  PropertiesService.getScriptProperties().setProperty("ALERTAS_STOCK_CFG", JSON.stringify(payload));
  return { success: true };
}

function eliminarAlertaStock(material) {
  const cfg = leerAlertasStockConfig();
  cfg.items = (cfg.items || []).filter(i => i.material !== material);
  cfg.updatedAt = new Date().toISOString();
  PropertiesService.getScriptProperties().setProperty("ALERTAS_STOCK_CFG", JSON.stringify(cfg));
  return { success: true };
}

function revisarAlertasStock(desdeOnOpen) {
  const alertas = obtenerAlertasDisparadas();
  if (!alertas.length) {
    if (!desdeOnOpen) SpreadsheetApp.getUi().alert("No hay materiales por debajo del stock mínimo configurado.");
    return { success: true, count: 0 };
  }

  const props = PropertiesService.getScriptProperties();
  props.setProperty("ALERTAS_STOCK_LAST", JSON.stringify({
    ts: Date.now(),
    rows: alertas
  }));

  const html = HtmlService.createTemplateFromFile('FormularioRevisionStocks');
  SpreadsheetApp.getUi().showModalDialog(
    html.evaluate().setWidth(980).setHeight(680),
    'Materiales con stock bajo'
  );
  return { success: true, count: alertas.length };
}

function obtenerUltimaRevisionStocks() {
  const raw = PropertiesService.getScriptProperties().getProperty("ALERTAS_STOCK_LAST");
  if (!raw) return [];
  try {
    const parsed = JSON.parse(raw);
    return parsed.rows || [];
  } catch (e) {
    return [];
  }
}

function obtenerAlertasDisparadas() {
  const materiales = obtenerCatalogoMaterialesAlertas();
  const cfg = leerAlertasStockConfig();
  const mapaMinimos = {};
  (cfg.items || []).forEach(i => {
    mapaMinimos[i.material] = parseFloat(i.minimo) || 0;
  });

  return materiales
    .filter(m => mapaMinimos[m.nombre] !== undefined)
    .map(m => ({
      nombre: m.nombre,
      categoria: m.categoria,
      proveedor: m.proveedor,
      detalle: m.detalle,
      stockMinimo: mapaMinimos[m.nombre],
      stockActual: m.stock
    }))
    .filter(r => r.stockActual <= r.stockMinimo);
}

function obtenerCatalogoMaterialesAlertas() {
  const ssCostos = obtenerSSCostos();
  const hojaMat = ssCostos.getSheetByName(CONFIG.HOJA_MATERIALES);
  if (!hojaMat) return [];

  const data = hojaMat.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0].map(h => h.toString().trim().toUpperCase());

  const colDetalle = headers.indexOf("DETALLES");
  const colCat = headers.indexOf("CATEGORIA");
  const colProv = headers.indexOf("PROVEEDOR");
  const colStock = headers.indexOf("STOCK");

  return data.slice(1)
    .filter(row => colDetalle !== -1 && row[colDetalle])
    .map(row => ({
      nombre: row[colDetalle].toString().trim(),
      detalle: row[colDetalle].toString().trim(),
      categoria: colCat !== -1 ? (row[colCat] || "").toString().trim() : "",
      proveedor: colProv !== -1 ? (row[colProv] || "").toString().trim() : "",
      stock: colStock !== -1 ? (parseFloat(row[colStock]) || 0) : 0
    }));
}

function leerAlertasStockConfig() {
  const raw = PropertiesService.getScriptProperties().getProperty("ALERTAS_STOCK_CFG");
  if (!raw) return { version: 1, updatedAt: null, items: [] };
  try {
    const parsed = JSON.parse(raw);
    return {
      version: parsed.version || 1,
      updatedAt: parsed.updatedAt || null,
      items: Array.isArray(parsed.items) ? parsed.items : []
    };
  } catch (e) {
    return { version: 1, updatedAt: null, items: [] };
  }
}

function construirLinkCliente(nombreCliente) {
  const nombre = (nombreCliente || "").toString().trim();
  if (!nombre) return null;
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hojaClientes = ss.getSheetByName(CONFIG.HOJA_CLIENTES);
    if (!hojaClientes || hojaClientes.getLastRow() < 2) return nombre;
    const gidClientes = hojaClientes.getSheetId();
    const clientesData = hojaClientes.getRange("A2:A" + hojaClientes.getLastRow()).getValues().flat();
    const filaIdx = clientesData.findIndex(c => c && c.toString().trim().toUpperCase() === nombre.toUpperCase());
    if (filaIdx === -1) return nombre;
    const filaCliente = filaIdx + 2;
    return `=HYPERLINK("#gid=${gidClientes}&range=A${filaCliente}"; "${nombre.toUpperCase()}")`;
  } catch (e) {
    return nombre;
  }
}

function extraerNombreCliente(valorCliente, formulaCliente) {
  if (valorCliente && valorCliente.toString().trim() !== "") {
    return valorCliente.toString().trim();
  }
  const formula = (formulaCliente || "").toString();
  if (!formula) return "";
  const m = formula.match(/;\s*"([^"]+)"\s*\)$/);
  if (m && m[1]) return m[1].trim();
  return "";
}

function registrarMovimientoLog({ accion, hojaOrigen, hojaDestino, fila, estado, detalle }) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hojaLog = obtenerHojaLog(ss);
    const email = Session.getActiveUser().getEmail() || "desconocido";
    hojaLog.appendRow([
      new Date(),
      accion || "",
      hojaOrigen || "",
      hojaDestino || "",
      fila || "",
      estado || "",
      detalle || "",
      email
    ]);
  } catch (e) {
    console.error("No se pudo registrar en log: " + e.message);
  }
}

function obtenerHojaLog(ss) {
  let hoja = ss.getSheetByName(CONFIG.HOJA_LOG);
  if (!hoja) {
    hoja = ss.insertSheet(CONFIG.HOJA_LOG);
    hoja.getRange(1, 1, 1, 8).setValues([[
      "FECHA_HORA",
      "ACCION",
      "HOJA_ORIGEN",
      "HOJA_DESTINO",
      "FILA",
      "ESTADO",
      "DETALLE",
      "USUARIO"
    ]]);
    hoja.getRange(1, 1, 1, 8).setFontWeight("bold");
    hoja.setFrozenRows(1);
  }
  return hoja;
}