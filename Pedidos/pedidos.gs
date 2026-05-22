// --- CONFIGURACIÓN Y CONEXIONES ---
const CONFIG = {
  ID_COSTOS: "1JHKK7tu30UyPeeq9hgwF3U2P8_84NtzJfINVRfa0WPA",
  HOJA_MATERIALES: "Materiales",
  HOJA_COSTOS: "Productos",
  HOJA_PEDIDOS: "Pedidos",
  HOJA_TERMINADOS: "Terminados",
  HOJA_ENTREGADOS: "Entregados",
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
    .addSeparator()
    .addItem('👤 Agregar Nuevo Cliente', 'abrirPromptNuevoCliente')
    .addItem('🔠 Convertir Clientes a Mayúsculas', 'normalizarRazonesSociales')
    .addToUi();
}

// Nueva función: Agregar cliente desde Prompt
function abrirPromptNuevoCliente() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.alert('Nuevo Cliente', '¿Desea ingresar un nuevo cliente?', ui.ButtonSet.YES_NO);
  if (res !== ui.Button.YES) return;

  const razon = ui.prompt('Razón Social', 'Ingrese Razón Social:', ui.ButtonSet.OK_CANCEL);
  const alias = ui.prompt('Alias', 'Ingrese Alias (separados por ;):', ui.ButtonSet.OK_CANCEL);
  const contacto = ui.prompt('Contacto', 'Ingrese Contacto (separados por ;):', ui.ButtonSet.OK_CANCEL);

  if (razon.getSelectedButton() !== ui.Button.OK || 
      (razon.getResponseText() === "" && alias.getResponseText() === "" && contacto.getResponseText() === "")) {
    ui.alert('❌ Error: Debe completar al menos un campo.');
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.HOJA_CLIENTES);
  sheet.appendRow([razon.getResponseText().toUpperCase(), alias.getResponseText(), contacto.getResponseText()]);
  ui.alert('✅ Cliente agregado correctamente.');
}


// Renombramos a onEditSimple para que Google no lo bloquee por seguridad
function onEditSimple(e) {
  if (!e?.range) return;
  const sheet = e.range.getSheet();
  const name = sheet.getName();
  const col = e.range.getColumn();
  const row = e.range.getRow();
  const val = e.value;

  // Solo procesar filas de datos (ignorar encabezados)
  if (row <= CONFIG.FILA_HEADERS) return;

  // 2. Lógica de Flujo de Pedidos: detectar columna ESTADO dinámicamente
  if (name === CONFIG.HOJA_PEDIDOS || name === CONFIG.HOJA_TERMINADOS) {
    // Leer los encabezados de la fila 3 para encontrar la columna ESTADO
    const headers = sheet.getRange(CONFIG.FILA_HEADERS, 1, 1, sheet.getLastColumn()).getValues()[0];
    const colEstado = headers.findIndex(h => h.toString().trim().toUpperCase() === "ESTADO") + 1;

    if (colEstado > 0 && col === colEstado) {
      if (name === CONFIG.HOJA_PEDIDOS) {
        procesarFlujoPedidos(sheet, row, col, val);
      } else if (name === CONFIG.HOJA_TERMINADOS) {
        procesarFlujoTerminados(sheet, row, col, val);
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
  onEditSimple(e);
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

    for (let i = filaInicioProds - 1; i < data.length; i++) {
      if (colId !== -1 && data[i][colId] && data[i][colId].toString().trim() !== "") {
        let recetaTexto = colMat !== -1 && data[i][colMat] ? data[i][colMat].toString() : "";
        let precioVal = colPrecio !== -1 && data[i][colPrecio] !== undefined ? parseFloat(data[i][colPrecio]) || 0 : 0;
        productos.push({
          id: data[i][colId],
          categoria: colCat !== -1 && data[i][colCat] ? data[i][colCat].toString().trim() : "Sin Categoría",
          nombre: ((colCat !== -1 && data[i][colCat] ? data[i][colCat] : "") + " - " + (colMod !== -1 && data[i][colMod] ? data[i][colMod] : "")).trim(),
          stock: colStock !== -1 ? parseFloat(data[i][colStock]) || 0 : 0,
          receta: parsearReceta(recetaTexto),
          precio: precioVal
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

    // Crear el texto del detalle del pedido: Cantidad | Categoria | Modelo | Precio Unitario
    let pedidoDetalle = datos.productos.map(p =>
      `${p.cantidad} | ${p.categoria} | ${p.modelo} | $${p.precio_unitario}`
    ).join('\n');

    // Parsear fecha para evitar el error de formato
    let dateObj = new Date(datos.fecha + "T12:00:00");

    sheet.getRange(filaNum, 1).setValue(dateObj).setNumberFormat('dd/mm/yyyy'); // Col A: FECHA
    sheet.getRange(filaNum, 2).setValue(datos.cliente); // Col B: CLIENTE
    sheet.getRange(filaNum, 3).setValue(pedidoDetalle); // Col C: PEDIDO

    // Col D (4): SUBTOTAL (Numérico)
    sheet.getRange(filaNum, 4).setValue(datos.subtotalBase);

    // Col E (5): TOTAL (Numérico)
    sheet.getRange(filaNum, 5).setValue(datos.totalGeneral);

    // Col F (6): CUENTA (Monto por el cual se pagó o 0)
    sheet.getRange(filaNum, 6).setValue(datos.montoPago || 0);

    // Col G (7): T.C
    sheet.getRange(filaNum, 7).setValue(datos.tc);

    // Col H (8): DTO (porcentaje de descuento en cascada)
    let dtoEquiv = datos.factorDesc === 1 ? 0 : (1 - datos.factorDesc);
    sheet.getRange(filaNum, 8).setValue(dtoEquiv).setNumberFormat("0.00%");

    // Col I (9): IMP (Porcentaje de impuestos)
    let impPctNum = parseFloat(datos.impuesto_pct) || 0;
    sheet.getRange(filaNum, 9).setValue(impPctNum / 100).setNumberFormat("0.00%");

    // Col J (10): DIAS TRANSCURRIDOS (Ignorado - Fórmula Excel)

    // Col K (11): PLAZO DE ENTREGA
    sheet.getRange(filaNum, 11).setValue(datos.plazo);

    // Col L (12): TIEMPO (Ignorado - Fórmula Excel)

    // Col M (13): FORMA ENTREGA (RETIRA | NOMBRE o ENVIO | DETALLE)
    let f_entrega_str = datos.f_entrega.toUpperCase() === "RETIRA"
      ? `RETIRA | ${datos.detalle_entrega || ''}`
      : `ENVIO | ${datos.detalle_entrega || ''}`;
    sheet.getRange(filaNum, 13).setValue(f_entrega_str);

    // Col N (14): ESTADO
    let estadoFinal = datos.estado || "No Encomendado";
    sheet.getRange(filaNum, 14).setValue(estadoFinal);

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

  const datosOrigen = hojaOrigen.getRange(filaOrigen, 1, 1, hojaOrigen.getLastColumn()).getValues()[0];

  // Construir la nueva fila para el destino
  let nuevaFila = headersDestino.map(header => {
    // 1. Prioridad: Valores extras pasados por parámetro (ej: FECHA ENTREGA)
    if (valoresExtras[header] !== undefined) return valoresExtras[header];

    // 2. Buscar por nombre en el origen
    const idxOrigen = headersOrigen.indexOf(header);
    if (idxOrigen !== -1) return datosOrigen[idxOrigen];

    // 3. Si no existe en el origen ni en extras, devolver vacío
    return "";
  });

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

  if (valNorm === "TERMINADO") {
    const destino = ss.getSheetByName(CONFIG.HOJA_TERMINADOS);
    if (!destino) return;
    transferirFilaDinamica(hoja, fila, destino);
    hoja.deleteRow(fila);
  }
  else if (valNorm === "ENTREGADO") {
    moverAEntregados(hoja, fila);
  }
  else if (valNorm === "ENTREGA PARCIAL") {
    hoja.getRange(fila, 12).setValue("En Produccion");
    PropertiesService.getScriptProperties().setProperty('FILA_ENTREGA_PARCIAL', fila.toString());
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
  if (!fila) return null;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.HOJA_PEDIDOS);

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
    const sheet = ss.getSheetByName(CONFIG.HOJA_PEDIDOS);
    const hEntregados = ss.getSheetByName(CONFIG.HOJA_ENTREGADOS);
    const fila = datos.fila;

    const headersOrigen = sheet.getRange(CONFIG.FILA_HEADERS, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => h.toString().trim().toUpperCase());
    const originalValues = sheet.getRange(fila, 1, 1, sheet.getLastColumn()).getValues()[0];

    let entregados = datos.productos_entregados.filter(p => p.cantidad_entregada > 0);
    if (!entregados.length) throw new Error('No se seleccionaron productos.');

    let subEnt = entregados.reduce((sum, p) => sum + (p.cantidad_entregada * p.precio_unitario), 0);
    let descPct = parseFloat(originalValues[headersOrigen.indexOf("DTO")]) || 0;
    let impPct = parseFloat(originalValues[headersOrigen.indexOf("IMP")]) || 0;
    let totalEnt = (subEnt * (1 - descPct)) * (1 + impPct);

    // 1. ENVIAR LO ENTREGADO A LA HOJA "ENTREGADOS"
    const headersDest = hEntregados.getRange(CONFIG.FILA_HEADERS, 1, 1, hEntregados.getLastColumn()).getValues()[0].map(h => h.toString().trim().toUpperCase());
    let filaParaEntregados = headersDest.map(header => {
      if (header === "PEDIDO") return entregados.map(p => `${p.cantidad_entregada} | ${p.categoria} | ${p.modelo} | $${p.precio_unitario}`).join('\n');
      if (header === "SUBTOTAL") return subEnt;
      if (header === "TOTAL") return totalEnt;
      if (header === "T.C") return datos.tc;
      if (header === "ESTADO") return "ENTREGADO";
      if (header === "CUENTA") return datos.pago_recibido || 0;
      if (header === "FECHA ENTREGA") return new Date();

      const idx = headersOrigen.indexOf(header);
      return idx !== -1 ? originalValues[idx] : "";
    });

    hEntregados.appendRow(filaParaEntregados);
    const lastRowEnt = hEntregados.getLastRow();
    headersDest.forEach((h, i) => { if (h.includes("FECHA")) hEntregados.getRange(lastRowEnt, i + 1).setNumberFormat('dd/mm/yyyy'); });

    // 2. ACTUALIZAR LO QUE QUEDA EN LA HOJA "PEDIDOS"
    let restantes = datos.productos_entregados.filter(p => (p.cantidad_original - p.cantidad_entregada) > 0);

    if (!restantes.length) {
      sheet.deleteRow(fila);
      return { success: true };
    }

    let detalleRestante = restantes.map(p => `${p.cantidad_original - p.cantidad_entregada} | ${p.categoria} | ${p.modelo} | $${p.precio_unitario}`).join('\n');
    let subRest = restantes.reduce((sum, p) => sum + ((p.cantidad_original - p.cantidad_entregada) * p.precio_unitario), 0);
    let totalRestante = (subRest * (1 - descPct)) * (1 + impPct);

    let pagoRecibido = parseFloat(datos.pago_recibido) || 0;
    let nuevoPagado = (datos.pagado_hasta_ahora || 0) + pagoRecibido;

    let nuevoSaldoTotal = totalRestante - pagoRecibido;
    let nuevoSaldoSubtotal = subRest - pagoRecibido;

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

    for (let i = 1; i < data.length; i++) {
      if (colDetalle !== -1 && data[i][colDetalle]) {
        materiales.push({
          nombre: data[i][colDetalle].toString().trim(),
          stock: colStock !== -1 ? parseFloat(data[i][colStock]) || 0 : 0,
          proveedor: colProv !== -1 ? data[i][colProv].toString().trim() : "Sin Proveedor"
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
    let colPrecio = headers.indexOf("PRECIO") + 1;
    if (colPrecio === 0 && hojaProds.getLastColumn() >= 14) {
      colPrecio = 14; // Fallback a columna N
    }

    if (colId === 0 || colPrecio === 0) return;

    const data = hojaProds.getDataRange().getValues();

    for (const prod of productosPedido) {
      if (!prod.id) continue;

      const filaIdx = data.findIndex(r => r[colId - 1] === prod.id);
      if (filaIdx !== -1) {
        const precioActual = parseFloat(data[filaIdx][colPrecio - 1]) || 0;
        if (prod.precio_unitario !== precioActual) {
          hojaProds.getRange(filaIdx + 1, colPrecio).setValue(prod.precio_unitario);
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
