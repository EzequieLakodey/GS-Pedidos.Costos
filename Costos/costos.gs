const CONFIG = {
    HOJA_MATERIALES: "Materiales",
    HOJA_PRODUCTOS: "Productos",
    FILA_ENCABEZADOS_PRODUCTOS: 3,
    FILA_INICIO_DATOS_PRODUCTOS: 4,
    FILA_ENCABEZADOS_MATERIALES: 1,
    FILA_INICIO_DATOS_MATERIALES: 2,
    FILA_INICIO_DATOS: 4,
};
// Nombres exactos de headers
const HEADERS_PRODUCTOS = {
    ID: "ID",
    CATEGORIA: "CATEGORIA",
    MODELO: "MODELO",
    MATERIALES: "MATERIALES",
    COSTO_MAT: "COSTO MAT",
    HORAS_MOBRA: "HORAS M.OBRA",
    MOBRA_USD: "M.OBRA USD",
    SUBTOTAL: "SUBTOTAL",
    DESPERDICIOS: "DESPERDICIOS",
    IMPUESTOS: "IMPUESTOS",
    IMPORTES_IMP: "IMPORTES IMPUESTOS",
    TOTAL: "TOTAL",
    STOCK: "STOCK",
    PRECIO: "PRECIO",
    ULTIMO_PRECIO: "ULTIMO PRECIO",
};
const HEADERS_MATERIALES = {
    ID: "ID",
    CATEGORIA: "Categoria",
    DETALLES: "Detalles",
    UNIDAD: "Unidad de Medida",
    PROVEEDOR: "Proveedor",
    PU_USD: "P/U USD",
    PU_ARS: "P/U ARS",
    IMP: "%IMP",
    DTO: "%DTO",
    TOTAL: "Total",
    STOCK: "STOCK",
};

function getColIndex(sheet, headerName, headerRow) {
    const row = headerRow || 1;
    const lastCol = sheet.getLastColumn();
    if (lastCol === 0) return -1;
    const vals = sheet.getRange(row, 1, 1, lastCol).getValues()[0];
    // Trim y limpieza de caracteres invisibles para evitar errores de tipeo
    const target = headerName.toString().trim().toUpperCase();
    const idx = vals.findIndex(
        (v) => v.toString().trim().toUpperCase() === target,
    );
    return idx === -1 ? -1 : idx + 1;
}

function getColMap(sheet, headersObj, headerRow) {
    if (!sheet) throw new Error("Hoja no válida o no encontrada");
    const map = {};
    const missing = [];
    Object.entries(headersObj).forEach(([key, label]) => {
        const col = getColIndex(sheet, label, headerRow);
        if (col === -1) missing.push(label);
        map[key] = col;
    });
    if (missing.length)
        throw new Error(
            'Columnas no encontradas en "' +
                sheet.getName() +
                '": ' +
                missing.join(", "),
        );
    return map;
}

function getColMapSafe(sheet, headersObj, headerRow) {
    const map = {};
    Object.entries(headersObj).forEach(([key, label]) => {
        map[key] = getColIndex(sheet, label, headerRow);
    });
    return map;
}

function onOpen() {
    SpreadsheetApp.getUi()
        .createMenu("⚙️ Gestión de Costos")
        .addItem("📦 Gestionar materiales", "gestionarMateriales")
        .addSeparator()
        .addItem("💵 Modificar Tipo de Cambio", "modificarTipoCambio")
        .addItem("📈 Aumentos de Materiales", "mostrarMenuAumentos")
        .addItem("📈 Aumento Jornales", "aplicarAumentoJornales")
        .addSeparator()
        .addItem("🆔 Migrar IDs", "migrarIdsExistentes")
        .addSeparator()
        .addItem("🗑️ Limpiar materiales", "limpiarMateriales")
        .addToUi();
}

function onEdit(e) {
    if (!e?.range) return;
    const sheet = e.range.getSheet();
    if (sheet.getName() !== CONFIG.HOJA_MATERIALES) return;
    const fila = e.range.getRow();
    if (fila < CONFIG.FILA_INICIO_DATOS) return;
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hojaCostos = ss.getSheetByName(CONFIG.HOJA_PRODUCTOS);
    if (!hojaCostos) return;
    const tc = parseFloat(hojaCostos.getRange("G1").getValue());
    if (!tc || isNaN(tc) || tc <= 0) return;
    const cols = getColMapSafe(
        sheet,
        HEADERS_MATERIALES,
        CONFIG.FILA_ENCABEZADOS_MATERIALES,
    );
    const COL_USD = cols.PU_USD,
        COL_ARS = cols.PU_ARS;
    if (COL_USD < 1 || COL_ARS < 1) return;
    let valE = parseFloat(sheet.getRange(fila, COL_USD).getValue()) || 0;
    let valF = parseFloat(sheet.getRange(fila, COL_ARS).getValue()) || 0;
    if (valE === 0 && valF === 0) return;
    const col = e.range.getColumn();
    if (col === COL_ARS || (valF > 0 && valE === 0)) {
        const conv = valF / tc;
        if (Math.abs(valE - conv) > 0.001)
            sheet.getRange(fila, COL_USD).setValue(conv);
    } else if (col === COL_USD || (valE > 0 && valF === 0)) {
        const conv = valE * tc;
        if (Math.abs(valF - conv) > 0.001)
            sheet.getRange(fila, COL_ARS).setValue(conv);
    }
}

function gestionarMateriales() {
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const sheet = ss.getActiveSheet();
        if (sheet.getName() !== CONFIG.HOJA_PRODUCTOS)
            throw new Error("Debes estar en la hoja " + CONFIG.HOJA_PRODUCTOS);
        const fila = sheet.getActiveCell().getRow();
        if (fila < CONFIG.FILA_INICIO_DATOS_PRODUCTOS)
            throw new Error("Selecciona una fila de productos válida.");

        PropertiesService.getScriptProperties().setProperty(
            "FILA_ACTIVA",
            fila.toString(),
        );
        const cols = getColMap(
            sheet,
            HEADERS_PRODUCTOS,
            CONFIG.FILA_ENCABEZADOS_PRODUCTOS,
        );

        const materialesExistentes = sheet
            .getRange(fila, cols.MATERIALES)
            .getValue();
        const esEdicion = materialesExistentes?.toString().trim() !== "";
        const html = HtmlService.createTemplateFromFile("Formulario");
        html.filaActiva = fila;
        html.modoEdicion = esEdicion;
        SpreadsheetApp.getUi().showModalDialog(
            html.evaluate().setWidth(700).setHeight(750),
            (esEdicion ? "✏️ Editando" : "➕ Agregando") +
                " materiales - Fila " +
                fila,
        );
    } catch (e) {
        SpreadsheetApp.getUi().alert(e.message);
    }
}
function obtenerDatosIniciales() {
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const fila = parseInt(
            PropertiesService.getScriptProperties().getProperty("FILA_ACTIVA"),
        );

        const sheetCostos = ss.getSheetByName(CONFIG.HOJA_PRODUCTOS);
        // CORRECCIÓN: Usar sheetCostos en lugar de sheet (que no existía)
        const colsCostos = getColMap(
            sheetCostos,
            HEADERS_PRODUCTOS,
            CONFIG.FILA_ENCABEZADOS_PRODUCTOS,
        );
        const categoriaActual = sheetCostos
            .getRange(fila, colsCostos.CATEGORIA)
            .getValue();

        const sheetCompras = ss.getSheetByName("Compras");
        const coeficientes = [{ nombre: "Ninguno", valor: 0 }];
        if (sheetCompras) {
            const ultima = sheetCompras.getLastRow();
            if (ultima >= 3)
                sheetCompras
                    .getRange("O3:R" + ultima)
                    .getValues()
                    .forEach((r) => {
                        if (r[0])
                            coeficientes.push({
                                nombre: r[0],
                                valor: parseFloat(r[3]) || 0,
                            });
                    });
        }

        const hojaMat = ss.getSheetByName(CONFIG.HOJA_MATERIALES);
        // CORRECCIÓN: Usar hojaMat y HEADERS_MATERIALES
        const colsMat = getColMap(
            hojaMat,
            HEADERS_MATERIALES,
            CONFIG.FILA_ENCABEZADOS_MATERIALES,
        );

        const catalogo = hojaMat
            .getDataRange()
            .getValues()
            .slice(CONFIG.FILA_INICIO_DATOS_MATERIALES - 1)
            .filter((r) => r[colsMat.DETALLES - 1])
            .map((r) => ({
                categoria: r[colsMat.CATEGORIA - 1],
                nombre: r[colsMat.DETALLES - 1],
                unidad: r[colsMat.UNIDAD - 1],
                proveedor: r[colsMat.PROVEEDOR - 1],
                precio: r[colsMat.PU_ARS - 1] || 0,
            }));

        return {
            success: true,
            catalogo,
            materialesFila: obtenerMaterialesDeFila(fila, coeficientes)
                .materiales,
            coeficientes,
            categoriaActual,
        };
    } catch (e) {
        return { success: false, message: e.message };
    }
}

function obtenerMaterialesDeFila(fila, listaCoef) {
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const sheetCostos = ss.getSheetByName(CONFIG.HOJA_PRODUCTOS);
        // CORRECCIÓN: Usar sheetCostos
        const colsCostos = getColMap(
            sheetCostos,
            HEADERS_PRODUCTOS,
            CONFIG.FILA_ENCABEZADOS_PRODUCTOS,
        );
        const texto = sheetCostos
            .getRange(fila, colsCostos.MATERIALES)
            .getValue();
        if (!texto) return { materiales: [] };

        const hojaMat = ss.getSheetByName(CONFIG.HOJA_MATERIALES);
        // CORRECCIÓN: Usar hojaMat y HEADERS_MATERIALES
        const colsMat = getColMap(
            hojaMat,
            HEADERS_MATERIALES,
            CONFIG.FILA_ENCABEZADOS_MATERIALES,
        );

        const mapaPrecios = {};
        hojaMat
            .getDataRange()
            .getValues()
            .forEach((r) => {
                if (r[colsMat.DETALLES - 1])
                    mapaPrecios[r[colsMat.DETALLES - 1]] = {
                        precio: r[colsMat.PU_ARS - 1],
                        cat: r[colsMat.CATEGORIA - 1],
                        prov: r[colsMat.PROVEEDOR - 1],
                        uni: r[colsMat.UNIDAD - 1],
                    };
            });

        const regex =
            /(.+?)\s*-\s*(.+?)\s*-\s*(.+?)\s*\(([\d.]+)\s*(.+?)\)(?:\s*\{W:(.+?)\})?/;
        const materiales = texto
            .toString()
            .split("\n")
            .map((linea) => {
                const m = linea.match(regex);
                if (!m) return null;
                const nombre = m[3].trim();
                const coefNombre = m[6] || "Ninguno";
                const infoCoef = listaCoef?.find(
                    (c) => c.nombre === coefNombre,
                );
                const infoPrecio = mapaPrecios[nombre] || {};
                return {
                    categoria: m[1].trim(),
                    proveedor: m[2].trim(),
                    nombre,
                    cantidad: parseFloat(m[4]),
                    unidad: m[5].trim(),
                    precio: infoPrecio.precio || 0,
                    coefNombre,
                    coefValor: infoCoef?.valor || 0,
                };
            })
            .filter(Boolean);
        return { materiales };
    } catch (e) {
        return { materiales: [] };
    }
}

function guardarMateriales(materiales) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const fila =
        PropertiesService.getScriptProperties().getProperty("FILA_ACTIVA");
    const sheet = ss.getSheetByName(CONFIG.HOJA_PRODUCTOS);
    const cols = getColMap(
        sheet,
        HEADERS_PRODUCTOS,
        CONFIG.FILA_ENCABEZADOS_PRODUCTOS,
    );
    const texto = materiales
        .map(
            (m) =>
                `${m.categoria || "S/C"} - ${m.proveedor || "S/P"} - ${m.nombre} (${m.cantidad} ${m.unidad}) {W:${m.coefNombre || "Ninguno"}}`,
        )
        .join("\n");
    let totalCostoFinal = 0,
        totalDesperdicio = 0;
    materiales.forEach((m) => {
        const base = m.precio * m.cantidad;
        const desp = base * (parseFloat(m.coefValor) || 0);
        totalDesperdicio += desp;
        totalCostoFinal += base + desp;
    });
    sheet.getRange(fila, cols.MATERIALES).setValue(texto);
    sheet.getRange(fila, cols.COSTO_MAT).setValue(totalCostoFinal);
    sheet.getRange(fila, cols.DESPERDICIOS).setValue(totalDesperdicio);
    return { success: true };
}

function guardarPreset(nombre, materiales) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const fila =
        PropertiesService.getScriptProperties().getProperty("FILA_ACTIVA");
    const sheet = ss.getSheetByName(CONFIG.HOJA_PRODUCTOS);
    const cols = getColMap(
        sheet,
        HEADERS_PRODUCTOS,
        CONFIG.FILA_ENCABEZADOS_PRODUCTOS,
    );
    const cat = sheet.getRange(fila, cols.CATEGORIA).getValue();
    const sh =
        ss.getSheetByName("Presets") || ss.insertSheet("Presets").hideSheet();
    if (sh.getLastRow() === 0) sh.appendRow(["Categoría", "Nombre", "JSON"]);
    sh.appendRow([cat, nombre, JSON.stringify(materiales)]);
    return { success: true };
}

function aplicarAumentoJornales() {
    const ui = SpreadsheetApp.getUi();
    const hojaJornales =
        SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Jornales");
    if (!hojaJornales)
        return ui.alert(
            "❌ Error",
            'No se encontró la hoja "Jornales".',
            ui.ButtonSet.OK,
        );
    const resp = ui.prompt(
        "📈 Aumento de Sueldos",
        "Ingresa el porcentaje de aumento (ej: 15):",
        ui.ButtonSet.OK_CANCEL,
    );
    if (resp.getSelectedButton() !== ui.Button.OK) return;
    const pct = parseFloat(resp.getResponseText().replace(",", ".").trim());
    if (isNaN(pct) || pct <= 0)
        return ui.alert(
            "❌ Error",
            "Ingresa un número válido mayor a 0.",
            ui.ButtonSet.OK,
        );
    const ultima = hojaJornales.getLastRow();
    if (ultima < 4) return;
    const rango = hojaJornales.getRange(4, 3, ultima - 3, 1);
    const vals = rango.getValues();
    const mult = 1 + pct / 100;
    let cnt = 0;
    vals.forEach((r, i) => {
        if (typeof r[0] === "number" && r[0] > 0) {
            vals[i][0] *= mult;
            cnt++;
        }
    });
    if (cnt > 0) {
        rango.setValues(vals);
        ui.alert(
            "✅ Éxito",
            `${pct}% aplicado a ${cnt} filas.`,
            ui.ButtonSet.OK,
        );
    }
}

function limpiarMateriales() {
    try {
        const ui = SpreadsheetApp.getUi();
        const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
        const fila = sheet.getActiveCell().getRow();
        if (fila < CONFIG.FILA_INICIO_DATOS)
            return ui.alert("Selecciona una fila de datos");
        if (
            ui.alert(
                "Confirmar",
                "¿Eliminar materiales de la fila " + fila + "?",
                ui.ButtonSet.YES_NO,
            ) !== ui.Button.YES
        )
            return;
        const cols = getColMap(
            sheet,
            HEADERS_PRODUCTOS,
            CONFIG.FILA_ENCABEZADOS_PRODUCTOS,
        );
        sheet.getRange(fila, cols.MATERIALES).setValue("");
        sheet.getRange(fila, cols.COSTO_MAT).setValue(0);
    } catch (e) {
        SpreadsheetApp.getUi().alert("Error: " + e.message);
    }
}

function configurarHojaIDs() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("ConfiguracionIDs");
    if (!sheet) {
        sheet = ss.insertSheet("ConfiguracionIDs");
        sheet.hideSheet();
        sheet.appendRow([
            "Tipo (Hoja)",
            "Categoria",
            "IntervaloInicio",
            "IntervaloFin",
        ]);
    }
}

function obtenerMapeoCategorias(hojaNombre) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(hojaNombre);
    const data = sheet.getDataRange().getValues();
    const isProd = hojaNombre === "Productos";
    const headerRowIdx = isProd
        ? CONFIG.FILA_ENCABEZADOS_PRODUCTOS - 1
        : CONFIG.FILA_ENCABEZADOS_MATERIALES - 1;
    const headers = data[headerRowIdx];
    const catCol = headers.findIndex(
        (h) =>
            h.toString().trim().toUpperCase() ===
            (isProd ? "CATEGORIA" : "CATEGORIA"),
    );

    const categorias = new Set();
    const startRow = isProd
        ? CONFIG.FILA_INICIO_DATOS_PRODUCTOS - 1
        : CONFIG.FILA_INICIO_DATOS_MATERIALES - 1;
    for (let i = startRow; i < data.length; i++) {
        if (data[i][catCol])
            categorias.add(data[i][catCol].toString().trim().toUpperCase());
    }
    return Array.from(categorias).sort();
}

function inicializarConfiguracionIDs() {
    configurarHojaIDs();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetConfig = ss.getSheetByName("ConfiguracionIDs");
    sheetConfig.getRange(2, 1, sheetConfig.getLastRow(), 4).clearContent();

    const hojas = ["Productos", "Materiales"];
    let nextStart = 1;
    const intervalo = 2000;

    hojas.forEach((h) => {
        const cats = obtenerMapeoCategorias(h);
        cats.forEach((cat) => {
            sheetConfig.appendRow([
                h,
                cat,
                nextStart,
                nextStart + intervalo - 1,
            ]);
            nextStart += intervalo;
        });
    });
}

function generarIdUnico(hojaNombre, categoria) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetConfig = ss.getSheetByName("ConfiguracionIDs");
    const configData = sheetConfig.getDataRange().getValues();
    const catUpper = categoria.toString().trim().toUpperCase();

    const config = configData.find(
        (row) =>
            row[0] === hojaNombre &&
            row[1].toString().toUpperCase() === catUpper,
    );
    if (!config) throw new Error("Categoría no configurada: " + categoria);

    const [inicio, fin] = [config[2], config[3]];
    const sheet = ss.getSheetByName(hojaNombre);
    const data = sheet.getDataRange().getValues();
    const isProd = hojaNombre === "Productos";
    const headers =
        data[
            isProd
                ? CONFIG.FILA_ENCABEZADOS_PRODUCTOS - 1
                : CONFIG.FILA_ENCABEZADOS_MATERIALES - 1
        ];
    const idCol = headers.indexOf("ID");
    const catCol = headers.findIndex(
        (h) =>
            h.toString().trim().toUpperCase() ===
            (isProd ? "CATEGORIA" : "CATEGORIA"),
    );

    let maxId = inicio - 1;
    for (
        let i = isProd
            ? CONFIG.FILA_INICIO_DATOS_PRODUCTOS - 1
            : CONFIG.FILA_INICIO_DATOS_MATERIALES - 1;
        i < data.length;
        i++
    ) {
        if (
            data[i][catCol]?.toString().trim().toUpperCase() === catUpper &&
            typeof data[i][idCol] === "number"
        ) {
            if (data[i][idCol] > maxId) maxId = data[i][idCol];
        }
    }
    return maxId + 1;
}

function migrarIdsExistentes() {
    try {
        inicializarConfiguracionIDs();
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const hojas = ["Productos", "Materiales"];

        hojas.forEach((hojaNombre) => {
            const sheet = ss.getSheetByName(hojaNombre);
            const data = sheet.getDataRange().getValues();
            const isProd = hojaNombre === "Productos";
            const startRowIdx = isProd
                ? CONFIG.FILA_INICIO_DATOS_PRODUCTOS - 1
                : CONFIG.FILA_INICIO_DATOS_MATERIALES - 1;
            const headers =
                data[
                    isProd
                        ? CONFIG.FILA_ENCABEZADOS_PRODUCTOS - 1
                        : CONFIG.FILA_ENCABEZADOS_MATERIALES - 1
                ];
            const idCol = headers.indexOf("ID");
            const catCol = headers.findIndex(
                (h) =>
                    h.toString().trim().toUpperCase() ===
                    (isProd ? "CATEGORIA" : "CATEGORIA"),
            );

            for (let i = startRowIdx; i < data.length; i++) {
                if (
                    data[i][catCol] &&
                    (typeof data[i][idCol] !== "number" || data[i][idCol] === 0)
                ) {
                    const newId = generarIdUnico(hojaNombre, data[i][catCol]);
                    sheet.getRange(i + 1, idCol + 1).setValue(newId);
                }
            }
        });
        SpreadsheetApp.getUi().alert("✅ Migración de IDs completada.");
    } catch (e) {
        SpreadsheetApp.getUi().alert("❌ Error en migración: " + e.message);
    }
}

function modificarTipoCambio() {
    const ui = SpreadsheetApp.getUi();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hojaCostos = ss.getSheetByName(CONFIG.HOJA_PRODUCTOS);
    const hojaMat = ss.getSheetByName(CONFIG.HOJA_MATERIALES);
    if (!hojaCostos || !hojaMat)
        return ui.alert(
            "❌ Error",
            "No se encontraron las hojas.",
            ui.ButtonSet.OK,
        );

    const tcActual = hojaCostos.getRange("G1").getValue();
    const resp = ui.prompt(
        "💵 Modificar Tipo de Cambio",
        `TC Actual: $${tcActual}\n\nIngresa el nuevo TC:`,
        ui.ButtonSet.OK_CANCEL,
    );
    if (resp.getSelectedButton() !== ui.Button.OK) return;
    const nuevoTc = parseFloat(resp.getResponseText().replace(",", ".").trim());
    if (isNaN(nuevoTc) || nuevoTc <= 0)
        return ui.alert(
            "❌ Error",
            "Ingresa un número válido.",
            ui.ButtonSet.OK,
        );

    hojaCostos.getRange("G1").setValue(nuevoTc);
    // CORRECCIÓN: Usar hojaMat y HEADERS_MATERIALES
    const colsMat = getColMap(
        hojaMat,
        HEADERS_MATERIALES,
        CONFIG.FILA_ENCABEZADOS_MATERIALES,
    );
    const datos = hojaMat.getDataRange().getValues();
    const rango = hojaMat.getRange(
        CONFIG.FILA_INICIO_DATOS_MATERIALES,
        colsMat.PU_USD,
        datos.length - (CONFIG.FILA_INICIO_DATOS_MATERIALES - 1),
        2,
    );
    const precios = rango.getValues();
    let cnt = 0;
    precios.forEach((r, i) => {
        const usd = parseFloat(r[0]) || 0,
            ars = parseFloat(r[1]) || 0;
        if (usd > 0) {
            precios[i][1] = usd * nuevoTc;
            cnt++;
        } else if (ars > 0) {
            precios[i][0] = ars / nuevoTc;
            cnt++;
        }
    });
    rango.setValues(precios);
    ui.alert(
        "✅ Éxito",
        `TC → $${nuevoTc}. ${cnt} materiales recalculados.`,
        ui.ButtonSet.OK,
    );
}

function mostrarMenuAumentos() {
    SpreadsheetApp.getUi().showModalDialog(
        HtmlService.createTemplateFromFile("FormularioAumentos")
            .evaluate()
            .setWidth(450)
            .setHeight(420),
        "📈 Aumento Masivo de Materiales",
    );
}

function obtenerDatosParaAumentos() {
    const hojaMat = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
        CONFIG.HOJA_MATERIALES,
    );
    if (!hojaMat)
        throw new Error(
            'No se encontró la hoja "' + CONFIG.HOJA_MATERIALES + '".',
        );
    const colsMat = getColMapSafe(
        hojaMat,
        HEADERS_MATERIALES,
        CONFIG.FILA_ENCABEZADOS_MATERIALES,
    );
    const datos = hojaMat
        .getDataRange()
        .getValues()
        .slice(CONFIG.FILA_ENCABEZADOS_MATERIALES);
    const proveedores = new Set(),
        categorias = new Set(),
        relacionPC = {},
        relacionCP = {};
    datos.forEach((r) => {
        const cat =
            colsMat.CATEGORIA > 0
                ? r[colsMat.CATEGORIA - 1]?.toString().trim() || ""
                : "";
        const prov =
            colsMat.PROVEEDOR > 0
                ? r[colsMat.PROVEEDOR - 1]?.toString().trim() || ""
                : "";
        if (prov) proveedores.add(prov);
        if (cat) categorias.add(cat);
        if (prov && cat) {
            (relacionPC[prov] = relacionPC[prov] || new Set()).add(cat);
            (relacionCP[cat] = relacionCP[cat] || new Set()).add(prov);
        }
    });
    const toArr = (obj) =>
        Object.fromEntries(
            Object.entries(obj).map(([k, v]) => [k, Array.from(v).sort()]),
        );
    return {
        proveedores: Array.from(proveedores).sort(),
        categorias: Array.from(categorias).sort(),
        relacionPC: toArr(relacionPC),
        relacionCP: toArr(relacionCP),
    };
}

function aplicarAumentoMasivo(criterios) {
    const { proveedor, categoria, porcentaje } = criterios;
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hojaMat = ss.getSheetByName(CONFIG.HOJA_MATERIALES);
    const hojaCostos = ss.getSheetByName(CONFIG.HOJA_PRODUCTOS);
    if (!hojaMat)
        throw new Error(
            'No se encontró la hoja "' + CONFIG.HOJA_MATERIALES + '".',
        );
    const tc = hojaCostos
        ? parseFloat(hojaCostos.getRange("G1").getValue()) || 1
        : 1;
    const colsMat = getColMap(
        hojaMat,
        HEADERS_MATERIALES,
        CONFIG.FILA_ENCABEZADOS_MATERIALES,
    );
    const datos = hojaMat.getDataRange().getValues();
    const filaInicio = CONFIG.FILA_INICIO_DATOS_MATERIALES;
    const mult = 1 + porcentaje / 100;
    const numFilasDatos = datos.length - (filaInicio - 1);
    if (numFilasDatos <= 0) return { success: true, actualizados: 0 };
    const rango = hojaMat.getRange(
        filaInicio,
        colsMat.PU_USD,
        numFilasDatos,
        2,
    );
    const precios = rango.getValues();
    let cnt = 0;
    datos.slice(filaInicio - 1).forEach((r, i) => {
        const cat =
            colsMat.CATEGORIA > 0
                ? r[colsMat.CATEGORIA - 1]?.toString().trim() || ""
                : "";
        const prov =
            colsMat.PROVEEDOR > 0
                ? r[colsMat.PROVEEDOR - 1]?.toString().trim() || ""
                : "";
        if (
            (proveedor === "TODOS" || proveedor === prov) &&
            (categoria === "TODAS" || categoria === cat)
        ) {
            const usd = parseFloat(precios[i][0]) || 0,
                ars = parseFloat(precios[i][1]) || 0;
            if (usd > 0) {
                precios[i][0] = usd * mult;
                precios[i][1] = precios[i][0] * tc;
            } else if (ars > 0) {
                precios[i][1] = ars * mult;
                precios[i][0] = precios[i][1] / tc;
            }
            cnt++;
        }
    });
    rango.setValues(precios);
    return { success: true, actualizados: cnt };
}

