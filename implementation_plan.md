# Implementation Plan - Spreadsheet Connection Fixes & Visual Polish of Forms

This plan addresses:
1. Fixing connection issues between the **Pedidos** spreadsheet and the **Costos** spreadsheet.
2. Polishing and unifying the visual style of all four HTML forms (`Formulario.html`, `FormularioAumentos.html`, `FormularioPedidos.html`, `FormularioEntregaParcial.html`) to deliver a modern, premium, and responsive user experience.
3. Fixing hidden reference and configuration bugs in `costos.gs` that break the "Aumento Masivo de Materiales" feature.

---

## User Review Required

> [!IMPORTANT]
> **Script Authorization Required**: 
> Because the scripts interact between two separate spreadsheets via ID, Google Apps Script requires explicit authorization. When deploying the updated code, please open the Apps Script editor for both spreadsheets and run a function (e.g., `onOpen` or `instalarTodoElSistema`) manually once. This will trigger the Google OAuth dialog to authorize cross-spreadsheet permissions.

> [!TIP]
> **Sheet Renaming Fallback**:
> We have added a dynamic sheet fallback. If your spreadsheet tab in the "Costos" file is still named `"Costos"` instead of `"Productos"`, the script will automatically detect and query it without throwing errors.

---

## Proposed Changes

### 1. Connection & Core Backend Fixes

#### [MODIFY] [pedidos.gs](file:///home/blank/Desktop/excel/Pedidos/pedidos.gs)
* **Robust Sheet Fallback**: In `obtenerDataInicialPedidos` and `descontarStockDePedido`, support query of both `"Productos"` and `"Costos"` sheets.
* **Auto-detect Header Row**: Replace the hardcoded `CONFIG.FILA_HEADERS` row number (3) for external sheet queries with a dynamic scan (looking for headers like `"ID"` and `"MODELO"` in the first 5 rows). This avoids mismatches when the Pedidos sheet is formatted differently than the Costos sheet.
* **Clients Sheet Fallback**: If the "Clientes" sheet is missing, query the sheet gracefully without throwing errors.

#### [MODIFY] [costos.gs](file:///home/blank/Desktop/excel/Costos/costos.gs)
* **Fix Reference Errors**: In `obtenerDatosParaAumentos` and `aplicarAumentoMasivo`, replace the undefined `sheet` variable with `hojaMat` and use `HEADERS_MATERIALES` rather than `HEADERS_COSTOS`.
* **Fix Configuration Bug**: Correct `CONFIG.FILA_ENCABEZADOS` references to `CONFIG.FILA_ENCABEZADOS_MATERIALES` or `CONFIG.FILA_INICIO_DATOS_MATERIALES` to prevent calculation crashes during mass increases.

---

### 2. Premium Design System for Forms
We will establish a unified modern design system using:
* **Typography**: Beautiful sans-serif modern look (Google Font `Outfit` or `Inter`).
* **Themes**: 
  * *Sales / Orders (`Pedidos`)*: Deep blue / violet gradients (`#2563eb` and `#7c3aed`), card glassmorphism with subtle borders, and smooth dark/slate totals box.
  * *Cost / Finance (`Costos`)*: Modern slate / emerald green theme (`#059669` and `#0f766e`) with gold accents for premium pricing visual highlights.
* **Layouts**: Fully responsive grid systems, rounded corners (`border-radius: 12px` to `16px`), elegant shadow transitions, custom hover effects on action buttons, and beautiful state badges (success, warnings, alerts).
* **Code Cleanliness**: Remove duplicate Javascript functions (like `actualizarPrecioInfo` in `Formulario.html`).

#### [MODIFY] [FormularioPedidos.html](file:///home/blank/Desktop/excel/Pedidos/FormularioPedidos.html)
* Restyle the order intake form with the premium deep-blue/violet gradient theme.
* Highlight important sections, style inputs, checkboxes, and buttons with custom focus rings and subtle hover translations.

#### [MODIFY] [FormularioEntregaParcial.html](file:///home/blank/Desktop/excel/Pedidos/FormularioEntregaParcial.html)
* Polish layout to match the master order intake styling, utilizing subtle indicators for pending/original items and styled total breakdowns.

#### [MODIFY] [Formulario.html](file:///home/blank/Desktop/excel/Costos/Formulario.html)
* Elevate the design of the product recipe form with the finance-focused emerald green design system.
* Use card styling for recipe items, clear preset-loader controls, and high-impact calculations displaying waste metrics.
* **Deduplicate Functions**: Remove the duplicate `actualizarPrecioInfo` definition which overrode correct calculations.

#### [MODIFY] [FormularioAumentos.html](file:///home/blank/Desktop/excel/Costos/FormularioAumentos.html)
* Apply a cohesive, polished UI card matching the green-emerald theme.
* Add transition effects for loading/form rendering states and style the percentage inputs with elegant clear layouts.

---

## Verification Plan

### Manual Verification
1. **Google Apps Script Verification**: 
   * Copy the updated [pedidos.gs](file:///home/blank/Desktop/excel/Pedidos/pedidos.gs) and [costos.gs](file:///home/blank/Desktop/excel/Costos/costos.gs) to their respective Google Sheets script editors.
   * Run `onOpen` manually to trigger permission screens.
   * Open the custom menus `⚙️ Gestión Integral > ➕ Cargar Nuevo Pedido` and check if the form loads immediately.
2. **Selectors & Connection Check**:
   * Confirm that the **Clients** selector (from local Sheet "Clientes"), the **Categories** filter, and the **Products** selector are fully populated.
   * Verify that **Tipo de Cambio (Ref)** displays the correct value from cell `G1` of the Costos sheet.
3. **Form Visual Inspection**:
   * Open all four dialog pages in Google Sheets. Check that colors, buttons, hover states, rounded corners, card styling, and tables scale cleanly and render in high premium quality.
