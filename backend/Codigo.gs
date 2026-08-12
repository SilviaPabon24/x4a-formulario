/**
 * BACKEND — X4A Formulario de Llaves
 * ============================================================
 * Este script convierte una Google Sheet en la "base de datos"
 * compartida del formulario, evitando que varias personas se
 * pisen los cambios como pasaba en el Excel.
 *
 * Las primeras 36 columnas (COLUMNAS_CRM) son EXACTAMENTE las
 * del export de X4A, en el mismo orden — son las que se ven en
 * la ventana "Ver tabla para copiar" y se pegan tal cual en X4A.
 *
 * "ID Interno" es la única columna de control, solo para tener
 * un identificador único por fila internamente; nunca aparece
 * en la tabla que se copia hacia X4A.
 *
 * INSTALACIÓN (ver README.md para el paso a paso completo):
 * 1. Crea una Google Sheet nueva, llámala "X4A - Oportunidades".
 * 2. Extensiones → Apps Script. Borra el contenido de Código.gs
 *    y pega este archivo completo.
 * 3. Implementar → Nueva implementación → Tipo: Aplicación web.
 *    - Ejecutar como: Yo
 *    - Quién tiene acceso: Cualquier usuario
 * 4. Copia la URL resultante (termina en /exec) y pégala en
 *    js/config.js como APPS_SCRIPT_URL.
 *
 * NOTA: si ya habías desplegado una versión anterior con otras
 * columnas, RENOMBRA la pestaña "Oportunidades" vieja (para no
 * perder los datos) antes de volver a probar, así se crea una
 * pestaña nueva con el esquema correcto.
 */

const HOJA_NOMBRE = "Oportunidades";

// Valores fijos que siempre van en el export a X4A.
const OWNER_FIJO = "Joel Ayala";
const CREATED_BY_FIJO = "Joel Ayala";
const EMAIL_FIJO = "Joel.Ayala@ingrammicro.com";

// Columnas EXACTAS del export del CRM X4A, en este orden.
const COLUMNAS_CRM = [
  "Government Number",
  "Owner",
  "Est. Close Date",
  "Expiry Date",
  "Created On",
  "Probability",
  "Stage",
  "AccountID",
  "End User Name",
  "Subject",
  "Source Campaign",
  "Est. Revenue",
  "Budget Amount",
  "Currency",
  "Current Situation",
  "Proposed Solution",
  "IM Calendar Month",
  "IM Calendar Quarter",
  "IM Calendar Week",
  "IM Calendar Year",
  "ID",
  "Created By",
  "Primary Email (Created By) (User)",
  "Contact",
  "Sales Stage2",
  "Status Reason (PTS)",
  "Status Reason",
  "Status",
  "Customer Need",
  "Country",
  "BCN",
  "Source of originating lead?",
  "Is Renewal Order",
  "Identifier Type",
  "Unique Identifier",
  "End User MST.",
];

// Única columna de control — SOLO uso interno, nunca se copia a X4A.
const COLUMNAS_CONTROL = ["ID Interno"];

const COLUMNAS = COLUMNAS_CRM.concat(COLUMNAS_CONTROL);

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(HOJA_NOMBRE);
  if (!sheet) {
    sheet = ss.insertSheet(HOJA_NOMBRE);
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(COLUMNAS);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function doGet(e) {
  const action = e.parameter.action;
  if (action === "list") {
    // "columns" son solo las de X4A — así la tabla para copiar
    // nunca incluye la columna interna de control.
    return jsonOut_({ ok: true, columns: COLUMNAS_CRM, rows: listRows_() });
  }
  return jsonOut_({ ok: false, error: "Acción no reconocida" });
}

function doPost(e) {
  let body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonOut_({ ok: false, error: "JSON inválido" });
  }

  if (body.action === "create") {
    return jsonOut_(crearFila_(body.row));
  }
  return jsonOut_({ ok: false, error: "Acción no reconocida" });
}

function crearFila_(row) {
  const sheet = getSheet_();
  const idInterno = "OPP-" + new Date().getTime();

  const computed = {
    Owner: OWNER_FIJO,
    "Created By": CREATED_BY_FIJO,
    "Primary Email (Created By) (User)": EMAIL_FIJO,
    Currency: "US Dollar",
    Country: "Colombia",
    "Status Reason": "In Progress",
    Status: "Open",
    "ID Interno": idInterno,
  };
  // Nota: "ID" (columna de X4A) queda vacía a propósito — la
  // asigna X4A al importar. "Expiry Date", "Stage", "AccountID",
  // "Proposed Solution", "IM Calendar Month/Quarter/Week/Year",
  // "Contact", "Sales Stage2", "Status Reason (PTS)",
  // "Customer Need", "BCN", "Identifier Type", "Unique Identifier"
  // y "End User MST." también quedan vacías porque nadie las
  // envía y no están en "computed".

  const values = COLUMNAS.map((col) => {
    if (computed[col] !== undefined) return computed[col];
    return row[col] !== undefined ? row[col] : "";
  });

  sheet.appendRow(values);
  return { ok: true, id: idInterno };
}

function listRows_() {
  const sheet = getSheet_();
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    const obj = {};
    headers.forEach((h, idx) => {
      let val = data[i][idx];
      if (val instanceof Date) {
        val = Utilities.formatDate(val, Session.getScriptTimeZone(), "dd/MM/yyyy");
      }
      obj[h] = val;
    });
    rows.push(obj);
  }
  return rows.reverse();
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(
    ContentService.MimeType.JSON
  );
}
