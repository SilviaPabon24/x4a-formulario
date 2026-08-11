/**
 * BACKEND — X4A Formulario de Llaves
 * ============================================================
 * Este script convierte una Google Sheet en la "base de datos"
 * compartida del formulario, evitando que varias personas se
 * pisen los cambios como pasaba en el Excel.
 *
 * INSTALACIÓN (ver README.md para el paso a paso completo):
 * 1. Crea una Google Sheet nueva, llámala "X4A - Oportunidades".
 * 2. Extensiones → Apps Script. Borra el contenido de Código.gs
 *    y pega este archivo completo.
 * 3. Ajusta HOJA_NOMBRE si usaste otro nombre de pestaña.
 * 4. Implementar → Nueva implementación → Tipo: Aplicación web.
 *    - Ejecutar como: Yo
 *    - Quién tiene acceso: Cualquier usuario
 * 5. Copia la URL resultante (termina en /exec) y pégala en
 *    js/config.js como APPS_SCRIPT_URL.
 */

const HOJA_NOMBRE = "Oportunidades";

const COLUMNAS = [
  "ID",
  "Created On",
  "Government Number",
  "Owner",
  "Est Close Date",
  "Expiry Date",
  "Probability",
  "AccountID",
  "End User Name",
  "Subject",
  "Source Campaign",
  "Est Revenue",
  "Budget Amount",
  "Currency",
  "Current Situation",
  "IM Calendar Month",
  "IM Calendar Quarter",
  "IM Calendar Year",
  "Created By",
  "Primary Email (Created By)",
  "Country",
  "Source of originating lead",
  "Is Renewal Order",
  "Notificado",
  "Fecha Notificación",
];

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
    return jsonOut_({ ok: true, rows: listRows_() });
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
  if (body.action === "notify") {
    return jsonOut_(notificar_(body.ids));
  }
  return jsonOut_({ ok: false, error: "Acción no reconocida" });
}

function crearFila_(row) {
  const sheet = getSheet_();
  const id = "OPP-" + new Date().getTime();
  const createdOn = new Date();

  const values = COLUMNAS.map((col) => {
    if (col === "ID") return id;
    if (col === "Created On") return createdOn;
    if (col === "Notificado") return "No";
    if (col === "Fecha Notificación") return "";
    return row[col] !== undefined ? row[col] : "";
  });

  sheet.appendRow(values);
  return { ok: true, id: id };
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
      if (val instanceof Date) val = Utilities.formatDate(val, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
      obj[h] = val;
    });
    rows.push(obj);
  }
  // Más recientes primero
  return rows.reverse();
}

function notificar_(ids) {
  const sheet = getSheet_();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idCol = headers.indexOf("ID");
  const emailCol = headers.indexOf("Primary Email (Created By)");
  const ownerCol = headers.indexOf("Owner");
  const subjectCol = headers.indexOf("Subject");
  const govCol = headers.indexOf("Government Number");
  const notifCol = headers.indexOf("Notificado");
  const notifFechaCol = headers.indexOf("Fecha Notificación");

  let sent = 0;
  for (let i = 1; i < data.length; i++) {
    const rowId = data[i][idCol];
    if (ids.indexOf(String(rowId)) === -1) continue;

    const email = data[i][emailCol];
    const owner = data[i][ownerCol];
    const subject = data[i][subjectCol];
    const gobierno = data[i][govCol];

    if (email) {
      MailApp.sendEmail({
        to: email,
        subject: "Tu oportunidad ya fue cargada en X4A",
        body:
          `Hola ${owner || ""},\n\n` +
          `Confirmamos que la siguiente oportunidad ya fue cargada en el CRM X4A:\n\n` +
          `TEMA: ${subject}\n` +
          `Gobierno: ${gobierno}\n\n` +
          `No es necesario volver a registrarla.\n\n` +
          `Este es un mensaje automático del Formulario de Llaves X4A.`,
      });
      sent++;
    }

    sheet.getRange(i + 1, notifCol + 1).setValue("Si");
    sheet.getRange(i + 1, notifFechaCol + 1).setValue(new Date());
  }

  return { ok: true, sent: sent };
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(
    ContentService.MimeType.JSON
  );
}
