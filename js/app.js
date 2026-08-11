// ============================================================
// X4A — Formulario de Llaves (réplica del Excel "Llaves_para_X4A")
// ============================================================

let EXPORT_TABLE = [];
let CANALES = [];
const SITUACIONES = [
  "Identificando necesidad",
  "Cotización enviada",
  "Aprobación de precios en curso",
  "Orden de compra esperada",
  "En negociación",
  "Ganada",
  "Perdida",
];
const ORIGENES = ["IBM", "RED HAT", "Traditional"];

// ---------- Utilidades ----------
const $ = (id) => document.getElementById(id);
const norm = (s) => (s || "").toString().trim();
const fmtMoney = (n) =>
  isFinite(n)
    ? n.toLocaleString("en-US", { style: "currency", currency: "USD" })
    : "";

function todayISO() {
  return new Date().toISOString().slice(0, 10);
}

// ---------- Carga de datos de referencia ----------
async function loadReferenceData() {
  const [exportRes, canalesRes] = await Promise.all([
    fetch("data/export.json"),
    fetch("data/canales.json"),
  ]);
  EXPORT_TABLE = await exportRes.json();
  CANALES = await canalesRes.json();

  // Índice por nombre normalizado para el lookup de BCN (equivalente a XLOOKUP)
  CANALES_INDEX = {};
  CANALES.forEach((c) => {
    const key = norm(c.nombre).toUpperCase();
    if (!CANALES_INDEX[key]) CANALES_INDEX[key] = [];
    CANALES_INDEX[key].push(c);
  });
}
let CANALES_INDEX = {};

// ---------- Sección 1: Canal / Cliente ----------
function setupCanalAutocomplete() {
  const input = $("nombreCanal");
  const list = $("canalDatalist");
  list.innerHTML = "";
  const seen = new Set();
  CANALES.slice(0, 8048).forEach((c) => {
    if (seen.has(c.nombre)) return;
    seen.add(c.nombre);
    const opt = document.createElement("option");
    opt.value = c.nombre;
    list.appendChild(opt);
  });

  input.addEventListener("input", () => {
    recalcBCN();
    if (!$("clienteFinal").dataset.userEdited) {
      $("clienteFinal").value = input.value;
    }
    validateAll();
  });
}

function recalcBCN() {
  const nombre = norm($("nombreCanal").value).toUpperCase();
  const bcnField = $("bcn");
  const bcnWarning = $("bcnWarning");
  if (!nombre) {
    bcnField.value = "";
    bcnWarning.textContent = "";
    return;
  }
  const matches = CANALES_INDEX[nombre];
  if (!matches || matches.length === 0) {
    bcnField.value = "Canal no encontrado en la hoja Canales";
    bcnField.classList.add("error-field");
    bcnWarning.textContent = "";
    return;
  }
  bcnField.classList.remove("error-field");
  bcnField.value = matches[0].bcn;
  if (matches.length > 1) {
    bcnWarning.textContent = `OJO: hay ${matches.length} registros con este nombre; se tomó el primer BCN. Verificar.`;
  } else {
    bcnWarning.textContent = "";
  }
}

// ---------- Sección 2: Llave de creación ----------
function populateOrigenSelect() {
  const sel = $("origen");
  sel.innerHTML = '<option value="">Selecciona…</option>';
  ORIGENES.forEach((o) => {
    const opt = document.createElement("option");
    opt.value = o;
    opt.textContent = o;
    sel.appendChild(opt);
  });
}

function populateGobiernoSelect() {
  const origen = $("origen").value;
  const sel = $("gobierno");
  sel.innerHTML = '<option value="">Selecciona…</option>';
  sel.disabled = !origen;
  if (!origen) return;
  EXPORT_TABLE.filter((r) => r.origen === origen).forEach((r) => {
    const opt = document.createElement("option");
    opt.value = r.gobierno;
    opt.textContent = r.gobierno;
    sel.appendChild(opt);
  });
}

function populateCampanaSelect() {
  const gobierno = $("gobierno").value;
  const sel = $("campana");
  sel.innerHTML = '<option value="">Selecciona…</option>';
  sel.disabled = !gobierno;
  if (!gobierno) return;
  const row = EXPORT_TABLE.find((r) => r.gobierno === gobierno);
  if (!row) return;
  row.campanas.forEach((c) => {
    const opt = document.createElement("option");
    opt.value = c;
    opt.textContent = c;
    sel.appendChild(opt);
  });
}

function recalcTema() {
  const gobierno = $("gobierno").value;
  const row = EXPORT_TABLE.find((r) => r.gobierno === gobierno);
  const temaPrefix = row ? row.tema : "";
  $("temaPrefix").value = temaPrefix;

  const detalle = norm($("detalle").value);
  let temaFinal = "";
  if (!temaPrefix) {
    temaFinal = "";
  } else if (!detalle) {
    temaFinal = `Falta el detalle de la oportunidad — el TEMA debe iniciar por: ${temaPrefix}`;
  } else {
    temaFinal = `${temaPrefix}_${detalle.replace(/\s+/g, "_")}`;
  }
  $("temaFinal").value = temaFinal;
  return temaFinal;
}

function validateLlave() {
  const origen = $("origen").value;
  const gobierno = $("gobierno").value;
  const campana = $("campana").value;
  const box = $("validacionLlave");
  const ambigBox = $("avisoAmbiguedad");

  if (!origen || !gobierno || !campana) {
    box.textContent = "Faltan campos de la llave";
    box.className = "status-box status-warn";
    ambigBox.textContent = "";
    return false;
  }

  const row = EXPORT_TABLE.find(
    (r) => r.origen === origen && r.gobierno === gobierno
  );
  if (!row) {
    box.textContent = "ERROR: el Gobierno no corresponde al Origen";
    box.className = "status-box status-error";
    ambigBox.textContent = "";
    return false;
  }
  if (!row.campanas.includes(campana)) {
    box.textContent = "ERROR: la Campaña no corresponde al Gobierno";
    box.className = "status-box status-error";
    ambigBox.textContent = "";
    return false;
  }
  box.textContent = "OK: llave válida";
  box.className = "status-box status-ok";

  // Aviso de ambigüedad: ¿la campaña aparece en más de un Gobierno?
  const countGobiernosConCampana = EXPORT_TABLE.filter((r) =>
    r.campanas.includes(campana)
  ).length;
  if (countGobiernosConCampana > 1) {
    ambigBox.textContent =
      "ATENCIÓN: esta campaña existe en más de un Gobierno — confirmar si es venta IBM directa o Traditional (Cloud).";
    ambigBox.className = "status-box status-warn";
  } else {
    ambigBox.textContent = "Sin ambigüedad";
    ambigBox.className = "status-box status-ok";
  }
  return true;
}

// ---------- Sección 3: Datos comerciales ----------
function recalcIngresoEstimado() {
  const valor = parseFloat($("valorCotizado").value);
  const trmDia = parseFloat($("trmDia").value);
  const trmBase = parseFloat($("trmBase").value);
  const out = $("ingresoEstimado");
  if (!valor || !trmDia || !trmBase) {
    out.value = "";
    return;
  }
  out.value = fmtMoney((valor * trmDia) / trmBase);
}

function populateSituacionSelect() {
  const sel = $("situacion");
  sel.innerHTML = '<option value="">Selecciona…</option>';
  SITUACIONES.forEach((s) => {
    const opt = document.createElement("option");
    opt.value = s;
    opt.textContent = s;
    sel.appendChild(opt);
  });
}

// ---------- Validación global + habilitar guardar ----------
function validateAll() {
  const temaFinal = recalcTema();
  const llaveOk = validateLlave();
  recalcIngresoEstimado();

  const temaOk = temaFinal && !temaFinal.startsWith("Falta el detalle");
  const bcnOk =
    $("bcn").value && !$("bcn").value.startsWith("Canal no encontrado");

  const allOk =
    llaveOk &&
    temaOk &&
    bcnOk &&
    norm($("fechaCierreIngram").value) &&
    norm($("valorCotizado").value) &&
    norm($("situacion").value) &&
    norm($("creadoPor").value) &&
    norm($("creadoPorEmail").value);

  $("btnGuardar").disabled = !allOk;
  return allOk;
}

// ---------- Guardar oportunidad ----------
async function guardarOportunidad() {
  if (!validateAll()) return;

  const btn = $("btnGuardar");
  btn.disabled = true;
  btn.textContent = "Guardando…";

  const fechaCierre = $("fechaCierreIngram").value; // YYYY-MM-DD
  const [y, m] = fechaCierre.split("-");
  const quarter = Math.ceil(parseInt(m, 10) / 3);

  const payload = {
    action: "create",
    row: {
      "Government Number": $("gobierno").value,
      Owner: $("creadoPor").value,
      "Est Close Date": fechaCierre,
      "Expiry Date": $("fechaCierreFabricante").value || fechaCierre,
      Probability: $("probabilidad").value,
      AccountID: $("bcn").value,
      "End User Name": $("clienteFinal").value,
      Subject: $("temaFinal").value,
      "Source Campaign": $("campana").value,
      "Est Revenue": (
        parseFloat($("valorCotizado").value) *
        parseFloat($("trmDia").value) /
        parseFloat($("trmBase").value)
      ).toFixed(2),
      "Budget Amount": $("compraEstimada").value,
      Currency: "US Dollar",
      "Current Situation": $("situacion").value,
      "IM Calendar Month": m,
      "IM Calendar Quarter": "Q" + quarter,
      "IM Calendar Year": y,
      "Created By": $("creadoPor").value,
      "Primary Email (Created By)": $("creadoPorEmail").value,
      Country: "Colombia",
      "Source of originating lead": $("origen").value,
      "Is Renewal Order": $("gobierno").value.toLowerCase().includes("renov") ||
        $("gobierno").value.toLowerCase().includes("renew")
        ? "Yes"
        : "No",
    },
  };

  try {
    const res = await postToBackend(payload);
    if (res && res.ok) {
      showToast("Oportunidad guardada en X4A ✔");
      resetForm();
      loadDashboard();
    } else {
      showToast(
        "No se pudo guardar: " + (res && res.error ? res.error : "error desconocido"),
        true
      );
    }
  } catch (err) {
    showToast(
      "El backend no está configurado todavía (ver README.md) o falló la conexión.",
      true
    );
    console.error(err);
  } finally {
    btn.disabled = false;
    btn.textContent = "Guardar oportunidad en X4A";
  }
}

async function postToBackend(payload) {
  if (!APPS_SCRIPT_URL || APPS_SCRIPT_URL.includes("PEGA_AQUI")) {
    throw new Error("APPS_SCRIPT_URL no configurada");
  }
  const res = await fetch(APPS_SCRIPT_URL, {
    method: "POST",
    // text/plain evita el preflight CORS con Apps Script
    headers: { "Content-Type": "text/plain;charset=utf-8" },
    body: JSON.stringify(payload),
  });
  return res.json();
}

function resetForm() {
  [
    "nombreCanal",
    "detalle",
    "valorCotizado",
    "compraEstimada",
    "probabilidad",
  ].forEach((id) => ($(id).value = ""));
  $("clienteFinal").value = "";
  delete $("clienteFinal").dataset.userEdited;
  $("bcn").value = "";
  $("origen").value = "";
  populateGobiernoSelect();
  populateCampanaSelect();
  $("temaPrefix").value = "";
  $("temaFinal").value = "";
  $("situacion").value = "";
  $("validacionLlave").textContent = "";
  $("validacionLlave").className = "status-box";
  $("avisoAmbiguedad").textContent = "";
  $("ingresoEstimado").value = "";
  validateAll();
}

function showToast(msg, isError = false) {
  const t = $("toast");
  t.textContent = msg;
  t.className = "toast show" + (isError ? " toast-error" : "");
  setTimeout(() => (t.className = "toast"), 4000);
}

// ---------- Dashboard ----------
async function loadDashboard() {
  const tbody = $("dashboardBody");
  tbody.innerHTML = `<tr><td colspan="8" class="muted">Cargando…</td></tr>`;
  try {
    const url = `${APPS_SCRIPT_URL}?action=list`;
    const res = await fetch(url);
    const data = await res.json();
    renderDashboard(data.rows || []);
  } catch (err) {
    tbody.innerHTML = `<tr><td colspan="8" class="muted">No se pudo cargar el dashboard (configura el backend en js/config.js — ver README.md).</td></tr>`;
  }
}

function renderDashboard(rows) {
  const tbody = $("dashboardBody");
  if (!rows.length) {
    tbody.innerHTML = `<tr><td colspan="8" class="muted">Aún no hay oportunidades cargadas.</td></tr>`;
    return;
  }
  tbody.innerHTML = "";
  rows.forEach((r) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td><input type="checkbox" class="rowCheck" data-id="${r.ID}" ${r.Notificado === "Si" ? "disabled" : ""}></td>
      <td>${r["End User Name"] || ""}</td>
      <td>${r["Subject"] || ""}</td>
      <td>${r["Government Number"] || ""}</td>
      <td>${r["Est Revenue"] ? fmtMoney(parseFloat(r["Est Revenue"])) : ""}</td>
      <td>${r["Created By"] || ""}</td>
      <td>${r["Created On"] || ""}</td>
      <td>${r.Notificado === "Si" ? '<span class="badge badge-ok">Notificado</span>' : '<span class="badge">Pendiente</span>'}</td>
    `;
    tbody.appendChild(tr);
  });
}

async function notificarSeleccionados() {
  const checks = Array.from(document.querySelectorAll(".rowCheck:checked"));
  if (!checks.length) {
    showToast("Selecciona al menos una oportunidad para notificar.", true);
    return;
  }
  const ids = checks.map((c) => c.dataset.id);
  const btn = $("btnNotificar");
  btn.disabled = true;
  btn.textContent = "Enviando…";
  try {
    const res = await postToBackend({ action: "notify", ids });
    if (res && res.ok) {
      showToast(`Correo enviado a ${res.sent} persona(s).`);
      loadDashboard();
    } else {
      showToast("No se pudo enviar: " + (res.error || "error"), true);
    }
  } catch (err) {
    showToast("Falló el envío de correos.", true);
  } finally {
    btn.disabled = false;
    btn.textContent = "Notificar seleccionados";
  }
}

// ---------- Navegación por pestañas ----------
function setupTabs() {
  document.querySelectorAll(".tab-btn").forEach((btn) => {
    btn.addEventListener("click", () => {
      document
        .querySelectorAll(".tab-btn")
        .forEach((b) => b.classList.remove("active"));
      document
        .querySelectorAll(".tab-panel")
        .forEach((p) => p.classList.remove("active"));
      btn.classList.add("active");
      $(btn.dataset.tab).classList.add("active");
      if (btn.dataset.tab === "panel-dashboard") loadDashboard();
    });
  });
}

// ---------- Inicialización ----------
async function init() {
  await loadReferenceData();
  populateOrigenSelect();
  populateSituacionSelect();
  setupCanalAutocomplete();
  setupTabs();

  $("origen").addEventListener("change", () => {
    populateGobiernoSelect();
    populateCampanaSelect();
    recalcTema();
    validateAll();
  });
  $("gobierno").addEventListener("change", () => {
    populateCampanaSelect();
    recalcTema();
    validateAll();
  });
  $("campana").addEventListener("change", validateAll);
  $("detalle").addEventListener("input", validateAll);
  $("clienteFinal").addEventListener("input", (e) => {
    e.target.dataset.userEdited = "1";
  });
  ["valorCotizado", "trmDia", "trmBase", "fechaCierreIngram", "fechaCierreFabricante", "compraEstimada", "probabilidad", "situacion", "creadoPor", "creadoPorEmail"].forEach(
    (id) => $(id).addEventListener("input", validateAll)
  );

  $("btnGuardar").addEventListener("click", guardarOportunidad);
  $("btnNotificar").addEventListener("click", notificarSeleccionados);
  $("btnRefrescarDashboard").addEventListener("click", loadDashboard);

  $("trmDia").value = "";
  $("trmBase").value = "3708";
  $("fechaCierreIngram").value = todayISO();

  validateAll();
}

document.addEventListener("DOMContentLoaded", init);
