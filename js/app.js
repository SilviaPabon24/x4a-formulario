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
  const list = $("situacionDatalist");
  list.innerHTML = "";
  SITUACIONES.forEach((s) => {
    const opt = document.createElement("option");
    opt.value = s;
    list.appendChild(opt);
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
    norm($("clienteFinal").value) &&
    norm($("fechaCierreIngram").value) &&
    norm($("valorCotizado").value) &&
    norm($("situacion").value);

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

  const payload = {
    action: "create",
    row: {
      "Government Number": $("gobierno").value,
      "Est. Close Date": fechaCierre,
      Probability: $("probabilidad").value,
      "End User Name": $("clienteFinal").value,
      Subject: $("temaFinal").value,
      "Source Campaign": $("campana").value,
      "Est. Revenue": (
        (parseFloat($("valorCotizado").value) *
          parseFloat($("trmDia").value)) /
        parseFloat($("trmBase").value)
      ).toFixed(2),
      "Budget Amount": $("compraEstimada").value,
      "Current Situation": $("situacion").value,
      "Source of originating lead?": $("origen").value,
      "Is Renewal Order":
        $("gobierno").value.toLowerCase().includes("renov") ||
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
let CRM_COLUMNS = [];
let LAST_ROWS = [];

async function loadDashboard() {
  const tbody = $("dashboardBody");
  tbody.innerHTML = `<tr><td colspan="4" class="muted">Cargando…</td></tr>`;
  try {
    const url = `${APPS_SCRIPT_URL}?action=list`;
    const res = await fetch(url);
    const data = await res.json();
    CRM_COLUMNS = data.columns || [];
    LAST_ROWS = data.rows || [];
    renderDashboard(LAST_ROWS);
  } catch (err) {
    tbody.innerHTML = `<tr><td colspan="4" class="muted">No se pudo cargar el dashboard (configura el backend en js/config.js — ver README.md).</td></tr>`;
  }
}

function renderDashboard(rows) {
  const tbody = $("dashboardBody");
  if (!rows.length) {
    tbody.innerHTML = `<tr><td colspan="4" class="muted">Aún no hay oportunidades cargadas.</td></tr>`;
    return;
  }
  tbody.innerHTML = "";
  rows.forEach((r) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${r["End User Name"] || ""}</td>
      <td>${r["Subject"] || ""}</td>
      <td>${r["Government Number"] || ""}</td>
      <td>${r["Est. Revenue"] ? fmtMoney(parseFloat(r["Est. Revenue"])) : ""}</td>
    `;
    tbody.appendChild(tr);
  });
}

// ---------- Ver / copiar tabla (formato listo para pegar en X4A) ----------
function abrirModalTabla() {
  if (!LAST_ROWS.length) {
    showToast("No hay oportunidades cargadas para mostrar.", true);
    return;
  }
  const cols = CRM_COLUMNS.length
    ? CRM_COLUMNS
    : Object.keys(LAST_ROWS[0]).filter(
        (k) => k !== "Notificado" && k !== "Fecha Notificación"
      );

  const tabla = $("tablaCopiable");
  let html = "<thead><tr>";
  cols.forEach((c) => (html += `<th>${c}</th>`));
  html += "</tr></thead><tbody>";
  LAST_ROWS.forEach((r) => {
    html += "<tr>";
    cols.forEach((c) => (html += `<td>${r[c] ?? ""}</td>`));
    html += "</tr>";
  });
  html += "</tbody>";
  tabla.innerHTML = html;

  $("modalTabla").classList.add("show");
}

function cerrarModalTabla() {
  $("modalTabla").classList.remove("show");
}

async function copiarTablaAlPortapapeles() {
  const cols = CRM_COLUMNS.length
    ? CRM_COLUMNS
    : Object.keys(LAST_ROWS[0] || {}).filter(
        (k) => k !== "Notificado" && k !== "Fecha Notificación"
      );

  const lineas = [cols.join("\t")];
  LAST_ROWS.forEach((r) => {
    lineas.push(cols.map((c) => (r[c] ?? "").toString().replace(/\t/g, " ")).join("\t"));
  });
  const tsv = lineas.join("\n");

  try {
    await navigator.clipboard.writeText(tsv);
    showToast("Tabla copiada ✔ — pégala en X4A o Excel con Ctrl+V");
  } catch (err) {
    // Respaldo si el navegador bloquea el portapapeles: seleccionar la tabla
    const tabla = $("tablaCopiable");
    const range = document.createRange();
    range.selectNodeContents(tabla);
    const sel = window.getSelection();
    sel.removeAllRanges();
    sel.addRange(range);
    showToast("Tabla seleccionada — usa Ctrl+C para copiarla");
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
  ["valorCotizado", "trmDia", "fechaCierreIngram", "fechaCierreFabricante", "compraEstimada", "probabilidad", "situacion", "clienteFinal"].forEach(
    (id) => $(id).addEventListener("input", validateAll)
  );

  $("btnGuardar").addEventListener("click", guardarOportunidad);
  $("btnRefrescarDashboard").addEventListener("click", loadDashboard);
  $("btnVerTabla").addEventListener("click", abrirModalTabla);
  $("btnCerrarModal").addEventListener("click", cerrarModalTabla);
  $("btnCopiarTabla").addEventListener("click", copiarTablaAlPortapapeles);
  $("modalTabla").addEventListener("click", (e) => {
    if (e.target.id === "modalTabla") cerrarModalTabla();
  });

  $("trmDia").value = "";
  $("trmBase").value = "3708";
  $("fechaCierreIngram").value =
