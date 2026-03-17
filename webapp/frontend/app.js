/**
 * app.js - Allokering WebApp frontend
 */

const API = (() => {
  // Default: same origin (works on Render/production and local when served by FastAPI).
  if (window.location.protocol === "http:" || window.location.protocol === "https:") {
    return "";
  }
  // Fallback if frontend is opened as file://
  return "http://localhost:8000";
})();
const ASK_URL = "https://noeffectui-frey.nowastelogistics.com/desktop";
const ASK_LOGIN_WAIT_SECONDS = 60 * 60; // 60 min
const ASK_OPEN_VIA = "shortcut";        // alltid v+o

// ---------------------------------------------------------------------------
// Fil-konfiguration
// ---------------------------------------------------------------------------

const FILE_SLOTS = [
  { key: "orders",     label: "Detalj Kundorder (Alla)", accept: ".csv,.txt" },
  { key: "buffer",     label: "Buffertpall",              accept: ".csv,.txt" },
  { key: "automation", label: "Saldo inkl. automation",   accept: ".csv,.txt" },
  { key: "item",       label: "Item option",              accept: ".csv,.txt" },
  { key: "overview",   label: "Order\u00f6versikt",       accept: ".csv,.txt" },
  { key: "dispatch",   label: "Dispatchpallar",           accept: ".csv,.txt" },
];

const PROG_SLOTS = [
  { key: "prognos",  label: "Prognos",        accept: ".xlsx,.xls" },
  { key: "campaign", label: "Kampanjvolymer", accept: ".xlsx,.xls" },
];

const WMS_SLOTS = [
  { key: "wms_receive", label: "Varumottagningslogg",                    accept: ".csv,.txt" },
  { key: "wms_buffer",  label: "Buffertpall",                            accept: ".csv,.txt" },
  { key: "wms_booking", label: "Ej inlagrade artiklar",                  accept: ".csv,.txt" },
  { key: "wms_trans",   label: "Translogg",                              accept: ".csv,.txt" },
  { key: "wms_pick",    label: "Plocklogg full",                         accept: ".csv,.txt" },
  { key: "wms_correct", label: "Saldojusteringar & inventeringsavvikelser", accept: ".csv,.txt" },
];

const ASK_FETCHABLE_KEYS = new Set([
  ...FILE_SLOTS.map(s => s.key),
  ...WMS_SLOTS.map(s => s.key),
]);
const FILTER_REFRESH_KEYS = new Set([
  ...FILE_SLOTS.map(s => s.key),
  ...WMS_SLOTS.map(s => s.key),
]);
const ORDERSALDO_REFRESH_KEYS = new Set(["orders"]);
const ENABLE_ASK_CSV = false;
// UI-toggle: hide ASK "Hämta CSV" buttons but keep backend/frontend logic intact.
const SHOW_ASK_FETCH_BUTTONS = false && ENABLE_ASK_CSV;

const ASK_VIEW_OVERRIDES = {
  overview: "Orderöversikt",
  buffer: "Buffertpall",
  wms_buffer: "Buffertpall",
};

const CLASSIFIER_DATA_SLOTS = [
  { key: "item_attribute", label: "item_attribute" },
  { key: "item_alias", label: "item_alias" },
  { key: "item", label: "item" },
  { key: "main_category", label: "main_category" },
];

// Resultatnycklar -> visningsnamn
const RESULT_LABELS = {
  "allokerade":       "Allokering",
  "nearmiss":         "Near-Miss",
  "pallplatser":      "Preliminärbokning",
  "refill":           "Refill",
  "hib-koppling":     "HIB Kontroll",
  "orderkontroll":    "Order Kontroll",
  "dispatchkontroll": "Dispatch Kontroll",
  "eftersok":         "Öppna eftersök",
  "prognos":          "Prognos",
  "sales":            "Öppna försäljningsinsikter",
};

const ACTION_REQUIREMENTS = {
  "allokering": [
    { type: "file", key: "orders", label: "Beställningslinjer (CSV)" },
    { type: "file", key: "buffer", label: "Buffertpallar (CSV)" },
  ],
  "hib-koppling": [
    { type: "file", key: "orders", label: "Beställningslinjer (CSV)" },
    { type: "file", key: "overview", label: "Orderöversikt (CSV)" },
  ],
  "orderkontroll": [
    { type: "file", key: "overview", label: "Orderöversikt (CSV)" },
  ],
  "dispatchkontroll": [
    { type: "file", key: "overview", label: "Orderöversikt (CSV)" },
    { type: "file", key: "dispatch", label: "Dispatchpallar (CSV)" },
  ],
  "eftersok": [
    { type: "file", key: "wms_receive", label: "Varumottagningslogg (CSV)" },
    { type: "file", key: "wms_buffer", label: "Buffertpall (CSV)" },
    { type: "input", key: "purchase-input", label: "Inköpsnummer" },
    { type: "input", key: "article-input", label: "Artikelnummer" },
  ],
};

const COPY_BUTTON_REQUIREMENTS = {
  list1: {
    buttonId: "btn-kompletta",
    requiredLabels: ["Beställningslinjer (CSV)"],
    lockedHint: "För att aktivera krävs: Beställningslinjer (CSV), och att listan innehåller värden.",
  },
  list2: {
    buttonId: "btn-pafyllning",
    requiredLabels: ["Beställningslinjer (CSV)"],
    lockedHint: "För att aktivera krävs: Beställningslinjer (CSV), och att listan innehåller värden.",
  },
};

const OPEN_RESULT_ORDER = [
  "allokerade",
  "nearmiss",
  "pallplatser",
  "prognos",
  "refill",
  "hib-koppling",
  "orderkontroll",
  "dispatchkontroll",
  "eftersok",
];

const OPEN_BUTTON_HINTS = {
  "allokerade": "Tryck först: Kör allokering",
  "nearmiss": "Tryck först: Kör allokering",
  "pallplatser": "Tryck först: Kör allokering",
  "refill": "Tryck först: Kör allokering",
  "hib-koppling": "Tryck först: Kör HIB-koppling",
  "orderkontroll": "Tryck först: Kontrollera orderöversikt",
  "dispatchkontroll": "Tryck först: Kontrollera dispatchpallar",
  "eftersok": "Tryck först: Kör Eftersök",
  "prognos": "Ladda upp Prognos eller Kampanjvolymer först",
};

const THEME_STORAGE_KEY = "allok_theme";

// ---------------------------------------------------------------------------
// State
// ---------------------------------------------------------------------------

let sessionId = null;
let sseSource = null;
let sseReconnectTimer = null;
let sseReconnectLogAt = 0;
let sessionRecoveryPromise = null;
let fileStatuses = {};  // key -> filename | null
let availableResults = new Set();
let cachedFilterOptions = {};  // { bolag: [...], ordertyp: [...] }
let selectedFilters = { bolag: [], ordertyp: [] };
let actionMissingByJob = {}; // { jobId: [missing labels] }
let classifierStatusPoll = null;
let classifierConfig = null;
let classifierSessionId = null;
let classifierDataFiles = {};

function getCurrentTheme() {
  return document.documentElement.getAttribute("data-theme") === "dark" ? "dark" : "light";
}

function applyTheme(theme) {
  const resolvedTheme = theme === "dark" ? "dark" : "light";
  document.documentElement.setAttribute("data-theme", resolvedTheme);
  try {
    localStorage.setItem(THEME_STORAGE_KEY, resolvedTheme);
  } catch (_e) {
    // Tyst fel
  }

  const btn = document.getElementById("theme-toggle-btn");
  if (btn) {
    const isDark = resolvedTheme === "dark";
    btn.textContent = isDark ? "Ljust läge" : "Mörkt läge";
    btn.setAttribute("aria-pressed", isDark ? "true" : "false");
    btn.title = isDark ? "Byt till ljust läge" : "Byt till mörkt läge";
  }
}

function toggleTheme() {
  applyTheme(getCurrentTheme() === "dark" ? "light" : "dark");
  initTooltips();
}

function resetFrontendState(options = {}) {
  const clearInputs = !!options.clearInputs;
  fileStatuses = {};
  cachedFilterOptions = {};
  selectedFilters = { bolag: [], ordertyp: [] };
  availableResults.clear();

  [...FILE_SLOTS, ...PROG_SLOTS, ...WMS_SLOTS].forEach(slot => {
    setBadge(slot.key, "Ej fil", false);
  });

  renderResultButtons();
  renderFilterCard({});
  clearKalltypSummary();

  if (clearInputs) {
    const purchaseInput = document.getElementById("purchase-input");
    const articleInput = document.getElementById("article-input");
    if (purchaseInput) purchaseInput.value = "";
    if (articleInput) articleInput.value = "";
  }

  ["allokering", "hib-koppling", "orderkontroll", "dispatchkontroll", "eftersok"].forEach(resetJobButton);
}

async function readErrorDetail(resp) {
  try {
    const data = await resp.clone().json();
    const detail = String(data?.detail || data?.message || "").trim();
    if (detail) {
      return detail;
    }
  } catch (_e) {
    // Faller tillbaka till text
  }

  try {
    const text = String(await resp.text()).trim();
    if (text) {
      return text;
    }
  } catch (_e) {
    // Ignorera
  }

  return resp.statusText || `HTTP ${resp.status}`;
}

function isSessionMissingDetail(detail) {
  return /Session saknas/i.test(String(detail || ""));
}

async function recoverExpiredSession(reason) {
  if (sessionRecoveryPromise) {
    return sessionRecoveryPromise;
  }

  sessionRecoveryPromise = (async () => {
    const previousSessionId = sessionId;
    sessionId = null;
    try {
      sessionStorage.removeItem("allok_session_id");
    } catch (_e) {
      // Tyst fel
    }

    resetFrontendState();
    await ensureSession();
    connectSSE();
    await refreshFileStatus();
    await refreshFilterOptions();
    await refreshResultStatus();
    await refreshOrdersaldoButtonState();
    syncActionButtonsState();

    const reasonText = reason ? ` (${reason})` : "";
    appendLog(`Sessionen återställdes${reasonText}. Ladda upp filerna igen.`);
    return previousSessionId !== sessionId;
  })().finally(() => {
    sessionRecoveryPromise = null;
  });

  return sessionRecoveryPromise;
}

async function fetchWithSessionRecovery(url, options = {}, context = "") {
  const resp = await fetch(url, options);
  if (resp.ok) {
    return resp;
  }

  const detail = await readErrorDetail(resp);
  if (resp.status === 404 && isSessionMissingDetail(detail)) {
    await recoverExpiredSession(context || "servern tappade sessionen");
    throw new Error("Sessionen återställdes. Ladda upp filerna igen.");
  }

  throw new Error(detail || resp.statusText || `HTTP ${resp.status}`);
}

// ---------------------------------------------------------------------------
// Init
// ---------------------------------------------------------------------------

async function ensureSession() {
  // Kontrollera om befintlig session fortfarande lever (servern kan ha startats om)
  const stored = sessionStorage.getItem("allok_session_id");
  if (stored) {
    try {
      const check = await fetch(`${API}/api/upload/${stored}`);
      if (check.ok) {
        sessionId = stored;
        return;
      }
    } catch (e) { /* server nere eller session borta */ }
  }
  // Skapa ny session
  const resp = await fetch(`${API}/api/session`, { method: "POST" });
  const data = await resp.json();
  sessionId = data.session_id;
  sessionStorage.setItem("allok_session_id", sessionId);
}

window.addEventListener("DOMContentLoaded", async () => {
  await ensureSession();

  applyTheme(getCurrentTheme());
  setupMainTabs();
  renderFileRows();
  renderProgRows();
  renderWmsRows();
  renderResultButtons();
  connectSSE();
  setupDragDrop();

  await refreshFileStatus();
  await refreshFilterOptions();
  await refreshResultStatus();
  await refreshOrdersaldoButtonState();
  if (ENABLE_ASK_CSV) {
    await checkAskCsvAgentHealth();
  }
  syncActionButtonsState();
  bindActionInputWatchers();
  // Initialisera Bootstrap-tooltips
  initTooltips();
});

// ---------------------------------------------------------------------------
// Rendera fil-rader
// ---------------------------------------------------------------------------

function renderFileRows() {
  const container = document.getElementById("file-rows");
  container.innerHTML = "";
  FILE_SLOTS.forEach(slot => {
    container.appendChild(buildFileRow(slot));
  });
}

function renderProgRows() {
  const container = document.getElementById("prog-rows");
  container.innerHTML = "";
  PROG_SLOTS.forEach(slot => {
    container.appendChild(buildFileRow(slot));
  });
}

function renderWmsRows() {
  const container = document.getElementById("wms-file-rows");
  container.innerHTML = "";
  WMS_SLOTS.forEach(slot => {
    container.appendChild(buildFileRow(slot));
  });
}

function buildFileRow(slot) {
  const row = document.createElement("div");
  row.className = "d-flex align-items-center gap-2 mb-1";
  row.style.minWidth = "460px";

  const label = document.createElement("span");
  label.textContent = slot.label;
  label.style.minWidth = "220px";
  label.style.fontSize = "13px";
  label.style.color = "var(--page-text)";

  const badge = document.createElement("span");
  badge.className = "file-status-badge";
  badge.id = `badge-${slot.key}`;
  badge.textContent = "Ej fil";
  badge.title = "Klicka för att välja fil";
  badge.addEventListener("click", () => triggerFilePicker(slot.key, slot.accept));

  const removeBtn = document.createElement("button");
  removeBtn.className = "btn btn-danger btn-sm py-0 px-1";
  removeBtn.innerHTML = "&#10005;";
  removeBtn.title = `Ta bort ${slot.label}`;
  removeBtn.setAttribute("data-bs-toggle", "tooltip");
  removeBtn.addEventListener("click", () => removeFile(slot.key));

  const askBtn = document.createElement("button");
  askBtn.className = "btn btn-outline-primary btn-sm py-0 px-2";
  askBtn.id = `btn-ask-fetch-${slot.key}`;
  askBtn.textContent = "H\u00e4mta CSV";
  askBtn.title = `H\u00e4mta ${viewNameFromSlotLabel(slot.label)} fr\u00e5n ASK och l\u00e4gg i denna slot`;
  askBtn.setAttribute("data-bs-toggle", "tooltip");
  askBtn.addEventListener("click", () => fetchAskCsvToSlot(slot.key, slot.label));

  // Hidden file input (multiple tillatet for att kunna valja flera filer)
  const input = document.createElement("input");
  input.type = "file";
  input.accept = slot.accept || "";
  input.multiple = true;
  input.style.display = "none";
  input.id = `input-${slot.key}`;
  input.addEventListener("change", async (e) => {
    const files = Array.from(e.target.files);
    const changedKeys = new Set();
    for (const file of files) {
      const detectedKey = matchFileToSlot(file.name) || slot.key;
      const uploadedKey = await uploadFile(detectedKey, file, { skipPostRefresh: true });
      if (uploadedKey) {
        changedKeys.add(uploadedKey);
      }
    }
    if (changedKeys.size > 0) {
      await refreshAfterFileChanges(changedKeys);
    }
    input.value = "";
  });

  row.appendChild(label);
  row.appendChild(badge);
  row.appendChild(removeBtn);
  if (isAskFetchableSlot(slot.key)) {
    row.appendChild(askBtn);
  }
  row.appendChild(input);
  return row;
}

function isAskFetchableSlot(fileKey) {
  return SHOW_ASK_FETCH_BUTTONS && ASK_FETCHABLE_KEYS.has(fileKey);
}

function viewNameFromSlotLabel(label) {
  return (label || "")
    .replace(/\s*\((csv|xlsx|xls)\)\s*$/i, "")
    .trim();
}

function askViewNameForSlot(fileKey, label) {
  return ASK_VIEW_OVERRIDES[fileKey] || viewNameFromSlotLabel(label);
}

function triggerFilePicker(key, accept) {
  const input = document.getElementById(`input-${key}`);
  if (input) input.click();
}

// ---------------------------------------------------------------------------
// Upload / Remove
// ---------------------------------------------------------------------------

function hasTrackedKey(changedKeys, trackedKeys) {
  for (const key of changedKeys || []) {
    if (trackedKeys.has(key)) {
      return true;
    }
  }
  return false;
}

async function refreshAfterFileChanges(changedKeys) {
  if (hasTrackedKey(changedKeys, FILTER_REFRESH_KEYS)) {
    await refreshFilterOptions();
  }
  if (hasTrackedKey(changedKeys, ORDERSALDO_REFRESH_KEYS)) {
    await refreshOrdersaldoButtonState();
  }
  syncActionButtonsState();
}

async function uploadFile(fileKey, file, options = {}) {
  const { skipPostRefresh = false } = options;
  const previousStatus = fileStatuses[fileKey] || null;
  setBadge(fileKey, "Laddar...", false);
  try {
    const formData = new FormData();
    formData.append("file_key", fileKey);
    formData.append("page_id", getActiveMainPage());
    formData.append("file", file);
    const resp = await fetchWithSessionRecovery(`${API}/api/upload/${sessionId}`, {
      method: "POST",
      body: formData,
    }, "uppladdning");
    const data = await resp.json();
    const actualKey = data.file_key || fileKey;
    if (actualKey !== fileKey) {
      fileStatuses[fileKey] = previousStatus;
      setBadge(fileKey, previousStatus || "Ej fil", !!previousStatus);
      appendLog(`Fil uppladdad: [${actualKey}] ${data.filename} (innehållsdetekterad från ${fileKey})`);
    } else {
      appendLog(`Fil uppladdad: [${actualKey}] ${data.filename}`);
    }
    fileStatuses[actualKey] = data.filename;
    setBadge(actualKey, data.filename, true);
    if (!skipPostRefresh) {
      await refreshAfterFileChanges(new Set([actualKey]));
    }
    return actualKey;
  } catch (e) {
    setBadge(fileKey, "FEL", false);
    appendLog(`Fel vid uppladdning av ${fileKey}: ${e}`);
    syncActionButtonsState();
    return null;
  }
}

async function removeFile(fileKey) {
  try {
    await fetchWithSessionRecovery(`${API}/api/upload/${sessionId}/${fileKey}`, { method: "DELETE" }, "borttagning av fil");
    fileStatuses[fileKey] = null;
    setBadge(fileKey, "Ej fil", false);
    appendLog(`Fil borttagen: [${fileKey}]`);
    await refreshAfterFileChanges(new Set([fileKey]));
  } catch (e) {
    appendLog(`Fel vid borttagning av ${fileKey}: ${e}`);
  }
}

async function refreshFileStatus() {
  try {
    const resp = await fetchWithSessionRecovery(`${API}/api/upload/${sessionId}`, {}, "hämtning av filstatus");
    const data = await resp.json();
    [...FILE_SLOTS, ...PROG_SLOTS, ...WMS_SLOTS].forEach(slot => {
      const filename = data.files?.[slot.key] || null;
      fileStatuses[slot.key] = filename;
      setBadge(slot.key, filename || "Ej fil", !!filename);
    });
    for (const [key, filename] of Object.entries(data.files || {})) {
      fileStatuses[key] = filename;
      setBadge(key, filename || "Ej fil", !!filename);
    }
    syncActionButtonsState();
  } catch (e) {
    // Tyst fel
  }
}

function setBadge(key, text, loaded) {
  const badge = document.getElementById(`badge-${key}`);
  if (!badge) return;
  const maxLen = 18;
  const displayText = text.length > maxLen ? "..." + text.slice(-maxLen + 3) : text;
  badge.textContent = displayText;
  badge.title = text;
  if (loaded) {
    badge.classList.add("loaded");
  } else {
    badge.classList.remove("loaded");
  }
}

// ---------------------------------------------------------------------------
// ASK CSV bridge
// ---------------------------------------------------------------------------

let askInitPromise = null;

async function checkAskCsvAgentHealth() {
  if (!ENABLE_ASK_CSV) return;
  try {
    const resp = await fetch(`${API}/api/ask-csv/health`);
    if (!resp.ok) throw new Error("agent ej nåbar");
    const data = await resp.json();
    if (data.ready) {
      appendLog("ASK CSV-agent: redo.");
    } else {
      appendLog("ASK CSV-agent: uppe men ej initierad. Init sker automatiskt vid första 'Hämta CSV'.");
    }
  } catch (e) {
    appendLog("ASK CSV-agent: inte nåbar. Starta start_ask_csv_listener.bat.");
  }
}

async function initAskCsvAgent(options = {}) {
  if (!ENABLE_ASK_CSV) return false;
  const silent = !!options.silent;
  if (askInitPromise) return askInitPromise;

  askInitPromise = (async () => {
    try {
      if (!silent) {
        appendLog(`Initierar ASK CSV-agent (loginfönster upp till ${ASK_LOGIN_WAIT_SECONDS}s)...`);
      }
      const resp = await fetch(`${API}/api/ask-csv/init`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          url: ASK_URL,
          login_wait: ASK_LOGIN_WAIT_SECONDS,
          headless: false,
          slow_mo: 80,
          goto_timeout: 60,
        }),
      });
      if (!resp.ok) {
        const errData = await resp.json().catch(() => ({}));
        throw new Error(errData.detail || resp.statusText);
      }
      const data = await resp.json().catch(() => ({}));
      if (data && data.ready === false) {
        appendLog("ASK CSV-agent: login hann inte bli klar inom väntetiden.");
        return false;
      }
      if (!silent) {
        appendLog("ASK CSV-agent init klar.");
      }
      return true;
    } catch (e) {
      appendLog(`FEL vid init av ASK CSV-agent: ${e}`);
      return false;
    } finally {
      askInitPromise = null;
    }
  })();

  return askInitPromise;
}

async function fetchAskCsvToSlot(fileKey, slotLabel) {
  if (!ENABLE_ASK_CSV) {
    appendLog("ASK CSV ar avstangt i denna build.");
    return;
  }
  const btn = document.getElementById(`btn-ask-fetch-${fileKey}`);
  const original = btn ? btn.textContent : "Hämta CSV";
  if (btn) {
    btn.disabled = true;
    btn.innerHTML = `<span class="spinner-border spinner-border-sm me-1"></span>${original}`;
  }

  try {
    let healthReady = false;
    try {
      const h = await fetch(`${API}/api/ask-csv/health`);
      if (h.ok) {
        const hs = await h.json();
        healthReady = !!hs.ready;
      }
    } catch (e) {
      // ignoreras här, fångas nedan vid init/fetch
    }

    if (!healthReady) {
      appendLog("ASK CSV-agent ej initierad. Startar login-flöde nu...");
      const ok = await initAskCsvAgent({ silent: true });
      if (!ok) {
        throw new Error("Initiering misslyckades. Se loggen ovan.");
      }
      appendLog("Login-flöde klart. Fortsätter med CSV-hämtning...");
    }

    const viewName = askViewNameForSlot(fileKey, slotLabel);

    appendLog(`Hämtar CSV från ASK: vy='${viewName}' -> slot='${fileKey}'...`);
    const resp = await fetchWithSessionRecovery(`${API}/api/ask-csv/fetch/${sessionId}`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        view_name: viewName,
        target_file_key: fileKey,
        open_via: ASK_OPEN_VIA,
        open_text: "Visa",
        export_text: "Exportera till CSV",
        grid_wait: 30,
        download_timeout: 120,
      }),
    }, "ASK CSV-hämtning");
    const data = await resp.json();

    fileStatuses[fileKey] = data.filename;
    setBadge(fileKey, data.filename, true);
    appendLog(`CSV hämtad och kopplad till [${fileKey}]: ${data.filename} (${data.bytes} bytes)`);
    await refreshAfterFileChanges(new Set([fileKey]));
  } catch (e) {
    appendLog(`FEL vid CSV-hämtning: ${e}`);
  } finally {
    if (btn) {
      btn.disabled = false;
      btn.textContent = original;
    }
  }
}
// ---------------------------------------------------------------------------
// Filter
// ---------------------------------------------------------------------------

async function refreshFilterOptions() {
  try {
    const resp = await fetchWithSessionRecovery(`${API}/api/filters/${sessionId}`, {}, "hämtning av filter");
    const data = await resp.json();
    cachedFilterOptions = data;
    renderFilterCard(data);
  } catch (e) {
    // Tyst fel
  }
}

function renderFilterCard(options) {
  const normalized = { ...(options || {}) };
  const card = document.getElementById("filter-card");
  const content = document.getElementById("filter-content");
  const hasBolag = normalized.bolag && normalized.bolag.length > 0;
  const hasOrdertyp = normalized.ordertyp && normalized.ordertyp.length > 0;

  if (!hasBolag && !hasOrdertyp) {
    card.style.display = "none";
    return;
  }
  card.style.display = "";
  content.innerHTML = "";

  const row = document.createElement("div");
  row.className = "d-flex gap-3 flex-wrap";

  if (hasBolag) {
    row.appendChild(buildFilterGroup("bolag", "Bolag", normalized.bolag));
  }
  if (hasOrdertyp) {
    row.appendChild(buildFilterGroup("ordertyp", "Ordertyp", normalized.ordertyp));
  }
  content.appendChild(row);

  // "Valj alla" / "Rensa" knappar
  const btnRow = document.createElement("div");
  btnRow.className = "d-flex gap-2 mt-2";
  const selectAll = document.createElement("button");
  selectAll.className = "btn btn-sm btn-outline-secondary";
  selectAll.textContent = "Valj alla";
  selectAll.onclick = () => {
    document.querySelectorAll(".filter-check").forEach(cb => { cb.checked = true; });
    saveFilters();
  };
  const clearAll = document.createElement("button");
  clearAll.className = "btn btn-sm btn-outline-secondary";
  clearAll.textContent = "Rensa filter";
  clearAll.onclick = () => {
    document.querySelectorAll(".filter-check").forEach(cb => { cb.checked = false; });
    saveFilters();
  };
  btnRow.appendChild(selectAll);
  btnRow.appendChild(clearAll);
  content.appendChild(btnRow);
}

function buildFilterGroup(key, title, values) {
  const group = document.createElement("div");
  group.innerHTML = `<strong>${title}</strong>`;
  values.forEach(val => {
    const id = `filter-${key}-${val.replace(/\W/g, "_")}`;
    const div = document.createElement("div");
    div.className = "form-check form-check-sm";
    const isChecked = selectedFilters[key].length === 0 || selectedFilters[key].includes(val);
    div.innerHTML = `
      <input class="form-check-input filter-check" type="checkbox" id="${id}"
             data-group="${key}" value="${val}" ${isChecked ? "checked" : ""}>
      <label class="form-check-label" for="${id}" style="font-size:12px">${val}</label>
    `;
    div.querySelector("input").addEventListener("change", saveFilters);
    group.appendChild(div);
  });
  return group;
}

async function saveFilters() {
  // Preserve hidden groups (e.g. if Ordertyp group is temporarily unavailable
  // after file refresh) so we don't accidentally reset those selections.
  const filters = {
    bolag: Array.isArray(selectedFilters.bolag) ? [...selectedFilters.bolag] : [],
    ordertyp: Array.isArray(selectedFilters.ordertyp) ? [...selectedFilters.ordertyp] : [],
  };
  const groupSeen = { bolag: false, ordertyp: false };
  const groupValues = { bolag: [], ordertyp: [] };

  document.querySelectorAll(".filter-check").forEach(cb => {
    if (cb.checked) {
      const group = cb.dataset.group;
      if (groupValues[group] !== undefined) {
        groupSeen[group] = true;
        groupValues[group].push(cb.value);
      }
    } else {
      const group = cb.dataset.group;
      if (groupValues[group] !== undefined) {
        groupSeen[group] = true;
      }
    }
  });

  for (const key of Object.keys(groupSeen)) {
    if (groupSeen[key]) {
      filters[key] = groupValues[key];
    }
  }

  selectedFilters = filters;
  try {
    await fetchWithSessionRecovery(`${API}/api/filters/${sessionId}`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(filters),
    }, "sparning av filter");
  } catch (_e) {
    // Tyst fel
  }
  await refreshOrdersaldoButtonState();
  appendLog(`Filter uppdaterat: bolag=[${filters.bolag.join(",")}], ordertyp=[${filters.ordertyp.join(",")}]`);
}

function setButtonLockState(btn, locked, hintText) {
  if (!btn) return;
  if (!btn.dataset.baseTitle) {
    btn.dataset.baseTitle = btn.title || "";
  }
  btn.dataset.locked = locked ? "1" : "0";
  btn.classList.toggle("btn-soft-disabled", !!locked);
  btn.setAttribute("aria-disabled", locked ? "true" : "false");
  if (locked) {
    btn.title = hintText || "Saknar indata";
  } else {
    btn.title = btn.dataset.baseTitle || "";
  }
}

function getRequiredLabelsForJob(job) {
  return (ACTION_REQUIREMENTS[job] || [])
    .map(req => req.label)
    .filter(Boolean);
}

function buildRequirementTitle(baseTitle, requiredLabels) {
  const base = String(baseTitle || "").trim();
  const requirementText = (requiredLabels || []).length > 0
    ? `Kräver: ${(requiredLabels || []).join(", ")}.`
    : "";
  return [base, requirementText].filter(Boolean).join(" ");
}

function getMissingRequirements(job) {
  const reqs = ACTION_REQUIREMENTS[job] || [];
  const missing = [];
  reqs.forEach(req => {
    if (req.type === "file") {
      if (!fileStatuses[req.key]) missing.push(req.label);
      return;
    }
    if (req.type === "input") {
      const el = document.getElementById(req.key);
      if (!el || !String(el.value || "").trim()) missing.push(req.label);
    }
  });
  return missing;
}

function syncActionButtonsState() {
  actionMissingByJob = {};
  Object.keys(ACTION_REQUIREMENTS).forEach(job => {
    const btn = document.getElementById(`btn-${job}`);
    if (!btn) return;
    const missing = getMissingRequirements(job);
    const readyTitle = buildRequirementTitle(btn.dataset.baseTitle || btn.title || "", getRequiredLabelsForJob(job));
    actionMissingByJob[job] = missing;
    if (missing.length > 0) {
      setButtonLockState(btn, true, `För att aktivera krävs: ${missing.join(", ")}`);
    } else {
      setButtonLockState(btn, false, "");
      btn.title = readyTitle || btn.title;
    }
  });
  initTooltips();
}

function bindActionInputWatchers() {
  ["purchase-input", "article-input"].forEach(id => {
    const el = document.getElementById(id);
    if (!el || el.dataset.boundActionWatcher === "1") return;
    el.addEventListener("input", () => syncActionButtonsState());
    el.dataset.boundActionWatcher = "1";
  });
}

// ---------------------------------------------------------------------------
// Korning
// ---------------------------------------------------------------------------

async function runJob(job) {
  const missing = getMissingRequirements(job);
  if (missing.length > 0) {
    appendLog(`Saknar filer: ${missing.join(", ")}`);
    syncActionButtonsState();
    return;
  }

  const btn = document.getElementById(`btn-${job}`);
  if (btn) {
    btn.disabled = true;
    btn.dataset.originalText = btn.textContent;
    btn.innerHTML = `<span class="spinner-border spinner-border-sm me-1"></span>${btn.dataset.originalText}`;
  }

  try {
    let url = `${API}/api/run/${job}/${sessionId}`;
    let body = undefined;
    let method = "POST";
    let headers = {};

    if (job === "eftersok") {
      body = JSON.stringify({
        purchase: document.getElementById("purchase-input").value.trim(),
        article: document.getElementById("article-input").value.trim(),
      });
      headers["Content-Type"] = "application/json";
    }

    await fetchWithSessionRecovery(url, { method, body, headers }, `start av ${job}`);
    appendLog(`Startade jobb: ${job}`);
  } catch (e) {
    appendLog(`FEL vid start av ${job}: ${e}`);
    if (btn) {
      btn.disabled = false;
      btn.textContent = btn.dataset.originalText || btn.textContent;
    }
  }
}

function resetJobButton(job) {
  const btn = document.getElementById(`btn-${job}`);
  if (btn) {
    btn.disabled = false;
    btn.textContent = btn.dataset.originalText || btn.textContent;
    initTooltips();
  }
  syncActionButtonsState();
}

// ---------------------------------------------------------------------------
// SSE-logg
// ---------------------------------------------------------------------------

let currentJob = null;

function connectSSE() {
  if (sseSource) {
    sseSource.close();
  }
  if (sseReconnectTimer) {
    clearTimeout(sseReconnectTimer);
    sseReconnectTimer = null;
  }
  sseSource = new EventSource(`${API}/api/log/stream/${sessionId}`);
  sseSource.onmessage = (e) => {
    const msg = e.data.replace(/\\n/g, "\n");

    // Kalltyp-summering
    if (msg.startsWith("__KALLTYP_SUMMARY:")) {
      try {
        const jsonStr = msg.replace("__KALLTYP_SUMMARY:", "").replace(/__$/, "");
        renderKalltypSummary(JSON.parse(jsonStr));
      } catch (_e) { /* ignore parse errors */ }
      return;
    }

    // Resultat-events
    if (msg.startsWith("__RESULT:")) {
      const key = msg.replace("__RESULT:", "").replace(/__/g, "").trim();
      activateResultButton(key);
      return;
    }

    // Klar/Fel
    if (msg === "__DONE__" || msg === "__ERROR__") {
      // Aterstall alla korningsknappar
      ["allokering", "hib-koppling", "orderkontroll", "dispatchkontroll", "eftersok"].forEach(resetJobButton);
      appendLog(msg === "__DONE__" ? "--- Klar ---" : "--- Avbrots med fel ---");
      refreshResultStatus().catch(() => {});
      if (msg === "__DONE__") {
        // Uppdatera ordersaldo-listor efter avslutad korning
        fetch(`${API}/api/ordersaldo/refresh/${sessionId}`, { method: "POST" })
          .then(() => refreshOrdersaldoButtonState().catch(() => {}))
          .catch(() => {});
      } else {
        refreshOrdersaldoButtonState().catch(() => {});
      }
      return;
    }

    appendLog(msg);
  };
  sseSource.onerror = async () => {
    // Servern kan ha startats om -> session-id kan vara ogiltigt.
    if (sseSource) {
      sseSource.close();
      sseSource = null;
    }
    if (sseReconnectTimer) {
      return;
    }
    sseReconnectTimer = setTimeout(async () => {
      sseReconnectTimer = null;
      try {
        await ensureSession();
      } catch (e) {
        const now = Date.now();
        if (now - sseReconnectLogAt > 30000) {
          appendLog(`SSE reconnect misslyckades: ${e}`);
          sseReconnectLogAt = now;
        }
      }
      connectSSE();
    }, 3000);
  };
}

// ---------------------------------------------------------------------------
// Logg
// ---------------------------------------------------------------------------

function appendLog(msg) {
  const el = document.getElementById("log-output");
  el.textContent += msg + "\n";
  el.scrollTop = el.scrollHeight;
}

function appendClassifierLog(msg) {
  const el = document.getElementById("classifier-log-output");
  if (!el) return;
  const ts = new Date().toLocaleTimeString("sv-SE");
  el.textContent += `[${ts}] ${msg}\n`;
  el.scrollTop = el.scrollHeight;
}

// ---------------------------------------------------------------------------
// Main tabs
// ---------------------------------------------------------------------------

function setupMainTabs() {
  const buttons = Array.from(document.querySelectorAll("#main-tabs .nav-link[data-page]"));
  if (buttons.length === 0) return;
  buttons.forEach(btn => {
    btn.addEventListener("click", () => {
      setMainPage(btn.dataset.page || "allokering-page");
    });
  });
  const activeButton = buttons.find(btn => btn.classList.contains("active"));
  setMainPage(activeButton?.dataset.page || "allokering-page");
}

function setMainPage(pageId) {
  document.querySelectorAll(".main-page-pane").forEach(pane => {
    pane.classList.toggle("active", pane.id === pageId);
  });
  document.querySelectorAll("#main-tabs .nav-link[data-page]").forEach(btn => {
    btn.classList.toggle("active", btn.dataset.page === pageId);
  });
}

function getActiveMainPage() {
  return document.querySelector(".main-page-pane.active")?.id || "allokering-page";
}

// ---------------------------------------------------------------------------
// Classifier web flow
// ---------------------------------------------------------------------------

function parseClassifierCategories() {
  const raw = document.getElementById("cls-categories")?.value || "";
  const lines = raw
    .split("\n")
    .map(v => v.trim())
    .filter(Boolean);
  const out = [];
  const seen = new Set();
  for (const line of lines) {
    const key = line.toLowerCase();
    if (seen.has(key)) continue;
    seen.add(key);
    out.push(line);
  }
  return out;
}

function setClassifierUiEnabled(enabled) {
  const ids = [
    "cls-test-name",
    "cls-image-dir",
    "cls-output-dir",
    "cls-categories",
    "cls-shuffle",
    "btn-classifier-session-start",
  ];
  ids.forEach(id => {
    const el = document.getElementById(id);
    if (el) el.disabled = !enabled;
  });
}

function setClassifierStatusBadge(kind, text) {
  const badge = document.getElementById("classifier-status-badge");
  if (!badge) return;
  if (kind === "running") badge.className = "badge text-bg-success";
  else if (kind === "done") badge.className = "badge text-bg-primary";
  else if (kind === "error") badge.className = "badge text-bg-danger";
  else badge.className = "badge text-bg-secondary";
  badge.textContent = text;
}

function buildClassifierDataRow(slot) {
  const row = document.createElement("div");
  row.className = "d-flex align-items-center gap-2 mb-1 classifier-data-row";

  const label = document.createElement("span");
  label.textContent = slot.label;
  label.style.minWidth = "110px";
  label.style.fontSize = "12px";
  label.style.color = "var(--page-text)";

  const badge = document.createElement("span");
  badge.className = "file-status-badge";
  badge.id = `cls-data-badge-${slot.key}`;
  badge.textContent = "Ej fil";
  badge.title = "Välj fil";
  badge.addEventListener("click", () => {
    const inp = document.getElementById(`cls-data-input-${slot.key}`);
    if (inp) inp.click();
  });

  const uploadBtn = document.createElement("button");
  uploadBtn.className = "btn btn-outline-primary btn-sm py-0 px-2";
  uploadBtn.textContent = "Ladda upp";
  uploadBtn.onclick = () => {
    const inp = document.getElementById(`cls-data-input-${slot.key}`);
    if (inp) inp.click();
  };

  const removeBtn = document.createElement("button");
  removeBtn.className = "btn btn-danger btn-sm py-0 px-1";
  removeBtn.innerHTML = "&#10005;";
  removeBtn.title = "Ta bort fil";
  removeBtn.onclick = () => removeClassifierDataFile(slot.key);

  const input = document.createElement("input");
  input.type = "file";
  input.accept = ".csv,.txt";
  input.style.display = "none";
  input.id = `cls-data-input-${slot.key}`;
  input.addEventListener("change", async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    await uploadClassifierDataFile(slot.key, file);
    input.value = "";
  });

  row.appendChild(label);
  row.appendChild(badge);
  row.appendChild(uploadBtn);
  row.appendChild(removeBtn);
  row.appendChild(input);
  return row;
}

function renderClassifierDataRows() {
  const container = document.getElementById("cls-data-file-rows");
  if (!container) return;
  container.innerHTML = "";
  CLASSIFIER_DATA_SLOTS.forEach(slot => {
    container.appendChild(buildClassifierDataRow(slot));
  });
}

function setClassifierDataBadge(key, text, loaded) {
  const badge = document.getElementById(`cls-data-badge-${key}`);
  if (!badge) return;
  const maxLen = 18;
  const raw = String(text || "");
  const displayText = raw.length > maxLen ? "..." + raw.slice(-maxLen + 3) : raw;
  badge.textContent = displayText || "Ej fil";
  badge.title = raw || "Ej fil";
  badge.classList.toggle("loaded", !!loaded);
}

function applyClassifierDataStatus(files) {
  classifierDataFiles = files || {};
  CLASSIFIER_DATA_SLOTS.forEach(slot => {
    const f = classifierDataFiles[slot.key];
    const loaded = !!(f && f.exists);
    const name = loaded ? (f.filename || "Laddad") : "Ej fil";
    setClassifierDataBadge(slot.key, name, loaded);
  });
}

async function refreshClassifierDataFiles() {
  const resp = await fetch(`${API}/api/classifier/web/data-files`);
  const data = await resp.json().catch(() => ({}));
  if (!resp.ok) throw new Error(data.detail || resp.statusText);
  applyClassifierDataStatus(data.files || {});
  if (typeof data.item_attribute_rows === "number") {
    appendClassifierLog(`item_attribute IMG-rader: ${data.item_attribute_rows}`);
  }
}

async function uploadClassifierDataFile(fileKey, file) {
  setClassifierDataBadge(fileKey, "Laddar...", false);
  try {
    const formData = new FormData();
    formData.append("file_key", fileKey);
    formData.append("file", file);
    const resp = await fetch(`${API}/api/classifier/web/data-files/upload`, {
      method: "POST",
      body: formData,
    });
    const data = await resp.json().catch(() => ({}));
    if (!resp.ok) throw new Error(data.detail || resp.statusText);
    applyClassifierDataStatus(data.files || {});
    if (fileKey === "item_attribute" && typeof data.item_attribute_rows === "number") {
      appendClassifierLog(`item_attribute uppladdad, IMG-rader: ${data.item_attribute_rows}`);
    } else {
      appendClassifierLog(`${fileKey} uppladdad: ${data.filename || file.name}`);
    }
  } catch (e) {
    setClassifierDataBadge(fileKey, "FEL", false);
    appendClassifierLog(`Kunde inte ladda upp ${fileKey}: ${e}`);
  }
}

async function removeClassifierDataFile(fileKey) {
  try {
    const resp = await fetch(`${API}/api/classifier/web/data-files/${fileKey}`, { method: "DELETE" });
    const data = await resp.json().catch(() => ({}));
    if (!resp.ok) throw new Error(data.detail || resp.statusText);
    applyClassifierDataStatus(data.files || {});
    appendClassifierLog(`${fileKey} borttagen.`);
  } catch (e) {
    appendClassifierLog(`Kunde inte ta bort ${fileKey}: ${e}`);
  }
}

function renderClassifierState(state, announce = false) {
  const statusText = document.getElementById("classifier-status-text");
  const sessionIdEl = document.getElementById("classifier-session-id");
  const fileEl = document.getElementById("classifier-current-file");
  const img = document.getElementById("classifier-current-image");
  const catWrap = document.getElementById("classifier-category-buttons");
  const skipBtn = document.getElementById("btn-classifier-skip");
  const finishBtn = document.getElementById("btn-classifier-session-finish");

  if (!statusText || !sessionIdEl || !fileEl || !img || !catWrap || !skipBtn || !finishBtn) return;

  if (!state) {
    setClassifierStatusBadge("idle", "Ej startad");
    statusText.textContent = "Starta en session för att börja.";
    sessionIdEl.textContent = "-";
    fileEl.textContent = "-";
    img.style.display = "none";
    img.removeAttribute("src");
    catWrap.innerHTML = "";
    skipBtn.disabled = true;
    finishBtn.disabled = true;
    setClassifierUiEnabled(true);
    return;
  }

  classifierSessionId = state.session_id || classifierSessionId;
  sessionIdEl.textContent = classifierSessionId || "-";
  const sourceLabel = state.image_source === "url" ? "källa: datafil-URL" : "källa: bildmapp";
  statusText.textContent = `${state.index || 0}/${state.total || 0} klassificerade, hoppade över: ${state.skipped || 0} (${sourceLabel})`;
  fileEl.textContent = state.current_article ? `${state.current_article} (${state.current_filename || "-"})` : (state.current_filename || "-");
  finishBtn.disabled = false;

  if (state.done) {
    setClassifierStatusBadge("done", "Klar");
    img.style.display = "none";
    img.removeAttribute("src");
    catWrap.innerHTML = "";
    skipBtn.disabled = true;
    setClassifierUiEnabled(true);
    if (announce) appendClassifierLog("Session klar.");
    return;
  }

  setClassifierStatusBadge("running", "Pågår");
  setClassifierUiEnabled(false);
  skipBtn.disabled = false;

  if (state.image_url) {
    img.src = `${API}${state.image_url}`;
    img.style.display = "block";
  } else {
    img.style.display = "none";
    img.removeAttribute("src");
  }

  catWrap.innerHTML = "";
  (state.categories || []).forEach(cat => {
    const btn = document.createElement("button");
    btn.className = "btn btn-success btn-sm";
    btn.textContent = cat;
    btn.onclick = () => classifyCurrentImage(cat);
    catWrap.appendChild(btn);
  });
}

async function initClassifierWeb() {
  renderClassifierState(null);
  renderClassifierDataRows();
  try {
    const resp = await fetch(`${API}/api/classifier/web/config`);
    if (!resp.ok) throw new Error(await resp.text());
    classifierConfig = await resp.json();
    const imageDir = document.getElementById("cls-image-dir");
    const outputDir = document.getElementById("cls-output-dir");
    const testName = document.getElementById("cls-test-name");
    const categories = document.getElementById("cls-categories");
    if (imageDir) imageDir.value = classifierConfig.default_image_dir || "";
    if (outputDir) outputDir.value = classifierConfig.default_output_dir || "";
    if (testName && !testName.value) testName.value = "test1";
    if (categories && !categories.value) categories.value = "Hund\nKatt\nÖvrigt";
    applyClassifierDataStatus(classifierConfig.data_files || {});
    await refreshClassifierDataFiles();
    appendClassifierLog("Classifier-web klar.");
  } catch (e) {
    setClassifierStatusBadge("error", "Fel");
    const statusText = document.getElementById("classifier-status-text");
    if (statusText) statusText.textContent = "Kunde inte ladda classifier-konfiguration";
    appendClassifierLog(`Init-fel: ${e}`);
  }
}

async function startClassifierSession() {
  const btn = document.getElementById("btn-classifier-session-start");
  const original = btn ? btn.textContent : "Starta klassificering";
  if (btn) {
    btn.disabled = true;
    btn.innerHTML = `<span class="spinner-border spinner-border-sm me-1"></span>${original}`;
  }

  try {
    const testName = (document.getElementById("cls-test-name")?.value || "").trim();
    const imageDir = (document.getElementById("cls-image-dir")?.value || "").trim();
    const outputDir = (document.getElementById("cls-output-dir")?.value || "").trim();
    const shuffle = !!document.getElementById("cls-shuffle")?.checked;
    const categories = parseClassifierCategories();
    if (!testName) throw new Error("Testnamn saknas.");
    if (categories.length === 0) throw new Error("Lägg in minst en kategori.");

    const resp = await fetch(`${API}/api/classifier/web/start`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        test_name: testName,
        categories,
        image_dir: imageDir,
        output_dir: outputDir,
        shuffle,
      }),
    });
    const data = await resp.json().catch(() => ({}));
    if (!resp.ok) throw new Error(data.detail || resp.statusText);
    renderClassifierState(data, true);
    appendClassifierLog(data.message || "Session startad.");
  } catch (e) {
    setClassifierStatusBadge("error", "Fel");
    appendClassifierLog(`Kunde inte starta: ${e}`);
  } finally {
    if (btn) {
      btn.disabled = false;
      btn.textContent = original;
    }
  }
}

async function refreshClassifierSessionState(announce = false) {
  if (!classifierSessionId) {
    if (announce) appendClassifierLog("Ingen aktiv session.");
    return;
  }
  try {
    const resp = await fetch(`${API}/api/classifier/web/${classifierSessionId}`);
    const data = await resp.json().catch(() => ({}));
    if (!resp.ok) throw new Error(data.detail || resp.statusText);
    renderClassifierState(data, announce);
  } catch (e) {
    appendClassifierLog(`Statusfel: ${e}`);
  }
}

async function classifyCurrentImage(category) {
  if (!classifierSessionId) return;
  try {
    const resp = await fetch(`${API}/api/classifier/web/${classifierSessionId}/classify`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ category }),
    });
    const data = await resp.json().catch(() => ({}));
    if (!resp.ok) throw new Error(data.detail || resp.statusText);
    renderClassifierState(data, false);
    appendClassifierLog(data.message || `Klassad som ${category}`);
  } catch (e) {
    appendClassifierLog(`Klassificering misslyckades: ${e}`);
  }
}

async function skipClassifierImage() {
  if (!classifierSessionId) return;
  try {
    const resp = await fetch(`${API}/api/classifier/web/${classifierSessionId}/skip`, {
      method: "POST",
    });
    const data = await resp.json().catch(() => ({}));
    if (!resp.ok) throw new Error(data.detail || resp.statusText);
    renderClassifierState(data, false);
    appendClassifierLog(data.message || "Bild hoppades över.");
  } catch (e) {
    appendClassifierLog(`Kunde inte hoppa över bild: ${e}`);
  }
}

async function finishClassifierSession() {
  if (!classifierSessionId) {
    appendClassifierLog("Ingen aktiv session att avsluta.");
    return;
  }
  try {
    const resp = await fetch(`${API}/api/classifier/web/${classifierSessionId}/finish`, {
      method: "POST",
    });
    const data = await resp.json().catch(() => ({}));
    if (!resp.ok) throw new Error(data.detail || resp.statusText);
    renderClassifierState(data, true);
    appendClassifierLog(data.message || "Session avslutad.");
  } catch (e) {
    appendClassifierLog(`Kunde inte avsluta session: ${e}`);
  }
}

// ---------------------------------------------------------------------------
// Resultat-knappar
// ---------------------------------------------------------------------------

function activateResultButton(key) {
  availableResults.add(key);
  renderResultButtons();
}

async function refreshResultStatus() {
  try {
    const resp = await fetchWithSessionRecovery(`${API}/api/run/status/${sessionId}`, {}, "hämtning av resultatstatus");
    const data = await resp.json();
    availableResults = new Set(data.results || []);
    renderResultButtons();
  } catch (e) {
    // Tyst fel
  }
}

function createResultButton(key, isReady) {
  const label = RESULT_LABELS[key] || `Öppna ${key}`;
  const btn = document.createElement("button");
  btn.className = "btn btn-sm action-btn action-btn-open";
  btn.textContent = label;
  btn.dataset.resultKey = key;
  btn.setAttribute("data-bs-toggle", "tooltip");
  btn.onclick = () => openResult(key);
  if (isReady) {
    btn.title = `Ladda ner och öppna ${label.replace(/^Öppna\\s+/i, "")} som Excel-fil`;
    setButtonLockState(btn, false, "");
  } else {
    setButtonLockState(btn, true, OPEN_BUTTON_HINTS[key] || "Resultatet är inte tillgängligt än.");
  }
  return btn;
}

function resultContainerForKey(key) {
  if (key === "eftersok") {
    return document.getElementById("result-buttons-eftersok");
  }
  return document.getElementById("result-buttons-allokering");
}

function updateResultContainerVisibility(container) {
  if (!container) return;
  const hasButtons = container.childElementCount > 0;
  const actionGroup = container.closest(".action-group");
  if (actionGroup) {
    actionGroup.style.display = hasButtons ? "" : "none";
    return;
  }
  container.style.display = hasButtons ? "" : "none";
}

function renderResultButtons() {
  const allokeringContainer = document.getElementById("result-buttons-allokering");
  const eftersokContainer = document.getElementById("result-buttons-eftersok");
  if (allokeringContainer) allokeringContainer.innerHTML = "";
  if (eftersokContainer) eftersokContainer.innerHTML = "";

  OPEN_RESULT_ORDER.forEach(key => {
    if (!availableResults.has(key)) {
      return;
    }
    const container = resultContainerForKey(key);
    if (!container) return;
    container.appendChild(createResultButton(key, true));
  });

  // Visar oväntade/extra resultatnycklar om backend returnerar fler.
  Array.from(availableResults)
    .filter(key => !OPEN_RESULT_ORDER.includes(key))
    .forEach(key => {
      const container = resultContainerForKey(key) || allokeringContainer || eftersokContainer;
      if (!container) return;
      container.appendChild(createResultButton(key, true));
    });

  updateResultContainerVisibility(allokeringContainer);
  updateResultContainerVisibility(eftersokContainer);
  initTooltips();
}

function openResult(key) {
  if (!availableResults.has(key)) {
    appendLog(OPEN_BUTTON_HINTS[key] || `Resultat '${key}' är inte tillgängligt än.`);
    return;
  }
  downloadResult(key);
}

function downloadResult(key) {
  window.open(`${API}/api/download/${sessionId}/${key}`, "_blank");
}

// ---------------------------------------------------------------------------
// Summering per Kalltyp
// ---------------------------------------------------------------------------

const KALLTYP_COLORS = {
  HELPALL: "#28a745",
  AUTOSTORE: "#007bff",
  HIB: "#6f42c1",
  HUVUDPLOCK: "#dc3545",
  SKRYMMANDE: "#fd7e14",
  EHANDEL: "#20c997",
};

function renderKalltypSummary(rows) {
  const card = document.getElementById("kalltyp-summary-card");
  const tbody = document.querySelector("#kalltyp-summary-table tbody");
  if (!card || !tbody) return;
  tbody.innerHTML = "";
  if (!Array.isArray(rows) || rows.length === 0) {
    card.style.display = "none";
    return;
  }
  for (const r of rows) {
    const color = KALLTYP_COLORS[r.kalltyp] || "inherit";
    const tr = document.createElement("tr");
    tr.innerHTML = `<td style="color:${color};font-weight:600">${r.kalltyp || ""}</td><td class="text-end">${r.antal_text || ""}</td><td class="text-end">${r.kolli ?? ""}</td>`;
    tbody.appendChild(tr);
  }
  card.style.display = "";
}

function clearKalltypSummary() {
  const card = document.getElementById("kalltyp-summary-card");
  const tbody = document.querySelector("#kalltyp-summary-table tbody");
  if (tbody) tbody.innerHTML = "";
  if (card) card.style.display = "none";
}

// ---------------------------------------------------------------------------
// Kopiera listor
// ---------------------------------------------------------------------------

function setOrdersaldoCopyButtonState(buttonId, count, emptyHint) {
  const btn = document.getElementById(buttonId);
  if (!btn) return;
  const copyConfig = Object.values(COPY_BUTTON_REQUIREMENTS).find(cfg => cfg.buttonId === buttonId) || null;
  const readyTitle = buildRequirementTitle(btn.dataset.baseTitle || btn.title || "", copyConfig?.requiredLabels || []);
  if (count > 0) {
    setButtonLockState(btn, false, "");
    btn.title = readyTitle || btn.title;
  } else {
    setButtonLockState(btn, true, emptyHint || copyConfig?.lockedHint || "Listan är tom.");
  }
}

async function refreshOrdersaldoButtonState() {
  try {
    const [r1, r2] = await Promise.all([
      fetchWithSessionRecovery(`${API}/api/ordersaldo/list1/${sessionId}`, {}, "hämtning av kompletta ordrar"),
      fetchWithSessionRecovery(`${API}/api/ordersaldo/list2/${sessionId}`, {}, "hämtning av påfyllningsbehov"),
    ]);
    const d1 = await r1.json().catch(() => ({}));
    const d2 = await r2.json().catch(() => ({}));
    const c1 = Array.isArray(d1.values) ? d1.values.length : 0;
    const c2 = Array.isArray(d2.values) ? d2.values.length : 0;
    setOrdersaldoCopyButtonState("btn-kompletta", c1, "Lista 1 är tom.");
    setOrdersaldoCopyButtonState("btn-pafyllning", c2, "Lista 2 är tom.");
  } catch (e) {
    setOrdersaldoCopyButtonState("btn-kompletta", 0, "Lista 1 är tom.");
    setOrdersaldoCopyButtonState("btn-pafyllning", 0, "Lista 2 är tom.");
  }
  initTooltips();
}

async function copyList(listId) {
  const copyConfig = COPY_BUTTON_REQUIREMENTS[listId];
  const btn = copyConfig ? document.getElementById(copyConfig.buttonId) : null;
  if (btn && btn.dataset.locked === "1") {
    appendLog(btn.title || "Listan är inte tillgänglig ännu.");
    return;
  }
  try {
    const resp = await fetchWithSessionRecovery(`${API}/api/ordersaldo/${listId}/${sessionId}`, {}, `kopiering av ${listId}`);
    const data = await resp.json();
    const values = data.values || [];
    if (values.length === 0) {
      appendLog(`${listId === "list1" ? "Kompletta ordrar" : "Påfyllningsbehov"}: listan är tom.`);
      return;
    }
    await navigator.clipboard.writeText(values.join("\n"));
    const label = listId === "list1" ? "ordernummer" : "artikelnummer";
    appendLog(`${values.length} ${label} kopierade.`);
  } catch (e) {
    appendLog(`FEL vid kopiering av ${listId}: ${e}`);
  }
}

// ---------------------------------------------------------------------------
// Rensa cache
// ---------------------------------------------------------------------------

async function resetCache() {
  try {
    await fetchWithSessionRecovery(`${API}/api/session/reset-all/${sessionId}`, { method: "POST" }, "rensning av cache");
    resetFrontendState({ clearInputs: true });
    document.getElementById("log-output").textContent = "";
    appendLog("Cache rensad (filer och resultat borttagna).");
    await refreshFilterOptions();
    await refreshOrdersaldoButtonState();
    syncActionButtonsState();
  } catch (e) {
    appendLog(`FEL vid rensning: ${e}`);
  }
}

// ---------------------------------------------------------------------------
// Chunked Excel
// ---------------------------------------------------------------------------

async function openChunkedExcel() {
  const rawText = document.getElementById("paste-values").value.trim();
  const values = rawText.split("\n").map(v => v.trim()).filter(v => v);
  if (values.length === 0) {
    appendLog("Inga värden att öppna.");
    return;
  }
  const chunkSize = parseInt(document.getElementById("chunk-size").value) || 2000;
  appendLog(`Skapar Excel med ${values.length} värden, ${chunkSize} per kolumn...`);
  try {
    const resp = await fetch(`${API}/api/chunked-excel`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ values, chunk_size: chunkSize }),
    });
    if (!resp.ok) throw new Error(await resp.text());
    const blob = await resp.blob();
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "chunked_values.xlsx";
    a.click();
    URL.revokeObjectURL(url);
    appendLog("Excel skapad och nedladdad.");
  } catch (e) {
    appendLog(`FEL vid skapande av Excel: ${e}`);
  }
}

// ---------------------------------------------------------------------------
// Kolumnknappar for Dela-sektionen
// ---------------------------------------------------------------------------

const MAX_CHUNK_COL_BUTTONS = 10;

function updateChunkColButtons() {
  const container = document.getElementById("chunk-col-buttons");
  if (!container) return;

  const rawText = document.getElementById("paste-values").value.trim();
  const values = rawText.split("\n").map(v => v.trim()).filter(v => v);
  const chunkSize = parseInt(document.getElementById("chunk-size").value) || 2000;
  const numCols = Math.ceil(values.length / chunkSize);

  container.innerHTML = "";

  if (numCols < 2) {
    container.style.display = "none";
    return;
  }

  container.style.display = "";
  const visible = Math.min(numCols, MAX_CHUNK_COL_BUTTONS);
  for (let i = 0; i < visible; i++) {
    const btn = document.createElement("button");
    btn.className = "btn btn-warning btn-sm";
    btn.textContent = `Kopiera ${i + 1}`;
    btn.title = `Kopiera alla värden från kolumn ${i + 1} till urklipp`;
    btn.setAttribute("data-bs-toggle", "tooltip");
    btn.dataset.colIndex = i;
    btn.addEventListener("click", () => copyChunkColumn(i));
    container.appendChild(btn);
  }
  initTooltips();
}

async function copyChunkColumn(colIndex) {
  const rawText = document.getElementById("paste-values").value.trim();
  const values = rawText.split("\n").map(v => v.trim()).filter(v => v);
  const chunkSize = parseInt(document.getElementById("chunk-size").value) || 2000;
  const start = colIndex * chunkSize;
  const chunk = values.slice(start, start + chunkSize);
  if (chunk.length === 0) {
    appendLog(`Kolumn ${colIndex + 1} är tom.`);
    return;
  }
  await navigator.clipboard.writeText(chunk.join("\n"));
  appendLog(`${chunk.length} värden från kolumn ${colIndex + 1} kopierade.`);
}

// Lyssna pa andringar i textarea och chunk-size
window.addEventListener("DOMContentLoaded", () => {
  const textarea = document.getElementById("paste-values");
  const chunkInput = document.getElementById("chunk-size");
  if (textarea) textarea.addEventListener("input", updateChunkColButtons);
  if (chunkInput) chunkInput.addEventListener("input", updateChunkColButtons);
});

// ---------------------------------------------------------------------------
// Bootstrap-tooltips
// ---------------------------------------------------------------------------

let _tooltipInstances = [];

function initTooltips() {
  // Förstör gamla instanser
  _tooltipInstances.forEach(t => { try { t.dispose(); } catch (e) { /* ignoreras */ } });
  _tooltipInstances = [];
  // Skapa nya för alla element med data-bs-toggle="tooltip"
  document.querySelectorAll('[data-bs-toggle="tooltip"]').forEach(el => {
    const instance = new bootstrap.Tooltip(el, { trigger: "hover", placement: "top" });
    _tooltipInstances.push(instance);
  });
}

// ---------------------------------------------------------------------------
// Drag & Drop
// ---------------------------------------------------------------------------

const ALL_SLOTS = [...FILE_SLOTS, ...PROG_SLOTS, ...WMS_SLOTS];

function setupDragDrop() {
  const body = document.body;

  body.addEventListener("dragover", (e) => {
    e.preventDefault();
    body.classList.add("drag-over");
  });

  body.addEventListener("dragleave", (e) => {
    if (!e.relatedTarget || e.relatedTarget === document.documentElement) {
      body.classList.remove("drag-over");
    }
  });

  body.addEventListener("drop", async (e) => {
    e.preventDefault();
    body.classList.remove("drag-over");
    const files = Array.from(e.dataTransfer.files);
    const changedKeys = new Set();
    for (const file of files) {
      const matched = matchFileToSlot(file.name);
      if (matched) {
        appendLog(`Drag & drop: ${file.name} -> [${matched}]`);
        const uploadedKey = await uploadFile(matched, file, { skipPostRefresh: true });
        if (uploadedKey) {
          changedKeys.add(uploadedKey);
        }
      } else {
        appendLog(`Kunde inte matcha "${file.name}" till en fil-slot. Ladda upp manuellt.`);
      }
    }
    if (changedKeys.size > 0) {
      await refreshAfterFileChanges(changedKeys);
    }
  });
}

function matchFileToSlot(filename) {
  const lower = filename.toLowerCase();
  const activePage = getActiveMainPage();

  // Exakta prefix-matchningar (speglar allokera11.py _detect_file_type)
  const exactMap = {
    "wms_receive": "v_ask_receive_log",
    "wms_booking": "v_ask_booking_putaway",
    "wms_trans":   "v_ask_trans_log",
    "wms_pick":    "v_ask_pick_log_full",
    "wms_correct": "v_ask_correct_log",
  };
  if (lower.includes("v_ask_article_buffertpallet")) {
    return activePage === "eftersok-page" ? "wms_buffer" : "buffer";
  }
  for (const [slot, hint] of Object.entries(exactMap)) {
    if (lower.includes(hint)) return slot;
  }

  // Generiska matchningar
  if (lower.includes("kampanj") && (lower.endsWith(".xlsx") || lower.endsWith(".xls"))) return "campaign";
  if (lower.includes("prognos") && (lower.endsWith(".xlsx") || lower.endsWith(".xls"))) return "prognos";
  if (lower.includes("bestall") || lower.includes("best\u00e4ll") || (lower.includes("order") && lower.includes("detail"))) return "orders";
  if (lower.includes("buffert") || lower.includes("buffer")) {
    return activePage === "eftersok-page" ? "wms_buffer" : "buffer";
  }
  if (lower.includes("saldo") || lower.includes("automation")) return "automation";
  if (lower.includes("item") || lower.includes("artikel_option")) return "item";
  if (lower.includes("overview") || lower.includes("oversikt") || lower.includes("\u00f6versikt") || lower.includes("order_overview")) return "overview";
  if (lower.includes("dispatch")) return "dispatch";
  return null;
}
