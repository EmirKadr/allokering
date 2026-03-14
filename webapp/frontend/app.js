/**
 * app.js – Allokering WebApp frontend
 */

const API = "http://localhost:8000";

// ---------------------------------------------------------------------------
// Fil-konfiguration
// ---------------------------------------------------------------------------

const FILE_SLOTS = [
  { key: "orders",     label: "Beställningslinjer (CSV)",   accept: ".csv,.txt" },
  { key: "buffer",     label: "Buffertpallar (CSV)",         accept: ".csv,.txt" },
  { key: "automation", label: "Saldo inkl. automation (CSV)", accept: ".csv,.txt" },
  { key: "item",       label: "Item option (CSV)",           accept: ".csv,.txt" },
  { key: "overview",   label: "Orderöversikt (CSV)",         accept: ".csv,.txt" },
  { key: "dispatch",   label: "Dispatchpallar (CSV)",        accept: ".csv,.txt" },
];

const PROG_SLOTS = [
  { key: "prognos",  label: "Prognos (XLSX)",            accept: ".xlsx,.xls" },
  { key: "campaign", label: "Kampanjvolymer (XLSX)",      accept: ".xlsx,.xls" },
];

const WMS_SLOTS = [
  { key: "wms_receive", label: "Mottagningslogg (CSV)",   accept: ".csv,.txt" },
  { key: "wms_booking", label: "Ej inlagrade (CSV)",       accept: ".csv,.txt" },
  { key: "wms_trans",   label: "Translogg (CSV)",          accept: ".csv,.txt" },
  { key: "wms_pick",    label: "Plocklogg (CSV)",          accept: ".csv,.txt" },
  { key: "wms_correct", label: "Saldojustering (CSV)",     accept: ".csv,.txt" },
];

// Resultatnycklar → visningsnamn
const RESULT_LABELS = {
  "allokerade":       "Öppna allokerade",
  "nearmiss":         "Öppna near-miss",
  "pallplatser":      "Öppna pallplatser",
  "refill":           "Öppna påfyllning",
  "hib-koppling":     "Öppna HIB-koppling",
  "orderkontroll":    "Öppna orderkontroll",
  "dispatchkontroll": "Öppna dispatchkontroll",
  "eftersok":         "Öppna eftersök",
  "prognos":          "Öppna prognos vs autoplock",
  "sales":            "Öppna försäljningsinsikter",
};

// ---------------------------------------------------------------------------
// State
// ---------------------------------------------------------------------------

let sessionId = null;
let sseSource = null;
let fileStatuses = {};  // key -> filename | null
let availableResults = new Set();
let cachedFilterOptions = {};  // { bolag: [...], ordertyp: [...] }
let selectedFilters = { bolag: [], ordertyp: [] };

// ---------------------------------------------------------------------------
// Init
// ---------------------------------------------------------------------------

window.addEventListener("DOMContentLoaded", async () => {
  sessionId = sessionStorage.getItem("allok_session_id");
  if (!sessionId) {
    const resp = await fetch(`${API}/api/session`, { method: "POST" });
    const data = await resp.json();
    sessionId = data.session_id;
    sessionStorage.setItem("allok_session_id", sessionId);
  }

  renderFileRows();
  renderProgRows();
  renderWmsRows();
  connectSSE();
  setupDragDrop();

  // Hämta befintlig filstatus
  await refreshFileStatus();
  await refreshFilterOptions();
  await refreshResultStatus();
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
  row.style.minWidth = "360px";

  const label = document.createElement("span");
  label.textContent = slot.label;
  label.style.minWidth = "220px";
  label.style.fontSize = "13px";

  const badge = document.createElement("span");
  badge.className = "file-status-badge";
  badge.id = `badge-${slot.key}`;
  badge.textContent = "Ej fil";
  badge.title = "Klicka för att välja fil";
  badge.addEventListener("click", () => triggerFilePicker(slot.key, slot.accept));

  const removeBtn = document.createElement("button");
  removeBtn.className = "btn btn-danger btn-sm py-0 px-1";
  removeBtn.innerHTML = "&#10005;";
  removeBtn.title = "Ta bort fil";
  removeBtn.addEventListener("click", () => removeFile(slot.key));

  // Hidden file input
  const input = document.createElement("input");
  input.type = "file";
  input.accept = slot.accept || "";
  input.style.display = "none";
  input.id = `input-${slot.key}`;
  input.addEventListener("change", async (e) => {
    const file = e.target.files[0];
    if (file) {
      await uploadFile(slot.key, file);
    }
    input.value = "";
  });

  row.appendChild(label);
  row.appendChild(badge);
  row.appendChild(removeBtn);
  row.appendChild(input);
  return row;
}

function triggerFilePicker(key, accept) {
  const input = document.getElementById(`input-${key}`);
  if (input) input.click();
}

// ---------------------------------------------------------------------------
// Upload / Remove
// ---------------------------------------------------------------------------

async function uploadFile(fileKey, file) {
  setBadge(fileKey, "Laddar...", false);
  try {
    const formData = new FormData();
    formData.append("file_key", fileKey);
    formData.append("file", file);
    const resp = await fetch(`${API}/api/upload/${sessionId}`, {
      method: "POST",
      body: formData,
    });
    if (!resp.ok) throw new Error(await resp.text());
    const data = await resp.json();
    fileStatuses[fileKey] = data.filename;
    setBadge(fileKey, data.filename, true);
    appendLog(`Fil uppladdad: [${fileKey}] ${data.filename}`);
    await refreshFilterOptions();
  } catch (e) {
    setBadge(fileKey, "FEL", false);
    appendLog(`Fel vid uppladdning av ${fileKey}: ${e}`);
  }
}

async function removeFile(fileKey) {
  try {
    await fetch(`${API}/api/upload/${sessionId}/${fileKey}`, { method: "DELETE" });
    fileStatuses[fileKey] = null;
    setBadge(fileKey, "Ej fil", false);
    appendLog(`Fil borttagen: [${fileKey}]`);
    await refreshFilterOptions();
  } catch (e) {
    appendLog(`Fel vid borttagning av ${fileKey}: ${e}`);
  }
}

async function refreshFileStatus() {
  try {
    const resp = await fetch(`${API}/api/upload/${sessionId}`);
    const data = await resp.json();
    for (const [key, filename] of Object.entries(data.files || {})) {
      fileStatuses[key] = filename;
      setBadge(key, filename || "Ej fil", !!filename);
    }
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
// Filter
// ---------------------------------------------------------------------------

async function refreshFilterOptions() {
  try {
    const resp = await fetch(`${API}/api/filters/${sessionId}`);
    const data = await resp.json();
    cachedFilterOptions = data;
    renderFilterCard(data);
  } catch (e) {
    // Tyst fel
  }
}

function renderFilterCard(options) {
  const card = document.getElementById("filter-card");
  const content = document.getElementById("filter-content");
  const hasBolag = options.bolag && options.bolag.length > 0;
  const hasOrdertyp = options.ordertyp && options.ordertyp.length > 0;

  if (!hasBolag && !hasOrdertyp) {
    card.style.display = "none";
    return;
  }
  card.style.display = "";
  content.innerHTML = "";

  const row = document.createElement("div");
  row.className = "d-flex gap-3 flex-wrap";

  if (hasBolag) {
    row.appendChild(buildFilterGroup("bolag", "Bolag", options.bolag));
  }
  if (hasOrdertyp) {
    row.appendChild(buildFilterGroup("ordertyp", "Ordertyp", options.ordertyp));
  }
  content.appendChild(row);

  // "Välj alla" / "Rensa" knappar
  const btnRow = document.createElement("div");
  btnRow.className = "d-flex gap-2 mt-2";
  const selectAll = document.createElement("button");
  selectAll.className = "btn btn-sm btn-outline-secondary";
  selectAll.textContent = "Välj alla";
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

function saveFilters() {
  const filters = { bolag: [], ordertyp: [] };
  document.querySelectorAll(".filter-check").forEach(cb => {
    if (cb.checked) {
      const group = cb.dataset.group;
      if (filters[group] !== undefined) {
        filters[group].push(cb.value);
      }
    }
  });
  selectedFilters = filters;
  fetch(`${API}/api/filters/${sessionId}`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(filters),
  }).catch(() => {});
  appendLog(`Filter uppdaterat: bolag=[${filters.bolag.join(",")}], ordertyp=[${filters.ordertyp.join(",")}]`);
}

// ---------------------------------------------------------------------------
// Körning
// ---------------------------------------------------------------------------

async function runJob(job) {
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

    const resp = await fetch(url, { method, body, headers });
    if (!resp.ok) {
      const errData = await resp.json().catch(() => ({}));
      throw new Error(errData.detail || resp.statusText);
    }
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
  }
}

// ---------------------------------------------------------------------------
// SSE-logg
// ---------------------------------------------------------------------------

let currentJob = null;

function connectSSE() {
  if (sseSource) {
    sseSource.close();
  }
  sseSource = new EventSource(`${API}/api/log/stream/${sessionId}`);
  sseSource.onmessage = (e) => {
    const msg = e.data.replace(/\\n/g, "\n");

    // Resultat-events
    if (msg.startsWith("__RESULT:")) {
      const key = msg.replace("__RESULT:", "").trim();
      activateResultButton(key);
      return;
    }

    // Klar/Fel
    if (msg === "__DONE__" || msg === "__ERROR__") {
      // Återställ alla körningsknappar
      ["allokering", "hib-koppling", "orderkontroll", "dispatchkontroll", "eftersok"].forEach(resetJobButton);
      appendLog(msg === "__DONE__" ? "--- Klar ---" : "--- Avbröts med fel ---");
      refreshResultStatus();
      return;
    }

    appendLog(msg);
  };
  sseSource.onerror = () => {
    // Försök återansluta efter 3 sekunder
    setTimeout(connectSSE, 3000);
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

// ---------------------------------------------------------------------------
// Resultat-knappar
// ---------------------------------------------------------------------------

function activateResultButton(key) {
  availableResults.add(key);
  renderResultButtons();
}

async function refreshResultStatus() {
  try {
    const resp = await fetch(`${API}/api/run/status/${sessionId}`);
    const data = await resp.json();
    (data.results || []).forEach(k => availableResults.add(k));
    renderResultButtons();
  } catch (e) {
    // Tyst fel
  }
}

function renderResultButtons() {
  const container = document.getElementById("result-buttons");
  container.innerHTML = "";
  availableResults.forEach(key => {
    const label = RESULT_LABELS[key] || `Öppna ${key}`;
    const btn = document.createElement("button");
    btn.className = "btn btn-success btn-sm";
    btn.textContent = label;
    btn.onclick = () => downloadResult(key);
    container.appendChild(btn);
  });
}

function downloadResult(key) {
  window.open(`${API}/api/download/${sessionId}/${key}`, "_blank");
}

// ---------------------------------------------------------------------------
// Kopiera listor
// ---------------------------------------------------------------------------

function copyList(listId) {
  appendLog(`Kopiering av ${listId} är inte implementerat i webbversionen.`);
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
    for (const file of files) {
      const matched = matchFileToSlot(file.name);
      if (matched) {
        appendLog(`Drag & drop: ${file.name} -> [${matched}]`);
        await uploadFile(matched, file);
      } else {
        appendLog(`Kunde inte matcha "${file.name}" till en fil-slot. Ladda upp manuellt.`);
      }
    }
  });
}

function matchFileToSlot(filename) {
  const lower = filename.toLowerCase();
  // Försök matcha på förväntade filnamn
  const wmsMap = {
    "wms_receive":  ["receive", "mottagning", "v_ask_receive"],
    "wms_booking":  ["booking", "putaway", "inlagrade", "v_ask_booking"],
    "wms_buffert":  ["buffert", "buffer", "v_ask_article_buffert"],
    "wms_trans":    ["trans_log", "translogg", "v_ask_trans"],
    "wms_pick":     ["pick_log", "plocklogg", "v_ask_pick"],
    "wms_correct":  ["correct", "saldojust", "v_ask_correct"],
  };
  for (const [key, patterns] of Object.entries(wmsMap)) {
    if (patterns.some(p => lower.includes(p))) return key;
  }
  if (lower.includes("prognos") && (lower.endsWith(".xlsx") || lower.endsWith(".xls"))) return "prognos";
  if (lower.includes("kampanj") && (lower.endsWith(".xlsx") || lower.endsWith(".xls"))) return "campaign";
  if (lower.includes("bestall") || lower.includes("beställ") || lower.includes("order") && lower.includes("detail")) return "orders";
  if (lower.includes("buffert") || lower.includes("buffer")) return "buffer";
  if (lower.includes("saldo") || lower.includes("automation")) return "automation";
  if (lower.includes("item") || lower.includes("artikel_option")) return "item";
  if (lower.includes("overview") || lower.includes("översikt")) return "overview";
  if (lower.includes("dispatch")) return "dispatch";
  return null;
}
