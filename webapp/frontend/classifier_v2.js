/* classifier_v2.js
 * Chromium/web v2 flow for Artikelklassificering.
 */

(function () {
  const API_BASE = typeof API !== "undefined" ? API : "http://localhost:8000";
  const V2 = `${API_BASE}/api/classifier/v2`;

  const CLS2_DATA_SLOTS = [
    { key: "item_attribute", label: "item_attribute" },
    { key: "item_alias", label: "item_alias" },
    { key: "item", label: "item" },
    { key: "main_category", label: "main_category" },
  ];

  const state = {
    mode: localStorage.getItem("classifier_mode") || "v2",
    sessionId: null,
    latest: null,
    eventSource: null,
    pollTimer: null,
    selectedCards: new Set(),
    keymap: [],
  };

  function $(id) {
    return document.getElementById(id);
  }

  function isClassifierPageVisible() {
    const pane = $("classifier-page");
    return !!pane && pane.classList.contains("active");
  }

  function log(msg) {
    const el = $("cls2-log-output");
    if (!el) return;
    const ts = new Date().toLocaleTimeString("sv-SE");
    el.textContent += `[${ts}] ${msg}\n`;
    el.scrollTop = el.scrollHeight;
  }

  function setMode(mode) {
    state.mode = mode === "v1" ? "v1" : "v2";
    localStorage.setItem("classifier_mode", state.mode);
    const v1 = $("classifier-v1-wrap");
    const v2 = $("classifier-v2-wrap");
    const b1 = $("btn-classifier-mode-v1");
    const b2 = $("btn-classifier-mode-v2");
    if (v1) v1.style.display = state.mode === "v1" ? "" : "none";
    if (v2) v2.style.display = state.mode === "v2" ? "" : "none";
    if (b1) b1.className = state.mode === "v1" ? "btn btn-secondary btn-sm" : "btn btn-outline-secondary btn-sm";
    if (b2) b2.className = state.mode === "v2" ? "btn btn-primary btn-sm" : "btn btn-outline-primary btn-sm";
  }

  function setStatus(kind, text) {
    const badge = $("cls2-status-badge");
    if (!badge) return;
    if (kind === "running") badge.className = "badge text-bg-success";
    else if (kind === "done") badge.className = "badge text-bg-primary";
    else if (kind === "error") badge.className = "badge text-bg-danger";
    else badge.className = "badge text-bg-secondary";
    badge.textContent = text;
    const txt = $("cls2-status-text");
    if (txt) txt.textContent = text || "";
  }

  async function fetchJson(url, opts = {}) {
    const resp = await fetch(url, opts);
    const data = await resp.json().catch(() => ({}));
    if (!resp.ok) {
      throw new Error(data.detail || resp.statusText || "API-fel");
    }
    return data;
  }

  function parseCategories() {
    const raw = ($("cls2-categories")?.value || "").trim();
    const out = [];
    const seen = new Set();
    raw.split("\n").map(s => s.trim()).filter(Boolean).forEach((line) => {
      const parts = line.split("|");
      const name = (parts[0] || "").trim();
      const description = (parts.slice(1).join("|") || "").trim();
      if (!name) return;
      const key = name.toLowerCase();
      if (seen.has(key)) return;
      seen.add(key);
      out.push({ name, description });
    });
    return out;
  }

  function renderDataBadge(key, loaded, label) {
    const badge = $(`cls2-data-badge-${key}`);
    if (!badge) return;
    const text = label || "Ej fil";
    const display = text.length > 20 ? `...${text.slice(-17)}` : text;
    badge.textContent = display;
    badge.title = text;
    badge.classList.toggle("loaded", !!loaded);
  }

  function buildDataRow(slot) {
    const row = document.createElement("div");
    row.className = "d-flex align-items-center gap-2 mb-1 classifier-data-row";

    const label = document.createElement("span");
    label.style.minWidth = "105px";
    label.style.fontSize = "12px";
    label.textContent = slot.label;

    const badge = document.createElement("span");
    badge.className = "file-status-badge";
    badge.id = `cls2-data-badge-${slot.key}`;
    badge.textContent = "Ej fil";
    badge.addEventListener("click", () => {
      const inp = $(`cls2-data-input-${slot.key}`);
      if (inp) inp.click();
    });

    const up = document.createElement("button");
    up.className = "btn btn-outline-primary btn-sm py-0 px-2";
    up.textContent = "Ladda upp";
    up.addEventListener("click", () => {
      const inp = $(`cls2-data-input-${slot.key}`);
      if (inp) inp.click();
    });

    const rm = document.createElement("button");
    rm.className = "btn btn-danger btn-sm py-0 px-1";
    rm.innerHTML = "&#10005;";
    rm.addEventListener("click", async () => {
      try {
        await fetchJson(`${V2}/data-files/${slot.key}`, { method: "DELETE" });
        renderDataBadge(slot.key, false, "Ej fil");
        log(`${slot.key} borttagen.`);
      } catch (e) {
        log(`Kunde inte ta bort ${slot.key}: ${e}`);
      }
    });

    const input = document.createElement("input");
    input.type = "file";
    input.accept = ".csv,.txt";
    input.style.display = "none";
    input.id = `cls2-data-input-${slot.key}`;
    input.addEventListener("change", async (ev) => {
      const file = ev.target.files?.[0];
      if (!file) return;
      try {
        const fd = new FormData();
        fd.append("file_key", slot.key);
        fd.append("file", file);
        const resp = await fetch(`${V2}/data-files/upload`, { method: "POST", body: fd });
        const data = await resp.json().catch(() => ({}));
        if (!resp.ok) throw new Error(data.detail || resp.statusText);
        const f = (data.files || {})[slot.key];
        renderDataBadge(slot.key, !!(f && f.exists), f?.filename || file.name);
        log(`${slot.key} uppladdad.`);
      } catch (e) {
        log(`Uppladdning misslyckades (${slot.key}): ${e}`);
      } finally {
        input.value = "";
      }
    });

    row.appendChild(label);
    row.appendChild(badge);
    row.appendChild(up);
    row.appendChild(rm);
    row.appendChild(input);
    return row;
  }

  async function refreshDataFiles() {
    try {
      const data = await fetchJson(`${V2}/data-files`);
      const files = data.files || {};
      CLS2_DATA_SLOTS.forEach((slot) => {
        const f = files[slot.key];
        renderDataBadge(slot.key, !!(f && f.exists), f?.filename || "Ej fil");
      });
      if (typeof data.item_attribute_rows === "number") {
        log(`item_attribute IMG-rader: ${data.item_attribute_rows}`);
      }
    } catch (e) {
      log(`Kunde inte hämta datafiler: ${e}`);
    }
  }

  function ensureDefaultCategories(value) {
    if (!value || value.trim()) return value;
    return "Hund\nKatt\nÖvrigt";
  }

  async function initConfig() {
    try {
      const cfg = await fetchJson(`${V2}/config`);
      if ($("cls2-test-name") && !$("cls2-test-name").value) $("cls2-test-name").value = cfg.default_test_name || "test1";
      if ($("cls2-syfte") && !$("cls2-syfte").value) $("cls2-syfte").value = cfg.default_syfte || "";
      if ($("cls2-image-dir") && !$("cls2-image-dir").value) $("cls2-image-dir").value = cfg.default_image_dir || "";
      if ($("cls2-output-dir") && !$("cls2-output-dir").value) $("cls2-output-dir").value = cfg.default_output_dir || "";
      if ($("cls2-ai-url") && !$("cls2-ai-url").value) $("cls2-ai-url").value = cfg.ai_defaults?.api_url || "http://localhost:1234/v1";
      if ($("cls2-categories")) {
        const current = $("cls2-categories").value || "";
        if (!current.trim()) {
          $("cls2-categories").value = ensureDefaultCategories(
            (cfg.default_categories || []).map(c => `${c.name}${c.description ? ` | ${c.description}` : ""}`).join("\n")
          );
        }
      }
      await refreshDataFiles();
      log("Classifier v2 initierad.");
    } catch (e) {
      setStatus("error", "Kunde inte ladda v2-config");
      log(`Init-fel: ${e}`);
    }
  }

  function connectEvents() {
    if (!state.sessionId) return;
    if (state.eventSource) {
      try { state.eventSource.close(); } catch (_) {}
      state.eventSource = null;
    }
    const src = new EventSource(`${V2}/session/${state.sessionId}/events`);
    src.onmessage = (ev) => {
      try {
        const data = JSON.parse(ev.data || "{}");
        if (data.type === "log" && data.message) log(data.message.replace(/^\[\d\d:\d\d:\d\d\]\s*/, ""));
      } catch (_) {}
    };
    src.onerror = () => {
      try { src.close(); } catch (_) {}
      state.eventSource = null;
      setTimeout(() => {
        if (state.sessionId && state.mode === "v2") connectEvents();
      }, 2000);
    };
    state.eventSource = src;
  }

  function renderCurrentImage(sessionId, current) {
    const img = $("cls2-current-image");
    if (!img) return;
    if (sessionId && current && Object.keys(current).length > 0) {
      img.src = `${V2}/session/${sessionId}/image?ts=${Date.now()}`;
      img.style.display = "block";
    } else {
      img.removeAttribute("src");
      img.style.display = "none";
    }
  }

  function clearCardSelection() {
    state.selectedCards.clear();
    document.querySelectorAll(".cls2-card.selected").forEach((el) => el.classList.remove("selected"));
  }

  async function reclassifyArticles(articleNumbers, targetCategory, reason = "") {
    if (!state.sessionId) return;
    if (!articleNumbers.length) return;
    try {
      await fetchJson(`${V2}/session/${state.sessionId}/reclassify`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          article_numbers: articleNumbers,
          target_category: targetCategory,
          reason,
        }),
      });
      await refreshState(false);
    } catch (e) {
      log(`Omklassificering misslyckades: ${e}`);
    }
  }

  function renderKanban(kanban) {
    const host = $("cls2-kanban");
    if (!host) return;
    host.innerHTML = "";
    const cols = kanban?.columns || [];
    cols.forEach((col) => {
      const colEl = document.createElement("div");
      colEl.className = "cls2-kanban-col";
      colEl.dataset.category = col.category;

      const header = document.createElement("div");
      header.className = "cls2-kanban-col-header";
      header.textContent = `${col.category} (${(col.cards || []).length})`;
      if (col.knowledge_ready) {
        const dot = document.createElement("span");
        dot.className = "cls2-knowledge-dot";
        header.appendChild(dot);
      }
      header.addEventListener("click", async () => {
        const current = state.latest?.cat_knowledge?.[col.category] || "";
        const next = window.prompt(`AI-analys för ${col.category}`, current);
        if (next === null) return;
        await callAction("knowledge_update", { category: col.category, knowledge: next });
      });
      colEl.appendChild(header);

      const list = document.createElement("div");
      list.className = "cls2-kanban-col-body";
      list.addEventListener("dragover", (ev) => ev.preventDefault());
      list.addEventListener("drop", async (ev) => {
        ev.preventDefault();
        const dragged = ev.dataTransfer.getData("text/plain");
        const selected = state.selectedCards.size > 0 ? Array.from(state.selectedCards) : (dragged ? [dragged] : []);
        if (!selected.length) return;
        await reclassifyArticles(selected, col.category, "Flyttad i kanban");
        clearCardSelection();
      });

      (col.cards || []).forEach((card) => {
        const cardEl = document.createElement("div");
        cardEl.className = "cls2-card";
        cardEl.draggable = true;
        cardEl.dataset.article = card.article_number;
        cardEl.innerHTML = `<div class="fw-semibold">${card.article_number}</div><div class="small text-muted">${card.reason || ""}</div>`;
        cardEl.addEventListener("dragstart", (ev) => {
          ev.dataTransfer.setData("text/plain", card.article_number);
        });
        cardEl.addEventListener("click", (ev) => {
          if (ev.ctrlKey || ev.metaKey) {
            if (state.selectedCards.has(card.article_number)) {
              state.selectedCards.delete(card.article_number);
              cardEl.classList.remove("selected");
            } else {
              state.selectedCards.add(card.article_number);
              cardEl.classList.add("selected");
            }
          } else {
            clearCardSelection();
            state.selectedCards.add(card.article_number);
            cardEl.classList.add("selected");
          }
        });
        cardEl.addEventListener("contextmenu", async (ev) => {
          ev.preventDefault();
          const target = window.prompt("Flytta till kategori (skriv exakt namn):", col.category);
          if (!target) return;
          const reason = window.prompt("Orsak (valfritt):", "Omklassificering");
          const selected = state.selectedCards.size > 0 ? Array.from(state.selectedCards) : [card.article_number];
          await reclassifyArticles(selected, target, reason || "");
          clearCardSelection();
        });
        list.appendChild(cardEl);
      });
      colEl.appendChild(list);
      host.appendChild(colEl);
    });
  }

  function buildCategoryButtons(stateData) {
    const wrap = $("cls2-category-buttons");
    if (!wrap) return;
    wrap.innerHTML = "";
    state.keymap = [];
    const cats = stateData?.categories || [];
    cats.forEach((c, idx) => {
      const catName = c.name || c;
      const hotkey = idx < 9 ? `${idx + 1}` : (idx === 9 ? "0" : "");
      const btn = document.createElement("button");
      btn.className = "btn btn-success btn-sm";
      btn.textContent = hotkey ? `${catName} (${hotkey})` : catName;
      btn.addEventListener("click", () => classify(catName));
      wrap.appendChild(btn);
      state.keymap.push({ hotkey, category: catName });
    });
  }

  function renderState(data) {
    state.latest = data;
    if (!data) {
      setStatus("idle", "Ej startad");
      return;
    }
    const sid = data.session_id || state.sessionId;
    if (sid) state.sessionId = sid;
    if ($("cls2-session-id")) $("cls2-session-id").textContent = state.sessionId || "-";
    if ($("cls2-step")) $("cls2-step").textContent = data.ai?.step || "idle";
    if ($("cls2-current-article")) {
      $("cls2-current-article").textContent = data.current?.article_number || "-";
    }

    const summary = `${data.classified || 0}/${data.total || 0} klassificerade, oklassificerade: ${data.unclassified || 0}, skip: ${data.skipped_count || 0}`;
    if (data.done) setStatus("done", `Klar - ${summary}`);
    else if (data.ai?.running) setStatus("running", `AI kör - ${summary}`);
    else setStatus("running", summary);

    renderCurrentImage(state.sessionId, data.current || {});
    buildCategoryButtons(data);
    renderKanban(data.kanban || { columns: [] });
  }

  async function refreshState(announce = false) {
    if (!state.sessionId) return;
    try {
      const data = await fetchJson(`${V2}/session/${state.sessionId}`);
      renderState(data);
      if (announce) log("Session uppdaterad.");
    } catch (e) {
      log(`Statusfel: ${e}`);
    }
  }

  async function startSession() {
    try {
      const payload = {
        test_name: ($("cls2-test-name")?.value || "").trim(),
        syfte: ($("cls2-syfte")?.value || "").trim(),
        categories: parseCategories(),
        image_dir: ($("cls2-image-dir")?.value || "").trim(),
        output_dir: ($("cls2-output-dir")?.value || "").trim(),
        shuffle: !!$("cls2-shuffle")?.checked,
        source_mode: ($("cls2-source-mode")?.value || "auto").trim(),
        ai_settings: {
          api_url: ($("cls2-ai-url")?.value || "").trim(),
          model: ($("cls2-ai-model")?.value || "").trim(),
          api_key: ($("cls2-ai-key")?.value || "").trim(),
          compress_images: true,
        },
      };
      if (!payload.test_name) throw new Error("Testnamn saknas.");
      if (!payload.categories.length) throw new Error("Minst en kategori krävs.");
      const data = await fetchJson(`${V2}/session/start`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
      });
      state.sessionId = data.session_id;
      renderState(data);
      connectEvents();
      log("Session startad.");
    } catch (e) {
      setStatus("error", `${e}`);
      log(`Kunde inte starta session: ${e}`);
    }
  }

  async function classify(category) {
    if (!state.sessionId) return;
    try {
      const data = await fetchJson(`${V2}/session/${state.sessionId}/classify`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ category }),
      });
      renderState(data);
    } catch (e) {
      log(`Klassificering misslyckades: ${e}`);
    }
  }

  async function skip() {
    if (!state.sessionId) return;
    try {
      const data = await fetchJson(`${V2}/session/${state.sessionId}/skip`, { method: "POST" });
      renderState(data);
    } catch (e) {
      log(`Skip misslyckades: ${e}`);
    }
  }

  async function back() {
    if (!state.sessionId) return;
    try {
      const data = await fetchJson(`${V2}/session/${state.sessionId}/back`, { method: "POST" });
      renderState(data);
    } catch (e) {
      log(`Back misslyckades: ${e}`);
    }
  }

  async function callAction(action, payload = {}) {
    if (!state.sessionId) return;
    const data = await fetchJson(`${V2}/session/${state.sessionId}/action`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ action, payload }),
    });
    renderState(data);
  }

  async function importSession(kind, file) {
    if (!file) return;
    const fd = new FormData();
    fd.append("file", file);
    const endpoint = kind === "zip" ? `${V2}/session/import/zip` : `${V2}/session/import/excel`;
    try {
      const resp = await fetch(endpoint, { method: "POST", body: fd });
      const data = await resp.json().catch(() => ({}));
      if (!resp.ok) throw new Error(data.detail || resp.statusText);
      state.sessionId = data.session_id;
      renderState(data);
      connectEvents();
      log(`${kind.toUpperCase()}-session importerad.`);
    } catch (e) {
      log(`Import ${kind} misslyckades: ${e}`);
    }
  }

  function wireButtons() {
    $("btn-classifier-mode-v2")?.addEventListener("click", () => setMode("v2"));
    $("btn-classifier-mode-v1")?.addEventListener("click", () => setMode("v1"));
    $("btn-cls2-start")?.addEventListener("click", startSession);
    $("btn-cls2-refresh")?.addEventListener("click", () => refreshState(true));
    $("btn-cls2-skip")?.addEventListener("click", skip);
    $("btn-cls2-back")?.addEventListener("click", back);
    $("btn-cls2-finish")?.addEventListener("click", () => callAction("finish"));

    $("btn-cls2-ai-start")?.addEventListener("click", () => callAction("ai_start"));
    $("btn-cls2-ai-start-step2")?.addEventListener("click", () => callAction("ai_start_step2"));
    $("btn-cls2-ai-pause")?.addEventListener("click", () => callAction("ai_pause"));
    $("btn-cls2-ai-resume")?.addEventListener("click", () => callAction("ai_resume"));
    $("btn-cls2-ai-analyze-all")?.addEventListener("click", () => callAction("ai_analyze_all"));
    $("btn-cls2-ai-stop")?.addEventListener("click", () => callAction("ai_stop"));

    $("btn-cls2-export-zip")?.addEventListener("click", () => {
      if (!state.sessionId) return;
      window.open(`${V2}/session/${state.sessionId}/export/zip`, "_blank");
    });
    $("btn-cls2-export-excel")?.addEventListener("click", () => {
      if (!state.sessionId) return;
      window.open(`${V2}/session/${state.sessionId}/export/excel`, "_blank");
    });
    $("btn-cls2-add-category")?.addEventListener("click", async () => {
      const name = window.prompt("Namn på ny kategori:");
      if (!name) return;
      const description = window.prompt("Beskrivning (valfritt):", "") || "";
      await callAction("add_category", { name, description });
    });

    $("btn-cls2-import-zip")?.addEventListener("click", () => $("cls2-import-zip-input")?.click());
    $("btn-cls2-import-excel")?.addEventListener("click", () => $("cls2-import-excel-input")?.click());
    $("cls2-import-zip-input")?.addEventListener("change", (ev) => {
      const f = ev.target.files?.[0];
      if (f) importSession("zip", f);
      ev.target.value = "";
    });
    $("cls2-import-excel-input")?.addEventListener("change", (ev) => {
      const f = ev.target.files?.[0];
      if (f) importSession("excel", f);
      ev.target.value = "";
    });

    document.addEventListener("keydown", async (ev) => {
      if (state.mode !== "v2" || !isClassifierPageVisible()) return;
      const tag = (ev.target?.tagName || "").toLowerCase();
      if (tag === "input" || tag === "textarea") return;
      if (ev.key === "ArrowLeft") {
        ev.preventDefault();
        await back();
        return;
      }
      if (ev.key === "ArrowRight") {
        ev.preventDefault();
        await skip();
        return;
      }
      const key = ev.key;
      const found = state.keymap.find((k) => k.hotkey && k.hotkey === key);
      if (found) {
        ev.preventDefault();
        await classify(found.category);
      }
    });
  }

  function renderDataRows() {
    const host = $("cls2-data-file-rows");
    if (!host) return;
    host.innerHTML = "";
    CLS2_DATA_SLOTS.forEach((slot) => host.appendChild(buildDataRow(slot)));
  }

  function startPolling() {
    if (state.pollTimer) clearInterval(state.pollTimer);
    state.pollTimer = setInterval(() => {
      if (state.mode === "v2" && state.sessionId) {
        refreshState(false).catch(() => {});
      }
    }, 5000);
  }

  function boot() {
    renderDataRows();
    wireButtons();
    setMode(state.mode);
    initConfig();
    startPolling();
  }

  window.addEventListener("DOMContentLoaded", boot);
})();

