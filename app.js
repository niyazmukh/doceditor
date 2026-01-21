(() => {
  const STORAGE_KEY = "qtp:v1"; // legacy (localStorage) for migration / fallback
  const APP_VERSION = "0.6.0";

  const DB_NAME = "qtp";
  const DB_VERSION = 2;
  const DB_STORE = "kv";
  const DB_STATE_KEY = "state"; // legacy singleton
  const DB_TEMPLATES_STORE = "templates";
  const DB_PRESETS_STORE = "presets";
  const DB_ACTIVE_TEMPLATE_KEY = "activeTemplateId";

  /** @type {Promise<IDBDatabase> | null} */
  let dbPromise = null;
  let saveTimer = /** @type {ReturnType<typeof setTimeout> | null} */ (null);

  /** @type {null | ReturnType<typeof getUi>} */
  let ui = null;

  /** @type {{templateId: string|null, templateName: string, docName: string|null, templateHtml: string, fields: Array<any>, valuesByFieldId: Record<string,string>, constants: Array<{name:string,value:string}>, pdfFileName: string, textDirection: "auto"|"ltr"|"rtl"}} */
  const state = {
    templateId: null,
    templateName: "Untitled",
    templateCreatedAt: "",
    docName: null,
    templateHtml: "",
    fields: [],
    valuesByFieldId: {},
    constants: [],
    pdfFileName: "",
    textDirection: "auto",
  };

  let pendingSelectionRange = /** @type {Range|null} */ (null);
  let selectionLocked = false;
  let lastSelectionIssue = "";

  function getUi() {
    const $ = (id) => {
      const el = document.getElementById(id);
      if (!el) throw new Error(`Missing element #${id}`);
      return el;
    };

    return {
      fileInput: /** @type {HTMLInputElement} */ ($("fileInput")),
      templateInput: /** @type {HTMLInputElement} */ ($("templateInput")),
      btnExportTemplate: /** @type {HTMLButtonElement} */ ($("btnExportTemplate")),
      btnClear: /** @type {HTMLButtonElement} */ ($("btnClear")),

      templateSelect: /** @type {HTMLSelectElement} */ ($("templateSelect")),
      btnNewTemplate: /** @type {HTMLButtonElement} */ ($("btnNewTemplate")),
      btnDeleteTemplate: /** @type {HTMLButtonElement} */ ($("btnDeleteTemplate")),
      presetSelect: /** @type {HTMLSelectElement} */ ($("presetSelect")),
      btnSavePreset: /** @type {HTMLButtonElement} */ ($("btnSavePreset")),
      btnLoadPreset: /** @type {HTMLButtonElement} */ ($("btnLoadPreset")),
      btnDeletePreset: /** @type {HTMLButtonElement} */ ($("btnDeletePreset")),

      viewerTemplate: /** @type {HTMLDivElement} */ ($("viewerTemplate")),
      viewerFilled: /** @type {HTMLDivElement} */ ($("viewerFilled")),

      btnCreateField: /** @type {HTMLButtonElement} */ ($("btnCreateField")),
      selectionHint: /** @type {HTMLDivElement} */ ($("selectionHint")),

      fieldsList: /** @type {HTMLDivElement} */ ($("fieldsList")),
      fieldsEmpty: /** @type {HTMLDivElement} */ ($("fieldsEmpty")),

      constName: /** @type {HTMLInputElement} */ ($("constName")),
      constValue: /** @type {HTMLInputElement} */ ($("constValue")),
      btnAddConst: /** @type {HTMLButtonElement} */ ($("btnAddConst")),
      constList: /** @type {HTMLDivElement} */ ($("constList")),

      pdfFileName: /** @type {HTMLInputElement} */ ($("pdfFileName")),
      textDirection: /** @type {HTMLSelectElement} */ ($("textDirection")),

      btnDownloadPdf: /** @type {HTMLButtonElement} */ ($("btnDownloadPdf")),
      tabTemplate: /** @type {HTMLButtonElement} */ ($("tabTemplate")),
      tabFilled: /** @type {HTMLButtonElement} */ ($("tabFilled")),
      toggleInlinePreview: /** @type {HTMLInputElement} */ ($("toggleInlinePreview")),

      docName: /** @type {HTMLSpanElement} */ ($("docName")),
      docStatus: /** @type {HTMLSpanElement} */ ($("docStatus")),

      fieldDialog: /** @type {HTMLDialogElement} */ ($("fieldDialog")),
      fieldName: /** @type {HTMLInputElement} */ ($("fieldName")),
      fieldType: /** @type {HTMLSelectElement} */ ($("fieldType")),
      fieldFormulaWrap: /** @type {HTMLLabelElement} */ ($("fieldFormulaWrap")),
      fieldFormula: /** @type {HTMLInputElement} */ ($("fieldFormula")),
      applyAllMatches: /** @type {HTMLInputElement} */ ($("applyAllMatches")),
      matchCaseSensitive: /** @type {HTMLInputElement} */ ($("matchCaseSensitive")),
    };
  }

  function escapeHtml(text) {
    return String(text ?? "")
      .replaceAll("&", "&amp;")
      .replaceAll("<", "&lt;")
      .replaceAll(">", "&gt;")
      .replaceAll('"', "&quot;")
      .replaceAll("'", "&#039;");
  }

  function toSafeName(text) {
    const cleaned = String(text ?? "")
      .trim()
      .toLowerCase()
      .replaceAll(/[^a-z0-9_ ]/g, "")
      .replaceAll(/\s+/g, "_")
      .replaceAll(/_+/g, "_")
      .replaceAll(/^_+|_+$/g, "");
    return cleaned || `field_${Math.floor(Math.random() * 10000)}`;
  }

  function uid() {
    return (crypto?.randomUUID?.() ?? `id_${Date.now()}_${Math.floor(Math.random() * 1e9)}`).replaceAll("-", "");
  }

  function parseMaybeNumber(value) {
    const trimmed = String(value ?? "").trim();
    if (!trimmed) return null;
    const normalized = trimmed.replaceAll(",", "");
    const num = Number(normalized);
    return Number.isFinite(num) ? num : null;
  }

  function formatDateForDoc(value) {
    const trimmed = String(value ?? "").trim();
    if (!trimmed) return "";
    const d = new Date(`${trimmed}T00:00:00`);
    if (Number.isNaN(d.getTime())) return trimmed;
    const day = String(d.getDate()).padStart(2, '0');
    const month = String(d.getMonth() + 1).padStart(2, '0');
    const year = d.getFullYear();
    return `${day}/${month}/${year}`;
  }

  function isValidIdentifier(name) {
    return /^[a-z_][a-z0-9_]*$/.test(name);
  }

  function hasTemplate() {
    return Boolean(state.templateHtml && state.templateHtml.trim());
  }

  function setDocMeta(name, status) {
    if (!ui) return;
    ui.docName.textContent = name ?? "No file imported";
    ui.docStatus.textContent = status ?? "";
    document.title = status ? `QuoteToPDF - ${status}` : "QuoteToPDF";
  }

  function safeFilenameBase(name) {
    const base = String(name ?? "").trim().replace(/\.[^.]+$/, "");
    const cleaned = base.replaceAll(/[^a-zA-Z0-9._-]+/g, "_").replaceAll(/_+/g, "_").replaceAll(/^_+|_+$/g, "");
    return cleaned || "template";
  }

  function normalizePdfFilename(name, fallbackBase) {
    const raw = String(name ?? "").trim();
    const base = raw ? raw : String(fallbackBase ?? "").trim();
    const cleanedBase = safeFilenameBase(base);
    return cleanedBase.toLowerCase().endsWith(".pdf") ? cleanedBase : `${cleanedBase}.pdf`;
  }

  function containsHebrew(text) {
    return /[\u0590-\u05FF]/.test(String(text ?? ""));
  }

  function getFilledPlainText() {
    if (!ui) return "";
    renderFilled();
    return String(ui.viewerFilled.innerText ?? ui.viewerFilled.textContent ?? "");
  }

  function openPrintDialogForPdf() {
    if (!ui) return;
    renderFilled();
    setTab("filled");
    setDocMeta(state.docName, "For RTL/Hebrew, use your browser Print → Save as PDF.");
    window.print();
  }

  async function downloadPdfViaCanvas(filename) {
    if (!ui) return;
    if (!window.html2canvas) throw new Error("RTL PDF exporter failed to load (html2canvas).");
    if (!window.jspdf?.jsPDF) throw new Error("PDF library failed to load (jsPDF).");

    renderFilled();
    const exportNode = /** @type {HTMLDivElement} */ (ui.viewerFilled.cloneNode(true));
    exportNode.querySelectorAll(".qtp-field").forEach((s) => {
      s.removeAttribute("title");
      s.removeAttribute("data-has-value");
      s.removeAttribute("data-error");
      s.classList.remove("qtp-field");
    });

    const mount = document.createElement("div");
    mount.style.position = "fixed";
    mount.style.left = "-10000px";
    mount.style.top = "0";
    mount.style.width = "794px"; // ~A4 at 96dpi
    mount.style.background = "white";
    mount.style.color = "black";
    mount.style.padding = "24px";
    mount.style.fontFamily = "Arial, sans-serif";
    mount.style.fontSize = "14px";
    mount.style.lineHeight = "1.4";
    mount.appendChild(exportNode);
    document.body.appendChild(mount);

    try {
      const canvas = await window.html2canvas(mount, {
        backgroundColor: "#ffffff",
        scale: 2,
        useCORS: true,
      });

      const { jsPDF } = window.jspdf;
      const doc = new jsPDF({ unit: "pt", format: "a4" });
      const pageWidth = doc.internal.pageSize.getWidth();
      const pageHeight = doc.internal.pageSize.getHeight();
      const margin = 36;
      const targetWidth = pageWidth - margin * 2;
      const ptPerPx = targetWidth / canvas.width;
      const sliceHeightPx = Math.floor((pageHeight - margin * 2) / ptPerPx);

      const sliceCanvas = document.createElement("canvas");
      sliceCanvas.width = canvas.width;
      sliceCanvas.height = sliceHeightPx;
      const ctx = sliceCanvas.getContext("2d");
      if (!ctx) throw new Error("Canvas context unavailable");

      let rendered = 0;
      let page = 0;
      while (rendered < canvas.height) {
        const remaining = canvas.height - rendered;
        const currentSliceHeight = Math.min(sliceHeightPx, remaining);
        sliceCanvas.height = currentSliceHeight;
        ctx.clearRect(0, 0, sliceCanvas.width, sliceCanvas.height);
        ctx.drawImage(canvas, 0, rendered, canvas.width, currentSliceHeight, 0, 0, canvas.width, currentSliceHeight);

        const imgData = sliceCanvas.toDataURL("image/png");
        if (page > 0) doc.addPage();
        const imgHeight = currentSliceHeight * ptPerPx;
        doc.addImage(imgData, "PNG", margin, margin, targetWidth, imgHeight);

        rendered += currentSliceHeight;
        page++;
      }

      doc.save(filename);
    } finally {
      mount.remove();
    }
  }

  function downloadBlob(blob, filename) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 0);
  }

  function exportTemplate() {
    const data = toSerializableState();
    const filename = `${safeFilenameBase(state.templateName || state.docName || "template")}.qtp.json`;
    const blob = new Blob([JSON.stringify(data, null, 2)], { type: "application/json" });
    downloadBlob(blob, filename);
    setDocMeta(state.docName, `Exported ${filename}`);
  }

  async function importTemplateObject(obj) {
    if (!obj || typeof obj !== "object") throw new Error("Invalid template file");
    applyLoadedState(obj);
    state.templateId = uid();
    state.templateCreatedAt = new Date().toISOString();
    state.templateName = String(obj.templateName || state.templateName || state.docName || "Imported template").trim() || "Imported template";
    await saveActiveTemplateNow();
    await refreshLibraryUI();
    renderTemplate();
    renderFieldsList();
    renderConstants();
    setTab("template");
    setExportEnabled(hasTemplate());
    updateSelectionUI();
    setDocMeta(state.docName, "Document imported");
  }

  function setTab(which) {
    if (!ui) return;
    const showTemplate = which === "template";
    ui.tabTemplate.classList.toggle("tab--active", showTemplate);
    ui.tabFilled.classList.toggle("tab--active", !showTemplate);
    ui.viewerTemplate.classList.toggle("hidden", !showTemplate);
    ui.viewerFilled.classList.toggle("hidden", showTemplate);
  }

  function setExportEnabled(enabled) {
    if (!ui) return;
    ui.btnDownloadPdf.setAttribute("aria-disabled", enabled ? "false" : "true");
    ui.tabFilled.setAttribute("aria-disabled", enabled ? "false" : "true");
  }

  function supportsIdb() {
    return typeof indexedDB !== "undefined";
  }

  function openDb() {
    if (!supportsIdb()) return Promise.reject(new Error("IndexedDB unavailable"));
    if (dbPromise) return dbPromise;

    dbPromise = new Promise((resolve, reject) => {
      const req = indexedDB.open(DB_NAME, DB_VERSION);
      req.onupgradeneeded = () => {
        const db = req.result;
        if (!db.objectStoreNames.contains(DB_STORE)) db.createObjectStore(DB_STORE);
        if (!db.objectStoreNames.contains(DB_TEMPLATES_STORE)) db.createObjectStore(DB_TEMPLATES_STORE, { keyPath: "id" });
        if (!db.objectStoreNames.contains(DB_PRESETS_STORE)) {
          const store = db.createObjectStore(DB_PRESETS_STORE, { keyPath: "id" });
          store.createIndex("templateId", "templateId", { unique: false });
        }
      };
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => reject(req.error || new Error("Failed to open IndexedDB"));
    });
    return dbPromise;
  }

  async function idbGetFrom(storeName, key) {
    const db = await openDb();
    return await new Promise((resolve, reject) => {
      const tx = db.transaction(storeName, "readonly");
      const store = tx.objectStore(storeName);
      const req = store.get(key);
      req.onsuccess = () => resolve(req.result ?? null);
      req.onerror = () => reject(req.error || new Error("IndexedDB read failed"));
    });
  }

  async function idbPutTo(storeName, value) {
    const db = await openDb();
    return await new Promise((resolve, reject) => {
      const tx = db.transaction(storeName, "readwrite");
      const store = tx.objectStore(storeName);
      const req = store.put(value);
      req.onsuccess = () => resolve(null);
      req.onerror = () => reject(req.error || new Error("IndexedDB write failed"));
    });
  }

  async function idbDelFrom(storeName, key) {
    const db = await openDb();
    return await new Promise((resolve, reject) => {
      const tx = db.transaction(storeName, "readwrite");
      const store = tx.objectStore(storeName);
      const req = store.delete(key);
      req.onsuccess = () => resolve(null);
      req.onerror = () => reject(req.error || new Error("IndexedDB delete failed"));
    });
  }

  async function idbGetAllFrom(storeName) {
    const db = await openDb();
    return await new Promise((resolve, reject) => {
      const tx = db.transaction(storeName, "readonly");
      const store = tx.objectStore(storeName);
      const req = store.getAll();
      req.onsuccess = () => resolve(req.result ?? []);
      req.onerror = () => reject(req.error || new Error("IndexedDB getAll failed"));
    });
  }

  async function idbGetAllByIndex(storeName, indexName, key) {
    const db = await openDb();
    return await new Promise((resolve, reject) => {
      const tx = db.transaction(storeName, "readonly");
      const store = tx.objectStore(storeName);
      const idx = store.index(indexName);
      const req = idx.getAll(key);
      req.onsuccess = () => resolve(req.result ?? []);
      req.onerror = () => reject(req.error || new Error("IndexedDB index getAll failed"));
    });
  }

  async function idbClearStore(storeName) {
    const db = await openDb();
    return await new Promise((resolve, reject) => {
      const tx = db.transaction(storeName, "readwrite");
      const store = tx.objectStore(storeName);
      const req = store.clear();
      req.onsuccess = () => resolve(null);
      req.onerror = () => reject(req.error || new Error("IndexedDB clear failed"));
    });
  }

  async function idbGet(key) {
    const db = await openDb();
    return await new Promise((resolve, reject) => {
      const tx = db.transaction(DB_STORE, "readonly");
      const store = tx.objectStore(DB_STORE);
      const req = store.get(key);
      req.onsuccess = () => resolve(req.result ?? null);
      req.onerror = () => reject(req.error || new Error("IndexedDB read failed"));
    });
  }

  async function idbPut(key, value) {
    const db = await openDb();
    return await new Promise((resolve, reject) => {
      const tx = db.transaction(DB_STORE, "readwrite");
      const store = tx.objectStore(DB_STORE);
      const req = store.put(value, key);
      req.onsuccess = () => resolve(null);
      req.onerror = () => reject(req.error || new Error("IndexedDB write failed"));
    });
  }

  async function idbDel(key) {
    const db = await openDb();
    return await new Promise((resolve, reject) => {
      const tx = db.transaction(DB_STORE, "readwrite");
      const store = tx.objectStore(DB_STORE);
      const req = store.delete(key);
      req.onsuccess = () => resolve(null);
      req.onerror = () => reject(req.error || new Error("IndexedDB delete failed"));
    });
  }

  function toSerializableState() {
    return {
      schemaVersion: 1,
      appVersion: APP_VERSION,
      savedAt: new Date().toISOString(),
      templateId: state.templateId,
      templateName: state.templateName,
      docName: state.docName,
      templateHtml: state.templateHtml,
      fields: state.fields,
      valuesByFieldId: state.valuesByFieldId,
      constants: state.constants,
      pdfFileName: state.pdfFileName,
      textDirection: state.textDirection,
    };
  }

  function applyLoadedState(parsed) {
    if (!parsed || typeof parsed !== "object") return false;
    state.templateId = typeof parsed.templateId === "string" ? parsed.templateId : null;
    state.templateName = typeof parsed.templateName === "string" ? parsed.templateName : "Untitled";
    state.docName = typeof parsed.docName === "string" ? parsed.docName : null;
    state.templateHtml = typeof parsed.templateHtml === "string" ? parsed.templateHtml : "";

    state.fields = Array.isArray(parsed.fields)
      ? parsed.fields
          .filter((f) => f && typeof f === "object")
          .map((f) => ({
            id: typeof f.id === "string" ? f.id : uid(),
            name: typeof f.name === "string" ? toSafeName(f.name) : uid(),
            type: f.type === "text" || f.type === "number" || f.type === "date" || f.type === "formula" ? f.type : "text",
            formula: typeof f.formula === "string" ? f.formula : undefined,
            matchText: typeof f.matchText === "string" ? f.matchText : undefined,
            matchCaseSensitive: typeof f.matchCaseSensitive === "boolean" ? f.matchCaseSensitive : undefined,
          }))
      : [];

    state.valuesByFieldId =
      parsed.valuesByFieldId && typeof parsed.valuesByFieldId === "object" ? parsed.valuesByFieldId : {};
    state.constants = Array.isArray(parsed.constants)
      ? parsed.constants
          .filter((c) => c && typeof c === "object")
          .map((c) => ({
            name: typeof c.name === "string" ? toSafeName(c.name) : "",
            value: typeof c.value === "string" ? c.value : "",
          }))
          .filter((c) => c.name && c.value !== "")
      : [];

    state.pdfFileName = typeof parsed.pdfFileName === "string" ? parsed.pdfFileName : "";
    state.textDirection =
      parsed.textDirection === "rtl" || parsed.textDirection === "ltr" || parsed.textDirection === "auto"
        ? parsed.textDirection
        : "auto";
    return true;
  }

  function applyTextDirectionToViewers() {
    if (!ui) return;
    const explicit = state.textDirection === "rtl" ? "rtl" : state.textDirection === "ltr" ? "ltr" : null;
    if (explicit) {
      ui.viewerTemplate.setAttribute("dir", explicit);
      ui.viewerFilled.setAttribute("dir", explicit);
      return;
    }

    const templateText = String(ui.viewerTemplate.innerText ?? ui.viewerTemplate.textContent ?? "");
    const filledText = String(ui.viewerFilled.innerText ?? ui.viewerFilled.textContent ?? "");
    const sample = (filledText.trim() ? filledText : templateText).trim();
    const detected = detectDirection(sample);
    ui.viewerTemplate.setAttribute("dir", detected);
    ui.viewerFilled.setAttribute("dir", detected);
  }

  function detectDirection(text) {
    const s = String(text ?? "");
    // First-strong direction detection (covers Hebrew/Arabic).
    for (const ch of s) {
      const code = ch.codePointAt(0) ?? 0;
      // Hebrew: 0590–05FF
      if (code >= 0x0590 && code <= 0x05ff) return "rtl";
      // Arabic and related blocks: 0600–06FF, 0750–077F, 08A0–08FF
      if (
        (code >= 0x0600 && code <= 0x06ff) ||
        (code >= 0x0750 && code <= 0x077f) ||
        (code >= 0x08a0 && code <= 0x08ff)
      )
        return "rtl";
      // Basic Latin letters
      if ((code >= 0x0041 && code <= 0x005a) || (code >= 0x0061 && code <= 0x007a)) return "ltr";
      // Latin-ish fallback
      if (code >= 0x00c0 && code <= 0x02af) return "ltr";
    }
    return "ltr";
  }

  function getActiveTemplateRecordFromState() {
    const now = new Date().toISOString();
    return {
      id: state.templateId ?? uid(),
      name: String(state.templateName || "Untitled"),
      createdAt: state.templateCreatedAt || now,
      updatedAt: now,
      docName: state.docName,
      templateHtml: state.templateHtml,
      fields: state.fields,
      valuesByFieldId: state.valuesByFieldId,
      constants: state.constants,
      pdfFileName: state.pdfFileName,
      textDirection: state.textDirection,
    };
  }

  function applyTemplateRecordToState(rec) {
    state.templateId = rec?.id ?? null;
    state.templateName = typeof rec?.name === "string" ? rec.name : "Untitled";
    state.templateCreatedAt = typeof rec?.createdAt === "string" ? rec.createdAt : "";
    state.docName = typeof rec?.docName === "string" ? rec.docName : null;
    state.templateHtml = typeof rec?.templateHtml === "string" ? rec.templateHtml : "";
    state.fields = Array.isArray(rec?.fields) ? rec.fields : [];
    state.valuesByFieldId = rec?.valuesByFieldId && typeof rec.valuesByFieldId === "object" ? rec.valuesByFieldId : {};
    state.constants = Array.isArray(rec?.constants) ? rec.constants : [];
    state.pdfFileName = typeof rec?.pdfFileName === "string" ? rec.pdfFileName : "";
    state.textDirection = rec?.textDirection === "rtl" || rec?.textDirection === "ltr" || rec?.textDirection === "auto" ? rec.textDirection : "auto";
  }

  async function listTemplates() {
    if (!supportsIdb()) return [];
    const all = /** @type {any[]} */ (await idbGetAllFrom(DB_TEMPLATES_STORE));
    return all
      .filter((t) => t && typeof t === "object" && typeof t.id === "string")
      .sort((a, b) => String(a.name ?? "").localeCompare(String(b.name ?? "")));
  }

  async function listPresets(templateId) {
    if (!supportsIdb() || !templateId) return [];
    const all = /** @type {any[]} */ (await idbGetAllByIndex(DB_PRESETS_STORE, "templateId", templateId));
    return all
      .filter((p) => p && typeof p === "object" && typeof p.id === "string")
      .sort((a, b) => String(a.name ?? "").localeCompare(String(b.name ?? "")));
  }

  async function deleteTemplateById(templateId) {
    if (!supportsIdb() || !templateId) return;
    const presets = await listPresets(templateId);
    for (const p of presets) {
      try {
        await idbDelFrom(DB_PRESETS_STORE, p.id);
      } catch {
        // ignore
      }
    }
    await idbDelFrom(DB_TEMPLATES_STORE, templateId);
    try {
      const active = await idbGet(DB_ACTIVE_TEMPLATE_KEY);
      if (active === templateId) await idbDel(DB_ACTIVE_TEMPLATE_KEY);
    } catch {
      // ignore
    }
  }

  async function savePresetNow(name) {
    if (!supportsIdb() || !state.templateId) return;
    const now = new Date().toISOString();
    const rec = {
      id: uid(),
      templateId: state.templateId,
      name: String(name || "Saved values").trim() || "Saved values",
      createdAt: now,
      updatedAt: now,
      valuesByFieldId: state.valuesByFieldId,
      constants: state.constants,
    };
    await idbPutTo(DB_PRESETS_STORE, rec);
  }

  async function loadPresetById(presetId) {
    if (!supportsIdb() || !state.templateId || !presetId) return;
    const rec = await idbGetFrom(DB_PRESETS_STORE, presetId);
    if (!rec || rec.templateId !== state.templateId) return;
    state.valuesByFieldId = rec.valuesByFieldId && typeof rec.valuesByFieldId === "object" ? rec.valuesByFieldId : {};
    state.constants = Array.isArray(rec.constants) ? rec.constants : [];
    saveState();
    renderFieldsList();
    renderConstants();
    renderFilled();
  }

  async function deletePresetById(presetId) {
    if (!supportsIdb() || !presetId) return;
    await idbDelFrom(DB_PRESETS_STORE, presetId);
  }

  async function saveActiveTemplateNow() {
    if (!supportsIdb()) return;
    if (!state.templateId) state.templateId = uid();
    const rec = getActiveTemplateRecordFromState();
    state.templateId = rec.id;
    state.templateCreatedAt = rec.createdAt;
    await idbPutTo(DB_TEMPLATES_STORE, rec);
    await idbPut(DB_ACTIVE_TEMPLATE_KEY, rec.id);
  }

  async function flushSave() {
    if (saveTimer) {
      clearTimeout(saveTimer);
      saveTimer = null;
      try {
        await saveActiveTemplateNow();
      } catch {
        // ignore
      }
    }
  }

  async function refreshLibraryUI() {
    if (!ui) return;
    const templates = await listTemplates();

    ui.templateSelect.innerHTML = "";
    if (templates.length === 0) {
      const opt = document.createElement("option");
      opt.value = "";
      opt.textContent = "No saved documents";
      ui.templateSelect.appendChild(opt);
      ui.templateSelect.disabled = true;
      ui.btnDeleteTemplate.disabled = true;
    } else {
      ui.templateSelect.disabled = false;
      ui.btnDeleteTemplate.disabled = false;
      for (const t of templates) {
        const opt = document.createElement("option");
        opt.value = t.id;
        opt.textContent = t.name || t.docName || t.id;
        ui.templateSelect.appendChild(opt);
      }
      if (state.templateId && templates.some((t) => t.id === state.templateId)) ui.templateSelect.value = state.templateId;
      else ui.templateSelect.value = templates[0].id;
    }

    ui.presetSelect.innerHTML = "";
    const placeholder = document.createElement("option");
    placeholder.value = "";
    placeholder.textContent = "No saved values selected";
    ui.presetSelect.appendChild(placeholder);

    const presets = state.templateId ? await listPresets(state.templateId) : [];
    for (const p of presets) {
      const opt = document.createElement("option");
      opt.value = p.id;
      opt.textContent = p.name || p.id;
      ui.presetSelect.appendChild(opt);
    }
    ui.btnLoadPreset.disabled = presets.length === 0;
    ui.btnDeletePreset.disabled = presets.length === 0;
  }

  async function loadTemplateById(templateId) {
    if (!supportsIdb()) return;
    const rec = await idbGetFrom(DB_TEMPLATES_STORE, templateId);
    if (!rec) return;
    applyTemplateRecordToState(rec);
    if (ui) {
      ui.pdfFileName.value = state.pdfFileName ?? "";
      ui.textDirection.value = state.textDirection ?? "auto";
    }
    renderTemplate();
    renderFieldsList();
    renderConstants();
    renderFilled();
    applyTextDirectionToViewers();
    setTab("template");
    setExportEnabled(hasTemplate());
    updateSelectionUI();
  }

  async function migrateLegacySingletonIfNeeded() {
    if (!supportsIdb()) return false;
    const templates = await listTemplates();
    if (templates.length > 0) return false;

    // Try legacy singleton state from IndexedDB kv, then localStorage.
    let legacy = null;
    try {
      legacy = await idbGet(DB_STATE_KEY);
    } catch {
      legacy = null;
    }
    if (!legacy) {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (raw) {
        try {
          legacy = JSON.parse(raw);
        } catch {
          legacy = null;
        }
      }
    }
    if (!legacy) return false;
    if (!applyLoadedState(legacy)) return false;

    // Create a template from the legacy state.
    state.templateId = uid();
    state.templateName = state.docName ? String(state.docName).replace(/\.[^.]+$/, "") : "Template 1";
    state.templateCreatedAt = new Date().toISOString();
    await saveActiveTemplateNow();
    try {
      await idbDel(DB_STATE_KEY);
    } catch {
      // ignore
    }
    localStorage.removeItem(STORAGE_KEY);
    return true;
  }

  async function loadState() {
    if (!supportsIdb()) return;

    await migrateLegacySingletonIfNeeded();
    const templates = await listTemplates();
    if (templates.length === 0) return;

    let activeId = null;
    try {
      activeId = await idbGet(DB_ACTIVE_TEMPLATE_KEY);
    } catch {
      activeId = null;
    }

    const chosen = (activeId && templates.some((t) => t.id === activeId) ? activeId : templates[0].id) ?? templates[0].id;
    await loadTemplateById(chosen);
  }

  function saveState() {
    // Debounce writes (typing in inputs should not spam storage).
    if (saveTimer) clearTimeout(saveTimer);
    saveTimer = setTimeout(async () => {
      saveTimer = null;
      try {
        await saveActiveTemplateNow();
        await refreshLibraryUI();
      } catch {
        // ignore
      }
    }, 150);
  }

  async function clearState() {
    localStorage.removeItem(STORAGE_KEY);
    if (supportsIdb()) {
      try {
        await idbClearStore(DB_TEMPLATES_STORE);
        await idbClearStore(DB_PRESETS_STORE);
        await idbDel(DB_ACTIVE_TEMPLATE_KEY);
        await idbDel(DB_STATE_KEY);
      } catch {
        // ignore
      }
    }
    state.templateId = null;
    state.templateName = "Untitled";
    state.templateCreatedAt = "";
    state.docName = null;
    state.templateHtml = "";
    state.fields = [];
    state.valuesByFieldId = {};
    state.constants = [];
    state.pdfFileName = "";
    state.textDirection = "auto";
  }

  function syncTemplateFromDom() {
    if (!ui) return;
    state.templateHtml = ui.viewerTemplate.innerHTML;
    saveState();
  }

  function getFormulaContext(options) {
    const skipFormulaFieldId = options?.skipFormulaFieldId ?? null;
    /** @type {Record<string, any>} */
    const ctx = {};

    for (const c of state.constants) {
      if (!isValidIdentifier(c.name)) continue;
      const maybeNum = parseMaybeNumber(c.value);
      ctx[c.name] = maybeNum ?? String(c.value ?? "");
    }

    for (const f of state.fields) {
      if (!isValidIdentifier(f.name)) continue;
      if (f.type === "formula") continue;
      const raw = state.valuesByFieldId[f.id] ?? "";
      if (f.type === "number") {
        ctx[f.name] = parseMaybeNumber(raw) ?? 0;
      } else if (f.type === "date") {
        ctx[f.name] = formatDateForDoc(raw);
      } else {
        ctx[f.name] = String(raw ?? "");
      }
    }

    // Allow formulas to reference other formulas by name.
    // Evaluate formulas in a few passes to resolve dependencies (A -> B -> C).
    if (typeof exprEval === "undefined" || !exprEval?.Parser) return ctx;

    const parser = new exprEval.Parser({ operators: { assignment: false } });
    const formulas = state.fields.filter(
      (f) =>
        f.type === "formula" &&
        f.id !== skipFormulaFieldId &&
        isValidIdentifier(f.name) &&
        String(f.formula ?? "").trim().length > 0,
    );

    const computedNames = new Set();
    for (let pass = 0; pass < formulas.length; pass++) {
      let progressed = false;
      for (const f of formulas) {
        if (computedNames.has(f.name)) continue;
        try {
          const expr = parser.parse(String(f.formula ?? "").trim());
          const result = expr.evaluate(ctx);
          if (typeof result === "number" && !Number.isFinite(result)) throw new Error("Formula evaluated to non-finite number");
          ctx[f.name] = result;
          computedNames.add(f.name);
          progressed = true;
        } catch {
          // unresolved dependency or invalid expression; retry in later passes
        }
      }
      if (!progressed) break;
    }

    return ctx;
  }

  function computeFieldValue(field) {
    const raw = state.valuesByFieldId[field.id] ?? "";

    if (field.type === "text") return String(raw ?? "");
    if (field.type === "number") {
      const num = parseMaybeNumber(raw);
      return num === null ? "" : String(num);
    }
    if (field.type === "date") return formatDateForDoc(raw);
    if (field.type === "formula") {
      const formula = String(field.formula ?? "").trim();
      if (!formula) return "";
      const ctx = getFormulaContext({ skipFormulaFieldId: field.id });
      try {
        if (typeof exprEval === "undefined" || !exprEval?.Parser) {
          return { error: true, message: "Formula engine failed to load (expr-eval)." };
        }
        const parser = new exprEval.Parser({ operators: { assignment: false } });
        const expr = parser.parse(formula);
        const result = expr.evaluate(ctx);
        if (typeof result === "number" && Number.isFinite(result)) return String(result);
        return String(result ?? "");
      } catch (e) {
        return { error: true, message: String(e?.message ?? e) };
      }
    }
    return "";
  }

  function setupCollapsiblePanels() {
    document.querySelectorAll(".panel__title--collapsible").forEach((title) => {
      title.addEventListener("click", () => {
        const panel = title.getAttribute("data-panel");
        const body = title.nextElementSibling;
        if (!body) return;
        
        const isCollapsed = title.classList.contains("panel__title--collapsed");
        if (isCollapsed) {
          title.classList.remove("panel__title--collapsed");
          body.classList.remove("panel__body--collapsed");
        } else {
          title.classList.add("panel__title--collapsed");
          body.classList.add("panel__body--collapsed");
        }
      });
    });
  }

  function setupRightClickMenu() {
    if (!ui) return;
    
    // Remove any existing context menu
    document.getElementById("contextMenu")?.remove();
    
    // Create context menu
    const menu = document.createElement("div");
    menu.id = "contextMenu";
    menu.className = "context-menu hidden";
    menu.innerHTML = '<div class="context-menu__item" data-action="createField">Create field here</div>';
    document.body.appendChild(menu);
    
    ui.viewerTemplate.addEventListener("contextmenu", (e) => {
      if (!hasTemplate()) return;
      
      const range = ensureSelectionInsideTemplate();
      if (!range) return;
      
      e.preventDefault();
      pendingSelectionRange = range.cloneRange();
      
      menu.style.left = `${e.pageX}px`;
      menu.style.top = `${e.pageY}px`;
      menu.classList.remove("hidden");
    });
    
    document.addEventListener("click", () => {
      menu.classList.add("hidden");
    });
    
    menu.querySelector('[data-action="createField"]')?.addEventListener("click", (e) => {
      e.stopPropagation();
      menu.classList.add("hidden");
      openCreateFieldDialog();
    });
  }

  function generateFieldMap() {
    if (!ui || !hasTemplate()) return [];
    
    const blocks = ui.viewerTemplate.querySelectorAll("p, h1, h2, h3, h4, h5, h6, li");
    const map = [];
    
    blocks.forEach((block, idx) => {
      const fields = block.querySelectorAll(".qtp-field");
      const text = String(block.textContent ?? "").trim().slice(0, 40);
      const type = block.tagName.toLowerCase();
      
      if (text) {
        map.push({
          index: idx,
          text: text.length > 40 ? text + "..." : text,
          fieldCount: fields.length,
          type: type,
          element: block
        });
      }
    });
    
    return map;
  }

  function renderFieldMap() {
    const container = document.getElementById("fieldMapList");
    if (!container) return;
    
    const map = generateFieldMap();
    container.innerHTML = "";
    
    if (map.length === 0) {
      container.innerHTML = '<div class="hint">Import a document to see structure</div>';
      return;
    }
    
    map.forEach((item) => {
      const div = document.createElement("div");
      div.className = "field-map__item";
      if (item.fieldCount > 0) div.classList.add("field-map__item--has-fields");
      
      const icon = item.type.startsWith("h") ? "#" : "▪";
      div.innerHTML = `
        <span class="field-map__icon">${icon}</span>
        <span class="field-map__text">${escapeHtml(item.text)}</span>
        ${item.fieldCount > 0 ? `<span class="field-map__badge">${item.fieldCount}</span>` : ""}
      `;
      
      div.addEventListener("click", () => {
        item.element.scrollIntoView({ behavior: "smooth", block: "center" });
        item.element.style.transition = "background 0.3s";
        item.element.style.background = "rgba(37, 99, 235, 0.1)";
        setTimeout(() => {
          item.element.style.background = "";
        }, 1000);
      });
      
      container.appendChild(div);
    });
  }

  function scrollToFieldInput(fieldId) {
    if (!ui) return;
    const fieldItems = ui.fieldsList.querySelectorAll(".item");
    for (const item of fieldItems) {
      const inputs = item.querySelectorAll("input");
      let found = false;
      for (const input of inputs) {
        if (input.dataset.fieldId === fieldId) {
          found = true;
          break;
        }
      }
      if (found) {
        item.scrollIntoView({ behavior: "smooth", block: "nearest" });
        item.classList.add("item--highlight");
        setTimeout(() => item.classList.remove("item--highlight"), 2000);
        break;
      }
    }
  }

  function highlightFieldInDocument(fieldId, highlight) {
    if (!ui) return;
    const spans = ui.viewerTemplate.querySelectorAll(`.qtp-field[data-field-id="${CSS.escape(fieldId)}"]`);
    for (const span of spans) {
      if (highlight) span.classList.add("qtp-field--highlight");
      else span.classList.remove("qtp-field--highlight");
    }
  }

  function updateInlinePreview() {
    if (!ui) return;
    const enabled = ui.toggleInlinePreview.checked;
    if (enabled) {
      ui.viewerTemplate.classList.add("viewer--inline-preview");
      const spans = ui.viewerTemplate.querySelectorAll(".qtp-field[data-field-id]");
      for (const span of spans) {
        const id = span.getAttribute("data-field-id");
        const field = state.fields.find((f) => f.id === id);
        if (!field) continue;
        const computed = computeFieldValue(field);
        if (typeof computed === "object" && computed?.error) {
          span.setAttribute("data-preview-value", "⚠️");
        } else {
          const text = String(computed ?? "");
          span.setAttribute("data-preview-value", text ? `→ ${text}` : "");
        }
      }
    } else {
      ui.viewerTemplate.classList.remove("viewer--inline-preview");
    }
  }

  function decorateFieldSpans(root, { showComputed }) {
    const spans = root.querySelectorAll(".qtp-field[data-field-id]");
    for (const span of spans) {
      const id = span.getAttribute("data-field-id");
      const field = state.fields.find((f) => f.id === id);
      if (!field) continue;

      if (!showComputed) {
        span.removeAttribute("data-has-value");
        span.removeAttribute("data-error");
        span.setAttribute("title", `${field.name} (${field.type})`);
        continue;
      }

      const computed = computeFieldValue(field);
      if (typeof computed === "object" && computed?.error) {
        span.textContent = "#ERR";
        span.setAttribute("data-error", "true");
        span.removeAttribute("data-has-value");
        span.setAttribute("title", computed.message || "Error");
        continue;
      }

      const text = String(computed ?? "");
      span.textContent = text;
      span.setAttribute("title", `${field.name} (${field.type})`);
      if (text.trim()) span.setAttribute("data-has-value", "true");
      else span.removeAttribute("data-has-value");
      span.removeAttribute("data-error");
    }
  }

  function renderTemplate() {
    if (!ui) return;
    ui.viewerTemplate.innerHTML = state.templateHtml || "";
    decorateFieldSpans(ui.viewerTemplate, { showComputed: false });
    
    // Add click handlers to fields for navigation
    ui.viewerTemplate.querySelectorAll(".qtp-field[data-field-id]").forEach((span) => {
      span.addEventListener("click", (e) => {
        e.stopPropagation();
        const fieldId = span.getAttribute("data-field-id");
        if (fieldId) scrollToFieldInput(fieldId);
      });
    });
    
    updateInlinePreview();
    applyTextDirectionToViewers();
    setExportEnabled(hasTemplate());
    renderFieldMap();
  }

  function renderFilled() {
    if (!ui) return;
    ui.viewerFilled.innerHTML = state.templateHtml || "";
    decorateFieldSpans(ui.viewerFilled, { showComputed: true });
    applyTextDirectionToViewers();
  }

  function updateSelectionUI() {
    if (!ui) return;
    const rangeOk = Boolean(pendingSelectionRange);
    ui.btnCreateField.disabled = !rangeOk || !hasTemplate();
    ui.selectionHint.textContent = hasTemplate()
      ? rangeOk
        ? "Selection ready. Create a field from the highlighted text."
        : lastSelectionIssue || "Highlight text (within one paragraph) to create a field."
      : `Import a .docx or .txt to begin (v${APP_VERSION}).`;
  }

  function ensureSelectionInsideTemplate() {
    if (!ui) return null;
    lastSelectionIssue = "";

    const sel = window.getSelection();
    if (!sel || sel.rangeCount === 0) return null;
    const range = sel.getRangeAt(0);
    if (range.collapsed) return null;

    const common = range.commonAncestorContainer;
    const commonEl = common.nodeType === Node.ELEMENT_NODE ? /** @type {Element} */ (common) : common.parentElement;
    if (!commonEl) return null;
    if (!ui.viewerTemplate.contains(commonEl)) return null;

    const startEl =
      range.startContainer.nodeType === Node.ELEMENT_NODE
        ? /** @type {Element} */ (range.startContainer)
        : range.startContainer.parentElement;
    const endEl =
      range.endContainer.nodeType === Node.ELEMENT_NODE
        ? /** @type {Element} */ (range.endContainer)
        : range.endContainer.parentElement;

    if ((startEl && startEl.closest(".qtp-field")) || (endEl && endEl.closest(".qtp-field"))) {
      lastSelectionIssue = "Selection is inside an existing field. Select outside it (or delete the field) first.";
      return null;
    }

    const blockSelector = "p,li,td,th";
    const startBlock = startEl?.closest(blockSelector);
    const endBlock = endEl?.closest(blockSelector);
    if (startBlock && endBlock && startBlock !== endBlock) {
      lastSelectionIssue = "Selection must stay within a single paragraph (MVP).";
      return null;
    }

    return range;
  }

  function normalizeTextForSearch(text) {
    // Keep length stable (for offset mapping), except trimming the needle itself.
    return String(text ?? "").replaceAll("\u00A0", " ");
  }

  function buildTextIndex(rootEl) {
    /** @type {Array<{node: Text, start: number, end: number}>} */
    const segments = [];
    let text = "";

    const walker = document.createTreeWalker(
      rootEl,
      NodeFilter.SHOW_TEXT,
      /** @type {any} */ ({
        acceptNode: (node) => {
          if (!node?.nodeValue) return NodeFilter.FILTER_REJECT;
          const parent = node.parentElement;
          if (!parent) return NodeFilter.FILTER_REJECT;
          if (parent.closest(".qtp-field")) return NodeFilter.FILTER_REJECT;
          return NodeFilter.FILTER_ACCEPT;
        },
      }),
    );

    /** @type {Text|null} */
    let n = /** @type {any} */ (walker.nextNode());
    while (n) {
      const start = text.length;
      const value = n.nodeValue ?? "";
      text += value;
      segments.push({ node: n, start, end: start + value.length });
      n = /** @type {any} */ (walker.nextNode());
    }

    return { text, segments };
  }

  function locateInSegments(segments, pos) {
    if (segments.length === 0) return null;
    if (pos <= 0) return { node: segments[0].node, offset: 0 };

    for (const seg of segments) {
      if (pos <= seg.end) return { node: seg.node, offset: Math.max(0, pos - seg.start) };
    }

    const last = segments[segments.length - 1];
    return { node: last.node, offset: last.node.nodeValue?.length ?? 0 };
  }

  function isWordChar(ch) {
    if (!ch) return false;
    return /[A-Za-z0-9_]/.test(ch);
  }

  function findAllOccurrences(haystack, needle, options) {
    /** @type {number[]} */
    const out = [];
    if (!needle) return out;

    const caseSensitive = options?.caseSensitive ?? true;
    const wholeWord = options?.wholeWord ?? true;

    const h = caseSensitive ? haystack : haystack.toLowerCase();
    const n = caseSensitive ? needle : needle.toLowerCase();
    let i = 0;
    while (true) {
      const idx = h.indexOf(n, i);
      if (idx === -1) break;
      if (wholeWord) {
        const before = idx > 0 ? haystack[idx - 1] : "";
        const after = idx + needle.length < haystack.length ? haystack[idx + needle.length] : "";
        if (!isWordChar(before) && !isWordChar(after)) out.push(idx);
      } else {
        out.push(idx);
      }
      i = idx + needle.length; // non-overlapping
    }
    return out;
  }

  function applyFieldToAllExactMatches(fieldId, rawSelectedText, options) {
    if (!ui) return 0;
    const needle = normalizeTextForSearch(rawSelectedText).trim();
    if (!needle) return 0;

    let wrapped = 0;
    const blocks = ui.viewerTemplate.querySelectorAll("p,li,td,th,h1,h2,h3,h4,h5,h6");
    const targets = blocks.length ? Array.from(blocks) : [ui.viewerTemplate];

    for (const block of targets) {
      const initial = buildTextIndex(block);
      const initialHay = normalizeTextForSearch(initial.text);
      const occurrences = findAllOccurrences(initialHay, needle, options);
      if (occurrences.length === 0) continue;

      for (let k = occurrences.length - 1; k >= 0; k--) {
        const startPos = occurrences[k];
        const endPos = startPos + needle.length;

        const cur = buildTextIndex(block);
        const curHay = normalizeTextForSearch(cur.text);
        if (endPos > curHay.length) continue;

        const caseSensitive = options?.caseSensitive ?? true;
        const candidate = curHay.slice(startPos, endPos);
        if ((caseSensitive ? candidate : candidate.toLowerCase()) !== (caseSensitive ? needle : needle.toLowerCase())) continue;

        const start = locateInSegments(cur.segments, startPos);
        const end = locateInSegments(cur.segments, endPos);
        if (!start || !end) continue;

        const range = document.createRange();
        range.setStart(start.node, start.offset);
        range.setEnd(end.node, end.offset);

        // Extra safety: don't wrap inside an existing field.
        const startEl = start.node.parentElement;
        const endEl = end.node.parentElement;
        if ((startEl && startEl.closest(".qtp-field")) || (endEl && endEl.closest(".qtp-field"))) continue;

        wrapRangeWithFieldSpan(range, fieldId);
        wrapped++;
      }
    }

    return wrapped;
  }

  function wrapRangeWithFieldSpan(range, fieldId) {
    const span = document.createElement("span");
    span.className = "qtp-field";
    span.setAttribute("data-field-id", fieldId);

    const frag = range.extractContents();
    span.appendChild(frag);
    range.insertNode(span);
  }

  function renderFieldsList() {
    if (!ui) return;
    ui.fieldsList.innerHTML = "";
    ui.fieldsEmpty.classList.toggle("hidden", state.fields.length > 0);

    for (const field of state.fields) {
      const item = document.createElement("div");
      item.className = "item item--row";

      const row = document.createElement("div");
      row.className = "fieldRow";

      const nameInput = document.createElement("input");
      nameInput.className = "input";
      nameInput.value = field.name;
      nameInput.dataset.fieldId = field.id;

      const typePill = document.createElement("span");
      typePill.className = "pill";
      typePill.textContent = field.type;

      const valueInput = document.createElement("input");
      valueInput.className = "input";
      valueInput.dataset.fieldId = field.id;
      if (field.type === "formula") {
        valueInput.placeholder = "formula (e.g., licenses * price_per_license)";
        valueInput.value = field.formula ?? "";
      } else {
        if (field.type === "date") valueInput.type = "date";
        valueInput.placeholder = field.type === "number" ? "value (e.g., 49.99)" : "value";
        valueInput.value = state.valuesByFieldId[field.id] ?? "";
      }

      const btnDel = document.createElement("button");
      btnDel.className = "btn btn--danger";
      btnDel.type = "button";
      btnDel.textContent = "Delete";

      valueInput.addEventListener("input", () => {
        if (field.type === "formula") field.formula = valueInput.value;
        else state.valuesByFieldId[field.id] = valueInput.value;
        saveState();
        renderFilled();
        updateInlinePreview();
      });

      nameInput.addEventListener("change", () => {
        nameInput.value = toSafeName(nameInput.value);
        field.name = nameInput.value;
        saveState();
        renderFilled();
      });

      // Add hover effect to highlight field in document
      item.addEventListener("mouseenter", () => highlightFieldInDocument(field.id, true));
      item.addEventListener("mouseleave", () => highlightFieldInDocument(field.id, false));

      btnDel.addEventListener("click", () => {
        deleteField(field.id);
      });

      row.appendChild(nameInput);
      row.appendChild(typePill);
      row.appendChild(valueInput);
      row.appendChild(btnDel);
      item.appendChild(row);
      ui.fieldsList.appendChild(item);
    }

    renderFilled();
  }

  function deleteField(fieldId) {
    if (!ui) return;
    state.fields = state.fields.filter((f) => f.id !== fieldId);
    delete state.valuesByFieldId[fieldId];

    const spans = ui.viewerTemplate.querySelectorAll(`.qtp-field[data-field-id="${CSS.escape(fieldId)}"]`);
    for (const span of spans) {
      const parent = span.parentNode;
      if (!parent) continue;
      while (span.firstChild) parent.insertBefore(span.firstChild, span);
      parent.removeChild(span);
    }

    syncTemplateFromDom();
    renderTemplate();
    renderFieldsList();
  }

  function renderConstants() {
    if (!ui) return;
    ui.constList.innerHTML = "";
    for (const c of state.constants) {
      const item = document.createElement("div");
      item.className = "item";

      const head = document.createElement("div");
      head.className = "item__head";

      const title = document.createElement("div");
      title.className = "item__title";

      const name = document.createElement("input");
      name.className = "input";
      name.value = c.name;
      name.placeholder = "name";

      const value = document.createElement("input");
      value.className = "input";
      value.value = c.value;
      value.placeholder = "value";

      title.appendChild(name);
      title.appendChild(value);

      const actions = document.createElement("div");
      actions.className = "item__actions";
      const btnDel = document.createElement("button");
      btnDel.className = "btn btn--danger";
      btnDel.type = "button";
      btnDel.textContent = "Delete";
      actions.appendChild(btnDel);

      head.appendChild(title);
      head.appendChild(actions);
      item.appendChild(head);
      ui.constList.appendChild(item);

      name.addEventListener("change", () => {
        name.value = toSafeName(name.value);
        c.name = name.value;
        saveState();
        renderFilled();
      });
      value.addEventListener("input", () => {
        c.value = value.value;
        saveState();
        renderFilled();
      });
      btnDel.addEventListener("click", () => {
        state.constants = state.constants.filter((x) => x !== c);
        saveState();
        renderConstants();
        renderFilled();
      });
    }
  }

  function addConstantFromInputs() {
    if (!ui) return;
    const name = toSafeName(ui.constName.value);
    const value = String(ui.constValue.value ?? "").trim();
    if (!name || !value) return;
    state.constants.push({ name, value });
    ui.constName.value = "";
    ui.constValue.value = "";
    saveState();
    renderConstants();
    renderFilled();
  }

  async function importTxt(name, text) {
    await flushSave();
    const normalized = String(text ?? "").replaceAll("\r\n", "\n").replaceAll("\r", "\n");
    const lines = normalized.split("\n");
    const paras = [];
    let buf = [];
    const flush = () => {
      if (buf.length === 0) return;
      const joined = buf.join("\n");
      paras.push(`<p>${escapeHtml(joined).replaceAll("\n", "<br/>")}</p>`);
      buf = [];
    };
    for (const line of lines) {
      if (line.trim() === "") {
        flush();
        continue;
      }
      buf.push(line);
    }
    flush();

    state.templateId = uid();
    state.templateCreatedAt = new Date().toISOString();
    state.templateName = String(name ?? "Template").replace(/\.[^.]+$/, "") || "Template";

    state.docName = name;
    state.templateHtml = paras.join("\n") || "<p></p>";
    state.fields = [];
    state.valuesByFieldId = {};
    state.constants = [];
    state.pdfFileName = "";
    state.textDirection = "auto";

    await saveActiveTemplateNow();
    await refreshLibraryUI();
    setDocMeta(name, "Imported text");
    renderTemplate();
    renderFieldsList();
    renderConstants();
    setTab("template");
  }

  async function importDocx(file) {
    await flushSave();
    if (!window.mammoth?.convertToHtml) throw new Error("DOCX importer failed to load (mammoth).");
    const arrayBuffer = await file.arrayBuffer();
    const result = await window.mammoth.convertToHtml({ arrayBuffer });
    const html = String(result.value ?? "").trim();

    state.templateId = uid();
    state.templateCreatedAt = new Date().toISOString();
    state.templateName = String(file.name ?? "Template").replace(/\.[^.]+$/, "") || "Template";

    state.docName = file.name;
    state.templateHtml = html || "<p></p>";
    state.fields = [];
    state.valuesByFieldId = {};
    state.constants = [];
    state.pdfFileName = "";
    state.textDirection = "auto";

    await saveActiveTemplateNow();
    await refreshLibraryUI();
    setDocMeta(file.name, "Imported DOCX");
    renderTemplate();
    renderFieldsList();
    renderConstants();
    setTab("template");
  }

  function openCreateFieldDialog() {
    if (!ui) return;
    const range = pendingSelectionRange ?? ensureSelectionInsideTemplate();
    if (!range) return;
    pendingSelectionRange = range.cloneRange();
    selectionLocked = true;
    updateSelectionUI();

    const selected = String(range.toString() ?? "").trim();
    ui.fieldName.value = toSafeName(selected.slice(0, 48));
    ui.fieldType.value = "text";
    ui.fieldFormula.value = "";
    ui.applyAllMatches.checked = true;
    ui.matchCaseSensitive.checked = false;
    ui.fieldFormulaWrap.classList.add("hidden");
    ui.fieldDialog.showModal();
    ui.fieldName.focus();
  }

  function createFieldFromDialog() {
    if (!ui) return;
    const range = pendingSelectionRange;
    if (!range) return;

    const id = uid();
    const name = toSafeName(ui.fieldName.value);
    const type = /** @type {"text"|"number"|"date"|"formula"} */ (ui.fieldType.value);
    const formula = type === "formula" ? String(ui.fieldFormula.value ?? "").trim() : undefined;
    const rawSelectedText = String(range.toString() ?? "");
    const applyAll = Boolean(ui.applyAllMatches.checked);
    const caseSensitive = Boolean(ui.matchCaseSensitive.checked);

    try {
      state.fields.push({ id, name, type, formula, matchText: rawSelectedText.trim(), matchCaseSensitive: caseSensitive });
      wrapRangeWithFieldSpan(range, id);
      const wrappedElsewhere = applyAll
        ? applyFieldToAllExactMatches(id, rawSelectedText, { caseSensitive, wholeWord: true })
        : 0;
      pendingSelectionRange = null;
      syncTemplateFromDom();
      renderTemplate();
      renderFieldsList();
      updateSelectionUI();
      setDocMeta(state.docName, wrappedElsewhere > 0 ? `Field created: ${name} (applied ${wrappedElsewhere + 1}x)` : `Field created: ${name}`);
    } catch (e) {
      state.fields = state.fields.filter((f) => f.id !== id);
      delete state.valuesByFieldId[id];
      saveState();
      setDocMeta(state.docName, `Could not create field: ${String(e?.message ?? e)}`);
    } finally {
      selectionLocked = false;
    }
  }

  function setUpSelectionTracking() {
    if (!ui) return;
    const update = () => {
      if (selectionLocked) return;
      pendingSelectionRange = ensureSelectionInsideTemplate();
      updateSelectionUI();
    };
    document.addEventListener("selectionchange", () => {
      if (!ui) return;
      if (ui.viewerTemplate.classList.contains("hidden")) return;
      update();
    });
    ui.viewerTemplate.addEventListener("mouseup", update);
    ui.viewerTemplate.addEventListener("keyup", update);
  }

  function downloadPdfSimple() {
    if (!ui) return;
    if (!hasTemplate()) {
      setDocMeta(state.docName, "Import a .docx or .txt first.");
      return;
    }

    const filledText = getFilledPlainText();
    const effectiveDir = state.textDirection;

    const defaultBase = state.docName ? state.docName.replace(/\.[^.]+$/, "") : "document";
    const filename = normalizePdfFilename(state.pdfFileName, defaultBase);

    const isRtl = effectiveDir === "rtl" || (effectiveDir === "auto" && detectDirection(filledText) === "rtl");
    if (isRtl) {
      const prev = ui.btnDownloadPdf.textContent;
      ui.btnDownloadPdf.disabled = true;
      ui.btnDownloadPdf.textContent = "Generating...";
      setDocMeta(state.docName, "Generating RTL PDF...");

      downloadPdfViaCanvas(filename)
        .then(() => setDocMeta(state.docName, `Saved ${filename}`))
        .catch((e) => {
          setDocMeta(
            state.docName,
            `RTL PDF export failed: ${String(e?.message ?? e)}. Try Print -> Save as PDF (disable headers/footers).`,
          );
          openPrintDialogForPdf();
        })
        .finally(() => {
          ui.btnDownloadPdf.textContent = prev;
          ui.btnDownloadPdf.disabled = false;
        });
      return;
    }

    renderFilled();
    setTab("filled");

    if (!window.jspdf?.jsPDF) {
      setDocMeta(state.docName, "PDF library failed to load (jsPDF).");
      return;
    }

    const { jsPDF } = window.jspdf;
    const doc = new jsPDF({ unit: "pt", format: "a4" });
    doc.setFont("times", "normal");
    doc.setFontSize(12);

    const pageWidth = doc.internal.pageSize.getWidth();
    const pageHeight = doc.internal.pageSize.getHeight();
    const marginX = 48;
    const marginY = 54;
    const maxWidth = pageWidth - marginX * 2;
    const lineHeight = 16;

    const blocks = Array.from(ui.viewerFilled.querySelectorAll("p, li")).map((el) =>
      String(el.innerText ?? el.textContent ?? "").replaceAll("\r\n", "\n").replaceAll("\r", "\n").trimEnd(),
    );

    let y = marginY;
    for (const block of blocks.length ? blocks : [String(ui.viewerFilled.innerText ?? "").trim()]) {
      const text = String(block ?? "").trim();
      if (!text) {
        y += lineHeight;
        continue;
      }

      const lines = doc.splitTextToSize(text, maxWidth);
      for (const line of lines) {
        if (y + lineHeight > pageHeight - marginY) {
          doc.addPage();
          y = marginY;
        }
        doc.text(line, marginX, y);
        y += lineHeight;
      }
      y += lineHeight * 0.6;
    }

    doc.save(filename);
    setDocMeta(state.docName, `Saved ${filename}`);
  }

  function wireEvents() {
    if (!ui) return;

    ui.templateSelect.addEventListener("change", async () => {
      const nextId = ui.templateSelect.value;
      if (!nextId || nextId === state.templateId) return;
      await flushSave();
      await loadTemplateById(nextId);
      await refreshLibraryUI();
      setDocMeta(state.docName, `Switched to document: ${state.templateName}`);
    });

    ui.btnNewTemplate.addEventListener("click", async () => {
      const name = prompt("Document name:", "New document") ?? "";
      const trimmed = name.trim();
      if (!trimmed) return;
      await flushSave();
      state.templateId = uid();
      state.templateCreatedAt = new Date().toISOString();
      state.templateName = trimmed;
      state.docName = null;
      state.templateHtml = "";
      state.fields = [];
      state.valuesByFieldId = {};
      state.constants = [];
      state.pdfFileName = "";
      state.textDirection = "auto";
      await saveActiveTemplateNow();
      await refreshLibraryUI();
      renderTemplate();
      renderFieldsList();
      renderConstants();
      setTab("template");
      setExportEnabled(false);
      updateSelectionUI();
      setDocMeta(null, `Created document: ${state.templateName}`);
    });

    ui.btnDeleteTemplate.addEventListener("click", async () => {
      if (!state.templateId) return;
      const ok = confirm(`Delete document "${state.templateName}"? This also deletes its saved values.`);
      if (!ok) return;
      const toDelete = state.templateId;
      await flushSave();
      await deleteTemplateById(toDelete);
      state.templateId = null;
      state.templateName = "Untitled";
      state.templateCreatedAt = "";
      state.docName = null;
      state.templateHtml = "";
      state.fields = [];
      state.valuesByFieldId = {};
      state.constants = [];
      state.pdfFileName = "";
      state.textDirection = "auto";

      const templates = await listTemplates();
      if (templates.length > 0) {
        await loadTemplateById(templates[0].id);
      }
      await refreshLibraryUI();
      renderTemplate();
      renderFieldsList();
      renderConstants();
      setTab("template");
      setExportEnabled(hasTemplate());
      updateSelectionUI();
      setDocMeta(null, "Document deleted");
    });

    ui.btnSavePreset.addEventListener("click", async () => {
      if (!state.templateId) return;
      const name = prompt("Saved values name:", "Values 1") ?? "";
      const trimmed = name.trim();
      if (!trimmed) return;
      await savePresetNow(trimmed);
      await refreshLibraryUI();
      setDocMeta(state.docName, `Saved values: ${trimmed}`);
    });

    ui.btnLoadPreset.addEventListener("click", async () => {
      const presetId = ui.presetSelect.value;
      if (!presetId) return;
      await loadPresetById(presetId);
      await refreshLibraryUI();
      setDocMeta(state.docName, "Saved values loaded");
    });

    ui.btnDeletePreset.addEventListener("click", async () => {
      const presetId = ui.presetSelect.value;
      if (!presetId) return;
      const ok = confirm("Delete selected saved values?");
      if (!ok) return;
      await deletePresetById(presetId);
      await refreshLibraryUI();
      setDocMeta(state.docName, "Saved values deleted");
    });

    ui.fileInput.addEventListener("change", async () => {
      const file = ui.fileInput.files?.[0];
      ui.fileInput.value = "";
      if (!file) return;

      const ext = file.name.toLowerCase().split(".").pop() ?? "";
      try {
        if (ext === "txt") {
          await importTxt(file.name, await file.text());
        } else if (ext === "docx") {
          await importDocx(file);
        } else if (ext === "doc") {
          setDocMeta(file.name, ".doc is not supported in-browser. Please convert to .docx.");
        } else if (ext === "rtf") {
          setDocMeta(file.name, ".rtf is not implemented yet. Please use .docx or .txt.");
        } else {
          setDocMeta(file.name, "Unsupported file type. Please use .docx or .txt.");
        }
      } catch (e) {
        setDocMeta(file.name, `Import failed: ${String(e?.message ?? e)}`);
      }
    });

    ui.btnExportTemplate.addEventListener("click", () => {
      exportTemplate();
    });

    ui.templateInput.addEventListener("change", async () => {
      const file = ui.templateInput.files?.[0];
      ui.templateInput.value = "";
      if (!file) return;
      try {
        const text = await file.text();
        const obj = JSON.parse(text);
        await importTemplateObject(obj);
      } catch (e) {
        setDocMeta(state.docName, `Import document failed: ${String(e?.message ?? e)}`);
      }
    });

    ui.btnClear.addEventListener("click", () => {
      void clearState();
      pendingSelectionRange = null;
      selectionLocked = false;
      lastSelectionIssue = "";
      setDocMeta(null, `Import a .docx or .txt to begin (v${APP_VERSION}).`);
      ui.viewerFilled.innerHTML = "";
      void refreshLibraryUI();
      renderTemplate();
      renderFieldsList();
      renderConstants();
      setTab("template");
      setExportEnabled(false);
      updateSelectionUI();
    });

    ui.btnCreateField.addEventListener("mousedown", (e) => e.preventDefault());
    ui.btnCreateField.addEventListener("click", () => openCreateFieldDialog());

    ui.fieldType.addEventListener("change", () => {
      const isFormula = ui.fieldType.value === "formula";
      ui.fieldFormulaWrap.classList.toggle("hidden", !isFormula);
    });

    ui.fieldDialog.addEventListener("close", () => {
      if (!ui) return;
      if (ui.fieldDialog.returnValue === "ok") {
        createFieldFromDialog();
        return;
      }
      selectionLocked = false;
      pendingSelectionRange = ensureSelectionInsideTemplate();
      updateSelectionUI();
    });

    ui.btnAddConst.addEventListener("click", addConstantFromInputs);
    ui.constValue.addEventListener("keydown", (e) => {
      if (e.key === "Enter") addConstantFromInputs();
    });

    ui.btnDownloadPdf.addEventListener("click", () => downloadPdfSimple());

    ui.pdfFileName.addEventListener("input", () => {
      state.pdfFileName = ui.pdfFileName.value;
      saveState();
    });
    ui.textDirection.addEventListener("change", () => {
      state.textDirection = ui.textDirection.value === "rtl" ? "rtl" : ui.textDirection.value === "ltr" ? "ltr" : "auto";
      saveState();
      applyTextDirectionToViewers();
      renderFilled();
    });

    ui.tabTemplate.addEventListener("click", () => setTab("template"));
    ui.tabFilled.addEventListener("click", () => {
      if (!hasTemplate()) {
        setDocMeta(state.docName, "Import a .docx or .txt first.");
        return;
      }
      renderFilled();
      setTab("filled");
    });

    ui.toggleInlinePreview.addEventListener("change", () => {
      updateInlinePreview();
    });
  }

  function init() {
    try {
      ui = getUi();
    } catch (e) {
      alert(`QuoteToPDF failed to start: ${String(e?.message ?? e)}`);
      return;
    }

    if ("serviceWorker" in navigator) {
      navigator.serviceWorker.register("./service-worker.js").catch(() => {
        // ignore
      });
    }

    window.addEventListener("error", (ev) => setDocMeta(state.docName, `Error: ${String(ev?.message ?? ev)}`));
    window.addEventListener("unhandledrejection", (ev) =>
      setDocMeta(state.docName, `Error: ${String(ev?.reason?.message ?? ev?.reason ?? ev)}`),
    );

    (async () => {
      setupCollapsiblePanels();
      setupRightClickMenu();
      setUpSelectionTracking();
      wireEvents();
      await loadState();
      await refreshLibraryUI();

      const persistence = supportsIdb() ? "IndexedDB" : "localStorage";
      setDocMeta(
        state.docName,
        hasTemplate()
          ? `Loaded "${state.templateName}" (${persistence}) (v${APP_VERSION}).`
          : `Import a .docx or .txt to begin (v${APP_VERSION}).`,
      );

      renderTemplate();
      renderFieldsList();
      renderConstants();
      ui.pdfFileName.value = state.pdfFileName ?? "";
      ui.textDirection.value = state.textDirection ?? "auto";
      applyTextDirectionToViewers();
      updateSelectionUI();
      setExportEnabled(hasTemplate());
    })();
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", init);
  } else {
    init();
  }
})();
