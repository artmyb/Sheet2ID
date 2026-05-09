const STORAGE_KEY = "sheet2id-config-v2";
const LIVE_PREVIEW_DEBOUNCE_MS = 250;

const PAGE_SIZE_PRESETS_MM = {
  a4: { width: 210, height: 297 },
  letter: { width: 215.9, height: 279.4 },
  legal: { width: 215.9, height: 355.6 },
};

const FONT_CHOICES = [
  "Helvetica",
  "Helvetica-Bold",
  "Helvetica-Oblique",
  "Helvetica-BoldOblique",
  "Arial",
  "Arial Bold",
  "Arial Italic",
  "Arial Bold Italic",
  "Arial Narrow",
  "Arial Narrow Bold",
  "Arial Narrow Italic",
  "Arial Narrow Bold Italic",
  "Segoe UI",
  "Segoe UI Bold",
  "Segoe UI Italic",
  "Segoe UI Bold Italic",
  "Calibri",
  "Calibri Bold",
  "Calibri Italic",
  "Calibri Bold Italic",
  "Verdana",
  "Verdana Bold",
  "Verdana Italic",
  "Verdana Bold Italic",
  "Tahoma",
  "Tahoma Bold",
  "Tahoma Italic",
  "Tahoma Bold Italic",
  "Trebuchet MS",
  "Trebuchet MS Bold",
  "Trebuchet MS Italic",
  "Trebuchet MS Bold Italic",
  "Century Gothic",
  "Century Gothic Bold",
  "Century Gothic Italic",
  "Century Gothic Bold Italic",
  "Candara",
  "Candara Bold",
  "Candara Italic",
  "Candara Bold Italic",
  "Corbel",
  "Corbel Bold",
  "Corbel Italic",
  "Corbel Bold Italic",
  "Franklin Gothic Medium",
  "Franklin Gothic Medium Italic",
  "Comic Sans MS",
  "Comic Sans MS Bold",
  "Impact",
  "Times-Roman",
  "Times-Bold",
  "Times-Italic",
  "Times-BoldItalic",
  "Times New Roman",
  "Times New Roman Bold",
  "Times New Roman Italic",
  "Times New Roman Bold Italic",
  "Georgia",
  "Georgia Bold",
  "Georgia Italic",
  "Georgia Bold Italic",
  "Palatino Linotype",
  "Palatino Linotype Bold",
  "Palatino Linotype Italic",
  "Palatino Linotype Bold Italic",
  "Book Antiqua",
  "Book Antiqua Bold",
  "Book Antiqua Italic",
  "Book Antiqua Bold Italic",
  "Cambria",
  "Cambria Bold",
  "Cambria Italic",
  "Cambria Bold Italic",
  "Constantia",
  "Constantia Bold",
  "Constantia Italic",
  "Constantia Bold Italic",
  "Garamond",
  "Garamond Bold",
  "Garamond Italic",
  "Garamond Bold Italic",
  "Rockwell",
  "Rockwell Bold",
  "Rockwell Italic",
  "Rockwell Bold Italic",
  "Courier",
  "Courier-Bold",
  "Courier-Oblique",
  "Courier-BoldOblique",
  "Courier New",
  "Courier New Bold",
  "Courier New Italic",
  "Courier New Bold Italic",
  "Consolas",
  "Consolas Bold",
  "Consolas Italic",
  "Consolas Bold Italic",
  "Lucida Console",
  "DejaVu Sans",
  "DejaVu Sans Bold",
  "DejaVu Sans Italic",
  "DejaVu Sans Bold Italic",
  "DejaVu Serif",
  "DejaVu Serif Bold",
  "DejaVu Serif Italic",
  "DejaVu Serif Bold Italic",
  "DejaVu Sans Mono",
  "DejaVu Sans Mono Bold",
  "DejaVu Sans Mono Italic",
  "DejaVu Sans Mono Bold Italic",
  "Liberation Sans",
  "Liberation Sans Bold",
  "Liberation Sans Italic",
  "Liberation Sans Bold Italic",
  "Liberation Serif",
  "Liberation Serif Bold",
  "Liberation Serif Italic",
  "Liberation Serif Bold Italic",
  "Liberation Mono",
  "Liberation Mono Bold",
  "Liberation Mono Italic",
  "Liberation Mono Bold Italic",
  "Noto Sans",
  "Noto Sans Bold",
  "Noto Sans Italic",
  "Noto Sans Bold Italic",
  "Noto Serif",
  "Noto Serif Bold",
  "Noto Serif Italic",
  "Noto Serif Bold Italic",
  "Noto Sans Mono",
  "Noto Sans Mono Bold",
  "Noto Sans Mono Italic",
  "Noto Sans Mono Bold Italic",
];

const DEFAULT_UI_OPEN = {
  sections: {
    spreadsheet: true,
    page: false,
    layout: false,
    card: false,
    photos: false,
    backgrounds: false,
    elements: true,
  },
  rules: {},
  elements: {},
};

const elements = {
  spreadsheet: document.getElementById("spreadsheet"),
  photos: document.getElementById("photos"),
  backgrounds: document.getElementById("backgrounds"),
  previewStart: document.getElementById("preview-start"),
  spreadsheetMeta: document.getElementById("spreadsheet-meta"),
  photosMeta: document.getElementById("photos-meta"),
  backgroundsMeta: document.getElementById("backgrounds-meta"),
  configGui: document.getElementById("config-gui"),
  configEditor: document.getElementById("config-editor"),
  generateButton: document.getElementById("generate-button"),
  loadDefault: document.getElementById("load-default"),
  downloadConfig: document.getElementById("download-config"),
  applyJson: document.getElementById("apply-json"),
  status: document.getElementById("status"),
  metrics: document.getElementById("summary-metrics"),
  previewHead: document.getElementById("preview-head"),
  previewBody: document.getElementById("preview-body"),
  previewFrame: document.getElementById("preview-frame"),
  previewDownloadLink: document.getElementById("preview-download-link"),
  previewEmpty: document.getElementById("preview-empty"),
  previewCaption: document.getElementById("preview-caption"),
  pageSetupPreview: document.getElementById("page-setup-preview"),
  pageSetupMeta: document.getElementById("page-setup-meta"),
  pdfFrame: document.getElementById("pdf-frame"),
  downloadLink: document.getElementById("download-link"),
  sheetOptions: document.getElementById("sheet-options"),
};

const state = {
  defaultConfig: null,
  currentConfig: null,
  availableSheets: [],
  activePreviewPdfUrl: "",
  activeGeneratedPdfUrl: "",
  uiOpen: cloneDeep(DEFAULT_UI_OPEN),
  livePreviewTimer: 0,
  livePreviewRequestId: 0,
  livePreviewController: null,
};

initialize();

async function initialize() {
  bindEvents();
  refreshInputMeta();
  clearSummaryPreview("Load a spreadsheet to populate the summary table.");
  clearLivePreview("Choose a spreadsheet to start the live ID preview.");
  renderPageSetupPreview();
  await loadDefaultConfig();
}

function bindEvents() {
  const handleAssetChange = () => {
    refreshInputMeta();
    queueLivePreview({ immediate: true });
  };

  elements.spreadsheet.addEventListener("change", handleAssetChange);
  elements.photos.addEventListener("change", handleAssetChange);
  elements.backgrounds.addEventListener("change", handleAssetChange);
  elements.previewStart.addEventListener("input", handlePreviewStartInput);
  elements.previewStart.addEventListener("change", handlePreviewStartChange);

  elements.generateButton.addEventListener("click", generateFullPdf);
  elements.loadDefault.addEventListener("click", resetToDefaultConfig);
  elements.downloadConfig.addEventListener("click", downloadConfig);
  elements.applyJson.addEventListener("click", applyJsonConfig);

  elements.configGui.addEventListener("input", handleConfigGuiInput);
  elements.configGui.addEventListener("change", handleConfigGuiInput);
  elements.configGui.addEventListener("click", handleConfigGuiClick);
  elements.configGui.addEventListener("toggle", handleCollapsibleToggle, true);

  document.addEventListener("wheel", handleNumberInputWheel, { passive: false });
}

async function loadDefaultConfig() {
  const response = await fetch("/api/default-config");
  const config = await response.json();
  state.defaultConfig = normalizeClientConfig(config);

  const savedConfig = parseStoredConfig();
  setCurrentConfig(savedConfig || state.defaultConfig, {
    render: true,
    save: false,
    resetUi: true,
    previewImmediate: true,
  });
}

function parseStoredConfig() {
  const savedConfig = localStorage.getItem(STORAGE_KEY);
  if (!savedConfig) {
    return null;
  }

  try {
    return JSON.parse(savedConfig);
  } catch {
    return null;
  }
}

function setCurrentConfig(config, options = {}) {
  state.currentConfig = normalizeClientConfig(config || state.defaultConfig || {});

  if (options.resetUi) {
    state.uiOpen = cloneDeep(DEFAULT_UI_OPEN);
  }

  syncJsonEditor();

  if (options.save !== false) {
    saveConfig();
  }

  if (options.render !== false) {
    renderConfigGui();
  }

  renderPageSetupPreview();

  if (options.preview !== false) {
    queueLivePreview({ immediate: Boolean(options.previewImmediate) });
  }
}

function saveConfig() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state.currentConfig));
}

function resetToDefaultConfig() {
  setCurrentConfig(state.defaultConfig, { render: true, save: true, resetUi: true, previewImmediate: true });
  setStatus("Default configuration restored. The live preview updates automatically.", "info");
}

function downloadConfig() {
  const blob = new Blob([JSON.stringify(state.currentConfig, null, 2)], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = "sheet2id-config.json";
  link.click();
  URL.revokeObjectURL(url);
}

function applyJsonConfig() {
  try {
    const parsed = JSON.parse(elements.configEditor.value || "{}");
    setCurrentConfig(parsed, { render: true, save: true, previewImmediate: true });
    setStatus("Advanced JSON applied. The live preview is refreshing automatically.", "success");
  } catch {
    setStatus("The advanced JSON is not valid.", "error");
  }
}

function syncJsonEditor() {
  elements.configEditor.value = JSON.stringify(serializeConfig(state.currentConfig), null, 2);
}

function refreshInputMeta() {
  const spreadsheet = elements.spreadsheet.files[0];
  elements.spreadsheetMeta.textContent = spreadsheet ? `${spreadsheet.name}` : "No spreadsheet selected.";
  elements.photosMeta.textContent = summarizeFiles(elements.photos.files, "photo");
  elements.backgroundsMeta.textContent = summarizeFiles(elements.backgrounds.files, "background");
}

function summarizeFiles(fileList, label) {
  const count = fileList.length;
  if (!count) {
    return `No ${label} folder selected.`;
  }

  return `${count} ${label}${count === 1 ? "" : "s"} selected.`;
}

function handlePreviewStartInput() {
  if (!elements.previewStart.value.trim()) {
    return;
  }

  queueLivePreview();
}

function handlePreviewStartChange() {
  normalizePreviewStart();
  queueLivePreview({ immediate: true });
}

async function generateFullPdf() {
  try {
    setBusy(true);

    setStatus("Generating the full PDF...", "info");
    const response = await fetch("/api/generate", {
      method: "POST",
      body: buildFormData(),
    });

    if (!response.ok) {
      throw new Error(await readErrorMessage(response));
    }

    const pdfBlob = await response.blob();
    attachGeneratedPdf(pdfBlob, "sheet2id-cards.pdf");
    setStatus("Full PDF generated successfully.", "success");
  } catch (error) {
    setStatus(error.message || "Something went wrong.", "error");
  } finally {
    setBusy(false);
  }
}

function buildFormData(options = {}) {
  const spreadsheet = elements.spreadsheet.files[0];
  if (!spreadsheet) {
    throw new Error("Choose a spreadsheet before previewing or generating.");
  }

  const previewStart = normalizePreviewStart();
  const previewMode = String(options.previewMode || "single-card");
  const formData = new FormData();
  formData.append("spreadsheet", spreadsheet, spreadsheet.name);
  appendFiles(formData, "photos", elements.photos.files);
  appendFiles(formData, "backgrounds", elements.backgrounds.files);
  formData.append("config", JSON.stringify(serializeConfig(state.currentConfig), null, 2));
  formData.append("previewMode", previewMode);
  formData.append("previewStart", String(previewStart));
  return formData;
}

function appendFiles(formData, fieldName, files) {
  Array.from(files).forEach((file) => {
    const filename = file.webkitRelativePath || file.name;
    formData.append(fieldName, file, filename);
  });
}

async function readErrorMessage(response) {
  const payload = await response.json().catch(() => ({ error: "Request failed." }));
  return payload.error || "Request failed.";
}

function normalizePreviewStart() {
  const previewStart = Math.max(1, Number.parseInt(elements.previewStart.value, 10) || 1);
  elements.previewStart.value = String(previewStart);
  return previewStart;
}

function renderPreviewSummary(preview) {
  renderMetrics(preview);
  renderTable(preview);
}

function clearSummaryPreview(message = "No preview yet.") {
  renderMetrics({
    totalRows: 0,
    matchedPhotos: 0,
    matchedBackgrounds: 0,
  });
  elements.previewHead.innerHTML = "<tr><th>Preview</th></tr>";
  elements.previewBody.innerHTML = `<tr><td>${escapeHtml(message)}</td></tr>`;
}

function renderMetrics(preview) {
  elements.metrics.innerHTML = `
    <div class="metric-card">
      <strong>${preview.totalRows}</strong>
      <span>Rows</span>
    </div>
    <div class="metric-card">
      <strong>${preview.matchedPhotos}</strong>
      <span>Photos matched</span>
    </div>
    <div class="metric-card">
      <strong>${preview.matchedBackgrounds}</strong>
      <span>Backgrounds matched</span>
    </div>
  `;
}

function renderTable(preview) {
  const fields = preview.displayFields || [];
  const headCells = [
    "<th>Row</th>",
    ...fields.map((field) => `<th>${escapeHtml(field)}</th>`),
    "<th>Photo</th>",
    "<th>Background</th>",
  ];
  elements.previewHead.innerHTML = `<tr>${headCells.join("")}</tr>`;

  if (!preview.cards.length) {
    elements.previewBody.innerHTML = `<tr><td colspan="${fields.length + 3}">No data rows found in the selected sheet.</td></tr>`;
    return;
  }

  elements.previewBody.innerHTML = preview.cards
    .map((card) => {
      const valueCells = fields.map((field) => `<td>${escapeHtml(card.values[field] || "")}</td>`).join("");
      return `
        <tr>
          <td>${card.rowNumber}</td>
          ${valueCells}
          <td>${escapeHtml(card.photo || "Missing")}</td>
          <td>${escapeHtml(card.background || "None")}</td>
        </tr>
      `;
    })
    .join("");
}

function attachPreviewPdf(blob, filename) {
  if (state.activePreviewPdfUrl) {
    URL.revokeObjectURL(state.activePreviewPdfUrl);
  }

  state.activePreviewPdfUrl = URL.createObjectURL(blob);
  elements.previewFrame.src = state.activePreviewPdfUrl;
  elements.previewDownloadLink.href = state.activePreviewPdfUrl;
  elements.previewDownloadLink.download = filename;
  elements.previewDownloadLink.classList.remove("disabled");
  elements.previewEmpty.hidden = true;
}

function clearLivePreview(message, caption = message) {
  if (state.activePreviewPdfUrl) {
    URL.revokeObjectURL(state.activePreviewPdfUrl);
    state.activePreviewPdfUrl = "";
  }

  elements.previewFrame.removeAttribute("src");
  elements.previewFrame.classList.remove("is-loading");
  elements.previewDownloadLink.href = "#";
  elements.previewDownloadLink.classList.add("disabled");
  elements.previewEmpty.textContent = message;
  elements.previewEmpty.hidden = false;
  elements.previewCaption.textContent = caption;
}

function setLivePreviewLoading(isLoading) {
  elements.previewFrame.classList.toggle("is-loading", isLoading);

  if (isLoading && !state.activePreviewPdfUrl) {
    elements.previewEmpty.textContent = "Building live ID preview...";
    elements.previewEmpty.hidden = false;
  }
}

function attachGeneratedPdf(blob, filename) {
  if (state.activeGeneratedPdfUrl) {
    URL.revokeObjectURL(state.activeGeneratedPdfUrl);
  }

  state.activeGeneratedPdfUrl = URL.createObjectURL(blob);
  elements.pdfFrame.src = state.activeGeneratedPdfUrl;
  elements.downloadLink.href = state.activeGeneratedPdfUrl;
  elements.downloadLink.download = filename;
  elements.downloadLink.classList.remove("disabled");
}

function setBusy(isBusy) {
  elements.generateButton.disabled = isBusy;
  elements.loadDefault.disabled = isBusy;
  elements.downloadConfig.disabled = isBusy;
  elements.applyJson.disabled = isBusy;
}

function setStatus(message, tone) {
  elements.status.textContent = message;
  elements.status.className = `status ${tone}`;
}

function queueLivePreview(options = {}) {
  renderPageSetupPreview();
  window.clearTimeout(state.livePreviewTimer);

  if (!elements.spreadsheet.files[0]) {
    state.availableSheets = [];
    renderSheetOptions();
    clearSummaryPreview("Load a spreadsheet to populate the summary table.");
    clearLivePreview("Choose a spreadsheet to start the live ID preview.");
    return;
  }

  const delay = options.immediate ? 0 : LIVE_PREVIEW_DEBOUNCE_MS;
  state.livePreviewTimer = window.setTimeout(() => {
    void refreshLivePreview();
  }, delay);
}

async function refreshLivePreview() {
  const requestId = ++state.livePreviewRequestId;
  if (state.livePreviewController) {
    state.livePreviewController.abort();
  }

  const controller = new AbortController();
  state.livePreviewController = controller;
  setLivePreviewLoading(true);

  try {
    const previewResponse = await fetch("/api/preview", {
      method: "POST",
      body: buildFormData({ previewMode: "single-card" }),
      signal: controller.signal,
    });

    if (!previewResponse.ok) {
      throw new Error(await readErrorMessage(previewResponse));
    }

    const preview = await previewResponse.json();
    if (requestId !== state.livePreviewRequestId) {
      return;
    }

    state.availableSheets = Array.isArray(preview.availableSheets) ? preview.availableSheets : [];
    renderSheetOptions();
    renderPreviewSummary(preview);

    if (!preview.totalRows) {
      clearLivePreview("No rows available for the live ID preview.", "No rows available in the selected sheet.");
      return;
    }

    const previewDocumentResponse = await fetch("/api/preview-document", {
      method: "POST",
      body: buildFormData({ previewMode: "single-card" }),
      signal: controller.signal,
    });

    if (!previewDocumentResponse.ok) {
      throw new Error(await readErrorMessage(previewDocumentResponse));
    }

    const previewBlob = await previewDocumentResponse.blob();
    if (requestId !== state.livePreviewRequestId) {
      return;
    }

    attachPreviewPdf(previewBlob, "sheet2id-preview-card.pdf");
    const actualRow = Math.min(normalizePreviewStart(), preview.totalRows);
    elements.previewCaption.textContent = `Live preview for row ${actualRow} of ${preview.totalRows}.`;

    if (elements.status.classList.contains("error")) {
      setStatus("Live preview refreshed. Generate when ready.", "success");
    }
  } catch (error) {
    if (controller.signal.aborted || requestId !== state.livePreviewRequestId) {
      return;
    }

    clearSummaryPreview("Preview summary unavailable.");
    clearLivePreview(error.message || "Unable to build the live ID preview.");
    setStatus(error.message || "Unable to build the live ID preview.", "error");
  } finally {
    if (requestId === state.livePreviewRequestId) {
      setLivePreviewLoading(false);
      if (state.livePreviewController === controller) {
        state.livePreviewController = null;
      }
    }
  }
}

function renderPageSetupPreview() {
  if (!elements.pageSetupPreview || !elements.pageSetupMeta) {
    return;
  }

  if (!state.currentConfig) {
    elements.pageSetupPreview.innerHTML = '<p class="page-setup-fallback">Adjust the page settings to see the page setup preview.</p>';
    elements.pageSetupMeta.textContent = "";
    return;
  }

  const pageSize = resolvePageSizeMm(state.currentConfig);
  const grid = resolveGridMm(state.currentConfig, pageSize);
  const slotCount = Math.max(0, grid.columns * grid.rows);

  if (!slotCount || pageSize.width <= 0 || pageSize.height <= 0) {
    elements.pageSetupPreview.innerHTML = '<p class="page-setup-fallback">Adjust the page settings to see the page setup preview.</p>';
    elements.pageSetupMeta.textContent = "";
    return;
  }

  const marginWidth = Math.max(0, pageSize.width - grid.margin * 2);
  const marginHeight = Math.max(0, pageSize.height - grid.margin * 2);
  const cardRects = Array.from({ length: slotCount }, (_unused, slot) => resolveCardRectTopMm(slot, grid));
  const overflowCount = cardRects.filter(
    (rect) =>
      rect.x + rect.width > pageSize.width - grid.margin + 0.001 ||
      rect.y + rect.height > pageSize.height - grid.margin + 0.001
  ).length;

  elements.pageSetupPreview.innerHTML = `
    <svg viewBox="0 0 ${pageSize.width} ${pageSize.height}" role="img" aria-label="Page setup preview">
      <rect x="0" y="0" width="${pageSize.width}" height="${pageSize.height}" rx="10" fill="#fffdf8" stroke="#c9b89e" stroke-width="1.2" />
      <rect
        x="${grid.margin}"
        y="${grid.margin}"
        width="${marginWidth}"
        height="${marginHeight}"
        rx="6"
        fill="none"
        stroke="#d6c3a6"
        stroke-width="0.9"
        stroke-dasharray="4 4"
      />
      ${cardRects
        .map((rect) => {
          const isOverflow =
            rect.x + rect.width > pageSize.width - grid.margin + 0.001 ||
            rect.y + rect.height > pageSize.height - grid.margin + 0.001;
          return `
            <rect
              x="${rect.x}"
              y="${rect.y}"
              width="${rect.width}"
              height="${rect.height}"
              rx="4"
              fill="${isOverflow ? "#f4d6cf" : "#ebc98d"}"
              fill-opacity="${isOverflow ? "0.82" : "0.38"}"
              stroke="${isOverflow ? "#a44d33" : "#935226"}"
              stroke-width="0.8"
            />
          `;
        })
        .join("")}
    </svg>
  `;

  const sizeLabel =
    String(state.currentConfig.page.size || "").toLowerCase() === "custom"
      ? `${formatMm(pageSize.width)} x ${formatMm(pageSize.height)} mm`
      : `${state.currentConfig.page.size} ${state.currentConfig.page.orientation}`;

  elements.pageSetupMeta.textContent = overflowCount
    ? `${slotCount} card slots shown. ${overflowCount} extend past the page margins with the current settings.`
    : `${slotCount} card slots shown on ${sizeLabel}. All rectangles fit inside the page margins.`;
}

function resolvePageSizeMm(config) {
  const preset = PAGE_SIZE_PRESETS_MM[String(config.page.size || "").toLowerCase()];
  let width = preset ? preset.width : numberOr(config.page.widthMm, 210);
  let height = preset ? preset.height : numberOr(config.page.heightMm, 297);

  if (String(config.page.orientation || "").toLowerCase() === "landscape") {
    return {
      width: Math.max(width, height),
      height: Math.min(width, height),
    };
  }

  return {
    width: Math.min(width, height),
    height: Math.max(width, height),
  };
}

function resolveGridMm(config, pageSize) {
  const margin = Math.max(0, numberOr(config.page.marginMm, 10));
  const gap = Math.max(0, numberOr(config.page.gapMm, 6));
  const cardWidth = Math.max(0.1, numberOr(config.card.widthMm, 85.6));
  const cardHeight = Math.max(0.1, numberOr(config.card.heightMm, 54));
  const columns = Math.max(1, integerOr(config.layout.columns, 1));
  const rows = Math.max(1, integerOr(config.layout.rows, 1));
  const usableWidth = pageSize.width - margin * 2;
  const usableHeight = pageSize.height - margin * 2;
  const gridWidth = columns * cardWidth + Math.max(0, columns - 1) * gap;
  const gridHeight = rows * cardHeight + Math.max(0, rows - 1) * gap;

  return {
    margin,
    gap,
    cardWidth,
    cardHeight,
    columns,
    rows,
    offsetX: margin + (config.layout.centerOnPage ? Math.max(0, (usableWidth - gridWidth) / 2) : 0),
    offsetY: margin + (config.layout.centerOnPage ? Math.max(0, (usableHeight - gridHeight) / 2) : 0),
  };
}

function resolveCardRectTopMm(slot, grid) {
  const column = slot % grid.columns;
  const row = Math.floor(slot / grid.columns);

  return {
    x: grid.offsetX + column * (grid.cardWidth + grid.gap),
    y: grid.offsetY + row * (grid.cardHeight + grid.gap),
    width: grid.cardWidth,
    height: grid.cardHeight,
  };
}

function formatMm(value) {
  return Number(value).toFixed(Number(value) >= 100 ? 0 : 1).replace(/\.0$/, "");
}

function handleNumberInputWheel(event) {
  const target = event.target;
  if (!(target instanceof HTMLInputElement) || target.type !== "number" || target.disabled || target.readOnly) {
    return;
  }

  if (event.deltaY === 0) {
    return;
  }

  event.preventDefault();

  if (!Number.isFinite(target.valueAsNumber) && !target.value.trim()) {
    target.value = target.min || "0";
  }

  if (event.deltaY < 0) {
    target.stepUp();
  } else {
    target.stepDown();
  }

  target.dispatchEvent(new Event("input", { bubbles: true }));
}

function handleCollapsibleToggle(event) {
  const target = event.target;
  if (!(target instanceof HTMLDetailsElement)) {
    return;
  }

  const group = target.dataset.uiGroup;
  const key = target.dataset.uiKey;
  if (!group || !key) {
    return;
  }

  if (!state.uiOpen[group]) {
    state.uiOpen[group] = {};
  }

  state.uiOpen[group][key] = target.open;
}

function captureUiState() {
  const detailsElements = elements.configGui.querySelectorAll("details[data-ui-group][data-ui-key]");
  detailsElements.forEach((detailsElement) => {
    const group = detailsElement.dataset.uiGroup;
    const key = detailsElement.dataset.uiKey;
    if (!state.uiOpen[group]) {
      state.uiOpen[group] = {};
    }
    state.uiOpen[group][key] = detailsElement.open;
  });
}

function isUiOpen(group, key, fallbackOpen = false) {
  if (!state.uiOpen[group]) {
    return fallbackOpen;
  }

  if (Object.prototype.hasOwnProperty.call(state.uiOpen[group], key)) {
    return Boolean(state.uiOpen[group][key]);
  }

  return fallbackOpen;
}

function handleConfigGuiInput(event) {
  const target = event.target;
  if (!(target instanceof HTMLInputElement || target instanceof HTMLSelectElement || target instanceof HTMLTextAreaElement)) {
    return;
  }

  if (target.dataset.path) {
    setValueAtPath(state.currentConfig, target.dataset.path, parseInputValue(target));
    finishConfigUpdate();
    return;
  }

  if (target.dataset.ruleIndex !== undefined) {
    updateRuleFromInput(target);
    finishConfigUpdate(target.dataset.ruleField === "mode");
    return;
  }

  if (target.dataset.elementIndex !== undefined) {
    updateElementFromInput(target);
    finishConfigUpdate(target.dataset.elementField === "type");
  }
}

function handleConfigGuiClick(event) {
  const actionButton = event.target.closest("[data-action]");
  if (!actionButton) {
    return;
  }

  const action = actionButton.dataset.action;

  if (action === "add-rule") {
    state.currentConfig.card.background.rules.push(createBackgroundRule());
    finishConfigUpdate(true);
    return;
  }

  if (action === "remove-rule") {
    const index = Number(actionButton.dataset.ruleIndex);
    state.currentConfig.card.background.rules.splice(index, 1);
    finishConfigUpdate(true);
    return;
  }

  if (action === "add-text-element") {
    state.currentConfig.elements.push(createElement("text", state.currentConfig.elements.length));
    finishConfigUpdate(true);
    return;
  }

  if (action === "add-photo-element") {
    state.currentConfig.elements.push(createElement("photo", state.currentConfig.elements.length));
    finishConfigUpdate(true);
    return;
  }

  if (action === "duplicate-element") {
    const index = Number(actionButton.dataset.elementIndex);
    const duplicate = cloneDeep(state.currentConfig.elements[index]);
    duplicate.id = `${duplicate.id || "element"}-copy`;
    state.currentConfig.elements.splice(index + 1, 0, normalizeElement(duplicate, index + 1));
    finishConfigUpdate(true);
    return;
  }

  if (action === "remove-element") {
    const index = Number(actionButton.dataset.elementIndex);
    state.currentConfig.elements.splice(index, 1);
    finishConfigUpdate(true);
  }
}

function finishConfigUpdate(shouldRender = false) {
  syncJsonEditor();
  saveConfig();

  if (shouldRender) {
    renderConfigGui();
  }

  renderPageSetupPreview();
  queueLivePreview();
}

function updateRuleFromInput(target) {
  const index = Number(target.dataset.ruleIndex);
  const field = target.dataset.ruleField;
  const rule = state.currentConfig.card.background.rules[index];

  if (!rule) {
    return;
  }

  if (field === "mode") {
    setRuleMode(rule, target.value);
    return;
  }

  if (field === "conditionValue") {
    const mode = inferRuleMode(rule);
    if (mode === "contains") {
      rule.contains = target.value;
    } else if (mode === "regex") {
      rule.regex = target.value;
    } else {
      rule.equals = target.value;
    }
    return;
  }

  if (field === "flags") {
    rule.flags = target.value;
    return;
  }

  rule[field] = parseInputValue(target);
}

function updateElementFromInput(target) {
  const index = Number(target.dataset.elementIndex);
  const field = target.dataset.elementField;
  const elementConfig = state.currentConfig.elements[index];

  if (!elementConfig) {
    return;
  }

  elementConfig[field] = parseInputValue(target);
}

function parseInputValue(target) {
  if (target.type === "checkbox") {
    return target.checked;
  }

  const dataType = target.dataset.type || target.type;

  if (dataType === "number") {
    const parsed = Number(target.value);
    return Number.isFinite(parsed) ? parsed : 0;
  }

  if (dataType === "integer") {
    const parsed = Number.parseInt(target.value, 10);
    return Number.isFinite(parsed) ? parsed : 0;
  }

  if (dataType === "csv") {
    return target.value
      .split(",")
      .map((item) => item.trim())
      .filter(Boolean);
  }

  return target.value;
}

function renderConfigGui() {
  if (!state.currentConfig) {
    return;
  }

  if (elements.configGui.childElementCount) {
    captureUiState();
  }

  renderSheetOptions();

  const config = state.currentConfig;
  const availableSheetCopy = state.availableSheets.length
    ? `Available sheets: ${state.availableSheets.map(escapeHtml).join(", ")}.`
    : "Leave blank to use the first sheet in the workbook.";

  elements.configGui.innerHTML = `
    <div class="builder-grid">
      ${renderCollapsibleSection(
        "spreadsheet",
        "Spreadsheet",
        "Choose which sheet to read and how to treat empty rows.",
        `
          <div class="builder-subgrid columns-2">
            ${renderPathField("spreadsheet.sheetName", "Sheet Name", config.spreadsheet.sheetName, {
              list: "sheet-options",
              hint: availableSheetCopy,
              placeholder: "First sheet by default",
            })}
            ${renderCheckboxPathField("spreadsheet.skipEmptyRows", "Skip Empty Rows", config.spreadsheet.skipEmptyRows, {
              hint: "Recommended for cleaner imports.",
            })}
          </div>
        `,
        true
      )}

      ${renderCollapsibleSection(
        "page",
        "Page",
        "Set the PDF paper size, orientation, and page spacing.",
        `
          <div class="builder-subgrid columns-3">
            ${renderSelectPathField("page.size", "Page Size", config.page.size, ["A4", "Letter", "Legal", "Custom"])}
            ${renderSelectPathField("page.orientation", "Orientation", config.page.orientation, ["portrait", "landscape"])}
            ${renderNumberPathField("page.marginMm", "Margin (mm)", config.page.marginMm, { min: 0, step: 0.1 })}
            ${renderNumberPathField("page.gapMm", "Gap (mm)", config.page.gapMm, { min: 0, step: 0.1 })}
            ${renderNumberPathField("page.widthMm", "Custom Width (mm)", config.page.widthMm, { min: 1, step: 0.1 })}
            ${renderNumberPathField("page.heightMm", "Custom Height (mm)", config.page.heightMm, { min: 1, step: 0.1 })}
          </div>
        `
      )}

      ${renderCollapsibleSection(
        "layout",
        "Grid Layout",
        "Control how many cards fit onto each PDF page.",
        `
          <div class="builder-subgrid columns-3">
            ${renderNumberPathField("layout.columns", "Columns", config.layout.columns, { min: 1, step: 1, type: "integer" })}
            ${renderNumberPathField("layout.rows", "Rows", config.layout.rows, { min: 1, step: 1, type: "integer" })}
            ${renderCheckboxPathField("layout.centerOnPage", "Center Grid", config.layout.centerOnPage, {
              hint: "Centers the card block inside the page margins.",
            })}
          </div>
        `
      )}

      ${renderCollapsibleSection(
        "card",
        "Card",
        "Set the base card size, fill color, and border styling.",
        `
          <div class="builder-subgrid columns-3">
            ${renderNumberPathField("card.widthMm", "Card Width (mm)", config.card.widthMm, { min: 1, step: 0.1 })}
            ${renderNumberPathField("card.heightMm", "Card Height (mm)", config.card.heightMm, { min: 1, step: 0.1 })}
            ${renderNumberPathField("card.borderWidthMm", "Border Width (mm)", config.card.borderWidthMm, { min: 0, step: 0.05 })}
            ${renderColorPathField("card.fillColor", "Card Fill", config.card.fillColor)}
            ${renderColorPathField("card.borderColor", "Border Color", config.card.borderColor)}
          </div>
        `
      )}

      ${renderCollapsibleSection(
        "photos",
        "Photo Matching",
        "Choose which spreadsheet columns should be used to find matching photo filenames.",
        `
          <div class="builder-subgrid columns-2">
            ${renderPathField("photos.fields", "Match Fields", (config.photos.fields || []).join(", "), {
              dataType: "csv",
              hint: "Comma-separated. Example: name, id",
            })}
            ${renderPathField("photos.placeholderLabel", "Missing Photo Label", config.photos.placeholderLabel, {
              hint: "Shown when no matching photo is found.",
            })}
          </div>
        `
      )}

      ${renderCollapsibleSection(
        "backgrounds",
        "Background Rules",
        "Swap card backgrounds based on row data like duty, team, or department.",
        `
          <div class="builder-actions section-actions">
            <button class="button ghost compact" type="button" data-action="add-rule">Add Rule</button>
          </div>
          <div class="builder-subgrid columns-2">
            ${renderPathField("card.background.default", "Default Background Name", config.card.background.default, {
              hint: "Match against a background filename without worrying about case.",
            })}
          </div>
          <div class="editor-list">
            ${
              config.card.background.rules.length
                ? config.card.background.rules.map((rule, index) => renderRuleCard(rule, index)).join("")
                : `<div class="empty-state">No background rules yet. Add one if different duties need different designs.</div>`
            }
          </div>
        `
      )}

      ${renderCollapsibleSection(
        "elements",
        "Card Elements",
        "These blocks decide what appears on each card and exactly where it sits.",
        `
          <div class="builder-actions section-actions">
            <button class="button ghost compact" type="button" data-action="add-text-element">Add Text</button>
            <button class="button ghost compact" type="button" data-action="add-photo-element">Add Photo</button>
          </div>
          <div class="editor-list">
            ${
              config.elements.length
                ? config.elements.map((item, index) => renderElementCard(item, index)).join("")
                : `<div class="empty-state">No card elements yet. Add a text or photo block to start composing the card.</div>`
            }
          </div>
        `,
        true
      )}
    </div>
  `;
}

function renderSheetOptions() {
  elements.sheetOptions.innerHTML = (state.availableSheets || [])
    .map((sheetName) => `<option value="${escapeAttribute(sheetName)}"></option>`)
    .join("");
}

function renderCollapsibleSection(key, title, description, body, defaultOpen = false) {
  return `
    <details class="builder-section collapsible-panel" data-ui-group="sections" data-ui-key="${escapeAttribute(key)}" ${
      isUiOpen("sections", key, defaultOpen) ? "open" : ""
    }>
      <summary class="collapsible-summary">
        <div class="collapsible-copy">
          <h3>${escapeHtml(title)}</h3>
          <p>${escapeHtml(description)}</p>
        </div>
      </summary>
      <div class="collapsible-body">
        ${body}
      </div>
    </details>
  `;
}

function renderRuleCard(rule, index) {
  const mode = inferRuleMode(rule);
  const conditionValue = getRuleConditionValue(rule);
  const key = `rule-${index}`;
  const summaryCopy = rule.field
    ? `${escapeHtml(rule.field)} -> ${escapeHtml(rule.background || "background")}`
    : `Rule ${index + 1}`;

  return `
    <details class="editor-card collapsible-card" data-ui-group="rules" data-ui-key="${escapeAttribute(key)}" ${
      isUiOpen("rules", key, false) ? "open" : ""
    }>
      <summary class="collapsible-summary card-summary">
        <div class="collapsible-copy">
          <h4>${summaryCopy}</h4>
          <p>Match type: ${escapeHtml(mode)}</p>
        </div>
      </summary>
      <div class="collapsible-body">
        <div class="builder-actions section-actions">
          <button class="button ghost compact danger" type="button" data-action="remove-rule" data-rule-index="${index}">Remove Rule</button>
        </div>
        <div class="builder-subgrid columns-2">
          ${renderRuleField(index, "field", "Spreadsheet Field", rule.field)}
          ${renderRuleSelect(index, "mode", "Match Type", mode, ["equals", "contains", "regex"])}
          ${renderRuleField(index, "conditionValue", mode === "regex" ? "Pattern" : "Match Value", conditionValue, {
            hint: mode === "regex" ? "Example: ^manager$" : "Example: Manager",
          })}
          ${renderRuleField(index, "background", "Background Name", rule.background, {
            hint: "Match a filename from the uploaded background folder.",
          })}
          ${
            mode === "regex"
              ? renderRuleField(index, "flags", "Regex Flags", rule.flags || "i", {
                  hint: "Typical value: i",
                })
              : ""
          }
        </div>
      </div>
    </details>
  `;
}

function renderElementCard(elementConfig, index) {
  const typeLabel = elementConfig.type === "photo" ? "Photo Block" : "Text Block";
  const key = `element-${index}-${elementConfig.id || "item"}`;
  const summaryTitle = elementConfig.id || `${elementConfig.type}-${index + 1}`;

  return `
    <details class="editor-card collapsible-card" data-ui-group="elements" data-ui-key="${escapeAttribute(key)}" ${
      isUiOpen("elements", key, index < 2) ? "open" : ""
    }>
      <summary class="collapsible-summary card-summary">
        <div class="collapsible-copy">
          <h4>${escapeHtml(summaryTitle)}</h4>
          <p>${escapeHtml(typeLabel)}</p>
        </div>
      </summary>
      <div class="collapsible-body">
        <div class="builder-actions section-actions">
          <button class="button ghost compact" type="button" data-action="duplicate-element" data-element-index="${index}">Duplicate</button>
          <button class="button ghost compact danger" type="button" data-action="remove-element" data-element-index="${index}">Remove</button>
        </div>

        <div class="builder-subgrid columns-3">
          ${renderElementField(index, "id", "Element ID", elementConfig.id)}
          ${renderElementSelect(index, "type", "Type", elementConfig.type, ["text", "photo"])}
          ${renderElementField(index, "field", "Spreadsheet Field", elementConfig.field)}
          ${renderNumberElementField(index, "xMm", "X (mm)", elementConfig.xMm)}
          ${renderNumberElementField(index, "yMm", "Y (mm)", elementConfig.yMm)}
          ${renderNumberElementField(index, "widthMm", "Width (mm)", elementConfig.widthMm)}
          ${renderNumberElementField(index, "heightMm", "Height (mm)", elementConfig.heightMm)}
          ${renderNumberElementField(index, "paddingMm", "Padding (mm)", elementConfig.paddingMm)}
          ${renderNumberElementField(index, "borderWidthMm", "Border Width (mm)", elementConfig.borderWidthMm, { step: 0.05, min: 0 })}
          ${renderColorElementField(index, "fillColor", "Fill Color", elementConfig.fillColor)}
          ${renderColorElementField(index, "borderColor", "Border Color", elementConfig.borderColor)}
          ${
            elementConfig.type === "photo"
              ? renderElementSelect(index, "fit", "Image Fit", elementConfig.fit, ["cover", "contain"])
              : renderColorElementField(index, "color", "Text Color", elementConfig.color)
          }
        </div>

        ${
          elementConfig.type === "text"
            ? `
              <div class="builder-subgrid columns-3 element-text-grid">
                ${renderElementField(index, "template", "Template", elementConfig.template, {
                  hint: "Optional. Example: {{name}} / {{id}}",
                  wide: true,
                })}
                ${renderElementField(index, "value", "Fixed Value", elementConfig.value)}
                ${renderElementField(index, "prefix", "Prefix", elementConfig.prefix)}
                ${renderElementField(index, "suffix", "Suffix", elementConfig.suffix)}
                ${renderFontElementField(index, "font", "Font", elementConfig.font)}
                ${renderNumberElementField(index, "fontSize", "Font Size", elementConfig.fontSize, { min: 1, step: 0.1 })}
                ${renderNumberElementField(index, "minFontSize", "Min Font Size", elementConfig.minFontSize, { min: 1, step: 0.1 })}
                ${renderNumberElementField(index, "lineHeight", "Line Height", elementConfig.lineHeight, { min: 0.9, step: 0.05 })}
                ${renderNumberElementField(index, "maxLines", "Max Lines", elementConfig.maxLines, { min: 1, step: 1, type: "integer" })}
                ${renderElementSelect(index, "align", "Horizontal Align", elementConfig.align, ["left", "center", "right"])}
                ${renderElementSelect(index, "valign", "Vertical Align", elementConfig.valign, ["top", "center", "bottom"])}
                ${renderCheckboxElementField(index, "uppercase", "Uppercase", elementConfig.uppercase)}
                ${renderCheckboxElementField(index, "hideWhenEmpty", "Hide When Empty", elementConfig.hideWhenEmpty)}
              </div>
            `
            : ""
        }
      </div>
    </details>
  `;
}

function renderPathField(path, label, value, options = {}) {
  return renderInputField(label, {
    input: "text",
    value,
    path,
    hint: options.hint,
    list: options.list,
    placeholder: options.placeholder,
    dataType: options.dataType,
    wide: options.wide,
  });
}

function renderNumberPathField(path, label, value, options = {}) {
  return renderInputField(label, {
    input: "number",
    value,
    path,
    hint: options.hint,
    min: options.min,
    max: options.max,
    step: options.step || 0.1,
    dataType: options.type || "number",
    wide: options.wide,
  });
}

function renderColorPathField(path, label, value, options = {}) {
  return renderInputField(label, {
    input: "color",
    value: safeColor(value),
    path,
    hint: options.hint,
  });
}

function renderCheckboxPathField(path, label, checked, options = {}) {
  return renderCheckboxField(label, checked, {
    path,
    hint: options.hint,
    wide: options.wide,
  });
}

function renderSelectPathField(path, label, value, choices, options = {}) {
  return renderSelectField(label, choices, value, {
    path,
    hint: options.hint,
    wide: options.wide,
  });
}

function renderRuleField(index, field, label, value, options = {}) {
  return renderInputField(label, {
    input: "text",
    value,
    hint: options.hint,
    wide: options.wide,
    attributes: {
      "data-rule-index": index,
      "data-rule-field": field,
    },
  });
}

function renderRuleSelect(index, field, label, value, choices, options = {}) {
  return renderSelectField(label, choices, value, {
    hint: options.hint,
    attributes: {
      "data-rule-index": index,
      "data-rule-field": field,
    },
  });
}

function renderElementField(index, field, label, value, options = {}) {
  return renderInputField(label, {
    input: "text",
    value,
    hint: options.hint,
    list: options.list,
    wide: options.wide,
    attributes: {
      "data-element-index": index,
      "data-element-field": field,
    },
  });
}

function renderFontElementField(index, field, label, value, options = {}) {
  return renderSelectField(label, buildFontChoices(value), value, {
    hint: options.hint || "Expanded list of supported families plus bold and italic variants.",
    attributes: {
      "data-element-index": index,
      "data-element-field": field,
    },
  });
}

function renderNumberElementField(index, field, label, value, options = {}) {
  return renderInputField(label, {
    input: "number",
    value,
    hint: options.hint,
    min: options.min,
    max: options.max,
    step: options.step || 0.1,
    dataType: options.type || "number",
    wide: options.wide,
    attributes: {
      "data-element-index": index,
      "data-element-field": field,
    },
  });
}

function renderColorElementField(index, field, label, value, options = {}) {
  return renderInputField(label, {
    input: "color",
    value: safeColor(value),
    hint: options.hint,
    attributes: {
      "data-element-index": index,
      "data-element-field": field,
    },
  });
}

function renderElementSelect(index, field, label, value, choices, options = {}) {
  return renderSelectField(label, choices, value, {
    hint: options.hint,
    attributes: {
      "data-element-index": index,
      "data-element-field": field,
    },
  });
}

function renderCheckboxElementField(index, field, label, checked, options = {}) {
  return renderCheckboxField(label, checked, {
    hint: options.hint,
    attributes: {
      "data-element-index": index,
      "data-element-field": field,
    },
  });
}

function renderInputField(label, options) {
  const classes = ["builder-field"];
  if (options.wide) {
    classes.push("wide");
  }

  return `
    <label class="${classes.join(" ")}">
      <span>${escapeHtml(label)}</span>
      <input
        class="builder-input"
        type="${options.input}"
        value="${escapeAttribute(options.value ?? "")}"
        ${options.path ? `data-path="${escapeAttribute(options.path)}"` : ""}
        ${options.dataType ? `data-type="${escapeAttribute(options.dataType)}"` : ""}
        ${options.list ? `list="${escapeAttribute(options.list)}"` : ""}
        ${options.placeholder ? `placeholder="${escapeAttribute(options.placeholder)}"` : ""}
        ${options.min !== undefined ? `min="${escapeAttribute(options.min)}"` : ""}
        ${options.max !== undefined ? `max="${escapeAttribute(options.max)}"` : ""}
        ${options.step !== undefined ? `step="${escapeAttribute(options.step)}"` : ""}
        ${renderAttributeMap(options.attributes)}
      />
      ${options.hint ? `<small>${escapeHtml(options.hint)}</small>` : ""}
    </label>
  `;
}

function renderSelectField(label, choices, value, options = {}) {
  const classes = ["builder-field"];
  if (options.wide) {
    classes.push("wide");
  }

  return `
    <label class="${classes.join(" ")}">
      <span>${escapeHtml(label)}</span>
      <select class="builder-input" ${options.path ? `data-path="${escapeAttribute(options.path)}"` : ""} ${renderAttributeMap(options.attributes)}>
        ${choices
          .map(
            (choice) =>
              `<option value="${escapeAttribute(choice)}" ${String(choice) === String(value) ? "selected" : ""}>${escapeHtml(choice)}</option>`
          )
          .join("")}
      </select>
      ${options.hint ? `<small>${escapeHtml(options.hint)}</small>` : ""}
    </label>
  `;
}

function renderCheckboxField(label, checked, options = {}) {
  const classes = ["builder-field", "toggle-field"];
  if (options.wide) {
    classes.push("wide");
  }

  return `
    <label class="${classes.join(" ")}">
      <span>${escapeHtml(label)}</span>
      <div class="toggle-shell">
        <input
          type="checkbox"
          ${checked ? "checked" : ""}
          ${options.path ? `data-path="${escapeAttribute(options.path)}"` : ""}
          ${renderAttributeMap(options.attributes)}
        />
        <em>Enabled</em>
      </div>
      ${options.hint ? `<small>${escapeHtml(options.hint)}</small>` : ""}
    </label>
  `;
}

function renderAttributeMap(attributes = {}) {
  return Object.entries(attributes)
    .map(([key, value]) => `${key}="${escapeAttribute(value)}"`)
    .join(" ");
}

function buildFontChoices(currentValue) {
  const normalizedCurrent = String(currentValue || "").trim();
  if (!normalizedCurrent) {
    return FONT_CHOICES;
  }

  return FONT_CHOICES.includes(normalizedCurrent) ? FONT_CHOICES : [normalizedCurrent, ...FONT_CHOICES];
}

function setRuleMode(rule, nextMode) {
  const rememberedValue = getRuleConditionValue(rule);
  delete rule.equals;
  delete rule.value;
  delete rule.contains;
  delete rule.regex;
  delete rule.flags;
  rule.mode = nextMode;

  if (nextMode === "contains") {
    rule.contains = rememberedValue;
    return;
  }

  if (nextMode === "regex") {
    rule.regex = rememberedValue;
    rule.flags = "i";
    return;
  }

  rule.equals = rememberedValue;
}

function getRuleConditionValue(rule) {
  return rule.equals ?? rule.value ?? rule.contains ?? rule.regex ?? "";
}

function inferRuleMode(rule) {
  if (rule.mode === "contains" || rule.mode === "regex" || rule.mode === "equals") {
    return rule.mode;
  }

  if (rule.regex !== undefined) {
    return "regex";
  }

  if (rule.contains !== undefined) {
    return "contains";
  }

  return "equals";
}

function serializeConfig(config) {
  const cloned = cloneDeep(config);

  cloned.card.background.rules = (cloned.card.background.rules || []).map((rule) => {
    const mode = inferRuleMode(rule);
    const cleaned = {
      field: String(rule.field || ""),
      background: String(rule.background || ""),
    };

    if (mode === "contains") {
      cleaned.contains = String(rule.contains || "");
      return cleaned;
    }

    if (mode === "regex") {
      cleaned.regex = String(rule.regex || "");
      cleaned.flags = String(rule.flags || "i");
      return cleaned;
    }

    cleaned.equals = String(rule.equals ?? rule.value ?? "");
    return cleaned;
  });

  cloned.photos.fields = Array.isArray(cloned.photos.fields) ? cloned.photos.fields.map((item) => String(item).trim()).filter(Boolean) : [];
  cloned.elements = Array.isArray(cloned.elements) ? cloned.elements.map((item, index) => normalizeElement(item, index)) : [];
  return cloned;
}

function normalizeClientConfig(config = {}) {
  const merged = deepMerge(cloneDeep(state.defaultConfig || {}), cloneDeep(config || {}));

  merged.spreadsheet = merged.spreadsheet || {};
  merged.page = merged.page || {};
  merged.layout = merged.layout || {};
  merged.card = merged.card || {};
  merged.card.background = merged.card.background || {};
  merged.photos = merged.photos || {};

  merged.spreadsheet.sheetName = String(merged.spreadsheet.sheetName || "");
  merged.spreadsheet.skipEmptyRows = merged.spreadsheet.skipEmptyRows !== false;

  merged.page.size = String(merged.page.size || "A4");
  merged.page.orientation = String(merged.page.orientation || "portrait");
  merged.page.widthMm = numberOr(merged.page.widthMm, 210);
  merged.page.heightMm = numberOr(merged.page.heightMm, 297);
  merged.page.marginMm = numberOr(merged.page.marginMm, 10);
  merged.page.gapMm = numberOr(merged.page.gapMm, 6);

  merged.layout.columns = integerOr(merged.layout.columns, 2);
  merged.layout.rows = integerOr(merged.layout.rows, 5);
  merged.layout.centerOnPage = merged.layout.centerOnPage !== false;

  merged.card.widthMm = numberOr(merged.card.widthMm, 85.6);
  merged.card.heightMm = numberOr(merged.card.heightMm, 54);
  merged.card.fillColor = safeColor(merged.card.fillColor, "#FFFFFF");
  merged.card.borderColor = safeColor(merged.card.borderColor, "#D6D0C4");
  merged.card.borderWidthMm = numberOr(merged.card.borderWidthMm, 0.35);
  merged.card.background.default = String(merged.card.background.default || "");
  merged.card.background.rules = Array.isArray(merged.card.background.rules)
    ? merged.card.background.rules.map((rule) => normalizeRule(rule))
    : [];

  merged.photos.fields = Array.isArray(merged.photos.fields)
    ? merged.photos.fields.map((item) => String(item).trim()).filter(Boolean)
    : [];
  merged.photos.placeholderLabel = String(merged.photos.placeholderLabel || "NO PHOTO");

  merged.elements = Array.isArray(merged.elements)
    ? merged.elements.map((item, index) => normalizeElement(item, index))
    : [];

  return merged;
}

function normalizeRule(rule = {}) {
  const next = { ...rule };
  next.field = String(next.field || "");
  next.background = String(next.background || "");
  next.mode = inferRuleMode(next);

  if (next.mode === "contains") {
    next.contains = String(next.contains || "");
  } else if (next.mode === "regex") {
    next.regex = String(next.regex || "");
    next.flags = String(next.flags || "i");
  } else {
    next.equals = String(next.equals ?? next.value ?? "");
  }

  return next;
}

function normalizeElement(element = {}, index = 0) {
  const type = String(element.type || "text");
  return {
    id: String(element.id || `${type}-${index + 1}`),
    type,
    field: String(element.field || ""),
    template: String(element.template || ""),
    value: String(element.value || ""),
    prefix: String(element.prefix || ""),
    suffix: String(element.suffix || ""),
    xMm: numberOr(element.xMm, 0),
    yMm: numberOr(element.yMm, 0),
    widthMm: numberOr(element.widthMm, type === "photo" ? 24 : 40),
    heightMm: numberOr(element.heightMm, type === "photo" ? 30 : 10),
    paddingMm: numberOr(element.paddingMm, 0),
    borderColor: safeColor(element.borderColor || "#FFFFFF", "#FFFFFF"),
    borderWidthMm: numberOr(element.borderWidthMm, 0),
    fillColor: safeColor(element.fillColor || "#FFFFFF", "#FFFFFF"),
    fit: String(element.fit || "cover"),
    font: String(element.font || "Helvetica"),
    fontSize: numberOr(element.fontSize, 11),
    minFontSize: numberOr(element.minFontSize, 7),
    lineHeight: numberOr(element.lineHeight, 1.15),
    color: safeColor(element.color || "#111111", "#111111"),
    align: String(element.align || "left"),
    valign: String(element.valign || "top"),
    uppercase: Boolean(element.uppercase),
    maxLines: integerOr(element.maxLines, 99),
    hideWhenEmpty: Boolean(element.hideWhenEmpty),
  };
}

function createBackgroundRule() {
  return {
    field: "",
    mode: "equals",
    equals: "",
    background: "",
  };
}

function createElement(type, index) {
  if (type === "photo") {
    return normalizeElement(
      {
        id: `photo-${index + 1}`,
        type: "photo",
        xMm: 6,
        yMm: 8,
        widthMm: 24,
        heightMm: 30,
        fit: "cover",
        borderColor: "#FFFFFF",
        borderWidthMm: 0.5,
        fillColor: "#F3F0EA",
      },
      index
    );
  }

  return normalizeElement(
    {
      id: `text-${index + 1}`,
      type: "text",
      field: "",
      xMm: 10,
      yMm: 10,
      widthMm: 40,
      heightMm: 10,
      font: "Helvetica",
      fontSize: 11,
      minFontSize: 8,
      lineHeight: 1.1,
      color: "#111111",
      align: "left",
      valign: "top",
      maxLines: 2,
    },
    index
  );
}

function setValueAtPath(object, path, value) {
  const parts = path.split(".");
  let cursor = object;

  for (let index = 0; index < parts.length - 1; index += 1) {
    const key = parts[index];
    if (!cursor[key] || typeof cursor[key] !== "object") {
      cursor[key] = {};
    }
    cursor = cursor[key];
  }

  cursor[parts[parts.length - 1]] = value;
}

function deepMerge(base, override) {
  if (Array.isArray(base) || Array.isArray(override)) {
    return Array.isArray(override) ? override : Array.isArray(base) ? base : [];
  }

  if (!isPlainObject(base) || !isPlainObject(override)) {
    return override === undefined ? base : override;
  }

  const merged = { ...base };
  for (const [key, value] of Object.entries(override)) {
    merged[key] = key in base ? deepMerge(base[key], value) : value;
  }
  return merged;
}

function isPlainObject(value) {
  return Boolean(value) && typeof value === "object" && !Array.isArray(value);
}

function cloneDeep(value) {
  return JSON.parse(JSON.stringify(value));
}

function numberOr(value, fallback) {
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : fallback;
}

function integerOr(value, fallback) {
  const parsed = Number.parseInt(value, 10);
  return Number.isFinite(parsed) ? parsed : fallback;
}

function safeColor(value, fallback = "#111111") {
  const candidate = String(value || fallback).trim();
  return /^#[0-9a-fA-F]{6}$/.test(candidate) ? candidate : fallback;
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}

function escapeAttribute(value) {
  return escapeHtml(value).replaceAll("'", "&#39;");
}
