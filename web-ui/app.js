const state = {
  workerReady: false,
  isRunning: false,
  selectedFiles: [],
  results: [],
  errors: [],
  logs: [],
  metrics: null,
};

const dom = {
  folderInput: document.getElementById("folderInput"),
  fileInput: document.getElementById("fileInput"),
  runBtn: document.getElementById("runBtn"),
  zipBtn: document.getElementById("zipBtn"),
  status: document.getElementById("statusText"),
  pairStats: document.getElementById("pairStats"),
  countStats: document.getElementById("countStats"),
  metrics: document.getElementById("metricCards"),
  logBox: document.getElementById("logBox"),
  resultTableBody: document.getElementById("resultTableBody"),
  selectedFilesMeta: document.getElementById("selectedFilesMeta"),
  selectedFilesList: document.getElementById("selectedFilesList"),
  failFast: document.getElementById("failFast"),
  method2Enabled: document.getElementById("method2Enabled"),
  method3Enabled: document.getElementById("method3Enabled"),
  baselineOn: document.getElementById("baselineOn"),
  filterOn: document.getElementById("filterOn"),
  processingOrder: document.getElementById("processingOrder"),
  filterDomain: document.getElementById("filterDomain"),
  baselineMethod: document.getElementById("baselineMethod"),
  filterConfig: document.getElementById("filterConfig"),
  filterType: document.getElementById("filterType"),
  fLowHz: document.getElementById("fLowHz"),
  fHighHz: document.getElementById("fHighHz"),
  filterOrder: document.getElementById("filterOrder"),
};

const worker = new Worker("./worker.js");

function appendLog(level, message) {
  const line = `[${new Date().toLocaleTimeString()}] ${level.toUpperCase()} ${message}`;
  state.logs.push(line);
  dom.logBox.textContent = state.logs.slice(-200).join("\n");
  dom.logBox.scrollTop = dom.logBox.scrollHeight;
}

function setStatus(text) {
  dom.status.textContent = text;
}

function normalizeCandidateName(name) {
  return (name || "").split(/[\\/]/).pop() || "";
}

function deriveYName(xName) {
  const replaced = xName.replace("_X_", "_Y_");
  return replaced.replace(/_H1(?=\.xlsx$)/i, "_H2");
}

function isInputCandidate(name) {
  const lower = name.toLowerCase();
  if (!lower.endsWith(".xlsx")) return false;
  if (lower.startsWith("output_")) return false;
  if (lower.startsWith("~$")) return false;
  if (lower.endsWith("-manip.xlsx")) return false;
  return true;
}

function detectPairs(files) {
  const names = files.map((file) => normalizeCandidateName(file.name));
  const candidates = new Set(names.filter((name) => isInputCandidate(name)));
  const xFiles = [...candidates].filter((name) => name.includes("_X_") && /_H1(?=\.xlsx$)/i.test(name));
  const used = new Set();

  let found = 0;
  let missing = 0;
  for (const xName of xFiles) {
    const yName = deriveYName(xName);
    if (candidates.has(yName)) {
      found += 1;
      used.add(xName.toLowerCase());
      used.add(yName.toLowerCase());
    } else {
      missing += 1;
    }
  }

  let singles = 0;
  for (const name of candidates) {
    if (!used.has(name.toLowerCase())) {
      singles += 1;
    }
  }

  return {
    candidates: candidates.size,
    pairs: found,
    missing,
    singles,
  };
}

function base64ToBlob(base64, mimeType) {
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i += 1) {
    bytes[i] = binary.charCodeAt(i);
  }
  return new Blob([bytes], { type: mimeType });
}

function downloadBlob(blob, fileName) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = fileName;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

function renderMetrics() {
  dom.metrics.innerHTML = "";
  const metrics = state.metrics;

  if (!metrics) return;

  const entries = [
    ["Method2 Enabled", metrics.method2Enabled ? "yes" : "no"],
    ["Method3 Enabled", metrics.method3Enabled ? "yes" : "no"],
    ["Pairs Detected", metrics.pairsDetected ?? 0],
    ["Singles Detected", metrics.singlesDetected ?? 0],
    ["Method2 Processed", metrics.method2Processed ?? 0],
    ["Method2 Failed", metrics.method2Failed ?? 0],
    ["Method3 Produced", metrics.method3Produced ?? 0],
    ["Processed Total", metrics.processedTotal ?? metrics.pairsProcessed ?? 0],
    ["Failed Total", metrics.failedTotal ?? metrics.pairsFailed ?? 0],
    ["Missing Y", metrics.pairsMissing ?? 0],
  ];

  for (const [label, value] of entries) {
    const card = document.createElement("div");
    card.className = "metric-card";
    card.innerHTML = `<span>${label}</span><strong>${value}</strong>`;
    dom.metrics.appendChild(card);
  }
}

function renderResults() {
  dom.resultTableBody.innerHTML = "";

  state.results.forEach((item, idx) => {
    const tr = document.createElement("tr");

    const layerCount = item.metrics?.layerCount ?? "-";
    const surfaceBase = item.metrics?.surfaceBaseTotal_m ?? NaN;
    const surfaceRss = item.metrics?.surfaceProfileRSS_m ?? NaN;

    tr.innerHTML = `
      <td>${item.pairKey}</td>
      <td>${item.outputFileName}</td>
      <td>${layerCount}</td>
      <td>${Number.isFinite(surfaceBase) ? surfaceBase.toFixed(6) : "-"}</td>
      <td>${Number.isFinite(surfaceRss) ? surfaceRss.toFixed(6) : "-"}</td>
      <td><button class="download-one" data-index="${idx}">Download</button></td>
    `;

    dom.resultTableBody.appendChild(tr);
  });

  dom.resultTableBody.querySelectorAll(".download-one").forEach((btn) => {
    btn.addEventListener("click", () => {
      const index = Number(btn.dataset.index);
      const item = state.results[index];
      if (!item) return;
      const blob = base64ToBlob(
        item.outputBytesB64,
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
      );
      downloadBlob(blob, item.outputFileName);
    });
  });

  dom.zipBtn.disabled = !state.results.length || !state.workerReady;
}

function updateSelectionStats() {
  const stats = detectPairs(state.selectedFiles);
  dom.pairStats.textContent = `Pairs: ${stats.pairs} | Singles: ${stats.singles} | Missing Y: ${stats.missing}`;
  dom.countStats.textContent = `Candidate XLSX: ${stats.candidates}`;
}

function refreshButtons() {
  const canRun = state.workerReady && !state.isRunning && state.selectedFiles.length > 0;
  dom.runBtn.disabled = !canRun;
  dom.zipBtn.disabled = state.isRunning || !state.workerReady || state.results.length === 0;
  if (dom.folderInput) dom.folderInput.disabled = state.isRunning;
  if (dom.fileInput) dom.fileInput.disabled = state.isRunning;
  if (dom.failFast) dom.failFast.disabled = state.isRunning;
  if (dom.method2Enabled) dom.method2Enabled.disabled = state.isRunning;
  if (dom.method3Enabled) dom.method3Enabled.disabled = state.isRunning;
  if (dom.baselineOn) dom.baselineOn.disabled = state.isRunning;
  if (dom.filterOn) dom.filterOn.disabled = state.isRunning;
}

function formatBytes(bytes) {
  if (!Number.isFinite(bytes) || bytes <= 0) return "0.00 MB";
  return `${(bytes / (1024 * 1024)).toFixed(2)} MB`;
}

function parseNumberInput(value, fallback) {
  const parsed = Number(value);
  if (!Number.isFinite(parsed)) return fallback;
  return parsed;
}

function syncFilterInputs() {
  const baselineOn = !!dom.baselineOn?.checked;
  const filterOn = !!dom.filterOn?.checked;
  const cfg = (dom.filterConfig?.value || "bandpass").toLowerCase();
  const isHighpass = cfg === "highpass" || cfg === "high";
  const isLowpass = cfg === "lowpass" || cfg === "low";

  if (dom.baselineMethod) dom.baselineMethod.disabled = !baselineOn;
  if (dom.processingOrder) dom.processingOrder.disabled = !filterOn;
  if (dom.filterDomain) dom.filterDomain.disabled = !filterOn;
  if (dom.filterConfig) dom.filterConfig.disabled = !filterOn;
  if (dom.filterType) dom.filterType.disabled = !filterOn;
  if (dom.filterOrder) dom.filterOrder.disabled = !filterOn;
  if (dom.fLowHz) dom.fLowHz.disabled = !filterOn || isHighpass;
  if (dom.fHighHz) dom.fHighHz.disabled = !filterOn || isLowpass;
}

function renderSelectedFiles() {
  const listEl = dom.selectedFilesList;
  listEl.innerHTML = "";

  const files = state.selectedFiles;
  const totalBytes = files.reduce((sum, file) => sum + (file.size || 0), 0);
  dom.selectedFilesMeta.textContent = `${files.length} files | ${formatBytes(totalBytes)}`;

  if (!files.length) {
    const li = document.createElement("li");
    li.className = "empty";
    li.textContent = "No files selected.";
    listEl.appendChild(li);
    return;
  }

  const maxVisible = 120;
  files.slice(0, maxVisible).forEach((file) => {
    const li = document.createElement("li");
    li.innerHTML = `<span class="name">${normalizeCandidateName(file.name)}</span><span class="size">${formatBytes(
      file.size || 0
    )}</span>`;
    listEl.appendChild(li);
  });

  if (files.length > maxVisible) {
    const li = document.createElement("li");
    li.className = "empty";
    li.textContent = `... ${files.length - maxVisible} more files`;
    listEl.appendChild(li);
  }
}

function collectSelectedFiles() {
  const combined = [...Array.from(dom.folderInput.files || []), ...Array.from(dom.fileInput.files || [])];
  const unique = new Map();

  combined.forEach((file) => {
    const key = `${normalizeCandidateName(file.name)}::${file.size}::${file.lastModified}`;
    if (!unique.has(key)) {
      unique.set(key, file);
    }
  });

  state.selectedFiles = Array.from(unique.values());
}

function handleSelectionChange(sourceLabel) {
  collectSelectedFiles();
  const folderCount = (dom.folderInput.files || []).length;
  const fileCount = (dom.fileInput.files || []).length;
  appendLog(
    "info",
    `${sourceLabel} updated | folder items: ${folderCount}, file items: ${fileCount}, unique total: ${state.selectedFiles.length}`
  );
  updateSelectionStats();
  renderSelectedFiles();
  refreshButtons();
}

async function runBatch() {
  if (!state.workerReady || !state.selectedFiles.length || state.isRunning) {
    return;
  }

  state.isRunning = true;
  state.results = [];
  state.errors = [];
  state.metrics = null;
  renderResults();
  renderMetrics();
  refreshButtons();

  setStatus("Preparing files...");
  appendLog("info", `Preparing ${state.selectedFiles.length} files`);

  const selectedCandidates = state.selectedFiles.filter((file) => {
    const name = normalizeCandidateName(file.name);
    return isInputCandidate(name);
  });

  const filesPayload = [];
  for (const file of selectedCandidates) {
    const bytes = await file.arrayBuffer();
    filesPayload.push({ name: normalizeCandidateName(file.name), bytes });
  }

  if (!filesPayload.length) {
    appendLog("warning", "No valid candidate XLSX files to process.");
    setStatus("No candidate XLSX files");
    state.isRunning = false;
    refreshButtons();
    return;
  }

  const transferables = filesPayload.map((item) => item.bytes);

  worker.postMessage(
    {
      type: "runBatch",
      payload: {
        files: filesPayload,
        options: {
          methods: "strain_legacy",
          failFast: dom.failFast.checked,
          includeManip: false,
          method2Enabled: dom.method2Enabled.checked,
          method3Enabled: dom.method3Enabled.checked,
          processingOrder: dom.processingOrder.value,
          filterDomain: dom.filterDomain.value,
          baselineMethod: dom.baselineMethod.value,
          filterOn: dom.filterOn.checked,
          baselineOn: dom.baselineOn.checked,
          filterConfig: dom.filterConfig.value,
          filterType: dom.filterType.value,
          fLowHz: parseNumberInput(dom.fLowHz.value, 0.1),
          fHighHz: parseNumberInput(dom.fHighHz.value, 25),
          filterOrder: Math.max(1, Math.round(parseNumberInput(dom.filterOrder.value, 4))),
          // Legacy fields kept for backward compatibility in older workers.
          highpassEnabled: dom.filterOn.checked && dom.filterConfig.value === "highpass",
          highpassCutoffHz: parseNumberInput(dom.fHighHz.value, 25),
          highpassTransitionHz: 0.02,
        },
      },
    },
    transferables
  );

  setStatus("Running...");
}

function requestZipDownload() {
  if (!state.results.length || !state.workerReady) return;
  setStatus("Building ZIP...");

  worker.postMessage({
    type: "zipOutputs",
    payload: {
      fileName: "deepsoil_total_displacement_outputs.zip",
      results: state.results.map((item) => ({
        outputFileName: item.outputFileName,
        outputBytesB64: item.outputBytesB64,
      })),
    },
  });
}

worker.addEventListener("message", (event) => {
  const { type, payload } = event.data || {};

  if (type === "status") {
    setStatus(payload.message || "...");
    appendLog(payload.phase || "info", payload.message || "");
    return;
  }

  if (type === "initialized") {
    state.workerReady = true;
    setStatus("Ready");
    appendLog("info", "Worker initialized");
    refreshButtons();
    return;
  }

  if (type === "runBatchResult") {
    state.isRunning = false;
    state.results = payload.results || [];
    state.errors = payload.errors || [];
    state.metrics = payload.metrics || null;

    (payload.logs || []).forEach((log) => appendLog(log.level || "info", log.message || ""));
    state.errors.forEach((err) => appendLog("error", `${err.pairKey}: ${err.reason}`));

    renderResults();
    renderMetrics();
    refreshButtons();

    const processed = state.metrics?.processedTotal ?? state.metrics?.pairsProcessed ?? state.results.length;
    const failed = state.metrics?.failedTotal ?? state.metrics?.pairsFailed ?? state.errors.length;
    setStatus(`Done | Processed: ${processed}, Failed: ${failed}`);
    return;
  }

  if (type === "zipReady") {
    const zipBlob = base64ToBlob(payload.zipBytesB64, "application/zip");
    downloadBlob(zipBlob, payload.fileName || "outputs.zip");
    setStatus("ZIP downloaded");
    appendLog("info", "ZIP created and downloaded");
    refreshButtons();
    return;
  }

  if (type === "error") {
    state.isRunning = false;
    const message = payload?.message || "Unknown worker error";
    setStatus(`Error: ${message}`);
    appendLog("error", message);
    refreshButtons();
  }
});

dom.folderInput.addEventListener("change", () => handleSelectionChange("Folder Select"));
dom.fileInput.addEventListener("change", () => handleSelectionChange("File Select"));

dom.filterConfig.addEventListener("change", syncFilterInputs);
dom.filterOn.addEventListener("change", syncFilterInputs);
dom.baselineOn.addEventListener("change", syncFilterInputs);

dom.runBtn.addEventListener("click", () => {
  runBatch().catch((error) => {
    state.isRunning = false;
    const message = error instanceof Error ? error.message : String(error);
    setStatus(`Error: ${message}`);
    appendLog("error", message);
    refreshButtons();
  });
});

dom.zipBtn.addEventListener("click", requestZipDownload);

setStatus("Initializing worker...");
appendLog("info", "Initializing Pyodide worker...");
renderSelectedFiles();
syncFilterInputs();
worker.postMessage({ type: "initialize" });
