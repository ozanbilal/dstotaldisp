const state = {
  workerReady: false,
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
  includeManip: document.getElementById("includeManip"),
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

function isInputCandidate(name, includeManip) {
  const lower = name.toLowerCase();
  if (!lower.endsWith(".xlsx")) return false;
  if (lower.startsWith("output_")) return false;
  if (lower.startsWith("~$")) return false;
  if (!includeManip && lower.endsWith("-manip.xlsx")) return false;
  return true;
}

function detectPairs(files, includeManip) {
  const names = files.map((file) => normalizeCandidateName(file.name));
  const candidates = new Set(names.filter((name) => isInputCandidate(name, includeManip)));
  const xFiles = [...candidates].filter((name) => name.includes("_X_") && /_H1(?=\.xlsx$)/i.test(name));

  let found = 0;
  let missing = 0;
  for (const xName of xFiles) {
    if (candidates.has(deriveYName(xName))) {
      found += 1;
    } else {
      missing += 1;
    }
  }

  return {
    candidates: candidates.size,
    pairs: found,
    missing,
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
    ["Pairs Detected", metrics.pairsDetected ?? 0],
    ["Processed", metrics.pairsProcessed ?? 0],
    ["Failed", metrics.pairsFailed ?? 0],
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
  const includeManip = dom.includeManip.checked;
  const stats = detectPairs(state.selectedFiles, includeManip);
  dom.pairStats.textContent = `Pairs: ${stats.pairs} | Missing Y: ${stats.missing}`;
  dom.countStats.textContent = `Candidate XLSX: ${stats.candidates}`;
}

function refreshButtons() {
  const canRun = state.workerReady && state.selectedFiles.length > 0;
  dom.runBtn.disabled = !canRun;
  dom.zipBtn.disabled = !state.workerReady || state.results.length === 0;
}

function formatBytes(bytes) {
  if (!Number.isFinite(bytes) || bytes <= 0) return "0.00 MB";
  return `${(bytes / (1024 * 1024)).toFixed(2)} MB`;
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
  if (!state.workerReady || !state.selectedFiles.length) {
    return;
  }

  state.results = [];
  state.errors = [];
  state.metrics = null;
  renderResults();
  renderMetrics();

  setStatus("Preparing files...");
  appendLog("info", `Preparing ${state.selectedFiles.length} files`);

  const filesPayload = await Promise.all(
    state.selectedFiles.map(async (file) => {
      const bytes = await file.arrayBuffer();
      return { name: normalizeCandidateName(file.name), bytes };
    })
  );

  const transferables = filesPayload.map((item) => item.bytes);

  worker.postMessage(
    {
      type: "runBatch",
      payload: {
        files: filesPayload,
        options: {
          methods: "strain_legacy",
          failFast: dom.failFast.checked,
          includeManip: dom.includeManip.checked,
        },
      },
    },
    transferables
  );

  dom.runBtn.disabled = true;
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
    state.results = payload.results || [];
    state.errors = payload.errors || [];
    state.metrics = payload.metrics || null;

    (payload.logs || []).forEach((log) => appendLog(log.level || "info", log.message || ""));
    state.errors.forEach((err) => appendLog("error", `${err.pairKey}: ${err.reason}`));

    renderResults();
    renderMetrics();
    refreshButtons();

    const processed = state.metrics?.pairsProcessed ?? state.results.length;
    const failed = state.metrics?.pairsFailed ?? state.errors.length;
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
    const message = payload?.message || "Unknown worker error";
    setStatus(`Error: ${message}`);
    appendLog("error", message);
    refreshButtons();
  }
});

dom.folderInput.addEventListener("change", () => handleSelectionChange("Folder Select"));
dom.fileInput.addEventListener("change", () => handleSelectionChange("File Select"));

dom.includeManip.addEventListener("change", updateSelectionStats);

dom.runBtn.addEventListener("click", () => {
  runBatch().catch((error) => {
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
worker.postMessage({ type: "initialize" });
