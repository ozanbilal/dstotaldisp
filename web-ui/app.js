const state = {
  workerReady: false,
  isRunning: false,
  selectedFiles: [],
  ignoredFilesCount: 0,
  manualPairs: [],
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
  clearFilesBtn: document.getElementById("clearFilesBtn"),
  manualPairPanel: document.getElementById("manualPairPanel"),
  manualPairMeta: document.getElementById("manualPairMeta"),
  manualPairList: document.getElementById("manualPairList"),
  manualPairingEnabled: document.getElementById("manualPairingEnabled"),
  manualXSelect: document.getElementById("manualXSelect"),
  manualYSelect: document.getElementById("manualYSelect"),
  manualStageX: document.getElementById("manualStageX"),
  manualStageY: document.getElementById("manualStageY"),
  clearManualStageBtn: document.getElementById("clearManualStageBtn"),
  addManualPairBtn: document.getElementById("addManualPairBtn"),
  clearPairsBtn: document.getElementById("clearPairsBtn"),
  failFast: document.getElementById("failFast"),
  method2Enabled: document.getElementById("method2Enabled"),
  method3Enabled: document.getElementById("method3Enabled"),
  useDb3Directly: document.getElementById("useDb3Directly"),
  integrationCompareEnabled: document.getElementById("integrationCompareEnabled"),
  includeResultantProfiles: document.getElementById("includeResultantProfiles"),
  baseReference: document.getElementById("baseReference"),
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

const worker = new Worker("./worker.js?v=20260316c");

function appendLog(level, message) {
  const line = `[${new Date().toLocaleTimeString()}] ${level.toUpperCase()} ${message}`;
  state.logs.push(line);
  dom.logBox.textContent = state.logs.slice(-200).join("\n");
  dom.logBox.scrollTop = dom.logBox.scrollHeight;
}

function setStatus(text) {
  dom.status.textContent = text;
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function baseNameFromPath(name) {
  return String(name || "")
    .split(/[\\/]/)
    .pop() || "";
}

function sourcePathForFile(file) {
  return String(file?.webkitRelativePath || file?.name || "").replaceAll("\\", "/");
}

function logicalCandidateName(fileOrName) {
  const rawPath = typeof fileOrName === "string" ? String(fileOrName) : sourcePathForFile(fileOrName);
  const normalizedPath = rawPath.replaceAll("\\", "/");
  const baseName = baseNameFromPath(normalizedPath);
  const dbExtMatch = baseName.match(/\.(db3|db)$/i);
  if (dbExtMatch) {
    const parts = normalizedPath.split("/").filter(Boolean);
    if (parts.length >= 2) {
      return `${parts[parts.length - 2]}.${dbExtMatch[1].toLowerCase()}`;
    }
  }
  return baseName;
}

function normalizeCandidateName(name) {
  return logicalCandidateName(name);
}

function deriveYName(xName) {
  const replaced = xName.replace("_X_", "_Y_");
  return replaced.replace(/_H1/i, "_H2");
}

function inferAxisLabel(name) {
  const upper = String(name || "").toUpperCase();
  if (upper.includes("_X_")) return "X";
  if (upper.includes("_Y_")) return "Y";
  return "SINGLE";
}

function matchesAxis(name, axis) {
  return inferAxisLabel(name) === axis;
}

function isSupportedInputFile(fileOrName) {
  const raw = typeof fileOrName === "string" ? String(fileOrName) : sourcePathForFile(fileOrName);
  const lower = baseNameFromPath(raw).toLowerCase();
  const allowed = lower.endsWith(".xlsx") || lower.endsWith(".db") || lower.endsWith(".db3");
  if (!allowed) return false;
  if (lower.startsWith("output_")) return false;
  if (lower.startsWith("~$")) return false;
  if (lower.endsWith("-manip.xlsx")) return false;
  return true;
}

function isInputCandidate(name, useDb3Directly = false) {
  if (!isSupportedInputFile(name)) {
    return false;
  }
  const lower = String(name || "").toLowerCase();
  return useDb3Directly ? lower.endsWith(".db") || lower.endsWith(".db3") : lower.endsWith(".xlsx");
}

function getCandidateNames(useDb3Directly = false) {
  const names = state.selectedFiles.map((file) => logicalCandidateName(file));
  return [...new Set(names.filter((name) => isInputCandidate(name, useDb3Directly)))].sort((a, b) =>
    a.localeCompare(b)
  );
}

function normalizeManualPairEntry(entry) {
  if (!entry) return null;

  if (Array.isArray(entry) && entry.length >= 2) {
    const xName = String(entry[0] || "").trim();
    const yName = String(entry[1] || "").trim();
    return xName && yName ? { xName, yName } : null;
  }

  if (typeof entry === "object") {
    const xName = String(entry.xName || entry.x || "").trim();
    const yName = String(entry.yName || entry.y || "").trim();
    return xName && yName ? { xName, yName } : null;
  }

  return null;
}

function sanitizeManualPairs(manualPairs, candidateNames) {
  const candidateLookup = new Map(candidateNames.map((name) => [name.toLowerCase(), name]));
  const used = new Set();
  const pairs = [];
  const removed = [];

  (manualPairs || []).forEach((entry) => {
    const pair = normalizeManualPairEntry(entry);
    if (!pair) {
      removed.push("invalid");
      return;
    }

    const xKey = pair.xName.toLowerCase();
    const yKey = pair.yName.toLowerCase();
    const xName = candidateLookup.get(xKey);
    const yName = candidateLookup.get(yKey);

    if (!xName || !yName) {
      removed.push(`${pair.xName} + ${pair.yName}`);
      return;
    }
    if (inferAxisLabel(xName) !== "X" || inferAxisLabel(yName) !== "Y") {
      removed.push(`${pair.xName} + ${pair.yName}`);
      return;
    }
    if (used.has(xName.toLowerCase()) || used.has(yName.toLowerCase())) {
      removed.push(`${pair.xName} + ${pair.yName}`);
      return;
    }

    used.add(xName.toLowerCase());
    used.add(yName.toLowerCase());
    pairs.push({ xName, yName });
  });

  return { pairs, removed };
}

function detectPairs(files, useDb3Directly = false, manualPairs = [], manualPairingEnabled = false) {
  const names = files.map((file) => logicalCandidateName(file));
  const candidates = new Set(names.filter((name) => isInputCandidate(name, useDb3Directly)));
  const xFiles = [...candidates].filter((name) => name.includes("_X_") && /_H1/i.test(name));
  const used = new Set();
  const xlsxCandidates = [...candidates].filter((name) => name.toLowerCase().endsWith(".xlsx")).length;
  const dbCandidates = [...candidates].filter((name) => name.toLowerCase().endsWith(".db") || name.toLowerCase().endsWith(".db3")).length;

  let found = 0;
  let missing = 0;

  if (manualPairingEnabled) {
    const normalized = sanitizeManualPairs(manualPairs, [...candidates]).pairs;
    normalized.forEach((pair) => {
      found += 1;
      used.add(pair.xName.toLowerCase());
      used.add(pair.yName.toLowerCase());
    });
  } else {
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
  }

  let singles = 0;
  for (const name of candidates) {
    if (!used.has(name.toLowerCase())) {
      singles += 1;
    }
  }

  return {
    candidates: candidates.size,
    xlsxCandidates,
    dbCandidates,
    pairs: found,
    missing,
    singles,
    manualPairsApplied: manualPairingEnabled ? found : 0,
  };
}

function summarizeLoadedTypes() {
  let xlsx = 0;
  let db = 0;
  state.selectedFiles.forEach((file) => {
    const lower = logicalCandidateName(file).toLowerCase();
    if (lower.endsWith(".xlsx")) {
      xlsx += 1;
    } else if (lower.endsWith(".db") || lower.endsWith(".db3")) {
      db += 1;
    }
  });
  return { xlsx, db, total: xlsx + db };
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
    ["Integration Primary", metrics.integrationPrimary ?? "cumtrapz"],
    ["Method2 Enabled", metrics.method2Enabled ? "yes" : "no"],
    ["Method3 Enabled", metrics.method3Enabled ? "yes" : "no"],
    ["Manual Pairing", metrics.manualPairingEnabled ? "yes" : "no"],
    ["Manual Pairs Applied", metrics.manualPairsApplied ?? 0],
    ["Use DB3 Directly", metrics.useDb3Directly ? "yes" : "no"],
    ["Depth Resultants (RSS)", metrics.includeResultantProfiles ? "yes" : "no"],
    ["Integration Compare", metrics.integrationCompareEnabled ? "yes" : "no"],
    ["Alt Method", metrics.altIntegrationMethod ?? "-"],
    ["Base Reference", metrics.baseReference ?? "input"],
    ["XLSX Candidates", metrics.xlsxCandidates ?? 0],
    ["DB Candidates", metrics.dbCandidates ?? 0],
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
    const modeRaw = item.metrics?.mode ?? "-";
    const pairKeyRaw = item.pairKey || "-";
    const outputFileNameRaw = item.outputFileName || "-";
    const mode = escapeHtml(modeRaw);
    const pairKey = escapeHtml(pairKeyRaw);
    const outputFileName = escapeHtml(outputFileNameRaw);

    tr.innerHTML = `
      <td title="${pairKey}">${pairKey}</td>
      <td>${mode}</td>
      <td title="${outputFileName}">${outputFileName}</td>
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
  const loaded = summarizeLoadedTypes();
  const stats = detectPairs(
    state.selectedFiles,
    !!dom.useDb3Directly?.checked,
    state.manualPairs,
    !!dom.manualPairingEnabled?.checked
  );
  dom.pairStats.textContent = `Pairs: ${stats.pairs} | Singles: ${stats.singles} | Missing Y: ${stats.missing}`;
  dom.countStats.textContent = `Active: ${stats.candidates} | Loaded XLSX: ${loaded.xlsx} | Loaded DB: ${loaded.db}`;
}

function getManualSelection() {
  return {
    xName: String(dom.manualXSelect?.value || "").trim(),
    yName: String(dom.manualYSelect?.value || "").trim(),
  };
}

function renderManualStage() {
  const { xName, yName } = getManualSelection();
  if (dom.manualStageX) dom.manualStageX.textContent = xName || "No X picked";
  if (dom.manualStageY) dom.manualStageY.textContent = yName || "No Y picked";
}

function refreshButtons() {
  const canRun = state.workerReady && !state.isRunning && state.selectedFiles.length > 0;
  const { xName, yName } = getManualSelection();
  dom.runBtn.disabled = !canRun;
  dom.zipBtn.disabled = state.isRunning || !state.workerReady || state.results.length === 0;
  if (dom.folderInput) dom.folderInput.disabled = state.isRunning;
  if (dom.fileInput) dom.fileInput.disabled = state.isRunning;
  if (dom.manualPairingEnabled) dom.manualPairingEnabled.disabled = state.isRunning;
  if (dom.manualXSelect) dom.manualXSelect.disabled = state.isRunning || !dom.manualPairingEnabled?.checked;
  if (dom.manualYSelect) dom.manualYSelect.disabled = state.isRunning || !dom.manualPairingEnabled?.checked;
  if (dom.addManualPairBtn) {
    dom.addManualPairBtn.disabled = state.isRunning || !dom.manualPairingEnabled?.checked || !xName || !yName;
  }
  if (dom.clearFilesBtn) dom.clearFilesBtn.disabled = state.isRunning || state.selectedFiles.length === 0;
  if (dom.clearPairsBtn) dom.clearPairsBtn.disabled = state.isRunning || state.manualPairs.length === 0;
  if (dom.clearManualStageBtn) {
    dom.clearManualStageBtn.disabled = state.isRunning || !dom.manualPairingEnabled?.checked || (!xName && !yName);
  }
  if (dom.failFast) dom.failFast.disabled = state.isRunning;
  if (dom.method2Enabled) dom.method2Enabled.disabled = state.isRunning;
  if (dom.method3Enabled) dom.method3Enabled.disabled = state.isRunning;
  if (dom.useDb3Directly) dom.useDb3Directly.disabled = state.isRunning;
  if (dom.integrationCompareEnabled) dom.integrationCompareEnabled.disabled = state.isRunning;
  if (dom.includeResultantProfiles) dom.includeResultantProfiles.disabled = state.isRunning;
  if (dom.baseReference) dom.baseReference.disabled = state.isRunning;
  if (dom.baselineOn) dom.baselineOn.disabled = state.isRunning;
  if (dom.filterOn) dom.filterOn.disabled = state.isRunning;
  dom.manualPairList?.querySelectorAll(".manual-pair-remove").forEach((btn) => {
    btn.disabled = state.isRunning;
  });
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
  const useDb3Directly = !!dom.useDb3Directly?.checked;
  const baselineOn = !!dom.baselineOn?.checked;
  const filterOn = !!dom.filterOn?.checked;
  const cfg = (dom.filterConfig?.value || "bandpass").toLowerCase();
  const isHighpass = cfg === "highpass" || cfg === "high";
  const isLowpass = cfg === "lowpass" || cfg === "low";

  if (dom.baseReference) dom.baseReference.disabled = useDb3Directly || state.isRunning;
  if (dom.integrationCompareEnabled) dom.integrationCompareEnabled.disabled = useDb3Directly || state.isRunning;
  if (dom.baselineOn) dom.baselineOn.disabled = useDb3Directly || state.isRunning;
  if (dom.filterOn) dom.filterOn.disabled = useDb3Directly || state.isRunning;
  if (dom.baselineMethod) dom.baselineMethod.disabled = useDb3Directly || !baselineOn || state.isRunning;
  if (dom.processingOrder) dom.processingOrder.disabled = useDb3Directly || !filterOn || state.isRunning;
  if (dom.filterDomain) dom.filterDomain.disabled = useDb3Directly || !filterOn || state.isRunning;
  if (dom.filterConfig) dom.filterConfig.disabled = useDb3Directly || !filterOn || state.isRunning;
  if (dom.filterType) dom.filterType.disabled = useDb3Directly || !filterOn || state.isRunning;
  if (dom.filterOrder) dom.filterOrder.disabled = useDb3Directly || !filterOn || state.isRunning;
  if (dom.fLowHz) dom.fLowHz.disabled = useDb3Directly || !filterOn || isHighpass || state.isRunning;
  if (dom.fHighHz) dom.fHighHz.disabled = useDb3Directly || !filterOn || isLowpass || state.isRunning;
}

function renderSelectedFiles() {
  const listEl = dom.selectedFilesList;
  listEl.innerHTML = "";

  const files = state.selectedFiles;
  const totalBytes = files.reduce((sum, file) => sum + (file.size || 0), 0);
  dom.selectedFilesMeta.textContent = `${files.length} files | ${formatBytes(totalBytes)}`;
  const manualEnabled = !!dom.manualPairingEnabled?.checked;
  const useDb3Directly = !!dom.useDb3Directly?.checked;
  const usedX = new Set(state.manualPairs.map((pair) => pair.xName.toLowerCase()));
  const usedY = new Set(state.manualPairs.map((pair) => pair.yName.toLowerCase()));
  const { xName: pickedX, yName: pickedY } = getManualSelection();
  renderManualStage();

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
    const logicalName = logicalCandidateName(file);
    const sourcePath = sourcePathForFile(file);
    const isCandidate = isInputCandidate(logicalName, useDb3Directly);
    const axis = inferAxisLabel(logicalName);
    li.title = sourcePath || logicalName;
    const isPicked = pickedX === logicalName || pickedY === logicalName;
    if (isPicked) li.classList.add("is-picked");

    const nameSpan = document.createElement("span");
    nameSpan.className = "name";
    nameSpan.textContent = logicalName;

    const metaSpan = document.createElement("span");
    metaSpan.className = "file-meta";

    const sizeSpan = document.createElement("span");
    sizeSpan.className = "size";
    sizeSpan.textContent = formatBytes(file.size || 0);
    metaSpan.appendChild(sizeSpan);

    if (manualEnabled && isCandidate && (axis === "X" || axis === "Y")) {
      const actionWrap = document.createElement("span");
      actionWrap.className = "file-actions";

      const alreadyPaired =
        axis === "X" ? usedX.has(logicalName.toLowerCase()) : usedY.has(logicalName.toLowerCase());

      if (alreadyPaired) {
        const chip = document.createElement("span");
        chip.className = "manual-chip";
        chip.textContent = axis === "X" ? "Paired X" : "Paired Y";
        actionWrap.appendChild(chip);
      } else {
        const btn = document.createElement("button");
        btn.type = "button";
        btn.className = "pick-manual";
        btn.dataset.role = axis.toLowerCase();
        btn.dataset.name = logicalName;
        if (axis === "X") {
          btn.textContent = pickedX === logicalName ? "Picked X" : pickedY ? "Pair with Y" : "Pick X";
        } else {
          btn.textContent = pickedY === logicalName ? "Picked Y" : pickedX ? "Pair with X" : "Pick Y";
        }
        actionWrap.appendChild(btn);
      }

      metaSpan.appendChild(actionWrap);
    }

    li.appendChild(nameSpan);
    li.appendChild(metaSpan);
    listEl.appendChild(li);
  });

  if (files.length > maxVisible) {
    const li = document.createElement("li");
    li.className = "empty";
    li.textContent = `... ${files.length - maxVisible} more files`;
    listEl.appendChild(li);
  }

  listEl.querySelectorAll(".pick-manual").forEach((btn) => {
    btn.addEventListener("click", () => {
      const role = String(btn.dataset.role || "");
      const logicalName = String(btn.dataset.name || "");
      if (!logicalName) return;
      stageManualCandidate(role, logicalName);
    });
  });
}

function fillSelectOptions(selectEl, values, placeholder) {
  if (!selectEl) return;
  const current = selectEl.value;
  selectEl.innerHTML = "";

  const placeholderOption = document.createElement("option");
  placeholderOption.value = "";
  placeholderOption.textContent = placeholder;
  selectEl.appendChild(placeholderOption);

  values.forEach((value) => {
    const option = document.createElement("option");
    option.value = value;
    option.textContent = value;
    selectEl.appendChild(option);
  });

  if (values.includes(current)) {
    selectEl.value = current;
  } else {
    selectEl.value = "";
  }
}

function renderManualPairing() {
  const manualEnabled = !!dom.manualPairingEnabled?.checked;
  const useDb3Directly = !!dom.useDb3Directly?.checked;
  const candidateNames = getCandidateNames(useDb3Directly);
  const { pairs, removed } = sanitizeManualPairs(state.manualPairs, candidateNames);
  const hadRemovals = removed.length > 0;
  state.manualPairs = pairs;

  if (hadRemovals) {
    appendLog("warning", `Manual pairs pruned after file/mode change: ${removed.length}`);
  }

  dom.manualPairPanel.classList.toggle("is-hidden", !manualEnabled);
  dom.manualPairMeta.textContent = `${state.manualPairs.length} pairs`;
  renderManualStage();

  const usedX = new Set(state.manualPairs.map((pair) => pair.xName.toLowerCase()));
  const usedY = new Set(state.manualPairs.map((pair) => pair.yName.toLowerCase()));
  const xCandidates = candidateNames.filter((name) => inferAxisLabel(name) === "X" && !usedX.has(name.toLowerCase()));
  const yCandidates = candidateNames.filter((name) => inferAxisLabel(name) === "Y" && !usedY.has(name.toLowerCase()));

  fillSelectOptions(dom.manualXSelect, xCandidates, "Select X file");
  fillSelectOptions(dom.manualYSelect, yCandidates, "Select Y file");

  dom.manualPairList.innerHTML = "";
  if (!manualEnabled) {
    const li = document.createElement("li");
    li.className = "empty";
    li.textContent = "Manual pairing kapali.";
    dom.manualPairList.appendChild(li);
    renderSelectedFiles();
    return;
  }

  if (!state.manualPairs.length) {
    const li = document.createElement("li");
    li.className = "empty";
    li.textContent = "Henuz manuel pair yok. X ve Y secip Add Pair ile ekleyin.";
    dom.manualPairList.appendChild(li);
    renderSelectedFiles();
    return;
  }

  state.manualPairs.forEach((pair, index) => {
    const li = document.createElement("li");
    li.innerHTML = `
      <div class="pair-name" title="${escapeHtml(pair.xName)} + ${escapeHtml(pair.yName)}">
        <span class="pair-tag">X</span> ${escapeHtml(pair.xName)} &nbsp;→&nbsp;
        <span class="pair-tag">Y</span> ${escapeHtml(pair.yName)}
      </div>
      <button class="manual-pair-remove" type="button" data-index="${index}">Remove</button>
    `;
    dom.manualPairList.appendChild(li);
  });

  dom.manualPairList.querySelectorAll(".manual-pair-remove").forEach((btn) => {
    btn.addEventListener("click", () => {
      const index = Number(btn.dataset.index);
      const removedPair = state.manualPairs[index];
      state.manualPairs.splice(index, 1);
      appendLog("info", `Manual pair removed: ${removedPair?.xName || "-"} + ${removedPair?.yName || "-"}`);
      renderManualPairing();
      updateSelectionStats();
      refreshButtons();
    });
  });

  renderSelectedFiles();
}

function collectSelectedFiles() {
  const combined = [...Array.from(dom.folderInput.files || []), ...Array.from(dom.fileInput.files || [])];
  const unique = new Map();
  let ignored = 0;

  combined.forEach((file) => {
    if (!isSupportedInputFile(file)) {
      ignored += 1;
      return;
    }
    const sourcePath = sourcePathForFile(file);
    const key = `${sourcePath || logicalCandidateName(file)}::${file.size}::${file.lastModified}`;
    if (!unique.has(key)) {
      unique.set(key, file);
    }
  });

  state.selectedFiles = Array.from(unique.values());
  state.ignoredFilesCount = ignored;
}

function handleSelectionChange(sourceLabel) {
  collectSelectedFiles();
  const folderCount = (dom.folderInput.files || []).length;
  const fileCount = (dom.fileInput.files || []).length;
  const loaded = summarizeLoadedTypes();
  if (loaded.db > 0 && loaded.xlsx === 0 && dom.useDb3Directly && !dom.useDb3Directly.checked) {
    dom.useDb3Directly.checked = true;
    appendLog("info", "Only DB inputs detected. Switched to DB3 direct mode automatically.");
  } else if (loaded.xlsx > 0 && loaded.db === 0 && dom.useDb3Directly && dom.useDb3Directly.checked) {
    dom.useDb3Directly.checked = false;
    appendLog("info", "Only XLSX inputs detected. Switched to Excel/strain mode automatically.");
  }
  appendLog(
    "info",
    `${sourceLabel} updated | folder items: ${folderCount}, file items: ${fileCount}, supported total: ${state.selectedFiles.length}, ignored unsupported: ${state.ignoredFilesCount}`
  );
  syncFilterInputs();
  renderManualPairing();
  updateSelectionStats();
  renderSelectedFiles();
  refreshButtons();
}

function handleModeToggle() {
  const useDb3Directly = !!dom.useDb3Directly?.checked;
  appendLog("info", `Input mode changed: ${useDb3Directly ? "DB3 direct" : "Excel/strain"}`);
  renderManualPairing();
  updateSelectionStats();
  syncFilterInputs();
  refreshButtons();
}

function handleManualPairingToggle() {
  appendLog("info", `Manual pairing ${dom.manualPairingEnabled?.checked ? "enabled" : "disabled"}`);
  renderManualPairing();
  updateSelectionStats();
  refreshButtons();
}

function handleManualSelectionChange() {
  renderSelectedFiles();
  renderManualStage();
  refreshButtons();
}

function stageManualCandidate(role, logicalName) {
  const normalizedRole = String(role || "").toLowerCase();
  if (!logicalName || (normalizedRole !== "x" && normalizedRole !== "y")) return;

  if (normalizedRole === "x" && dom.manualXSelect) {
    dom.manualXSelect.value = logicalName;
  } else if (normalizedRole === "y" && dom.manualYSelect) {
    dom.manualYSelect.value = logicalName;
  }

  const { xName, yName } = getManualSelection();
  if (matchesAxis(xName, "X") && matchesAxis(yName, "Y")) {
    addManualPair(xName, yName, { source: "list" });
    return;
  }

  appendLog("info", `Manual candidate picked from list: ${logicalName}`);
  renderManualPairing();
  updateSelectionStats();
  refreshButtons();
}

function clearManualStage() {
  if (dom.manualXSelect) dom.manualXSelect.value = "";
  if (dom.manualYSelect) dom.manualYSelect.value = "";
  renderManualPairing();
  refreshButtons();
}

function clearManualPairs() {
  if (!state.manualPairs.length) return;
  state.manualPairs = [];
  clearManualStage();
  appendLog("info", "All manual pairs cleared.");
  renderManualPairing();
  updateSelectionStats();
  refreshButtons();
}

function clearSelectedFiles() {
  if (!state.selectedFiles.length) return;
  state.selectedFiles = [];
  state.ignoredFilesCount = 0;
  state.manualPairs = [];
  state.results = [];
  state.errors = [];
  state.metrics = null;
  if (dom.folderInput) dom.folderInput.value = "";
  if (dom.fileInput) dom.fileInput.value = "";
  clearManualStage();
  renderResults();
  renderMetrics();
  updateSelectionStats();
  renderSelectedFiles();
  renderManualPairing();
  setStatus(state.workerReady ? "Ready" : "Initializing...");
  appendLog("info", "Selected files, pairs and current outputs cleared.");
  refreshButtons();
}

function addManualPair(xOverride = null, yOverride = null, options = {}) {
  const xName = String(xOverride || dom.manualXSelect?.value || "").trim();
  const yName = String(yOverride || dom.manualYSelect?.value || "").trim();
  const source = options.source || "controls";

  if (!xName || !yName) {
    appendLog("warning", "Select both X and Y candidates before adding a manual pair.");
    return false;
  }
  if (inferAxisLabel(xName) !== "X" || inferAxisLabel(yName) !== "Y") {
    appendLog("warning", "Manual pair rejected due to axis mismatch.");
    return false;
  }
  if (state.manualPairs.some((pair) => pair.xName.toLowerCase() === xName.toLowerCase())) {
    appendLog("warning", `X candidate already used in manual pairs: ${xName}`);
    return false;
  }
  if (state.manualPairs.some((pair) => pair.yName.toLowerCase() === yName.toLowerCase())) {
    appendLog("warning", `Y candidate already used in manual pairs: ${yName}`);
    return false;
  }

  state.manualPairs.push({ xName, yName });
  if (dom.manualXSelect) dom.manualXSelect.value = "";
  if (dom.manualYSelect) dom.manualYSelect.value = "";
  appendLog("info", `Manual pair added (${source}): ${xName} + ${yName}`);
  renderManualPairing();
  updateSelectionStats();
  refreshButtons();
  return true;
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
    const name = logicalCandidateName(file);
    return isInputCandidate(name, !!dom.useDb3Directly?.checked);
  });

  const filesPayload = selectedCandidates.map((file) => ({
    name: logicalCandidateName(file),
    file,
  }));

  if (!filesPayload.length) {
    appendLog("warning", "No valid candidate input files for the selected mode.");
    setStatus("No candidate input files");
    state.isRunning = false;
    refreshButtons();
    return;
  }

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
          manualPairingEnabled: dom.manualPairingEnabled.checked,
          manualPairs: dom.manualPairingEnabled.checked ? state.manualPairs.map((pair) => ({ ...pair })) : [],
          useDb3Directly: dom.useDb3Directly.checked,
          integrationCompareEnabled: dom.integrationCompareEnabled.checked,
          includeResultantProfiles: dom.includeResultantProfiles.checked,
          altIntegrationMethod: "fft_regularized",
          baseReference: dom.baseReference?.value || "input",
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
    }
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
dom.useDb3Directly.addEventListener("change", handleModeToggle);
dom.manualPairingEnabled.addEventListener("change", handleManualPairingToggle);
dom.addManualPairBtn.addEventListener("click", addManualPair);
if (dom.clearFilesBtn) dom.clearFilesBtn.addEventListener("click", clearSelectedFiles);
if (dom.clearPairsBtn) dom.clearPairsBtn.addEventListener("click", clearManualPairs);
dom.manualXSelect.addEventListener("change", handleManualSelectionChange);
dom.manualYSelect.addEventListener("change", handleManualSelectionChange);
if (dom.clearManualStageBtn) dom.clearManualStageBtn.addEventListener("click", clearManualStage);

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
renderManualPairing();
syncFilterInputs();
worker.postMessage({ type: "initialize" });
