import {
  buildSourceViewerScene,
  classifyViewerResult,
  cycleSceneItem,
  groupViewerResults,
  normalizeShellMode,
  normalizeViewerChartMode,
  summarizeViewerResults,
} from "./viewer.mjs?v=20260408b";
import { detectDirectionInfo, PAIR_SCORE_AUTO_THRESHOLD, resolveDeepsoilPairingCandidates } from "./pairing.mjs?v=20260408b";

const APP_VERSION = "20260408b";
const MANUAL_PAIR_STORAGE_KEY = "deepsoil-total-disp.manual-pairs.v1";
const SHELL_MODE_STORAGE_KEY = "deepsoil-total-disp.shell-mode.v1";
const VIEWER_PREFS_STORAGE_KEY = "deepsoil-total-disp.viewer-prefs.v1";
const RUN_SNAPSHOT_DB_NAME = "deepsoil-total-disp";
const RUN_SNAPSHOT_STORE = "runSnapshots";
const RUN_SNAPSHOT_KEY = "latest-v1";
const MAX_PARALLEL_WORKERS = 4;

const PLOT_COLORS = ["#0d684d", "#ce6433", "#3559a6", "#9a4a31", "#6e7681", "#5f7c2f", "#764a95"];

const state = {
  workerReady: false,
  isRunning: false,
  selectedFiles: [],
  ignoredFilesCount: 0,
  manualPairs: [],
  manualStage: { xName: "", yName: "" },
  pairSuggestions: [],
  results: [],
  sourceCatalog: [],
  errors: [],
  logs: [],
  metrics: null,
  progress: { visible: true, label: "Initializing worker...", percent: 8, indeterminate: true },
  snapshotPersistPromise: null,
  shellMode: "viewer",
  viewerChartMode: "focus",
  activeSourceId: "",
  activeFamilyKey: "",
  activeChartKey: "",
  activeLayerIndex: 0,
  compareSourceIds: [],
  seriesVisibilityMap: {},
  spectrumPeriodMax: "",
  logX: false,
  logY: false,
};

const dom = {
  folderInput: document.getElementById("folderInput"),
  fileInput: document.getElementById("fileInput"),
  runBtn: document.getElementById("runBtn"),
  zipBtn: document.getElementById("zipBtn"),
  clearAllBtn: document.getElementById("clearAllBtn"),
  status: document.getElementById("statusText"),
  progressPanel: document.getElementById("progressPanel"),
  progressLabel: document.getElementById("progressLabel"),
  progressValue: document.getElementById("progressValue"),
  progressBar: document.getElementById("progressBar"),
  pairStats: document.getElementById("pairStats"),
  countStats: document.getElementById("countStats"),
  metrics: document.getElementById("metricCards"),
  metricMeta: document.getElementById("metricMeta"),
  sourceList: document.getElementById("sourceList"),
  sourceMeta: document.getElementById("sourceMeta"),
  familyList: document.getElementById("familyList"),
  familyMeta: document.getElementById("familyMeta"),
  chartMeta: document.getElementById("chartMeta"),
  chartList: document.getElementById("chartList"),
  viewerFocusBtn: document.getElementById("viewer-focus-btn"),
  viewerCompareBtn: document.getElementById("viewer-compare-btn"),
  compareSourceList: document.getElementById("compareSourceList"),
  compareMeta: document.getElementById("compareMeta"),
  compareAxisXToggle: document.getElementById("compareAxisXToggle"),
  compareAxisYToggle: document.getElementById("compareAxisYToggle"),
  compareClearBtn: document.getElementById("compareClearBtn"),
  periodMaxInput: document.getElementById("periodMaxInput"),
  periodMaxResetBtn: document.getElementById("periodMaxResetBtn"),
  logXToggle: document.getElementById("logXToggle"),
  logYToggle: document.getElementById("logYToggle"),
  seriesToggleList: document.getElementById("seriesToggleList"),
  seriesMeta: document.getElementById("seriesMeta"),
  plotHost: document.getElementById("plotHost"),
  catalogMeta: document.getElementById("catalogMeta"),
  stageTitle: document.getElementById("stageTitle"),
  stageSubtitle: document.getElementById("stageSubtitle"),
  activeSourceLabel: document.getElementById("activeSourceLabel"),
  activeFamilyLabel: document.getElementById("activeFamilyLabel"),
  activeChartLabel: document.getElementById("activeChartLabel"),
  activeLayerLabel: document.getElementById("activeLayerLabel"),
  sourcePrevBtn: document.getElementById("sourcePrevBtn"),
  sourceNextBtn: document.getElementById("sourceNextBtn"),
  chartPrevBtn: document.getElementById("chartPrevBtn"),
  chartNextBtn: document.getElementById("chartNextBtn"),
  layerPrevBtn: document.getElementById("layerPrevBtn"),
  layerNextBtn: document.getElementById("layerNextBtn"),
  sourceCounter: document.getElementById("sourceCounter"),
  chartCounter: document.getElementById("chartCounter"),
  layerCounter: document.getElementById("layerCounter"),
  logBox: document.getElementById("logBox"),
  resultTableBody: document.getElementById("resultTableBody"),
  resultPreviewMeta: document.getElementById("resultPreviewMeta"),
  resultPreviewDeck: document.getElementById("resultPreviewDeck"),
  selectedFilesMeta: document.getElementById("selectedFilesMeta"),
  selectedFilesList: document.getElementById("selectedFilesList"),
  clearFilesBtn: document.getElementById("clearFilesBtn"),
  legacyDrawer: document.getElementById("legacyDrawer"),
  viewerModeBtn: document.getElementById("viewer-mode-btn"),
  legacyModeBtn: document.getElementById("legacy-mode-btn"),
  viewerSwitcherHint: document.getElementById("viewer-switcher-hint"),
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
  autoPairBtn: document.getElementById("autoPairBtn"),
  clearPairsBtn: document.getElementById("clearPairsBtn"),
  pairSuggestionsPanel: document.getElementById("pairSuggestionsPanel"),
  pairSuggestionsMeta: document.getElementById("pairSuggestionsMeta"),
  pairSuggestionsList: document.getElementById("pairSuggestionsList"),
  applyAllSuggestionsBtn: document.getElementById("applyAllSuggestionsBtn"),
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

const workerSlots = [];
let plotStabilizeTimer = 0;

function workerTag(slot) {
  return `W${slot.id}`;
}

function clipMiddle(text, maxLength = 88, tailLength = 28) {
  const value = String(text || "");
  if (value.length <= maxLength) return value;
  const headLength = Math.max(12, maxLength - tailLength - 1);
  return `${value.slice(0, headLength)}…${value.slice(-tailLength)}`;
}

function clampProgress(value) {
  const num = Number(value);
  if (!Number.isFinite(num)) return 0;
  return Math.max(0, Math.min(100, num));
}

function compactProgressMessage(message) {
  const text = String(message || "").trim();
  if (!text) return "Running...";
  const match = text.match(/^([^:]+:\s*)(.+)$/);
  if (!match) return clipMiddle(text, 108, 36);
  return `${match[1]}${clipMiddle(match[2], 68, 24)}`;
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function formatBytes(bytes) {
  if (!Number.isFinite(bytes) || bytes <= 0) return "0.00 MB";
  return `${(bytes / (1024 * 1024)).toFixed(2)} MB`;
}

function formatNumber(value, digits = 4) {
  const num = Number(value);
  return Number.isFinite(num) ? num.toFixed(digits) : "-";
}

function setStatus(text) {
  if (dom.status) dom.status.textContent = text;
}

function renderProgress() {
  if (!dom.progressPanel || !dom.progressLabel || !dom.progressValue || !dom.progressBar) return;
  const { visible, label, percent, indeterminate } = state.progress;
  dom.progressPanel.classList.toggle("is-hidden", !visible);
  dom.progressLabel.textContent = label || "";
  dom.progressValue.textContent = indeterminate ? "..." : `${Math.round(clampProgress(percent))}%`;
  dom.progressBar.classList.toggle("is-indeterminate", !!indeterminate);
  dom.progressBar.style.width = `${indeterminate ? 35 : clampProgress(percent)}%`;
}

function setProgressState(next = {}) {
  state.progress = {
    ...state.progress,
    ...next,
    percent: next.percent == null ? state.progress.percent : clampProgress(next.percent),
  };
  renderProgress();
}

function appendLog(level, message) {
  const line = `[${new Date().toLocaleTimeString()}] ${String(level || "info").toUpperCase()} ${message}`;
  state.logs.push(line);
  if (dom.logBox) {
    dom.logBox.textContent = state.logs.slice(-220).join("\n");
    dom.logBox.scrollTop = dom.logBox.scrollHeight;
  }
}

function createWorkerInstance(slot) {
  const worker = new Worker(`./worker.js?v=${APP_VERSION}`);
  worker.addEventListener("message", (event) => handleWorkerMessage(slot, event.data || {}));
  return worker;
}

function createWorkerSlot(id) {
  const slot = {
    id,
    worker: null,
    ready: false,
    initializing: false,
    busy: false,
    initPromise: null,
    pendingInit: null,
    pendingJob: null,
    progress: { message: "", progress: null, indeterminate: true, phase: "boot" },
    currentJob: null,
  };
  slot.worker = createWorkerInstance(slot);
  workerSlots.push(slot);
  return slot;
}

function getOrCreateWorkerSlot(index) {
  return workerSlots[index - 1] || createWorkerSlot(index);
}

const primaryWorkerSlot = getOrCreateWorkerSlot(1);

function retireWorkerSlot(slot) {
  if (slot.worker) {
    try {
      slot.worker.terminate();
    } catch {
      // Ignore worker termination failures during recovery.
    }
  }
  slot.ready = false;
  slot.initializing = false;
  slot.busy = false;
  slot.initPromise = null;
  slot.pendingInit = null;
  slot.pendingJob = null;
  slot.currentJob = null;
  slot.progress = { message: "Worker restarting...", progress: null, indeterminate: true, phase: "boot" };
  slot.worker = createWorkerInstance(slot);

  if (slot === primaryWorkerSlot) {
    state.workerReady = false;
    refreshButtons();
  }
}

function updateWorkerAggregateProgress() {
  const active = workerSlots.filter((slot) => slot.initializing || slot.busy);
  if (!active.length) return;

  let weightedTotal = 0;
  let weightedKnown = 0;
  let indeterminate = false;
  let latestMessage = "";

  active.forEach((slot) => {
    const weight = Math.max(1, Number(slot.currentJob?.weight) || 1);
    const progressValue = Number(slot.progress?.progress);
    if (Number.isFinite(progressValue)) {
      weightedTotal += progressValue * weight;
      weightedKnown += weight;
    } else {
      indeterminate = true;
    }
    if (!latestMessage && slot.progress?.message) latestMessage = slot.progress.message;
  });

  const percent = weightedKnown > 0 ? weightedTotal / weightedKnown : state.progress.percent;
  const shortMessage = compactProgressMessage(latestMessage || "Running...");
  const label = active.length > 1 ? `Parallel run (${active.length} workers): ${shortMessage}` : shortMessage;
  setStatus(label);
  setProgressState({ visible: true, label, percent, indeterminate: indeterminate || weightedKnown === 0 });
}

function handleWorkerFailure(slot, message) {
  slot.busy = false;
  slot.initializing = false;
  slot.currentJob = null;
  slot.progress = { message, progress: 0, indeterminate: false, phase: "error" };

  if (slot.pendingInit) {
    slot.pendingInit.reject(new Error(message));
    slot.pendingInit = null;
    slot.initPromise = null;
  }
  if (slot.pendingJob) {
    slot.pendingJob.reject(new Error(message));
    slot.pendingJob = null;
  }

  retireWorkerSlot(slot);
  updateWorkerAggregateProgress();
}

function handleWorkerMessage(slot, data) {
  const { type, payload } = data || {};

  if (type === "status") {
    slot.progress = {
      message: payload?.message || "",
      progress: Number.isFinite(Number(payload?.progress)) ? Number(payload.progress) : null,
      indeterminate: !!payload?.indeterminate,
      phase: payload?.phase || "info",
    };
    appendLog(payload?.phase || "info", `[${workerTag(slot)}] ${payload?.message || ""}`);

    if (slot.initializing || slot.busy) {
      updateWorkerAggregateProgress();
    } else if (payload?.phase === "ready") {
      setStatus("Ready");
      setProgressState({ visible: false, label: "", percent: 100, indeterminate: false });
    }
    return;
  }

  if (type === "initialized") {
    slot.ready = true;
    slot.initializing = false;
    slot.progress = { message: "Ready", progress: 100, indeterminate: false, phase: "ready" };
    if (slot.pendingInit) {
      slot.pendingInit.resolve(slot);
      slot.pendingInit = null;
    }
    if (slot === primaryWorkerSlot) {
      state.workerReady = true;
      setStatus("Ready");
      setProgressState({ visible: false, label: "", percent: 100, indeterminate: false });
      appendLog("info", "Worker initialized");
      refreshButtons();
    } else {
      appendLog("info", `${workerTag(slot)} initialized`);
    }
    return;
  }

  if (type === "runBatchResult") {
    slot.busy = false;
    slot.progress = { message: "Done", progress: 100, indeterminate: false, phase: "ready" };
    const pending = slot.pendingJob;
    slot.pendingJob = null;
    slot.currentJob = null;
    if (pending) pending.resolve(payload);
    updateWorkerAggregateProgress();
    return;
  }

  if (type === "error") {
    handleWorkerFailure(slot, payload?.message || "Unknown worker error");
  }
}

function ensureWorkerReady(slot) {
  if (slot.ready) return Promise.resolve(slot);
  if (slot.initPromise) return slot.initPromise;

  slot.initializing = true;
  slot.progress = { message: `Initializing ${workerTag(slot)}...`, progress: 8, indeterminate: true, phase: "boot" };
  updateWorkerAggregateProgress();
  slot.initPromise = new Promise((resolve, reject) => {
    slot.pendingInit = { resolve, reject };
    slot.worker.postMessage({ type: "initialize" });
  }).finally(() => {
    slot.initPromise = null;
  });
  return slot.initPromise;
}

async function ensureWorkerPool(size) {
  const safeSize = Math.max(1, size);
  const slots = [];
  for (let index = 1; index <= safeSize; index += 1) slots.push(getOrCreateWorkerSlot(index));
  await Promise.all(slots.map((slot) => ensureWorkerReady(slot)));
  return slots;
}

async function runBatchOnSlot(slot, job) {
  await ensureWorkerReady(slot);
  slot.busy = true;
  slot.currentJob = job;
  slot.progress = { message: job.label || "Running...", progress: 4, indeterminate: true, phase: "run" };
  updateWorkerAggregateProgress();
  return new Promise((resolve, reject) => {
    slot.pendingJob = { resolve, reject };
    slot.worker.postMessage({ type: "runBatch", payload: { files: job.files, options: job.options } });
  });
}

function updateShellModeUi() {
  state.shellMode = normalizeShellMode(state.shellMode);
  document.body.dataset.shellMode = state.shellMode;

  if (dom.viewerModeBtn) {
    const active = state.shellMode === "viewer";
    dom.viewerModeBtn.classList.toggle("active", active);
    dom.viewerModeBtn.setAttribute("aria-selected", active ? "true" : "false");
  }
  if (dom.legacyModeBtn) {
    const active = state.shellMode === "legacy";
    dom.legacyModeBtn.classList.toggle("active", active);
    dom.legacyModeBtn.setAttribute("aria-selected", active ? "true" : "false");
  }
  if (dom.viewerSwitcherHint) {
    dom.viewerSwitcherHint.textContent =
      state.shellMode === "viewer"
        ? "Viewer modu kaynak eğriler ve sade karşılaştırma için varsayılan yüzeydir."
        : "Legacy modu pairleme, filtre ve batch analysis seçeneklerini açar.";
  }
  if (dom.legacyDrawer) dom.legacyDrawer.open = state.shellMode === "legacy";
}

function loadShellModePreference() {
  try {
    return normalizeShellMode(globalThis.localStorage?.getItem(SHELL_MODE_STORAGE_KEY));
  } catch {
    return "viewer";
  }
}

function getViewerPrefsSnapshot() {
  return {
    chartMode: state.viewerChartMode,
    activeSourceId: String(state.activeSourceId || ""),
    activeFamilyKey: String(state.activeFamilyKey || ""),
    activeChartKey: String(state.activeChartKey || ""),
    activeLayerIndex: Number.isFinite(Number(state.activeLayerIndex)) ? Math.max(0, Number(state.activeLayerIndex)) : 0,
    compareSourceIds: Array.isArray(state.compareSourceIds) ? [...state.compareSourceIds] : [],
    seriesVisibilityMap:
      state.seriesVisibilityMap && typeof state.seriesVisibilityMap === "object" ? { ...state.seriesVisibilityMap } : {},
    spectrumPeriodMax: String(state.spectrumPeriodMax || "").trim(),
    logX: !!state.logX,
    logY: !!state.logY,
  };
}

function loadViewerPrefs() {
  try {
    const raw = globalThis.localStorage?.getItem(VIEWER_PREFS_STORAGE_KEY);
    if (!raw) return null;
    const parsed = JSON.parse(raw);
    return {
      chartMode: normalizeViewerChartMode(parsed?.chartMode),
      activeSourceId: String(parsed?.activeSourceId || "").trim(),
      activeFamilyKey: String(parsed?.activeFamilyKey || "").trim(),
      activeChartKey: String(parsed?.activeChartKey || "").trim(),
      activeLayerIndex: Number.isFinite(Number(parsed?.activeLayerIndex)) ? Math.max(0, Number(parsed.activeLayerIndex)) : 0,
      compareSourceIds: Array.isArray(parsed?.compareSourceIds)
        ? parsed.compareSourceIds.map((value) => String(value || "").trim()).filter(Boolean)
        : [],
      seriesVisibilityMap: parsed?.seriesVisibilityMap && typeof parsed.seriesVisibilityMap === "object" ? parsed.seriesVisibilityMap : {},
      spectrumPeriodMax: String(parsed?.spectrumPeriodMax || "").trim(),
      logX: !!parsed?.logX,
      logY: !!parsed?.logY,
    };
  } catch {
    return null;
  }
}

function persistViewerPrefs() {
  try {
    globalThis.localStorage?.setItem(VIEWER_PREFS_STORAGE_KEY, JSON.stringify(getViewerPrefsSnapshot()));
  } catch {
    // Ignore local storage failures.
  }
}

function updateViewerModeUi() {
  state.viewerChartMode = normalizeViewerChartMode(state.viewerChartMode);
  document.body.dataset.viewerChartMode = state.viewerChartMode;

  const buttonMap = [
    [dom.viewerFocusBtn, "focus"],
    [dom.viewerCompareBtn, "compare"],
  ];
  buttonMap.forEach(([button, mode]) => {
    if (!button) return;
    const active = state.viewerChartMode === mode;
    button.classList.toggle("active", active);
    button.setAttribute("aria-selected", active ? "true" : "false");
  });
}

function applyViewerPrefs(next = {}, { persist = true } = {}) {
  if (Object.prototype.hasOwnProperty.call(next, "chartMode")) state.viewerChartMode = normalizeViewerChartMode(next.chartMode);
  if (Object.prototype.hasOwnProperty.call(next, "activeSourceId")) state.activeSourceId = String(next.activeSourceId || "").trim();
  if (Object.prototype.hasOwnProperty.call(next, "activeFamilyKey")) state.activeFamilyKey = String(next.activeFamilyKey || "").trim();
  if (Object.prototype.hasOwnProperty.call(next, "activeChartKey")) state.activeChartKey = String(next.activeChartKey || "").trim();
  if (Object.prototype.hasOwnProperty.call(next, "activeLayerIndex")) {
    const value = Number(next.activeLayerIndex);
    state.activeLayerIndex = Number.isFinite(value) && value >= 0 ? value : 0;
  }
  if (Object.prototype.hasOwnProperty.call(next, "compareSourceIds")) {
    state.compareSourceIds = Array.isArray(next.compareSourceIds)
      ? next.compareSourceIds.map((value) => String(value || "").trim()).filter(Boolean)
      : [];
  }
  if (Object.prototype.hasOwnProperty.call(next, "seriesVisibilityMap")) {
    state.seriesVisibilityMap =
      next.seriesVisibilityMap && typeof next.seriesVisibilityMap === "object" ? { ...next.seriesVisibilityMap } : {};
  }
  if (Object.prototype.hasOwnProperty.call(next, "spectrumPeriodMax")) {
    const value = String(next.spectrumPeriodMax || "").trim();
    state.spectrumPeriodMax = value;
  }
  if (Object.prototype.hasOwnProperty.call(next, "logX")) state.logX = !!next.logX;
  if (Object.prototype.hasOwnProperty.call(next, "logY")) state.logY = !!next.logY;
  updateViewerModeUi();
  if (persist) persistViewerPrefs();
}

function syncSourceViewerSelection() {
  const scene = buildSourceViewerScene(state.sourceCatalog, getViewerPrefsSnapshot());
  state.viewerChartMode = scene.chartMode;
  state.activeSourceId = scene.activeSourceId || "";
  state.activeFamilyKey = scene.activeFamilyKey || "";
  state.activeChartKey = scene.activeChartKey || "";
  state.activeLayerIndex = Number.isFinite(Number(scene.activeLayerIndex)) ? Number(scene.activeLayerIndex) : 0;
  state.compareSourceIds = Array.isArray(scene.compareSourceIds) ? [...scene.compareSourceIds] : [];
  updateViewerModeUi();
  return scene;
}

function setShellMode(nextMode, { persist = true } = {}) {
  state.shellMode = normalizeShellMode(nextMode);
  updateShellModeUi();
  syncFilterInputs();
  renderAll();
  if (!persist) return;
  try {
    globalThis.localStorage?.setItem(SHELL_MODE_STORAGE_KEY, state.shellMode);
  } catch {
    // Ignore local persistence failures.
  }
}

function setViewerChartMode(nextMode) {
  applyViewerPrefs({ chartMode: nextMode }, { persist: true });
  renderAll();
}

function setActiveSource(sourceId) {
  applyViewerPrefs(
    {
      activeSourceId: sourceId,
      compareSourceIds: [],
    },
    { persist: true }
  );
  renderAll();
}

function setActiveFamily(familyKey) {
  applyViewerPrefs(
    {
      activeFamilyKey: familyKey,
      activeChartKey: "",
      activeLayerIndex: 0,
      compareSourceIds: [],
      seriesVisibilityMap: {},
    },
    { persist: true }
  );
  renderAll();
}

function setActiveChart(chartKey) {
  applyViewerPrefs(
    {
      activeChartKey: chartKey,
      activeLayerIndex: 0,
      compareSourceIds: [],
      seriesVisibilityMap: {},
    },
    { persist: true }
  );
  renderAll();
}

function setActiveLayer(layerIndex) {
  applyViewerPrefs({ activeLayerIndex: Math.max(0, Number(layerIndex) || 0) }, { persist: true });
  renderAll();
}

function toggleCompareSource(sourceId, checked) {
  const current = new Set((state.compareSourceIds || []).map((value) => String(value || "").trim()).filter(Boolean));
  const normalized = String(sourceId || "").trim();
  if (!normalized) return;
  if (checked) current.add(normalized);
  else current.delete(normalized);
  applyViewerPrefs({ compareSourceIds: [...current] }, { persist: true });
  renderAll();
}

function toggleCompareAxisBulk(axisLabel, checked) {
  const normalizedAxis = String(axisLabel || "").trim().toUpperCase();
  if (normalizedAxis !== "X" && normalizedAxis !== "Y") return;
  const scene = syncSourceViewerSelection();
  const axisSourceIds = (scene.compareCandidates || [])
    .filter((source) => String(source?.axis || "").trim().toUpperCase() === normalizedAxis)
    .map((source) => String(source.sourceId || ""))
    .filter(Boolean);
  if (!axisSourceIds.length) return;

  const current = new Set((state.compareSourceIds || []).map((value) => String(value || "").trim()).filter(Boolean));
  axisSourceIds.forEach((sourceId) => {
    if (checked) current.add(sourceId);
    else current.delete(sourceId);
  });
  applyViewerPrefs({ compareSourceIds: [...current] }, { persist: true });
  renderAll();
}

function clearCompareSources() {
  if (!state.compareSourceIds.length) return;
  applyViewerPrefs({ compareSourceIds: [] }, { persist: true });
  renderAll();
}

function setSpectrumPeriodMax(value) {
  const raw = String(value ?? "").trim();
  if (!raw) {
    applyViewerPrefs({ spectrumPeriodMax: "" }, { persist: true });
    renderAll();
    return;
  }
  const parsed = Number(raw);
  if (!Number.isFinite(parsed) || parsed <= 0) return;
  applyViewerPrefs({ spectrumPeriodMax: String(parsed) }, { persist: true });
  renderAll();
}

function setLogAxis(axis, enabled) {
  if (axis === "x") applyViewerPrefs({ logX: !!enabled }, { persist: true });
  if (axis === "y") applyViewerPrefs({ logY: !!enabled }, { persist: true });
  renderAll();
}

function toggleSeriesVisibility(entryKey, checked) {
  const next = { ...(state.seriesVisibilityMap || {}) };
  next[String(entryKey || "")] = !!checked;
  applyViewerPrefs({ seriesVisibilityMap: next }, { persist: true });
  renderAll();
}

function baseNameFromPath(name) {
  return String(name || "").split(/[\\/]/).pop() || "";
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
    if (parts.length >= 2) return `${parts[parts.length - 2]}.${dbExtMatch[1].toLowerCase()}`;
  }
  return baseName;
}

function stripOutputSuffix(stem) {
  return String(stem || "").replace(/[_.-](ACC|VEL|DISP|SA|SV|SD|TH)$/i, "");
}

function detectDirection(name) {
  const stem = stripOutputSuffix(String(name || "").replace(/\.[^.]+$/, ""));
  const f = stem.toUpperCase();
  if (/[_.-]X([_.-]|$)/.test(f) || f.endsWith("_X") || f.startsWith("X_")) return "X";
  if (/[_.-]Y([_.-]|$)/.test(f) || f.endsWith("_Y") || f.startsWith("Y_")) return "Y";
  if (/(HN1|H1|HNE|EW|000|180|270|360|225|210)$/.test(f) || /[_.-](HN1|H1|HNE|E|W|EW|000|180|270|360|225|210)(?=[_.-]|$)/.test(f)) return "X";
  if (/(HN2|H2|HNN|NS|090|045|135|315|300)$/.test(f) || /[_.-](HN2|H2|HNN|N|S|NS|090|045|135|315|300)(?=[_.-]|$)/.test(f)) return "Y";
  if (/\d(E|EW|W|X)$/.test(f)) return "X";
  if (/\d(N|NS|S|Y)$/.test(f)) return "Y";
  if (/[A-Za-z]X\d+$/u.test(stem)) return "X";
  if (/[A-Za-z]Y\d+$/u.test(stem)) return "Y";
  return "SINGLE";
}

function inferAxisLabel(name) {
  const upper = String(name || "").toUpperCase();
  if (upper.includes("_X_")) return "X";
  if (upper.includes("_Y_")) return "Y";
  const detected = detectDirectionInfo(name);
  if (detected?.side === "X" || detected?.side === "Y") return detected.side;
  return detectDirection(name);
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
  if (!isSupportedInputFile(name)) return false;
  const lower = String(name || "").toLowerCase();
  return useDb3Directly ? lower.endsWith(".db") || lower.endsWith(".db3") : lower.endsWith(".xlsx");
}

function canManualAssign(name, role, useDb3Directly = false) {
  if (!isInputCandidate(name, useDb3Directly)) return false;
  if (useDb3Directly) return true;
  return inferAxisLabel(name) === String(role || "").toUpperCase();
}

function getCandidateNames(useDb3Directly = false) {
  const names = state.selectedFiles.map((file) => logicalCandidateName(file));
  return [...new Set(names.filter((name) => isInputCandidate(name, useDb3Directly)))].sort((a, b) => a.localeCompare(b));
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

function normalizePersistedManualPairs(manualPairs) {
  const usedX = new Set();
  const usedY = new Set();
  const pairs = [];
  (manualPairs || []).forEach((entry) => {
    const pair = normalizeManualPairEntry(entry);
    if (!pair) return;
    const xKey = pair.xName.toLowerCase();
    const yKey = pair.yName.toLowerCase();
    if (xKey === yKey || usedX.has(xKey) || usedY.has(yKey)) return;
    usedX.add(xKey);
    usedY.add(yKey);
    pairs.push({ xName: pair.xName, yName: pair.yName });
  });
  return pairs;
}

function loadManualPairPrefs() {
  try {
    const raw = globalThis.localStorage?.getItem(MANUAL_PAIR_STORAGE_KEY);
    if (!raw) return { manualPairingEnabled: false, manualPairs: [] };
    const parsed = JSON.parse(raw);
    return {
      manualPairingEnabled: !!parsed?.manualPairingEnabled,
      manualPairs: normalizePersistedManualPairs(parsed?.manualPairs),
    };
  } catch {
    return { manualPairingEnabled: false, manualPairs: [] };
  }
}

function persistManualPairPrefs() {
  try {
    if (!globalThis.localStorage) return;
    const payload = {
      manualPairingEnabled: !!dom.manualPairingEnabled?.checked,
      manualPairs: state.manualPairs.map((pair) => ({ xName: pair.xName, yName: pair.yName })),
    };
    if (!payload.manualPairingEnabled && payload.manualPairs.length === 0) {
      globalThis.localStorage.removeItem(MANUAL_PAIR_STORAGE_KEY);
      return;
    }
    globalThis.localStorage.setItem(MANUAL_PAIR_STORAGE_KEY, JSON.stringify(payload));
  } catch {
    // Ignore local persistence failures.
  }
}

function openRunSnapshotDb() {
  if (!globalThis.indexedDB) return Promise.resolve(null);
  return new Promise((resolve, reject) => {
    const request = globalThis.indexedDB.open(RUN_SNAPSHOT_DB_NAME, 1);
    request.onupgradeneeded = () => {
      const db = request.result;
      if (!db.objectStoreNames.contains(RUN_SNAPSHOT_STORE)) db.createObjectStore(RUN_SNAPSHOT_STORE);
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error || new Error("IndexedDB open failed."));
  });
}

async function loadRunSnapshot() {
  const db = await openRunSnapshotDb();
  if (!db) return null;
  return new Promise((resolve, reject) => {
    const tx = db.transaction(RUN_SNAPSHOT_STORE, "readonly");
    const request = tx.objectStore(RUN_SNAPSHOT_STORE).get(RUN_SNAPSHOT_KEY);
    request.onsuccess = () => resolve(request.result || null);
    request.onerror = () => reject(request.error || new Error("IndexedDB read failed."));
    tx.oncomplete = () => db.close();
    tx.onabort = () => {
      db.close();
      reject(tx.error || new Error("IndexedDB transaction aborted."));
    };
  });
}

async function persistRunSnapshot() {
  const db = await openRunSnapshotDb();
  if (!db) return;
  const payload = {
    savedAt: new Date().toISOString(),
    results: state.results,
    sourceCatalog: state.sourceCatalog,
    errors: state.errors,
    metrics: state.metrics,
    viewerPrefs: getViewerPrefsSnapshot(),
  };
  return new Promise((resolve, reject) => {
    const tx = db.transaction(RUN_SNAPSHOT_STORE, "readwrite");
    const request = tx.objectStore(RUN_SNAPSHOT_STORE).put(payload, RUN_SNAPSHOT_KEY);
    request.onsuccess = () => resolve();
    request.onerror = () => reject(request.error || new Error("IndexedDB write failed."));
    tx.oncomplete = () => db.close();
    tx.onabort = () => {
      db.close();
      reject(tx.error || new Error("IndexedDB transaction aborted."));
    };
  });
}

async function clearRunSnapshot() {
  const db = await openRunSnapshotDb();
  if (!db) return;
  return new Promise((resolve, reject) => {
    const tx = db.transaction(RUN_SNAPSHOT_STORE, "readwrite");
    const request = tx.objectStore(RUN_SNAPSHOT_STORE).delete(RUN_SNAPSHOT_KEY);
    request.onsuccess = () => resolve();
    request.onerror = () => reject(request.error || new Error("IndexedDB delete failed."));
    tx.oncomplete = () => db.close();
    tx.onabort = () => {
      db.close();
      reject(tx.error || new Error("IndexedDB transaction aborted."));
    };
  });
}

function sanitizeManualPairs(manualPairs, candidateNames, useDb3Directly = false) {
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
    const xName = candidateLookup.get(pair.xName.toLowerCase());
    const yName = candidateLookup.get(pair.yName.toLowerCase());
    if (!xName || !yName || xName.toLowerCase() === yName.toLowerCase()) {
      removed.push(`${pair.xName} + ${pair.yName}`);
      return;
    }
    if (!canManualAssign(xName, "X", useDb3Directly) || !canManualAssign(yName, "Y", useDb3Directly)) {
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

function resolveBatchPairingPlan(filesPayload, baseOptions = {}) {
  const candidateNames = [...new Set(filesPayload.map((item) => item.name))];
  const byName = new Map(filesPayload.map((item) => [String(item.name || "").toLowerCase(), item]));

  if (baseOptions.manualPairingEnabled) {
    const sanitized = sanitizeManualPairs(baseOptions.manualPairs, candidateNames, false).pairs;
    const pairs = sanitized
      .map((pair) => {
        const xItem = byName.get(pair.xName.toLowerCase()) || null;
        const yItem = byName.get(pair.yName.toLowerCase()) || null;
        if (!xItem || !yItem) return null;
        return [xItem, yItem, clipMiddle(xItem.name || yItem.name || "", 64, 24), null];
      })
      .filter(Boolean);
    const used = new Set(
      pairs
        .flatMap((pair) => [pair[0]?.name, pair[1]?.name])
        .filter(Boolean)
        .map((name) => String(name).toLowerCase())
    );
    const leftovers = filesPayload.filter((item) => !used.has(String(item.name || "").toLowerCase()));
    return { pairs, leftovers, suggestions: [], manualPairs: sanitized };
  }

  const resolved = resolveDeepsoilPairingCandidates(filesPayload);
  return {
    ...resolved,
    manualPairs: resolved.pairs
      .map(([xFile, yFile]) => ({ xName: xFile?.name || "", yName: yFile?.name || "" }))
      .filter((pair) => pair.xName && pair.yName),
  };
}

function buildPairSuggestions(candidateNames, useDb3Directly = false, existingPairs = []) {
  const available = candidateNames.map((name) => ({ name }));
  const resolved = resolveDeepsoilPairingCandidates(available, { existingPairs });
  const suggestions = [
    ...resolved.pairs.map(([xFile, yFile, base]) => ({
      id: `${xFile?.name || ""}||${yFile?.name || ""}`,
      xName: xFile?.name || "",
      yName: yFile?.name || "",
      base: base || clipMiddle(xFile?.name || "", 64, 18),
      score: PAIR_SCORE_AUTO_THRESHOLD,
      auto: true,
      reason: "deepsoil pair",
    })),
    ...resolved.suggestions.map((item) => ({
      id: item.id || `${item.x?.name || ""}||${item.y?.name || ""}`,
      xName: item.x?.name || "",
      yName: item.y?.name || "",
      base: item.base || clipMiddle(item.x?.name || "", 64, 18),
      score: Number(item.score) || 0,
      auto: false,
      reason: item.reason || "",
    })),
  ];
  return suggestions.sort((a, b) => b.score - a.score || a.base.localeCompare(b.base));
}

function applyPairSuggestion(suggestion, options = {}) {
  if (!suggestion?.xName || !suggestion?.yName) return false;
  const added = addManualPair(suggestion.xName, suggestion.yName, {
    source: "suggestion",
    allowAxisFallback: true,
  });
  if (added && !options.silent) appendLog("info", `Pair suggestion applied: ${suggestion.base || suggestion.xName}`);
  return added;
}

function applyAllPairSuggestions() {
  let count = 0;
  for (const suggestion of state.pairSuggestions || []) {
    if (applyPairSuggestion(suggestion, { silent: true })) count += 1;
  }
  appendLog("info", count > 0 ? `${count} pair suggestion(s) applied.` : "No valid pair suggestions to apply.");
  renderAll();
}

function autoPair() {
  let count = 0;
  for (const suggestion of (state.pairSuggestions || []).filter((item) => item.auto)) {
    if (applyPairSuggestion(suggestion, { silent: true })) count += 1;
  }
  appendLog("info", count > 0 ? `${count} pair(s) auto-paired.` : "No high-confidence pairs found.");
  renderAll();
}

function detectPairs(files, useDb3Directly = false, manualPairs = [], manualPairingEnabled = false) {
  const names = files.map((file) => logicalCandidateName(file));
  const candidates = new Set(names.filter((name) => isInputCandidate(name, useDb3Directly)));
  const xlsxCandidates = [...candidates].filter((name) => name.toLowerCase().endsWith(".xlsx")).length;
  const dbCandidates = [...candidates].filter((name) => name.toLowerCase().endsWith(".db") || name.toLowerCase().endsWith(".db3")).length;

  if (useDb3Directly) {
    return {
      candidates: candidates.size,
      xlsxCandidates,
      dbCandidates,
      pairs: 0,
      missing: 0,
      singles: candidates.size,
      manualPairsApplied: 0,
    };
  }

  let found = 0;
  let missing = 0;
  let singles = 0;
  if (manualPairingEnabled) {
    const normalized = sanitizeManualPairs(manualPairs, [...candidates], useDb3Directly).pairs;
    found = normalized.length;
    const used = new Set(normalized.flatMap((pair) => [pair.xName, pair.yName]).map((name) => name.toLowerCase()));
    for (const name of candidates) {
      if (!used.has(name.toLowerCase())) singles += 1;
    }
  } else {
    const pairingPlan = resolveBatchPairingPlan([...candidates].map((name) => ({ name })), {
      manualPairingEnabled: false,
      manualPairs: [],
    });
    found = pairingPlan.pairs.length;
    singles = pairingPlan.leftovers.length;
    missing = pairingPlan.leftovers.filter((item) => inferAxisLabel(item.name) === "X").length;
  }

  return { candidates: candidates.size, xlsxCandidates, dbCandidates, pairs: found, missing, singles, manualPairsApplied: manualPairingEnabled ? found : 0 };
}

function summarizeLoadedTypes() {
  let xlsx = 0;
  let db = 0;
  state.selectedFiles.forEach((file) => {
    const lower = logicalCandidateName(file).toLowerCase();
    if (lower.endsWith(".xlsx")) xlsx += 1;
    else if (lower.endsWith(".db") || lower.endsWith(".db3")) db += 1;
  });
  return { xlsx, db, total: xlsx + db };
}

function base64ToBlob(base64, mimeType) {
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index += 1) bytes[index] = binary.charCodeAt(index);
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

function openArtifact(result) {
  if (!result?.outputBytesB64) return;
  const blob = base64ToBlob(result.outputBytesB64, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  downloadBlob(blob, result.outputFileName || "output.xlsx");
}

function findSourceForResult(result, scene) {
  const pairKey = String(result?.pairKey || "");
  if (!pairKey) return null;
  return scene.sources.find((source) => Array.isArray(source.artifactPairKeys) && source.artifactPairKeys.includes(pairKey)) || null;
}

function buildArtifactFocusPrefs(result, scene) {
  const source = findSourceForResult(result, scene);
  if (!source) return null;
  const mode = String(result?.metrics?.mode || "").toLowerCase();
  let familyKey = "";
  let chartKey = "";
  if (mode === "method3" || mode === "method3_aggregate") {
    familyKey = "method3-aggregate";
  } else if (mode === "method2_single") {
    familyKey = "layer-series";
    chartKey = "layer-tbdy-total";
  } else if (mode === "single") {
    familyKey = "derived-profiles";
    chartKey = "derived-total-profile";
  } else if (mode === "pair") {
    familyKey = "derived-profiles";
    chartKey = "derived-resultant-profile";
  }
  return {
    chartMode: "focus",
    activeSourceId: source.sourceId,
    activeFamilyKey: familyKey,
    activeChartKey: chartKey,
    activeLayerIndex: 0,
    compareSourceIds: [],
    seriesVisibilityMap: {},
  };
}

function focusArtifact(result) {
  const scene = syncSourceViewerSelection();
  const prefs = buildArtifactFocusPrefs(result, scene);
  if (!prefs) return;
  applyViewerPrefs(prefs, { persist: true });
  renderAll();
}

function renderSourceList(scene) {
  if (!dom.sourceList) return;
  dom.sourceList.innerHTML = "";
  dom.sourceMeta.textContent = `${scene.sourceCount} sources`;
  if (!scene.sources.length) {
    dom.sourceList.innerHTML = '<div class="empty-block">Henüz source catalog yok. Dosyaları seçip batch çalıştır.</div>';
    return;
  }

  scene.sources.forEach((source) => {
    const chartCount = source.families.reduce((sum, family) => sum + family.charts.length, 0);
    const button = document.createElement("button");
    button.type = "button";
    button.className = `list-item${source.sourceId === scene.activeSourceId ? " is-active" : ""}`;
    button.title = source.sourceLabel;
    button.innerHTML = `
      <small>${escapeHtml(source.sourceKind)}${source.axis ? ` · ${escapeHtml(source.axis)}` : ""}</small>
      <strong>${escapeHtml(clipMiddle(source.sourceLabel, 72, 18))}</strong>
      <span>${source.families.length} families · ${chartCount} charts</span>
    `;
    button.addEventListener("click", () => setActiveSource(source.sourceId));
    dom.sourceList.appendChild(button);
  });
}

function renderFamilyList(scene) {
  if (!dom.familyList) return;
  dom.familyList.innerHTML = "";
  const families = scene.activeSource?.families || [];
  dom.familyMeta.textContent = `${families.length} families`;
  if (!families.length) {
    dom.familyList.innerHTML = '<div class="empty-block">Aktif source seçildiğinde family listesi burada görünür.</div>';
    return;
  }

  families.forEach((family) => {
    const layerInfo = family.supportsLayerSelection && family.layers?.length ? `${family.layers.length} layers` : `${family.charts.length} charts`;
    const button = document.createElement("button");
    button.type = "button";
    button.className = `list-item${family.familyKey === scene.activeFamilyKey ? " is-active" : ""}`;
    button.innerHTML = `
      <small>${escapeHtml(family.chartType)}${family.supportsOverlay ? " · overlay" : ""}</small>
      <strong>${escapeHtml(family.familyLabel)}</strong>
      <span>${family.charts.length} charts · ${layerInfo}</span>
    `;
    button.addEventListener("click", () => setActiveFamily(family.familyKey));
    dom.familyList.appendChild(button);
  });
}

function renderChartList(scene) {
  if (!dom.chartList) return;
  dom.chartList.innerHTML = "";
  const charts = scene.activeFamily?.charts || [];
  dom.chartMeta.textContent = `${charts.length} charts`;
  if (!charts.length) {
    dom.chartList.innerHTML = '<div class="empty-block">Aktif source seçildiğinde chart listesi burada görünür.</div>';
    return;
  }

  charts.forEach((chart) => {
    const layerCount = Array.isArray(chart.layerViews) ? chart.layerViews.length : 0;
    const seriesCount = layerCount ? chart.layerViews?.[0]?.series?.length || 0 : chart.series?.length || 0;
    const button = document.createElement("button");
    button.type = "button";
    button.className = `list-item${chart.chartKey === scene.activeChartKey ? " is-active" : ""}`;
    button.innerHTML = `
      <small>${escapeHtml(chart.sheetName)}${layerCount ? ` · ${layerCount} layers` : ""}</small>
      <strong>${escapeHtml(chart.chartLabel)}</strong>
      <span>${escapeHtml(chart.chartType)} · ${seriesCount} traces</span>
    `;
    button.addEventListener("click", () => setActiveChart(chart.chartKey));
    dom.chartList.appendChild(button);
  });
}

function renderCompareList(scene) {
  if (!dom.compareSourceList) return;
  dom.compareSourceList.innerHTML = "";
  dom.compareMeta.textContent = `${scene.compareSourceIds.length} selected`;

  const xCandidateIds = (scene.compareCandidates || [])
    .filter((source) => String(source?.axis || "").trim().toUpperCase() === "X")
    .map((source) => String(source.sourceId || ""))
    .filter(Boolean);
  const yCandidateIds = (scene.compareCandidates || [])
    .filter((source) => String(source?.axis || "").trim().toUpperCase() === "Y")
    .map((source) => String(source.sourceId || ""))
    .filter(Boolean);

  if (dom.compareAxisXToggle) {
    dom.compareAxisXToggle.disabled = !xCandidateIds.length;
    dom.compareAxisXToggle.checked =
      xCandidateIds.length > 0 && xCandidateIds.every((sourceId) => scene.compareSourceIds.includes(sourceId));
  }
  if (dom.compareAxisYToggle) {
    dom.compareAxisYToggle.disabled = !yCandidateIds.length;
    dom.compareAxisYToggle.checked =
      yCandidateIds.length > 0 && yCandidateIds.every((sourceId) => scene.compareSourceIds.includes(sourceId));
  }
  if (dom.compareClearBtn) dom.compareClearBtn.disabled = scene.compareSourceIds.length === 0;

  if (!scene.compareCandidates.length) {
    dom.compareSourceList.innerHTML = '<div class="empty-block">Karşılaştırılabilir başka source yok.</div>';
    return;
  }

  scene.compareCandidates.forEach((source) => {
    const line = document.createElement("label");
    line.className = "toggle-line";
    line.innerHTML = `
      <input type="checkbox" ${scene.compareSourceIds.includes(source.sourceId) ? "checked" : ""} />
      <span class="toggle-line__body">
        <strong title="${escapeHtml(source.sourceLabel)}">${escapeHtml(clipMiddle(source.sourceLabel, 56, 16))}</strong>
        <span>${escapeHtml(source.sourceKind)}${source.axis ? ` · ${escapeHtml(source.axis)}` : ""}</span>
      </span>
    `;
    const checkbox = line.querySelector("input");
    checkbox.addEventListener("change", () => toggleCompareSource(source.sourceId, checkbox.checked));
    dom.compareSourceList.appendChild(line);
  });
}

function renderAxisControls(scene) {
  const chartType = String(scene.activeChart?.chartType || "").toLowerCase();
  const isSpectrum = chartType === "spectrum";

  if (dom.periodMaxInput) {
    dom.periodMaxInput.disabled = !isSpectrum;
    dom.periodMaxInput.value = String(state.spectrumPeriodMax || "");
    if (!isSpectrum) dom.periodMaxInput.placeholder = "spectrum only";
    else dom.periodMaxInput.placeholder = "auto";
  }
  if (dom.periodMaxResetBtn) dom.periodMaxResetBtn.disabled = !isSpectrum || !String(state.spectrumPeriodMax || "").trim();
  if (dom.logXToggle) {
    dom.logXToggle.checked = !!state.logX;
    dom.logXToggle.disabled = !scene.activeChart;
  }
  if (dom.logYToggle) {
    dom.logYToggle.checked = !!state.logY;
    dom.logYToggle.disabled = !scene.activeChart;
  }
}

function renderSeriesToggleList(scene) {
  if (!dom.seriesToggleList) return;
  dom.seriesToggleList.innerHTML = "";
  const entries = scene.plotSeriesEntries || [];
  const visible = entries.filter((entry) => entry.visible !== false).length;
  dom.seriesMeta.textContent = `${visible}/${entries.length} visible`;
  if (!entries.length) {
    dom.seriesToggleList.innerHTML = '<div class="empty-block">Aktif chart için trace toggle listesi burada görünür.</div>';
    return;
  }

  entries.forEach((entry, index) => {
    const line = document.createElement("label");
    line.className = "toggle-line";
    line.innerHTML = `
      <input type="checkbox" ${entry.visible !== false ? "checked" : ""} />
      <span class="toggle-line__body">
        <strong>${escapeHtml(clipMiddle(entry.displayLabel, 54, 14))}</strong>
        <span>${escapeHtml(entry.layerLabel || scene.activeChart?.chartType || "series")}</span>
      </span>
    `;
    line.style.borderLeft = `4px solid ${PLOT_COLORS[index % PLOT_COLORS.length]}`;
    const checkbox = line.querySelector("input");
    checkbox.addEventListener("change", () => toggleSeriesVisibility(entry.entryKey, checkbox.checked));
    dom.seriesToggleList.appendChild(line);
  });
}

function renderStage(scene) {
  if (dom.catalogMeta) dom.catalogMeta.textContent = `${scene.sourceCount} sources | ${scene.familyCount} families | ${scene.chartCount} charts`;
  if (dom.stageTitle) dom.stageTitle.textContent = scene.activeChart?.chartLabel || "Focused chart";
  if (dom.stageSubtitle) {
    dom.stageSubtitle.textContent = scene.activeSource
      ? `${clipMiddle(scene.activeSource.sourceLabel, 96, 22)} · ${scene.activeFamily?.familyLabel || "-"} · ${scene.activeChart?.sheetName || "-"}`
      : "Kaynak eğriler için hazır.";
  }
  if (dom.activeSourceLabel) {
    dom.activeSourceLabel.textContent = scene.activeSource?.sourceLabel || "-";
    dom.activeSourceLabel.title = scene.activeSource?.sourceLabel || "";
  }
  if (dom.activeFamilyLabel) dom.activeFamilyLabel.textContent = scene.activeFamily?.familyLabel || "-";
  if (dom.activeChartLabel) dom.activeChartLabel.textContent = scene.activeChart?.chartLabel || "-";
  if (dom.activeLayerLabel) dom.activeLayerLabel.textContent = scene.activeLayerCount > 0 ? scene.activeLayerLabel || "Layer 1" : "-";

  const sourceCount = scene.sources.length;
  const chartCount = scene.activeFamily?.charts?.length || 0;
  if (dom.sourceCounter) dom.sourceCounter.textContent = sourceCount ? `${scene.activeSourceIndex + 1} / ${sourceCount}` : "0 / 0";
  if (dom.chartCounter) {
    const activeChartIndex = chartCount ? scene.activeFamily.charts.findIndex((chart) => chart.chartKey === scene.activeChartKey) + 1 : 0;
    dom.chartCounter.textContent = chartCount ? `${activeChartIndex} / ${chartCount}` : "0 / 0";
  }
  if (dom.layerCounter) dom.layerCounter.textContent = scene.activeLayerCount > 0 ? `${scene.activeLayerIndex + 1} / ${scene.activeLayerCount}` : "-";

  if (dom.sourcePrevBtn) dom.sourcePrevBtn.disabled = sourceCount <= 1;
  if (dom.sourceNextBtn) dom.sourceNextBtn.disabled = sourceCount <= 1;
  if (dom.chartPrevBtn) dom.chartPrevBtn.disabled = chartCount <= 1;
  if (dom.chartNextBtn) dom.chartNextBtn.disabled = chartCount <= 1;
  if (dom.layerPrevBtn) dom.layerPrevBtn.disabled = scene.activeLayerCount <= 1;
  if (dom.layerNextBtn) dom.layerNextBtn.disabled = scene.activeLayerCount <= 1;
}

function schedulePlotStabilize(Plotly, traces, layout, config, attempt = 0) {
  if (plotStabilizeTimer) globalThis.clearTimeout(plotStabilizeTimer);
  plotStabilizeTimer = globalThis.setTimeout(() => {
    try {
      Plotly.Plots.resize(dom.plotHost);
    } catch {
      // Ignore resize failures while the plot host settles.
    }

    const renderedTraceCount = dom.plotHost?.querySelectorAll(".scatterlayer .trace, .barlayer .trace").length || 0;
    if (renderedTraceCount > 0 || attempt >= 2) return;

    Promise.resolve(Plotly.newPlot(dom.plotHost, traces, layout, config))
      .then(() => schedulePlotStabilize(Plotly, traces, layout, config, attempt + 1))
      .catch(() => {
        // Ignore replot failures here; the next UI event will retry.
      });
  }, attempt === 0 ? 60 : 140);
}

function colorForEntry(index) {
  return PLOT_COLORS[index % PLOT_COLORS.length];
}

function buildPlotTrace(entry, index, chartType, spectrumPeriodMax) {
  let points = Array.isArray(entry.points) ? entry.points.slice() : [];
  if (chartType === "spectrum" || chartType === "fourier") {
    points = points.sort((left, right) => Number(left?.x || 0) - Number(right?.x || 0));
  }
  if (chartType === "spectrum" && Number.isFinite(spectrumPeriodMax) && spectrumPeriodMax > 0) {
    const clipped = points.filter((point) => Number(point?.x) <= spectrumPeriodMax);
    if (clipped.length >= 2) points = clipped;
  }
  return {
    type: "scatter",
    mode: "lines",
    name: entry.displayLabel,
    x: points.map((point) => point.x),
    y: points.map((point) => point.y),
    line: { color: colorForEntry(index), width: entry.sourceId === state.activeSourceId ? 3 : 2 },
    hovertemplate: "%{x}<br>%{y}<extra></extra>",
  };
}

function renderPlot(scene) {
  if (!dom.plotHost) return;
  const Plotly = globalThis.Plotly;
  const entries = (scene.plotSeriesEntries || []).filter((entry) => entry.visible !== false);
  if (!scene.activeChart || !scene.activeFamily || !scene.activeSource || !entries.length || !Plotly) {
    try {
      Plotly?.purge(dom.plotHost);
    } catch {
      // Ignore Plotly purge errors when no plot exists yet.
    }
    dom.plotHost.innerHTML = `<div class="empty-block empty-block--plot">${
      scene.activeChart ? "Bu chart için görünür trace kalmadı." : "Grafik sahnesi burada görünecek."
    }</div>`;
    return;
  }

  dom.plotHost.innerHTML = "";
  const chartType = String(scene.activeChart?.chartType || "").toLowerCase();
  const requestedPeriodMax = Number(String(state.spectrumPeriodMax || "").trim());
  const spectrumPeriodMax = chartType === "spectrum" && Number.isFinite(requestedPeriodMax) && requestedPeriodMax > 0 ? requestedPeriodMax : null;
  const traces = entries.map((entry, index) => buildPlotTrace(entry, index, chartType, spectrumPeriodMax));

  const allX = traces.flatMap((trace) => (Array.isArray(trace.x) ? trace.x : [])).map((value) => Number(value));
  const allY = traces.flatMap((trace) => (Array.isArray(trace.y) ? trace.y : [])).map((value) => Number(value));
  const positiveX = allX.filter((value) => Number.isFinite(value) && value > 0);
  const positiveY = allY.filter((value) => Number.isFinite(value) && value > 0);
  const canLogX = !!state.logX && positiveX.length > 0;
  const canLogY = !!state.logY && positiveY.length > 0;
  const layout = {
    paper_bgcolor: "rgba(0,0,0,0)",
    plot_bgcolor: "rgba(255,255,255,0.82)",
    margin: { l: 88, r: 28, t: 18, b: 78 },
    showlegend: false,
    hovermode: "closest",
    dragmode: "pan",
    uirevision: `${scene.activeSourceId}|${scene.activeFamilyKey}|${scene.activeChartKey}|${scene.activeLayerIndex}|${scene.chartMode}`,
    font: { family: "IBM Plex Mono, monospace", size: 12, color: "#161a1f" },
    xaxis: {
      title: { text: scene.activeChart.xLabel || "X" },
      type: canLogX ? "log" : "linear",
      automargin: true,
      showline: true,
      linecolor: "rgba(22,26,31,0.4)",
      gridcolor: "rgba(22,26,31,0.08)",
      zeroline: false,
      tickfont: { size: 11 },
    },
    yaxis: {
      title: { text: scene.activeChart.yLabel || "Y" },
      type: canLogY ? "log" : "linear",
      automargin: true,
      showline: true,
      linecolor: "rgba(22,26,31,0.4)",
      gridcolor: "rgba(22,26,31,0.08)",
      zeroline: false,
      tickfont: { size: 11 },
      autorange: scene.activeChart.invertY ? "reversed" : true,
    },
  };

  if (chartType === "spectrum" && spectrumPeriodMax) {
    if (canLogX) {
      const minPositive = Math.min(...positiveX);
      if (Number.isFinite(minPositive) && minPositive > 0 && spectrumPeriodMax > minPositive) {
        layout.xaxis.range = [Math.log10(minPositive), Math.log10(spectrumPeriodMax)];
      }
    } else {
      layout.xaxis.range = [0, spectrumPeriodMax];
    }
  }
  const config = {
    responsive: true,
    displaylogo: false,
    scrollZoom: true,
    modeBarButtonsToRemove: ["lasso2d", "select2d", "autoScale2d"],
  };
  Promise.resolve(Plotly.newPlot(dom.plotHost, traces, layout, config))
    .then(() => {
      schedulePlotStabilize(Plotly, traces, layout, config);
    })
    .catch(() => {
      // Ignore render promise failures here; the next state change will retry.
    });
}

function renderMetricCards(scene) {
  if (!dom.metrics) return;
  const loaded = summarizeLoadedTypes();
  const pairStats = detectPairs(
    state.selectedFiles,
    !!dom.useDb3Directly?.checked,
    state.manualPairs,
    state.shellMode === "legacy" && !!dom.manualPairingEnabled?.checked
  );
  const metrics = state.metrics || {};
  const cards = [
    ["Shell mode", state.shellMode],
    ["View mode", state.viewerChartMode],
    ["Active source", scene.activeSource?.sourceLabel || "-"],
    ["Active chart", scene.activeChart?.chartLabel || "-"],
    ["Loaded XLSX", String(loaded.xlsx)],
    ["Loaded DB", String(loaded.db)],
    ["Sources", String(scene.sourceCount)],
    ["Families", String(scene.familyCount)],
    ["Charts", String(scene.chartCount)],
    ["Pairs", String(pairStats.pairs)],
    ["Singles", String(pairStats.singles)],
    ["Processed", String(metrics.processedTotal ?? state.results.length ?? 0)],
    ["Failed", String(metrics.failedTotal ?? state.errors.length ?? 0)],
  ];

  dom.metricMeta.textContent = `${cards.length} stats`;
  dom.metrics.innerHTML = cards
    .map(
      ([label, value]) => `
        <div class="metric-card">
          <span>${escapeHtml(label)}</span>
          <strong title="${escapeHtml(String(value))}">${escapeHtml(clipMiddle(String(value), 44, 14))}</strong>
        </div>
      `
    )
    .join("");
}

function renderArtifacts() {
  const summary = summarizeViewerResults(state.results);
  if (dom.resultPreviewMeta) dom.resultPreviewMeta.textContent = `${summary.totalResults} outputs`;
  if (dom.resultPreviewDeck) {
    dom.resultPreviewDeck.innerHTML = summary.groups.length
      ? summary.groups
          .map(
            (group) =>
              `<div class="artifact-badge">${escapeHtml(group.label)} · ${group.results.length} outputs · ${group.chartCount} preview charts</div>`
          )
          .join("")
      : '<div class="empty-block">Henüz artifact yok.</div>';
  }

  if (!dom.resultTableBody) return;
  dom.resultTableBody.innerHTML = "";
  if (!state.results.length) {
    dom.resultTableBody.innerHTML = '<tr><td colspan="5">No batch outputs yet.</td></tr>';
    return;
  }

  state.results.forEach((result, index) => {
    const row = document.createElement("tr");
    const outputFileName = String(result.outputFileName || "");
    const mode = String(result.metrics?.mode || result.viewerKind || "-");
    row.innerHTML = `
      <td>${escapeHtml(result.viewerGroup || "-")}</td>
      <td title="${escapeHtml(result.pairKey || "-")}">${escapeHtml(clipMiddle(result.pairKey || "-", 48, 18))}</td>
      <td>${escapeHtml(mode)}</td>
      <td title="${escapeHtml(outputFileName)}">${escapeHtml(clipMiddle(outputFileName, 56, 18))}</td>
      <td>
        <div class="result-actions">
          <button class="focus-artifact" type="button" data-index="${index}">Focus</button>
          <button class="download-artifact" type="button" data-index="${index}">Download</button>
        </div>
      </td>
    `;
    dom.resultTableBody.appendChild(row);
  });

  dom.resultTableBody.querySelectorAll(".focus-artifact").forEach((button) => {
    button.addEventListener("click", () => {
      const result = state.results[Number(button.dataset.index)];
      if (result) focusArtifact(result);
    });
  });

  dom.resultTableBody.querySelectorAll(".download-artifact").forEach((button) => {
    button.addEventListener("click", () => {
      const result = state.results[Number(button.dataset.index)];
      if (result) openArtifact(result);
    });
  });
}

function getManualSelection() {
  return {
    xName: String(state.manualStage?.xName || dom.manualXSelect?.value || "").trim(),
    yName: String(state.manualStage?.yName || dom.manualYSelect?.value || "").trim(),
  };
}

function renderManualStage() {
  const selection = getManualSelection();
  if (dom.manualStageX) dom.manualStageX.textContent = selection.xName || "No X picked";
  if (dom.manualStageY) dom.manualStageY.textContent = selection.yName || "No Y picked";
}

function renderSelectedFiles() {
  if (!dom.selectedFilesList) return;
  dom.selectedFilesList.innerHTML = "";
  const totalBytes = state.selectedFiles.reduce((sum, file) => sum + (file.size || 0), 0);
  dom.selectedFilesMeta.textContent = `${state.selectedFiles.length} files | ${formatBytes(totalBytes)}`;

  if (!state.selectedFiles.length) {
    dom.selectedFilesList.innerHTML = '<li class="empty">No files selected.</li>';
    return;
  }

  const manualEnabled = state.shellMode === "legacy" && !!dom.manualPairingEnabled?.checked;
  const useDb3Directly = !!dom.useDb3Directly?.checked;
  const usedX = new Set(state.manualPairs.map((pair) => pair.xName.toLowerCase()));
  const usedY = new Set(state.manualPairs.map((pair) => pair.yName.toLowerCase()));
  const pickedX = String(state.manualStage?.xName || "");
  const pickedY = String(state.manualStage?.yName || "");

  state.selectedFiles.slice(0, 120).forEach((file) => {
    const logicalName = logicalCandidateName(file);
    const li = document.createElement("li");
    li.title = sourcePathForFile(file) || logicalName;
    li.classList.toggle("is-picked", pickedX === logicalName || pickedY === logicalName);

    const metaChunks = [formatBytes(file.size || 0)];
    const axis = inferAxisLabel(logicalName);
    if (axis && axis !== "SINGLE") metaChunks.push(axis);
    li.innerHTML = `
      <span class="name">${escapeHtml(clipMiddle(logicalName, 72, 18))}</span>
      <span class="file-meta"><span class="size">${escapeHtml(metaChunks.join(" · "))}</span></span>
    `;

    if (manualEnabled && isInputCandidate(logicalName, useDb3Directly)) {
      const metaEl = li.querySelector(".file-meta");
      const actionWrap = document.createElement("span");
      actionWrap.className = "file-actions";

      if (canManualAssign(logicalName, "X", useDb3Directly) && !usedX.has(logicalName.toLowerCase())) {
        const button = document.createElement("button");
        button.type = "button";
        button.className = "pick-manual";
        button.dataset.role = "x";
        button.dataset.name = logicalName;
        button.textContent = pickedX === logicalName ? "Picked X" : "Pick X";
        actionWrap.appendChild(button);
      }
      if (canManualAssign(logicalName, "Y", useDb3Directly) && !usedY.has(logicalName.toLowerCase())) {
        const button = document.createElement("button");
        button.type = "button";
        button.className = "pick-manual";
        button.dataset.role = "y";
        button.dataset.name = logicalName;
        button.textContent = pickedY === logicalName ? "Picked Y" : "Pick Y";
        actionWrap.appendChild(button);
      }
      if (usedX.has(logicalName.toLowerCase())) {
        const chip = document.createElement("span");
        chip.className = "manual-chip";
        chip.textContent = "Paired X";
        actionWrap.appendChild(chip);
      }
      if (usedY.has(logicalName.toLowerCase())) {
        const chip = document.createElement("span");
        chip.className = "manual-chip";
        chip.textContent = "Paired Y";
        actionWrap.appendChild(chip);
      }
      metaEl?.appendChild(actionWrap);
    }

    dom.selectedFilesList.appendChild(li);
  });

  if (state.selectedFiles.length > 120) {
    const li = document.createElement("li");
    li.className = "empty";
    li.textContent = `... ${state.selectedFiles.length - 120} more files`;
    dom.selectedFilesList.appendChild(li);
  }

  dom.selectedFilesList.querySelectorAll(".pick-manual").forEach((button) => {
    button.addEventListener("click", () => stageManualCandidate(button.dataset.role || "", button.dataset.name || ""));
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
  selectEl.value = values.includes(current) ? current : "";
}

function syncManualSelectValues() {
  if (dom.manualXSelect) {
    dom.manualXSelect.value = Array.from(dom.manualXSelect.options).some((opt) => opt.value === state.manualStage.xName)
      ? state.manualStage.xName
      : "";
  }
  if (dom.manualYSelect) {
    dom.manualYSelect.value = Array.from(dom.manualYSelect.options).some((opt) => opt.value === state.manualStage.yName)
      ? state.manualStage.yName
      : "";
  }
}

function renderPairSuggestions() {
  if (!dom.pairSuggestionsList || !dom.pairSuggestionsMeta || !dom.pairSuggestionsPanel) return;
  const manualEnabled = state.shellMode === "legacy" && !!dom.manualPairingEnabled?.checked;
  const suggestions = manualEnabled ? state.pairSuggestions || [] : [];
  dom.pairSuggestionsPanel.classList.toggle("is-hidden", !manualEnabled);
  dom.pairSuggestionsMeta.textContent = `${suggestions.length} suggestions`;
  if (dom.applyAllSuggestionsBtn) dom.applyAllSuggestionsBtn.disabled = state.isRunning || suggestions.length === 0;
  if (dom.autoPairBtn) dom.autoPairBtn.disabled = state.isRunning || suggestions.filter((item) => item.auto).length === 0;
  dom.pairSuggestionsList.innerHTML = "";

  if (!manualEnabled) {
    dom.pairSuggestionsList.innerHTML = '<li class="empty">Manual pairing kapalı.</li>';
    return;
  }
  if (!suggestions.length) {
    dom.pairSuggestionsList.innerHTML = '<li class="empty">No pair suggestions.</li>';
    return;
  }

  suggestions.forEach((suggestion) => {
    const li = document.createElement("li");
    li.className = "pair-suggestion-item";
    li.innerHTML = `
      <div class="pair-suggestion-main">
        <div class="pair-suggestion-title" title="${escapeHtml(`${suggestion.base} | ${suggestion.xName} -> ${suggestion.yName}`)}">
          ${escapeHtml(suggestion.base)} — ${escapeHtml(clipMiddle(suggestion.xName, 36, 12))} → ${escapeHtml(clipMiddle(suggestion.yName, 36, 12))}
        </div>
        <div class="pair-suggestion-meta">Score ${escapeHtml(String(suggestion.score))}${suggestion.auto ? " · auto" : ""}${suggestion.reason ? ` · ${escapeHtml(suggestion.reason)}` : ""}</div>
      </div>
      <button type="button" class="apply-pair-suggestion">Apply</button>
    `;
    const applyBtn = li.querySelector(".apply-pair-suggestion");
    if (applyBtn) {
      applyBtn.dataset.x = suggestion.xName;
      applyBtn.dataset.y = suggestion.yName;
    }
    dom.pairSuggestionsList.appendChild(li);
  });

  dom.pairSuggestionsList.querySelectorAll(".apply-pair-suggestion").forEach((button) => {
    button.addEventListener("click", () => {
      const suggestion = (state.pairSuggestions || []).find(
        (item) => item.xName === String(button.getAttribute("data-x") || "") && item.yName === String(button.getAttribute("data-y") || "")
      );
      if (!suggestion) return;
      applyPairSuggestion(suggestion);
      renderAll();
    });
  });
}

function renderManualPairing() {
  if (!dom.manualPairPanel || !dom.manualPairList || !dom.manualPairMeta) return;
  const manualEnabled = state.shellMode === "legacy" && !!dom.manualPairingEnabled?.checked;
  const useDb3Directly = !!dom.useDb3Directly?.checked;
  const candidateNames = getCandidateNames(useDb3Directly);
  const { pairs, removed } = candidateNames.length > 0
    ? sanitizeManualPairs(state.manualPairs, candidateNames, useDb3Directly)
    : { pairs: normalizePersistedManualPairs(state.manualPairs), removed: [] };
  state.manualPairs = pairs;
  persistManualPairPrefs();
  if (removed.length > 0) appendLog("warning", `Manual pairs pruned after file/mode change: ${removed.length}`);

  dom.manualPairPanel.classList.toggle("is-hidden", !manualEnabled);
  dom.manualPairMeta.textContent = `${state.manualPairs.length} pairs`;
  renderManualStage();

  const usedX = new Set(state.manualPairs.map((pair) => pair.xName.toLowerCase()));
  const usedY = new Set(state.manualPairs.map((pair) => pair.yName.toLowerCase()));
  const xCandidates = candidateNames.filter((name) => canManualAssign(name, "X", useDb3Directly) && !usedX.has(name.toLowerCase()));
  const yCandidates = candidateNames.filter((name) => canManualAssign(name, "Y", useDb3Directly) && !usedY.has(name.toLowerCase()));
  fillSelectOptions(dom.manualXSelect, xCandidates, "Select X file");
  fillSelectOptions(dom.manualYSelect, yCandidates, "Select Y file");
  syncManualSelectValues();

  state.pairSuggestions = manualEnabled ? buildPairSuggestions(candidateNames, useDb3Directly, state.manualPairs) : [];
  renderPairSuggestions();

  dom.manualPairList.innerHTML = "";
  if (!manualEnabled) {
    dom.manualPairList.innerHTML = '<li class="empty">Manual pairing kapalı.</li>';
    return;
  }
  if (!state.manualPairs.length) {
    dom.manualPairList.innerHTML = '<li class="empty">Henüz manuel pair yok.</li>';
    return;
  }

  state.manualPairs.forEach((pair, index) => {
    const li = document.createElement("li");
    li.innerHTML = `
      <div class="pair-name" title="${escapeHtml(pair.xName)} + ${escapeHtml(pair.yName)}">
        <span class="pair-tag">X</span> ${escapeHtml(pair.xName)} → <span class="pair-tag">Y</span> ${escapeHtml(pair.yName)}
      </div>
      <button class="manual-pair-remove" type="button" data-index="${index}">Remove</button>
    `;
    dom.manualPairList.appendChild(li);
  });

  dom.manualPairList.querySelectorAll(".manual-pair-remove").forEach((button) => {
    button.addEventListener("click", () => {
      const index = Number(button.dataset.index);
      const removedPair = state.manualPairs[index];
      state.manualPairs.splice(index, 1);
      appendLog("info", `Manual pair removed: ${removedPair?.xName || "-"} + ${removedPair?.yName || "-"}`);
      renderAll();
    });
  });
}

function updateSelectionStats() {
  const loaded = summarizeLoadedTypes();
  const manualActive = state.shellMode === "legacy" && !!dom.manualPairingEnabled?.checked;
  const stats = detectPairs(state.selectedFiles, !!dom.useDb3Directly?.checked, state.manualPairs, manualActive);
  if (dom.pairStats) dom.pairStats.textContent = `Pairs: ${stats.pairs} | Singles: ${stats.singles} | Missing Y: ${stats.missing}`;
  if (dom.countStats) dom.countStats.textContent = `Active: ${stats.candidates} | Loaded XLSX: ${loaded.xlsx} | Loaded DB: ${loaded.db}`;
}

function refreshButtons() {
  const canRun = state.workerReady && !state.isRunning && state.selectedFiles.length > 0;
  const selection = getManualSelection();
  const legacyMode = state.shellMode === "legacy";

  if (dom.runBtn) dom.runBtn.disabled = !canRun;
  if (dom.zipBtn) dom.zipBtn.disabled = state.isRunning || state.results.length === 0;
  if (dom.clearAllBtn) dom.clearAllBtn.disabled = state.isRunning || (!state.selectedFiles.length && !state.results.length && !state.sourceCatalog.length && !state.logs.length);
  if (dom.folderInput) dom.folderInput.disabled = state.isRunning;
  if (dom.fileInput) dom.fileInput.disabled = state.isRunning;
  if (dom.manualPairingEnabled) dom.manualPairingEnabled.disabled = state.isRunning || !legacyMode;
  if (dom.manualXSelect) dom.manualXSelect.disabled = state.isRunning || !legacyMode || !dom.manualPairingEnabled?.checked;
  if (dom.manualYSelect) dom.manualYSelect.disabled = state.isRunning || !legacyMode || !dom.manualPairingEnabled?.checked;
  if (dom.addManualPairBtn) dom.addManualPairBtn.disabled = state.isRunning || !legacyMode || !selection.xName || !selection.yName;
  if (dom.clearFilesBtn) dom.clearFilesBtn.disabled = state.isRunning || state.selectedFiles.length === 0;
  if (dom.clearPairsBtn) dom.clearPairsBtn.disabled = state.isRunning || !legacyMode || state.manualPairs.length === 0;
  if (dom.clearManualStageBtn) dom.clearManualStageBtn.disabled = state.isRunning || !legacyMode || (!selection.xName && !selection.yName);
  if (dom.failFast) dom.failFast.disabled = state.isRunning || !legacyMode;
  if (dom.method2Enabled) dom.method2Enabled.disabled = state.isRunning;
  if (dom.method3Enabled) dom.method3Enabled.disabled = state.isRunning;
  if (dom.useDb3Directly) dom.useDb3Directly.disabled = state.isRunning || !legacyMode;
  if (dom.integrationCompareEnabled) dom.integrationCompareEnabled.disabled = state.isRunning || !legacyMode;
  if (dom.includeResultantProfiles) dom.includeResultantProfiles.disabled = state.isRunning;
}

function syncFilterInputs() {
  const useDb3Directly = !!dom.useDb3Directly?.checked;
  const baselineOn = !!dom.baselineOn?.checked;
  const filterOn = !!dom.filterOn?.checked;
  const cfg = (dom.filterConfig?.value || "bandpass").toLowerCase();
  const isHighpass = cfg === "highpass" || cfg === "high";
  const isLowpass = cfg === "lowpass" || cfg === "low";
  const shellLocked = state.shellMode !== "legacy";

  if (dom.baseReference) dom.baseReference.disabled = shellLocked || useDb3Directly || state.isRunning;
  if (dom.integrationCompareEnabled) dom.integrationCompareEnabled.disabled = shellLocked || useDb3Directly || state.isRunning;
  if (dom.baselineOn) dom.baselineOn.disabled = shellLocked || useDb3Directly || state.isRunning;
  if (dom.filterOn) dom.filterOn.disabled = shellLocked || useDb3Directly || state.isRunning;
  if (dom.baselineMethod) dom.baselineMethod.disabled = shellLocked || useDb3Directly || !baselineOn || state.isRunning;
  if (dom.processingOrder) dom.processingOrder.disabled = shellLocked || useDb3Directly || !filterOn || state.isRunning;
  if (dom.filterDomain) dom.filterDomain.disabled = shellLocked || useDb3Directly || !filterOn || state.isRunning;
  if (dom.filterConfig) dom.filterConfig.disabled = shellLocked || useDb3Directly || !filterOn || state.isRunning;
  if (dom.filterType) dom.filterType.disabled = shellLocked || useDb3Directly || !filterOn || state.isRunning;
  if (dom.filterOrder) dom.filterOrder.disabled = shellLocked || useDb3Directly || !filterOn || state.isRunning;
  if (dom.fLowHz) dom.fLowHz.disabled = shellLocked || useDb3Directly || !filterOn || isHighpass || state.isRunning;
  if (dom.fHighHz) dom.fHighHz.disabled = shellLocked || useDb3Directly || !filterOn || isLowpass || state.isRunning;
}

function renderAll() {
  const scene = syncSourceViewerSelection();
  renderSourceList(scene);
  renderFamilyList(scene);
  renderChartList(scene);
  renderCompareList(scene);
  renderAxisControls(scene);
  renderSeriesToggleList(scene);
  renderStage(scene);
  renderPlot(scene);
  renderMetricCards(scene);
  renderArtifacts();
  renderSelectedFiles();
  renderManualPairing();
  updateSelectionStats();
  refreshButtons();
}

function restoreCachedRunIfAvailable() {
  return loadRunSnapshot()
    .then((snapshot) => {
      if (!snapshot || !Array.isArray(snapshot.results) || !snapshot.metrics) return;
      state.results = snapshot.results.map((item, index) => classifyViewerResult(item, index));
      state.sourceCatalog = Array.isArray(snapshot.sourceCatalog) ? snapshot.sourceCatalog : [];
      state.errors = Array.isArray(snapshot.errors) ? snapshot.errors : [];
      state.metrics = snapshot.metrics;
      if (snapshot.viewerPrefs) applyViewerPrefs(snapshot.viewerPrefs, { persist: false });
      renderAll();
      appendLog("info", `Restored cached batch results${snapshot.savedAt ? ` from ${snapshot.savedAt}` : ""}.`);
      setStatus("Restored cached batch results");
    })
    .catch((error) => {
      appendLog("warning", `Cached result restore failed: ${error instanceof Error ? error.message : String(error)}`);
    });
}

function collectSelectedFiles() {
  const combined = [...Array.from(dom.folderInput?.files || []), ...Array.from(dom.fileInput?.files || [])];
  const unique = new Map();
  let ignored = 0;
  combined.forEach((file) => {
    if (!isSupportedInputFile(file)) {
      ignored += 1;
      return;
    }
    const sourcePath = sourcePathForFile(file);
    const key = `${sourcePath || logicalCandidateName(file)}::${file.size}::${file.lastModified}`;
    if (!unique.has(key)) unique.set(key, file);
  });
  state.selectedFiles = Array.from(unique.values());
  state.ignoredFilesCount = ignored;
}

function handleSelectionChange(sourceLabel) {
  collectSelectedFiles();
  const loaded = summarizeLoadedTypes();
  if (loaded.db > 0 && loaded.xlsx === 0 && dom.useDb3Directly && !dom.useDb3Directly.checked) {
    dom.useDb3Directly.checked = true;
    appendLog("info", "Only DB inputs detected. Switched to DB3 direct mode automatically.");
  } else if (loaded.xlsx > 0 && loaded.db === 0 && dom.useDb3Directly && dom.useDb3Directly.checked) {
    dom.useDb3Directly.checked = false;
    appendLog("info", "Only XLSX inputs detected. Switched to Excel/strain mode automatically.");
  }
  appendLog("info", `${sourceLabel} updated | supported total: ${state.selectedFiles.length}, ignored unsupported: ${state.ignoredFilesCount}`);
  syncFilterInputs();
  renderAll();
}

function handleModeToggle() {
  appendLog("info", `Input mode changed: ${dom.useDb3Directly?.checked ? "DB3 direct" : "Excel/strain"}`);
  renderAll();
}

function handleManualPairingToggle() {
  appendLog("info", `Manual pairing ${dom.manualPairingEnabled?.checked ? "enabled" : "disabled"}`);
  renderAll();
}

function handleManualSelectionChange() {
  state.manualStage.xName = String(dom.manualXSelect?.value || "").trim();
  state.manualStage.yName = String(dom.manualYSelect?.value || "").trim();
  renderAll();
}

function stageManualCandidate(role, logicalName) {
  const normalizedRole = String(role || "").toLowerCase();
  if (!logicalName || (normalizedRole !== "x" && normalizedRole !== "y")) return;
  if (normalizedRole === "x") state.manualStage.xName = logicalName;
  else state.manualStage.yName = logicalName;
  syncManualSelectValues();
  const { xName, yName } = getManualSelection();
  const useDb3Directly = !!dom.useDb3Directly?.checked;
  if (
    xName &&
    yName &&
    xName.toLowerCase() !== yName.toLowerCase() &&
    canManualAssign(xName, "X", useDb3Directly) &&
    canManualAssign(yName, "Y", useDb3Directly)
  ) {
    addManualPair(xName, yName, { source: "list" });
    return;
  }
  renderAll();
}

function clearManualStage() {
  state.manualStage.xName = "";
  state.manualStage.yName = "";
  if (dom.manualXSelect) dom.manualXSelect.value = "";
  if (dom.manualYSelect) dom.manualYSelect.value = "";
  renderAll();
}

function addManualPair(xOverride = null, yOverride = null, options = {}) {
  const xName = String(xOverride || state.manualStage?.xName || dom.manualXSelect?.value || "").trim();
  const yName = String(yOverride || state.manualStage?.yName || dom.manualYSelect?.value || "").trim();
  const source = options.source || "controls";
  const allowAxisFallback = !!options.allowAxisFallback;
  const useDb3Directly = !!dom.useDb3Directly?.checked;

  if (!xName || !yName) {
    appendLog("warning", "Select both X and Y candidates before adding a manual pair.");
    return false;
  }
  if (xName.toLowerCase() === yName.toLowerCase()) {
    appendLog("warning", "Manual pair rejected: same candidate cannot be used as both X and Y.");
    return false;
  }
  const strictAxisOk = canManualAssign(xName, "X", useDb3Directly) && canManualAssign(yName, "Y", useDb3Directly);
  const xAxis = inferAxisLabel(xName);
  const yAxis = inferAxisLabel(yName);
  const fallbackAxisOk =
    allowAxisFallback &&
    isInputCandidate(xName, useDb3Directly) &&
    isInputCandidate(yName, useDb3Directly) &&
    xAxis !== "Y" &&
    yAxis !== "X";
  if (!strictAxisOk && !fallbackAxisOk) {
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
  state.manualStage.xName = "";
  state.manualStage.yName = "";
  if (dom.manualXSelect) dom.manualXSelect.value = "";
  if (dom.manualYSelect) dom.manualYSelect.value = "";
  appendLog("info", `Manual pair added (${source}): ${xName} + ${yName}`);
  renderAll();
  return true;
}

function clearManualPairs() {
  if (!state.manualPairs.length) return;
  state.manualPairs = [];
  clearManualStage();
  appendLog("info", "All manual pairs cleared.");
  renderAll();
}

function clearSelectedFiles(options = {}) {
  const { clearLogs = false, resetViewerPrefs = false } = options;
  if (!state.selectedFiles.length && !state.results.length && !state.sourceCatalog.length && !(clearLogs && state.logs.length)) return;
  state.selectedFiles = [];
  state.ignoredFilesCount = 0;
  state.manualPairs = [];
  state.pairSuggestions = [];
  state.results = [];
  state.sourceCatalog = [];
  state.errors = [];
  state.metrics = null;
  state.activeSourceId = "";
  state.activeFamilyKey = "";
  state.activeChartKey = "";
  state.activeLayerIndex = 0;
  state.compareSourceIds = [];
  state.seriesVisibilityMap = {};
  state.spectrumPeriodMax = "";
  state.logX = false;
  state.logY = false;
  if (dom.folderInput) dom.folderInput.value = "";
  if (dom.fileInput) dom.fileInput.value = "";
  if (dom.manualPairingEnabled) dom.manualPairingEnabled.checked = false;
  clearManualStage();
  if (clearLogs) {
    state.logs = [];
    if (dom.logBox) dom.logBox.textContent = "";
  }
  if (resetViewerPrefs) {
    try {
      globalThis.localStorage?.removeItem(VIEWER_PREFS_STORAGE_KEY);
      globalThis.localStorage?.removeItem(MANUAL_PAIR_STORAGE_KEY);
    } catch {
      // Ignore local storage cleanup failures.
    }
  }
  setStatus(clearLogs ? "Workspace cleared" : state.workerReady ? "Ready" : "Initializing...");
  if (!clearLogs) appendLog("info", "Selected files, pairs and current outputs cleared.");
  state.snapshotPersistPromise = clearRunSnapshot().catch(() => undefined);
  renderAll();
}

function clearEverything() {
  clearSelectedFiles({ clearLogs: true, resetViewerPrefs: true });
}

function parseNumberInput(value, fallback) {
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : fallback;
}

function buildRunOptions() {
  const legacyMode = state.shellMode === "legacy";
  return {
    methods: "strain_legacy",
    failFast: legacyMode && !!dom.failFast?.checked,
    includeManip: false,
    method2Enabled: !!dom.method2Enabled?.checked,
    method3Enabled: !!dom.method3Enabled?.checked,
    manualPairingEnabled: legacyMode && !!dom.manualPairingEnabled?.checked,
    manualPairs: legacyMode && !!dom.manualPairingEnabled?.checked ? state.manualPairs.map((pair) => ({ ...pair })) : [],
    useDb3Directly: !!dom.useDb3Directly?.checked,
    integrationCompareEnabled: legacyMode && !!dom.integrationCompareEnabled?.checked,
    includeResultantProfiles: !!dom.includeResultantProfiles?.checked,
    altIntegrationMethod: "fft_regularized",
    baseReference: legacyMode ? dom.baseReference?.value || "input" : "input",
    processingOrder: legacyMode ? dom.processingOrder?.value || "filter_then_baseline" : "filter_then_baseline",
    filterDomain: legacyMode ? dom.filterDomain?.value || "time" : "time",
    baselineMethod: legacyMode ? dom.baselineMethod?.value || "poly4" : "poly4",
    filterOn: legacyMode && !!dom.filterOn?.checked,
    baselineOn: legacyMode && !!dom.baselineOn?.checked,
    filterConfig: legacyMode ? dom.filterConfig?.value || "bandpass" : "bandpass",
    filterType: legacyMode ? dom.filterType?.value || "butter" : "butter",
    fLowHz: legacyMode ? parseNumberInput(dom.fLowHz?.value, 0.1) : 0.1,
    fHighHz: legacyMode ? parseNumberInput(dom.fHighHz?.value, 25) : 25,
    filterOrder: legacyMode ? Math.max(1, Math.round(parseNumberInput(dom.filterOrder?.value, 4))) : 4,
    highpassEnabled: legacyMode && !!dom.filterOn?.checked && dom.filterConfig?.value === "highpass",
    highpassCutoffHz: legacyMode ? parseNumberInput(dom.fHighHz?.value, 25) : 25,
    highpassTransitionHz: 0.02,
  };
}

function buildWorkerFilesPayload(files) {
  const used = new Set();
  return files.map((file, index) => {
    const original = logicalCandidateName(file);
    const cleaned = String(original || `input_${index}.xlsx`).replace(/[\x00-\x1f<>:"|?*]/g, "_").trim() || `input_${index}.xlsx`;
    let finalName = cleaned;
    let suffix = 1;
    while (used.has(finalName.toLowerCase())) {
      const dot = cleaned.lastIndexOf(".");
      finalName = dot >= 0 ? `${cleaned.slice(0, dot)}_${suffix}${cleaned.slice(dot)}` : `${cleaned}_${suffix}`;
      suffix += 1;
    }
    used.add(finalName.toLowerCase());
    return { name: finalName, file };
  });
}

function buildPrimaryJobs(filesPayload, baseOptions) {
  const pairingPlan = resolveBatchPairingPlan(filesPayload, baseOptions);
  const used = new Set();
  const jobs = [];
  let order = 0;

  pairingPlan.pairs.forEach((pair) => {
    const [xItem, yItem] = Array.isArray(pair) ? pair : [];
    if (!xItem || !yItem) return;
    const xKey = String(xItem.name || "").toLowerCase();
    const yKey = String(yItem.name || "").toLowerCase();
    if (used.has(xKey) || used.has(yKey)) return;
    used.add(xKey);
    used.add(yKey);
    jobs.push({
      order: order++,
      dispatchOrder: order,
      kind: "primary",
      label: `Pair ${xItem.name}`,
      weight: 2,
      files: [xItem, yItem],
      options: {
        ...baseOptions,
        method2Enabled: false,
        method3Enabled: false,
        skipMethod23Outputs: true,
        manualPairingEnabled: true,
        manualPairs: [{ xName: xItem.name, yName: yItem.name }],
      },
    });
  });

  pairingPlan.leftovers.forEach((item) => {
    const key = String(item?.name || "").toLowerCase();
    if (!item || used.has(key)) return;
    used.add(key);
    jobs.push({
      order: order++,
      dispatchOrder: order,
      kind: "primary",
      label: `Single ${item.name}`,
      weight: 1,
      files: [item],
      options: {
        ...baseOptions,
        method2Enabled: false,
        method3Enabled: false,
        skipMethod23Outputs: true,
        manualPairingEnabled: false,
        manualPairs: [],
      },
    });
  });

  return { jobs, pairingPlan };
}

function buildBatchJobs(filesPayload, baseOptions) {
  if (baseOptions.useDb3Directly) {
    return [
      {
        order: 0,
        kind: "full",
        label: "DB batch",
        weight: filesPayload.length,
        files: filesPayload,
        options: baseOptions,
      },
    ];
  }

  const primaryPlan = buildPrimaryJobs(filesPayload, baseOptions);
  const jobs = [...primaryPlan.jobs];
  if (baseOptions.method2Enabled || baseOptions.method3Enabled) {
    const aggregateManualPairs = (primaryPlan.pairingPlan?.pairs || [])
      .map((pair) => {
        const [xItem, yItem] = Array.isArray(pair) ? pair : [];
        return xItem?.name && yItem?.name ? { xName: xItem.name, yName: yItem.name } : null;
      })
      .filter(Boolean);
    jobs.push({
      order: jobs.length,
      dispatchOrder: 0,
      kind: "method23",
      label: "Method-2/3 aggregate",
      weight: Math.max(1, filesPayload.length),
      files: filesPayload,
      options: {
        ...baseOptions,
        skipPrimaryOutputs: true,
        manualPairingEnabled: aggregateManualPairs.length > 0,
        manualPairs: aggregateManualPairs,
      },
    });
  }
  return jobs;
}

function computeDesiredWorkerCount(jobCount, options = {}) {
  if (String(options.phase || "primary") === "method23") return Math.max(1, Math.min(1, jobCount));
  const hardware = Math.max(1, Number(globalThis.navigator?.hardwareConcurrency) || 2);
  const totalBytes = Math.max(0, Number(options.totalBytes) || 0);
  const megabytes = totalBytes / (1024 * 1024);
  let dataCap = MAX_PARALLEL_WORKERS;
  if (megabytes >= 750) dataCap = 2;
  else if (megabytes >= 350) dataCap = 3;
  const hardwareCap = Math.max(2, Math.floor(hardware / 8));
  return Math.max(1, Math.min(MAX_PARALLEL_WORKERS, hardwareCap, dataCap, jobCount));
}

function appendSummaryWarnings(jobLabel, logs) {
  (logs || []).forEach((log) => {
    const level = String(log?.level || "info").toLowerCase();
    if (level === "warning") appendLog(level, `[${jobLabel}] ${log?.message || ""}`);
  });
}

async function executeJobsInPool(jobs) {
  const totalBytes = jobs.reduce((sum, job) => sum + job.files.reduce((fileSum, item) => fileSum + (item.file?.size || 0), 0), 0);
  const workerCount = computeDesiredWorkerCount(jobs.length, { totalBytes });
  const slots = await ensureWorkerPool(workerCount);
  const orderedJobs = [...jobs].sort((a, b) => (a.dispatchOrder ?? a.order) - (b.dispatchOrder ?? b.order));
  const outputs = [];
  let nextIndex = 0;
  let firstError = null;

  async function workLoop(slot) {
    while (true) {
      if (firstError) return;
      const currentIndex = nextIndex;
      nextIndex += 1;
      if (currentIndex >= orderedJobs.length) return;
      const job = orderedJobs[currentIndex];
      try {
        appendLog("info", `[${workerTag(slot)}] Starting ${job.label}`);
        const summary = await runBatchOnSlot(slot, job);
        outputs.push({ job, summary });
      } catch (error) {
        firstError = error;
        return;
      }
    }
  }

  await Promise.all(slots.map((slot) => workLoop(slot)));
  if (firstError) throw firstError;
  return outputs.sort((a, b) => a.job.order - b.job.order);
}

function mergeJobSummaries(jobOutputs, baseOptions, detectedStats) {
  const results = [];
  const sourceCatalog = [];
  const errors = [];
  let method23Metrics = null;
  let pairsProcessed = 0;
  let pairsFailed = 0;
  let singlesProcessed = 0;
  let singlesFailed = 0;

  jobOutputs.forEach(({ job, summary }) => {
    (summary?.results || []).forEach((item) => results.push(item));
    (summary?.sourceCatalog || []).forEach((item) => sourceCatalog.push(item));
    (summary?.errors || []).forEach((item) => errors.push(item));
    appendSummaryWarnings(job.label, summary?.logs || []);
    const metrics = summary?.metrics || {};
    if (job.kind === "method23") {
      method23Metrics = metrics;
    } else {
      pairsProcessed += metrics.pairsProcessed || 0;
      pairsFailed += metrics.pairsFailed || 0;
      singlesProcessed += metrics.singlesProcessed || 0;
      singlesFailed += metrics.singlesFailed || 0;
    }
  });

  const method2Processed = method23Metrics?.method2Processed || 0;
  const method2Failed = method23Metrics?.method2Failed || 0;
  const method3Produced = method23Metrics?.method3Produced || 0;
  const method23Failed = method23Metrics?.failedTotal || 0;

  return {
    results,
    sourceCatalog,
    errors,
    metrics: {
      pairsDetected: detectedStats.pairs,
      pairsProcessed,
      pairsFailed,
      pairsMissing: detectedStats.missing,
      xlsxCandidates: detectedStats.xlsxCandidates,
      dbCandidates: detectedStats.dbCandidates,
      singlesDetected: detectedStats.singles,
      singlesProcessed,
      singlesFailed,
      method2Enabled: !!baseOptions.method2Enabled,
      method3Enabled: !!baseOptions.method3Enabled,
      includeResultantProfiles: !!baseOptions.includeResultantProfiles,
      baseReference: baseOptions.baseReference,
      integrationPrimary: method23Metrics?.integrationPrimary || "cumtrapz",
      integrationCompareEnabled: !!baseOptions.integrationCompareEnabled,
      altIntegrationMethod: baseOptions.integrationCompareEnabled ? baseOptions.altIntegrationMethod : null,
      altLowCutPolicy: method23Metrics?.altLowCutPolicy ?? null,
      manualPairingEnabled: !!baseOptions.manualPairingEnabled,
      manualPairsApplied: detectedStats.manualPairsApplied ?? 0,
      method2Detected: baseOptions.method2Enabled || baseOptions.method3Enabled ? detectedStats.xlsxCandidates : 0,
      method2Processed,
      method2Failed,
      method3Produced,
      processedTotal: pairsProcessed + singlesProcessed + method2Processed + method3Produced,
      failedTotal: pairsFailed + singlesFailed + method23Failed,
      useDb3Directly: !!baseOptions.useDb3Directly,
    },
  };
}

async function runBatch() {
  if (!state.workerReady || !state.selectedFiles.length || state.isRunning) return;
  state.isRunning = true;
  state.results = [];
  state.errors = [];
  state.metrics = null;
  setProgressState({ visible: true, label: "Preparing files...", percent: 4, indeterminate: true });
  setStatus("Preparing files...");
  appendLog("info", `Preparing ${state.selectedFiles.length} files`);
  renderAll();

  const selectedCandidates = state.selectedFiles.filter((file) => isInputCandidate(logicalCandidateName(file), !!dom.useDb3Directly?.checked));
  const filesPayload = buildWorkerFilesPayload(selectedCandidates);
  if (!filesPayload.length) {
    appendLog("warning", "No valid candidate input files for the selected mode.");
    setStatus("No candidate input files");
    state.isRunning = false;
    setProgressState({ visible: false, label: "", percent: 0, indeterminate: false });
    refreshButtons();
    return;
  }

  const baseOptions = buildRunOptions();
  const detectedStats = detectPairs(filesPayload.map((item) => item.name), baseOptions.useDb3Directly, baseOptions.manualPairs, baseOptions.manualPairingEnabled);
  const jobs = buildBatchJobs(filesPayload, baseOptions);
  const parallelCount = computeDesiredWorkerCount(jobs.length, {
    totalBytes: filesPayload.reduce((sum, item) => sum + (item.file?.size || 0), 0),
  });
  appendLog("info", `Dispatching ${jobs.length} job(s) across ${parallelCount} worker(s).`);
  setStatus(jobs.length > 1 ? `Running in parallel (${parallelCount} workers)...` : "Running...");

  const jobOutputs = await executeJobsInPool(jobs);
  const merged = baseOptions.useDb3Directly
    ? (() => {
        const [{ summary }] = jobOutputs;
        appendSummaryWarnings("DB batch", summary?.logs || []);
        return {
          results: summary?.results || [],
          sourceCatalog: summary?.sourceCatalog || [],
          errors: summary?.errors || [],
          metrics: summary?.metrics || null,
        };
      })()
    : mergeJobSummaries(jobOutputs, baseOptions, detectedStats);

  state.isRunning = false;
  state.results = (merged.results || []).map((item, index) => classifyViewerResult(item, index));
  state.sourceCatalog = Array.isArray(merged.sourceCatalog) ? merged.sourceCatalog : [];
  state.errors = merged.errors || [];
  state.metrics = merged.metrics || null;
  state.snapshotPersistPromise = persistRunSnapshot()
    .then(() => appendLog("info", "Batch results cached locally for crash/refresh recovery."))
    .catch((error) => appendLog("warning", `Local result cache failed: ${error instanceof Error ? error.message : String(error)}`));
  state.errors.forEach((err) => appendLog("error", `${err.pairKey}: ${err.reason}`));
  setProgressState({ visible: false, label: "", percent: 100, indeterminate: false });
  setStatus(`Done | Processed: ${state.metrics?.processedTotal ?? state.results.length}, Failed: ${state.metrics?.failedTotal ?? state.errors.length}`);
  renderAll();
}

async function requestZipDownload() {
  if (!state.results.length) return;
  if (state.snapshotPersistPromise) await state.snapshotPersistPromise.catch(() => undefined);
  setStatus("ZIP olusturuluyor...");
  setProgressState({ visible: true, label: "ZIP olusturuluyor...", percent: 88, indeterminate: true });
  const { buildZipBlobFromResults } = await import(`./zip.mjs?v=${APP_VERSION}`);
  const zipBlob = await buildZipBlobFromResults(state.results, {
    yieldEvery: 1,
    onProgress: ({ completed, total, percent, fileName }) => {
      const safeTotal = Math.max(0, Number(total) || 0);
      const safeCompleted = Math.max(0, Number(completed) || 0);
      const label = safeTotal > 0
        ? `ZIP olusturuluyor (${safeCompleted}/${safeTotal}): ${clipMiddle(fileName || "outputs", 72, 28)}`
        : "ZIP olusturuluyor...";
      setStatus(label);
      setProgressState({
        visible: true,
        label,
        percent: Number.isFinite(Number(percent)) ? Number(percent) : 88,
        indeterminate: false,
      });
    },
  });
  downloadBlob(zipBlob, "deepsoil_total_displacement_outputs.zip");
  setStatus("ZIP downloaded");
  setProgressState({ visible: false, label: "", percent: 100, indeterminate: false });
  appendLog("info", "ZIP created and downloaded");
  refreshButtons();
}

function bindStageNavigation() {
  dom.sourcePrevBtn?.addEventListener("click", () => {
    const scene = syncSourceViewerSelection();
    const item = cycleSceneItem(scene.sources, scene.activeSourceId, -1, (source) => source.sourceId);
    if (item) setActiveSource(item.sourceId);
  });
  dom.sourceNextBtn?.addEventListener("click", () => {
    const scene = syncSourceViewerSelection();
    const item = cycleSceneItem(scene.sources, scene.activeSourceId, 1, (source) => source.sourceId);
    if (item) setActiveSource(item.sourceId);
  });
  dom.chartPrevBtn?.addEventListener("click", () => {
    const scene = syncSourceViewerSelection();
    const charts = scene.activeFamily?.charts || [];
    const item = cycleSceneItem(charts, scene.activeChartKey, -1, (chart) => chart.chartKey);
    if (item) setActiveChart(item.chartKey);
  });
  dom.chartNextBtn?.addEventListener("click", () => {
    const scene = syncSourceViewerSelection();
    const charts = scene.activeFamily?.charts || [];
    const item = cycleSceneItem(charts, scene.activeChartKey, 1, (chart) => chart.chartKey);
    if (item) setActiveChart(item.chartKey);
  });
  dom.layerPrevBtn?.addEventListener("click", () => {
    const scene = syncSourceViewerSelection();
    if (scene.activeLayerCount <= 1) return;
    setActiveLayer((scene.activeLayerIndex - 1 + scene.activeLayerCount) % scene.activeLayerCount);
  });
  dom.layerNextBtn?.addEventListener("click", () => {
    const scene = syncSourceViewerSelection();
    if (scene.activeLayerCount <= 1) return;
    setActiveLayer((scene.activeLayerIndex + 1) % scene.activeLayerCount);
  });
}

function attachEventListeners() {
  dom.folderInput?.addEventListener("change", () => handleSelectionChange("Folder Select"));
  dom.fileInput?.addEventListener("change", () => handleSelectionChange("File Select"));
  dom.filterConfig?.addEventListener("change", syncFilterInputs);
  dom.filterOn?.addEventListener("change", syncFilterInputs);
  dom.baselineOn?.addEventListener("change", syncFilterInputs);
  dom.useDb3Directly?.addEventListener("change", handleModeToggle);
  dom.manualPairingEnabled?.addEventListener("change", handleManualPairingToggle);
  dom.addManualPairBtn?.addEventListener("click", () => addManualPair());
  dom.autoPairBtn?.addEventListener("click", autoPair);
  dom.applyAllSuggestionsBtn?.addEventListener("click", applyAllPairSuggestions);
  dom.clearFilesBtn?.addEventListener("click", clearSelectedFiles);
  dom.clearAllBtn?.addEventListener("click", clearEverything);
  dom.clearPairsBtn?.addEventListener("click", clearManualPairs);
  dom.manualXSelect?.addEventListener("change", handleManualSelectionChange);
  dom.manualYSelect?.addEventListener("change", handleManualSelectionChange);
  dom.clearManualStageBtn?.addEventListener("click", clearManualStage);
  dom.viewerModeBtn?.addEventListener("click", () => setShellMode("viewer"));
  dom.legacyModeBtn?.addEventListener("click", () => setShellMode("legacy"));
  dom.viewerFocusBtn?.addEventListener("click", () => setViewerChartMode("focus"));
  dom.viewerCompareBtn?.addEventListener("click", () => setViewerChartMode("compare"));
  dom.compareAxisXToggle?.addEventListener("change", () => toggleCompareAxisBulk("X", !!dom.compareAxisXToggle?.checked));
  dom.compareAxisYToggle?.addEventListener("change", () => toggleCompareAxisBulk("Y", !!dom.compareAxisYToggle?.checked));
  dom.compareClearBtn?.addEventListener("click", clearCompareSources);
  dom.periodMaxInput?.addEventListener("change", () => setSpectrumPeriodMax(dom.periodMaxInput?.value || ""));
  dom.periodMaxResetBtn?.addEventListener("click", () => setSpectrumPeriodMax(""));
  dom.logXToggle?.addEventListener("change", () => setLogAxis("x", !!dom.logXToggle?.checked));
  dom.logYToggle?.addEventListener("change", () => setLogAxis("y", !!dom.logYToggle?.checked));
  dom.runBtn?.addEventListener("click", () => {
    runBatch().catch((error) => {
      state.isRunning = false;
      const message = error instanceof Error ? error.message : String(error);
      setStatus(`Error: ${message}`);
      setProgressState({ visible: false, label: "", percent: 0, indeterminate: false });
      appendLog("error", message);
      refreshButtons();
    });
  });
  dom.zipBtn?.addEventListener("click", () => {
    requestZipDownload().catch((error) => {
      const message = error instanceof Error ? error.message : String(error);
      setStatus(`Error: ${message}`);
      setProgressState({ visible: false, label: "", percent: 0, indeterminate: false });
      appendLog("error", message);
      refreshButtons();
    });
  });
  globalThis.addEventListener("resize", () => {
    try {
      globalThis.Plotly?.Plots?.resize(dom.plotHost);
    } catch {
      // Ignore resize failures when no plot exists yet.
    }
  });
  globalThis.addEventListener("load", () => {
    renderAll();
    try {
      globalThis.Plotly?.Plots?.resize(dom.plotHost);
    } catch {
      // Ignore resize failures when no plot exists yet.
    }
  });
  bindStageNavigation();
}

function bootstrap() {
  setStatus("Initializing worker...");
  appendLog("info", "Initializing Pyodide worker...");

  const restoredManualPrefs = loadManualPairPrefs();
  state.manualPairs = restoredManualPrefs.manualPairs;
  if (dom.manualPairingEnabled) dom.manualPairingEnabled.checked = restoredManualPrefs.manualPairingEnabled;
  if (state.manualPairs.length > 0) appendLog("info", `Restored ${state.manualPairs.length} saved manual pair(s).`);

  const restoredViewerPrefs = loadViewerPrefs();
  if (restoredViewerPrefs) applyViewerPrefs(restoredViewerPrefs, { persist: false });

  setShellMode(loadShellModePreference(), { persist: false });
  syncFilterInputs();
  renderProgress();
  renderAll();
  restoreCachedRunIfAvailable().catch(() => undefined);
  ensureWorkerReady(primaryWorkerSlot).catch((error) => {
    const message = error instanceof Error ? error.message : String(error);
    setStatus(`Error: ${message}`);
    setProgressState({ visible: false, label: "", percent: 0, indeterminate: false });
    appendLog("error", message);
    refreshButtons();
  });
}

attachEventListeners();
bootstrap();
