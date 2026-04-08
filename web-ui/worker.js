const PYODIDE_BASE = "https://cdn.jsdelivr.net/pyodide/v0.27.3/full/";
const APP_VERSION = "20260408b";

importScripts(`${PYODIDE_BASE}pyodide.js`);

let pyodide = null;
let runBatchFromFs = null;
let initPromise = null;

function postStatus(message, phase = "info", progress = null, indeterminate = false) {
  self.postMessage({ type: "status", payload: { phase, message, progress, indeterminate } });
}

function sanitizeFileName(name, index) {
  const base = (name || `input_${index}.xlsx`).split(/[\\/]/).pop();
  const cleaned = base.replace(/[\x00-\x1f<>:"|?*]/g, "_").trim();
  return cleaned || `input_${index}.xlsx`;
}

function resetInputFs() {
  if (!pyodide.FS.analyzePath("/input").exists) {
    pyodide.FS.mkdir("/input");
    return;
  }

  const entries = pyodide.FS.readdir("/input").filter((name) => name !== "." && name !== "..");
  for (const entry of entries) {
    pyodide.FS.unlink(`/input/${entry}`);
  }
}

async function toUint8Array(fileItem) {
  if (fileItem && fileItem.bytes != null) {
    return new Uint8Array(fileItem.bytes);
  }

  const blob = fileItem?.file ?? fileItem?.blob ?? null;
  if (blob && typeof blob.arrayBuffer === "function") {
    const buffer = await blob.arrayBuffer();
    return new Uint8Array(buffer);
  }

  throw new Error("Invalid file payload: expected bytes or File/Blob.");
}

async function ensureInitialized() {
  if (initPromise) {
    return initPromise;
  }

  initPromise = (async () => {
    postStatus("Pyodide yukleniyor...", "boot", 8);
    pyodide = await loadPyodide({ indexURL: PYODIDE_BASE });

    postStatus("Numpy/Pandas/SQLite paketleri yukleniyor...", "boot", 24);
    try {
      await pyodide.loadPackage(["numpy", "pandas", "sqlite3"]);
    } catch (err) {
      postStatus("Temel paketler yuklenemedi, tekrar deniyorum...", "boot", 28, true);
      await pyodide.loadPackage(["numpy", "pandas", "sqlite3", "micropip"]);
    }

    try {
      await pyodide.loadPackage(["openpyxl"]);
    } catch (err) {
      postStatus("openpyxl paketini micropip ile yukluyorum...", "boot", 42, true);
      await pyodide.loadPackage(["micropip"]);
      await pyodide.runPythonAsync(`
import micropip
await micropip.install("openpyxl")
`);
    }

    postStatus("Python modulleri yukleniyor...", "boot", 62);

    const [coreResp, entryResp] = await Promise.all([
      fetch(`../disp_core.py?v=${APP_VERSION}`),
      fetch(`./py/pyodide_entry.py?v=${APP_VERSION}`),
    ]);

    if (!coreResp.ok) {
      throw new Error(`disp_core.py fetch failed (${coreResp.status})`);
    }
    if (!entryResp.ok) {
      throw new Error(`pyodide_entry.py fetch failed (${entryResp.status})`);
    }

    const coreCode = await coreResp.text();
    const entryCode = await entryResp.text();

    if (!pyodide.FS.analyzePath("/app").exists) {
      pyodide.FS.mkdir("/app");
    }

    pyodide.FS.writeFile("/app/disp_core.py", coreCode, { encoding: "utf8" });
    pyodide.FS.writeFile("/app/pyodide_entry.py", entryCode, { encoding: "utf8" });

    await pyodide.runPythonAsync(`
import sys
if "/app" not in sys.path:
    sys.path.insert(0, "/app")
from pyodide_entry import run_batch_from_fs
`);

    runBatchFromFs = pyodide.globals.get("run_batch_from_fs");

    postStatus("Hazir", "ready", 100);
  })();

  return initPromise;
}

async function handleRunBatch(payload) {
  await ensureInitialized();

  const files = Array.isArray(payload?.files) ? payload.files : [];
  const options = payload?.options || {};

  if (!files.length) {
    self.postMessage({
      type: "runBatchResult",
      payload: {
        results: [],
        logs: [{ level: "warning", message: "No files received." }],
        errors: [],
        metrics: {
          pairsDetected: 0,
          pairsProcessed: 0,
          pairsFailed: 0,
          pairsMissing: 0,
          xlsxCandidates: 0,
          dbCandidates: 0,
          singlesDetected: 0,
          singlesProcessed: 0,
          singlesFailed: 0,
          integrationPrimary: "cumtrapz",
          method2Enabled: false,
          method3Enabled: false,
          includeResultantProfiles: true,
          integrationCompareEnabled: false,
          altIntegrationMethod: null,
          altLowCutPolicy: null,
          baseReference: "input",
          useDb3Directly: false,
          manualPairingEnabled: false,
          manualPairsApplied: 0,
          method2Detected: 0,
          method2Processed: 0,
          method2Failed: 0,
          method3Produced: 0,
          processedTotal: 0,
          failedTotal: 0,
        },
      },
    });
    return;
  }

  postStatus(`Dosyalar worker FS'e yaziliyor (0/${files.length})...`, "run", 10);
  resetInputFs();

  const usedNames = new Set();
  for (let idx = 0; idx < files.length; idx += 1) {
    const file = files[idx];
    const normalized = sanitizeFileName(file?.name, idx);
    let finalName = normalized;
    let suffix = 1;
    while (usedNames.has(finalName.toLowerCase())) {
      const dot = normalized.lastIndexOf(".");
      if (dot >= 0) {
        finalName = `${normalized.slice(0, dot)}_${suffix}${normalized.slice(dot)}`;
      } else {
        finalName = `${normalized}_${suffix}`;
      }
      suffix += 1;
    }
    usedNames.add(finalName.toLowerCase());

    const bytes = await toUint8Array(file);
    pyodide.FS.writeFile(`/input/${finalName}`, bytes);
    if ((idx + 1) % 10 === 0 || idx + 1 === files.length) {
      const progress = 10 + Math.round(((idx + 1) / files.length) * 35);
      postStatus(`Dosyalar worker FS'e yaziliyor (${idx + 1}/${files.length})...`, "run", progress);
    }
  }

  postStatus("Hesaplama calisiyor...", "run", 55, true);

  const pyOptions = pyodide.toPy(options);
  const progressCallback = (message, phase = "run", progress = null, indeterminate = false) => {
    const numericProgress = Number(progress);
    postStatus(
      String(message || "Hesaplama calisiyor..."),
      String(phase || "run"),
      Number.isFinite(numericProgress) ? numericProgress : null,
      !!indeterminate
    );
  };

  let pyResult = null;
  let jsResult = null;
  try {
    pyResult = runBatchFromFs("/input", pyOptions, progressCallback);
    jsResult = pyResult.toJs({ dict_converter: Object.fromEntries });
    postStatus("Sonuclar hazirlaniyor...", "run", 92);
  } finally {
    try {
      if (pyOptions && typeof pyOptions.destroy === "function") {
        pyOptions.destroy();
      }
    } catch (err) {
      // Best-effort cleanup only. Pyodide can already be in a bad state here.
    }
    try {
      if (pyResult && typeof pyResult.destroy === "function") {
        pyResult.destroy();
      }
    } catch (err) {
      // Best-effort cleanup only. Pyodide can already be in a bad state here.
    }
  }

  try {
    pyodide.runPython("import gc; gc.collect()");
  } catch (err) {
    // Optional cleanup best effort.
  }

  self.postMessage({ type: "runBatchResult", payload: jsResult });
  postStatus("Batch tamamlandi", "ready", 100);
}

self.onmessage = async (event) => {
  const { type, payload } = event.data || {};

  try {
    if (type === "initialize") {
      await ensureInitialized();
      self.postMessage({ type: "initialized" });
      return;
    }

    if (type === "runBatch") {
      await handleRunBatch(payload);
      return;
    }

    self.postMessage({ type: "error", payload: { message: `Unknown worker message type: ${type}` } });
  } catch (error) {
    const rawMessage = error instanceof Error ? error.message : String(error);
    const message = rawMessage.toLowerCase().includes("memory access out of bounds")
      ? "memory access out of bounds (WASM bellek limiti asildi). Daha az dosya secip partlar halinde calistirin."
      : rawMessage;
    self.postMessage({ type: "error", payload: { message } });
    postStatus(`Hata: ${message}`, "error", 0);
  }
};
