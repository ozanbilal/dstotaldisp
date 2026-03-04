const PYODIDE_BASE = "https://cdn.jsdelivr.net/pyodide/v0.27.3/full/";

importScripts(`${PYODIDE_BASE}pyodide.js`);

let pyodide = null;
let runBatchFromFs = null;
let buildZipFromResults = null;
let initPromise = null;

function postStatus(message, phase = "info") {
  self.postMessage({ type: "status", payload: { phase, message } });
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

async function ensureInitialized() {
  if (initPromise) {
    return initPromise;
  }

  initPromise = (async () => {
    postStatus("Pyodide yukleniyor...", "boot");
    pyodide = await loadPyodide({ indexURL: PYODIDE_BASE });

    postStatus("Numpy/Pandas paketleri yukleniyor...", "boot");
    try {
      await pyodide.loadPackage(["numpy", "pandas", "openpyxl"]);
    } catch (err) {
      postStatus("openpyxl paketini micropip ile yukluyorum...", "boot");
      await pyodide.loadPackage(["numpy", "pandas", "micropip"]);
      await pyodide.runPythonAsync(`
import micropip
await micropip.install("openpyxl")
`);
    }

    postStatus("Python modulleri yukleniyor...", "boot");

    const [coreResp, entryResp] = await Promise.all([
      fetch("../disp_core.py"),
      fetch("./py/pyodide_entry.py"),
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
from pyodide_entry import run_batch_from_fs, build_zip_from_results
`);

    runBatchFromFs = pyodide.globals.get("run_batch_from_fs");
    buildZipFromResults = pyodide.globals.get("build_zip_from_results");

    postStatus("Hazir", "ready");
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
        metrics: { pairsDetected: 0, pairsProcessed: 0, pairsFailed: 0, pairsMissing: 0 },
      },
    });
    return;
  }

  postStatus(`Dosyalar worker FS'e yaziliyor (${files.length})...`, "run");
  resetInputFs();

  const usedNames = new Set();
  files.forEach((file, idx) => {
    const normalized = sanitizeFileName(file.name, idx);
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

    const bytes = new Uint8Array(file.bytes);
    pyodide.FS.writeFile(`/input/${finalName}`, bytes);
  });

  postStatus("Hesaplama calisiyor...", "run");

  const pyOptions = pyodide.toPy(options);
  const pyResult = runBatchFromFs("/input", pyOptions);
  const jsResult = pyResult.toJs({ dict_converter: Object.fromEntries });

  if (pyOptions && typeof pyOptions.destroy === "function") {
    pyOptions.destroy();
  }
  if (pyResult && typeof pyResult.destroy === "function") {
    pyResult.destroy();
  }

  self.postMessage({ type: "runBatchResult", payload: jsResult });
  postStatus("Batch tamamlandi", "ready");
}

async function handleZipOutputs(payload) {
  await ensureInitialized();

  const items = Array.isArray(payload?.results) ? payload.results : [];
  postStatus("ZIP olusturuluyor...", "zip");

  const pyItems = pyodide.toPy(items);
  const zipB64 = buildZipFromResults(pyItems);
  const zipString = typeof zipB64 === "string" ? zipB64 : zipB64.toString();

  if (pyItems && typeof pyItems.destroy === "function") {
    pyItems.destroy();
  }
  if (zipB64 && typeof zipB64.destroy === "function") {
    zipB64.destroy();
  }

  self.postMessage({ type: "zipReady", payload: { fileName: payload?.fileName || "deepsoil_outputs.zip", zipBytesB64: zipString } });
  postStatus("ZIP hazir", "ready");
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

    if (type === "zipOutputs") {
      await handleZipOutputs(payload);
      return;
    }

    self.postMessage({ type: "error", payload: { message: `Unknown worker message type: ${type}` } });
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    self.postMessage({ type: "error", payload: { message } });
    postStatus(`Hata: ${message}`, "error");
  }
};
