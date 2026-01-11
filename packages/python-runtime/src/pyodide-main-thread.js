/**
 * Main-thread Pyodide helpers.
 *
 * When SharedArrayBuffer/crossOriginIsolated is unavailable we can't use the
 * worker-based synchronous RPC bridge, so we fall back to running Pyodide on the
 * main thread and calling the spreadsheet bridge synchronously.
 */
 
const DEFAULT_INDEX_URL = "https://cdn.jsdelivr.net/pyodide/v0.25.1/full/";
 
/**
 * @param {string | undefined | null} indexURL
 */
export function resolveIndexURL(indexURL) {
  let resolved = typeof indexURL === "string" && indexURL.length > 0 ? indexURL : DEFAULT_INDEX_URL;
  if (!resolved.endsWith("/")) resolved += "/";
  return resolved;
}
 
const scriptLoadPromises = new Map();
 
function ensureDocumentAvailable() {
  if (typeof document === "undefined") {
    throw new Error(
      "PyodideRuntime mainThread mode requires a DOM (document) or a preloaded globalThis.loadPyodide",
    );
  }
}
 
/**
 * Ensure `pyodide.js` has been loaded and `globalThis.loadPyodide` is present.
 *
 * @param {string} resolvedIndexURL
 */
export async function ensurePyodideScript(resolvedIndexURL) {
  if (typeof globalThis.loadPyodide === "function") return;
 
  ensureDocumentAvailable();
 
  const src = `${resolvedIndexURL}pyodide.js`;
  const existing = document.querySelector(`script[src="${src}"]`);
 
  let promise = scriptLoadPromises.get(src);
  if (promise) return await promise;
 
  promise = new Promise((resolve, reject) => {
    const script = existing ?? document.createElement("script");
    if (!existing) {
      script.src = src;
      script.async = true;
      script.dataset.formulaPyodide = "true";
      (document.head ?? document.documentElement ?? document.body).appendChild(script);
    }
 
    const cleanup = () => {
      script.removeEventListener("load", onLoad);
      script.removeEventListener("error", onError);
    };
 
    const onLoad = () => {
      cleanup();
      if (typeof globalThis.loadPyodide !== "function") {
        reject(new Error(`Pyodide script loaded from ${src} but globalThis.loadPyodide is missing`));
        return;
      }
      resolve();
    };
 
    const onError = () => {
      cleanup();
      reject(new Error(`Failed to load Pyodide script from ${src}`));
    };
 
    script.addEventListener("load", onLoad);
    script.addEventListener("error", onError);
 
    // If the script tag was already present (and possibly already loaded),
    // resolve immediately once `loadPyodide` is available.
    if (existing && typeof globalThis.loadPyodide === "function") {
      cleanup();
      resolve();
    }
  });
 
  scriptLoadPromises.set(src, promise);
  promise.catch(() => scriptLoadPromises.delete(src));
  return await promise;
}
 
/**
 * Load Pyodide in the main thread.
 *
 * @param {{ indexURL?: string }} [options]
 */
export async function loadPyodideMainThread(options = {}) {
  const resolvedIndexURL = resolveIndexURL(options.indexURL);

  if (typeof globalThis.loadPyodide !== "function") {
    await ensurePyodideScript(resolvedIndexURL);
  }

  const loader = globalThis.loadPyodide;
  if (typeof loader !== "function") {
    throw new Error("PyodideRuntime mainThread mode could not find globalThis.loadPyodide after loading pyodide.js");
  }

  return await loader({ indexURL: resolvedIndexURL });
}
 
// --- Runtime helpers ---------------------------------------------------------
 
export function getWasmMemoryBytes(runtime) {
  const mod = runtime?._module;
  const buf = mod?.wasmMemory?.buffer ?? mod?.HEAP8?.buffer;
  return buf?.byteLength ?? null;
}
 
export function withCapturedOutput(runtime, fn) {
  let stdout = "";
  let stderr = "";
 
  const canCapture = typeof runtime?.setStdout === "function" && typeof runtime?.setStderr === "function";
  if (canCapture) {
    try {
      runtime.setStdout({ batched: (text) => (stdout += text) });
    } catch {
      try {
        runtime.setStdout({ write: (text) => (stdout += text) });
      } catch {
        // ignore
      }
    }
 
    try {
      runtime.setStderr({ batched: (text) => (stderr += text) });
    } catch {
      try {
        runtime.setStderr({ write: (text) => (stderr += text) });
      } catch {
        // ignore
      }
    }
  }
 
  const restore = () => {
    if (!canCapture) return;
    try {
      runtime.setStdout({ batched: (text) => console.log(text) });
    } catch {
      try {
        runtime.setStdout({ write: (text) => console.log(text) });
      } catch {
        // ignore
      }
    }
 
    try {
      runtime.setStderr({ batched: (text) => console.error(text) });
    } catch {
      try {
        runtime.setStderr({ write: (text) => console.error(text) });
      } catch {
        // ignore
      }
    }
  };
 
  return Promise.resolve()
    .then(fn)
    .then(
      (value) => ({ value, stdout, stderr }),
      (err) => {
        if (err && (typeof err === "object" || typeof err === "function")) {
          err.stdout = stdout;
          err.stderr = stderr;
        }
        throw err;
      },
    )
    .finally(restore);
}
 
export function applyNetworkSandbox(permissions) {
  const mode = permissions?.network ?? "none";
  const previousFetch = globalThis.fetch;
  const previousWebSocket = globalThis.WebSocket;
 
  if (mode === "none") {
    globalThis.fetch = async () => {
      throw new Error("Network access is not permitted");
    };
 
    globalThis.WebSocket = class BlockedWebSocket {
      constructor() {
        throw new Error("Network access is not permitted");
      }
    };
 
    return () => {
      globalThis.fetch = previousFetch;
      globalThis.WebSocket = previousWebSocket;
    };
  }
 
  if (mode === "allowlist") {
    const allowlist = new Set(permissions?.networkAllowlist ?? []);
    globalThis.fetch = async (input, init) => {
      const url = typeof input === "string" ? input : input?.url;
      const hostname = new URL(url, globalThis.location?.href ?? "https://localhost").hostname;
      if (!allowlist.has(hostname)) {
        throw new Error(`Network access to ${hostname} is not permitted`);
      }
      return previousFetch(input, init);
    };
 
    globalThis.WebSocket = class AllowlistWebSocket {
      constructor(url, protocols) {
        const hostname = new URL(url, globalThis.location?.href ?? "https://localhost").hostname;
        if (!allowlist.has(hostname)) {
          throw new Error(`Network access to ${hostname} is not permitted`);
        }
        return new previousWebSocket(url, protocols);
      }
    };
 
    return () => {
      globalThis.fetch = previousFetch;
      globalThis.WebSocket = previousWebSocket;
    };
  }
 
  // full access
  globalThis.fetch = previousFetch;
  globalThis.WebSocket = previousWebSocket;
  return () => {
    globalThis.fetch = previousFetch;
    globalThis.WebSocket = previousWebSocket;
  };
}
 
export async function applyPythonSandbox(runtime, permissions) {
  await runtime.runPythonAsync(`
from formula.runtime.sandbox import apply_sandbox
apply_sandbox(${JSON.stringify(permissions ?? {})})
`);
}
 
export async function runWithTimeout(runtime, code, timeoutMs, interruptView) {
  if (!interruptView || !Number.isFinite(timeoutMs) || timeoutMs <= 0) {
    return await runtime.runPythonAsync(code);
  }
 
  interruptView[0] = 0;
  const timer = setTimeout(() => {
    interruptView[0] = 2;
  }, timeoutMs);
 
  try {
    return await runtime.runPythonAsync(code);
  } finally {
    clearTimeout(timer);
    interruptView[0] = 0;
  }
}
 
const BRIDGE_API_KEY = "__formulaPyodideBridgeApi";
const BRIDGE_REGISTERED_KEY = "__formulaPyodideBridgeRegistered";

export function setFormulaBridgeApi(runtime, api) {
  runtime[BRIDGE_API_KEY] = api;
}

function coercePyProxy(value) {
  if (!value || (typeof value !== "object" && typeof value !== "function")) return value;
  if (typeof value.toJs !== "function") return value;
 
  let converted;
  try {
    converted = value.toJs({ dict_converter: Object.fromEntries });
  } catch {
    converted = value.toJs();
  }
 
  if (typeof value.destroy === "function") {
    try {
      value.destroy();
    } catch {
      // ignore
    }
  }
 
  return converted;
}
 
function normalizeRpcParams(params) {
  const maybeConverted = coercePyProxy(params);
  if (maybeConverted === null || typeof maybeConverted !== "object") {
    return maybeConverted;
  }
 
  if (maybeConverted instanceof Map) {
    const out = {};
    for (const [key, value] of maybeConverted.entries()) {
      out[String(key)] = normalizeRpcParams(value);
    }
    return out;
  }
 
  if (maybeConverted instanceof Set) {
    return Array.from(maybeConverted, (entry) => normalizeRpcParams(entry));
  }
 
  if (Array.isArray(maybeConverted)) {
    return maybeConverted.map((entry) => normalizeRpcParams(entry));
  }
 
  const out = {};
  for (const [key, value] of Object.entries(maybeConverted)) {
    out[key] = normalizeRpcParams(value);
  }
  return out;
}
 
function dispatchRpcSync(api, method, params) {
  if (api && typeof api[method] === "function") {
    const result = api[method](params);
    if (result && typeof result.then === "function") {
      throw new Error(
        `Spreadsheet RPC "${method}" returned a Promise; async spreadsheet APIs are not supported in Pyodide mainThread mode`,
      );
    }
    return result;
  }
  if (api && typeof api.call === "function") {
    const result = api.call(method, params);
    if (result && typeof result.then === "function") {
      throw new Error(
        `Spreadsheet RPC "${method}" returned a Promise; async spreadsheet APIs are not supported in Pyodide mainThread mode`,
      );
    }
    return result;
  }
  throw new Error(`Spreadsheet API does not implement RPC method "${method}"`);
}
 
export function registerFormulaBridge(runtime) {
  if (runtime[BRIDGE_REGISTERED_KEY]) return;
  const call = (method, params) => {
    const normalizedParams = normalizeRpcParams(params);
    const api = runtime[BRIDGE_API_KEY];
    if (!api) {
      throw new Error("PyodideRuntime has no spreadsheet API configured");
    }
    return dispatchRpcSync(api, method, normalizedParams);
  };
 
  runtime.registerJsModule("formula_bridge", {
    get_active_sheet_id: () => call("get_active_sheet_id", null),
    get_sheet_id: (name) => call("get_sheet_id", { name }),
    create_sheet: (name) => call("create_sheet", { name }),
    get_sheet_name: (sheet_id) => call("get_sheet_name", { sheet_id }),
    rename_sheet: (sheet_id, name) => call("rename_sheet", { sheet_id, name }),
 
    get_selection: () => call("get_selection", null),
    set_selection: (selection) => call("set_selection", { selection }),
 
    get_range_values: (range) => call("get_range_values", { range }),
    set_range_values: (range, values) => call("set_range_values", { range, values }),
    set_cell_value: (range, value) => call("set_cell_value", { range, value }),
    get_cell_formula: (range) => call("get_cell_formula", { range }),
    set_cell_formula: (range, formula) => call("set_cell_formula", { range, formula }),
    clear_range: (range) => call("clear_range", { range }),
 
    set_range_format: (range, format) => call("set_range_format", { range, format }),
    get_range_format: (range) => call("get_range_format", { range }),
  });
  runtime[BRIDGE_REGISTERED_KEY] = true;
}
 
export function installFormulaFiles(runtime, formulaFiles) {
  const rootDir = "/formula_api";
  runtime.FS.mkdirTree(rootDir);
 
  for (const [relPath, contents] of Object.entries(formulaFiles)) {
    const absPath = `${rootDir}/${relPath}`;
    const dir = absPath.slice(0, absPath.lastIndexOf("/"));
    runtime.FS.mkdirTree(dir);
    runtime.FS.writeFile(absPath, contents);
  }
 
  return rootDir;
}
 
export async function bootstrapFormulaBridge(runtime, rootDir) {
  await runtime.runPythonAsync(`
import sys
sys.path.insert(0, ${JSON.stringify(rootDir)})
import formula
from formula._js_bridge import JsBridge
formula.set_bridge(JsBridge())
`);
}
 
