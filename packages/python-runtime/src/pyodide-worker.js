// This file runs in a classic browser Worker context.
//
// Responsibilities:
// - Load Pyodide
// - Install the in-repo `formula` Python package into the Pyodide filesystem
// - Expose a synchronous spreadsheet RPC bridge via `pyodide.registerJsModule`
// - Execute user scripts with best-effort time + memory constraints
//
// Note: this file is intentionally not exercised in Node tests.

let pyodide = null;
let interruptView = null;

const encoder = new TextEncoder();
const decoder = new TextDecoder();

const originalFetch = self.fetch;
const originalWebSocket = self.WebSocket;

let rpcTimeoutMs = 2_000;
let rpcBufferBytes = 256 * 1024;

async function loadPyodideOnce({ indexURL } = {}) {
  if (pyodide) return pyodide;

  const resolvedIndexUrl = indexURL ?? "https://cdn.jsdelivr.net/pyodide/v0.25.1/full/";

  // Load Pyodide from the official CDN by default. Integrators can host this
  // locally and override `indexURL` when bundling.
  // eslint-disable-next-line no-undef
  importScripts(`${resolvedIndexUrl}pyodide.js`);

  // eslint-disable-next-line no-undef
  pyodide = await self.loadPyodide({
    indexURL: resolvedIndexUrl,
  });

  if (typeof SharedArrayBuffer !== "undefined" && typeof pyodide.setInterruptBuffer === "function") {
    const interruptBuffer = new SharedArrayBuffer(4);
    interruptView = new Int32Array(interruptBuffer);
    pyodide.setInterruptBuffer(interruptView);
  }

  return pyodide;
}

function getWasmMemoryBytes(runtime) {
  const mod = runtime?._module;
  const buf = mod?.wasmMemory?.buffer ?? mod?.HEAP8?.buffer;
  return buf?.byteLength ?? null;
}

function applyNetworkSandbox(permissions) {
  const mode = permissions?.network ?? "none";

  if (mode === "none") {
    self.fetch = async () => {
      throw new Error("Network access is not permitted");
    };

    self.WebSocket = class BlockedWebSocket {
      constructor() {
        throw new Error("Network access is not permitted");
      }
    };
    return;
  }

  if (mode === "allowlist") {
    const allowlist = new Set(permissions?.networkAllowlist ?? []);
    self.fetch = async (input, init) => {
      const url = typeof input === "string" ? input : input?.url;
      const hostname = new URL(url, self.location?.href ?? "https://localhost").hostname;
      if (!allowlist.has(hostname)) {
        throw new Error(`Network access to ${hostname} is not permitted`);
      }
      return originalFetch(input, init);
    };

    self.WebSocket = class AllowlistWebSocket {
      constructor(url, protocols) {
        const hostname = new URL(url, self.location?.href ?? "https://localhost").hostname;
        if (!allowlist.has(hostname)) {
          throw new Error(`Network access to ${hostname} is not permitted`);
        }
        return new originalWebSocket(url, protocols);
      }
    };
    return;
  }

  // full access
  self.fetch = originalFetch;
  self.WebSocket = originalWebSocket;
}

function rpcCallSync(method, params) {
  if (typeof SharedArrayBuffer === "undefined") {
    throw new Error("SharedArrayBuffer is required for the Pyodide formula bridge (enable crossOriginIsolated)");
  }

  const responseBuffer = new SharedArrayBuffer(8 + rpcBufferBytes);
  const header = new Int32Array(responseBuffer, 0, 2);
  const payload = new Uint8Array(responseBuffer, 8);

  header[0] = 0;
  header[1] = 0;

  self.postMessage({ type: "rpc", method, params, responseBuffer });

  const waitResult = Atomics.wait(header, 0, 0, rpcTimeoutMs);
  if (waitResult === "timed-out") {
    throw new Error(`Spreadsheet RPC "${method}" timed out after ${rpcTimeoutMs}ms`);
  }

  const length = Atomics.load(header, 1);
  const text = decoder.decode(payload.subarray(0, length));
  const parsed = JSON.parse(text);
  if (parsed.error) {
    throw new Error(parsed.error);
  }
  return parsed.result;
}

function registerFormulaBridge(runtime) {
  runtime.registerJsModule("formula_bridge", {
    get_active_sheet_id: () => rpcCallSync("get_active_sheet_id", null),
    get_sheet_id: (name) => rpcCallSync("get_sheet_id", { name }),
    create_sheet: (name) => rpcCallSync("create_sheet", { name }),
    get_sheet_name: (sheet_id) => rpcCallSync("get_sheet_name", { sheet_id }),
    rename_sheet: (sheet_id, name) => rpcCallSync("rename_sheet", { sheet_id, name }),

    get_range_values: (range) => rpcCallSync("get_range_values", { range }),
    set_range_values: (range, values) => rpcCallSync("set_range_values", { range, values }),
    set_cell_value: (range, value) => rpcCallSync("set_cell_value", { range, value }),
    get_cell_formula: (range) => rpcCallSync("get_cell_formula", { range }),
    set_cell_formula: (range, formula) => rpcCallSync("set_cell_formula", { range, formula }),
    clear_range: (range) => rpcCallSync("clear_range", { range }),
  });
}

function installFormulaFiles(runtime, formulaFiles) {
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

async function bootstrapFormulaBridge(runtime, rootDir) {
  // Ensure the formula API is importable and has an active Bridge configured.
  await runtime.runPythonAsync(`
import sys
sys.path.insert(0, ${JSON.stringify(rootDir)})
import formula
from formula._js_bridge import JsBridge
formula.set_bridge(JsBridge())
`);
}

async function applyPythonSandbox(runtime, permissions) {
  // We apply sandboxing in Python as well (blocking `open()`, `input()`, and
  // common filesystem/network imports). JS network access is separately
  // controlled via `applyNetworkSandbox`.
  await runtime.runPythonAsync(`
from formula.runtime.sandbox import apply_sandbox
apply_sandbox(${JSON.stringify(permissions ?? {})})
`);
}

async function runWithTimeout(runtime, code, timeoutMs) {
  if (!interruptView || !Number.isFinite(timeoutMs) || timeoutMs <= 0) {
    return await runtime.runPythonAsync(code);
  }

  interruptView[0] = 0;
  const timer = setTimeout(() => {
    interruptView[0] = 2;
    Atomics.notify(interruptView, 0);
  }, timeoutMs);

  try {
    return await runtime.runPythonAsync(code);
  } finally {
    clearTimeout(timer);
    interruptView[0] = 0;
  }
}

self.onmessage = async (event) => {
  const msg = event.data;
  if (!msg || typeof msg.type !== "string") return;

  if (msg.type === "init") {
    try {
      const runtime = await loadPyodideOnce({ indexURL: msg.indexURL });

      rpcTimeoutMs = Number.isFinite(msg.rpcTimeoutMs) ? msg.rpcTimeoutMs : rpcTimeoutMs;
      rpcBufferBytes = Number.isFinite(msg.rpcBufferBytes) ? msg.rpcBufferBytes : rpcBufferBytes;

      if (!msg.formulaFiles || typeof msg.formulaFiles !== "object") {
        throw new Error("Pyodide init requires `formulaFiles` to install the formula Python API");
      }

      const rootDir = installFormulaFiles(runtime, msg.formulaFiles);
      registerFormulaBridge(runtime);
      await bootstrapFormulaBridge(runtime, rootDir);

      // Apply default sandbox for subsequent executions.
      applyNetworkSandbox(msg.permissions ?? {});
      await applyPythonSandbox(runtime, msg.permissions ?? {});

      self.postMessage({ type: "ready" });
    } catch (err) {
      self.postMessage({ type: "ready", error: err?.message ?? String(err) });
    }
    return;
  }

  if (msg.type === "execute") {
    const requestId = msg.requestId;
    try {
      const runtime = await loadPyodideOnce({ indexURL: msg.indexURL });

      const permissions = msg.permissions ?? {};
      applyNetworkSandbox(permissions);
      await applyPythonSandbox(runtime, permissions);

      const maxMemoryBytes = msg.maxMemoryBytes;
      const beforeMem = getWasmMemoryBytes(runtime);

      const result = await runWithTimeout(runtime, msg.code, msg.timeoutMs);

      const afterMem = getWasmMemoryBytes(runtime);
      if (Number.isFinite(maxMemoryBytes) && maxMemoryBytes > 0 && afterMem != null && afterMem > maxMemoryBytes) {
        throw new Error(`Pyodide memory limit exceeded: ${afterMem} bytes > ${maxMemoryBytes} bytes`);
      }

      // If memory grew substantially during the run, still return the result but
      // include some debugging metadata.
      self.postMessage({
        type: "result",
        requestId,
        success: true,
        result,
        memory: { before: beforeMem, after: afterMem },
      });
    } catch (err) {
      self.postMessage({ type: "result", requestId, success: false, error: err?.message ?? String(err) });
    }
  }
};
