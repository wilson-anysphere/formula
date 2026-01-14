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

let executeQueue = Promise.resolve();

async function loadPyodideOnce({ indexURL } = {}) {
  if (pyodide) return pyodide;

  let resolvedIndexUrl = indexURL ?? "https://cdn.jsdelivr.net/pyodide/v0.25.1/full/";
  if (!resolvedIndexUrl.endsWith("/")) {
    resolvedIndexUrl += "/";
  }

  // Load Pyodide from the official CDN by default. Integrators can host this
  // locally and override `indexURL` when bundling.
  // eslint-disable-next-line no-undef
  importScripts(`${resolvedIndexUrl}pyodide.js`);

  // eslint-disable-next-line no-undef
  pyodide = await self.loadPyodide({
    indexURL: resolvedIndexUrl,
  });

  if (typeof SharedArrayBuffer !== "undefined" && typeof pyodide.setInterruptBuffer === "function") {
    // Follow Pyodide's recommended interrupt buffer format: a single byte that
    // can be set to 2 to raise KeyboardInterrupt.
    const interruptBuffer = new SharedArrayBuffer(1);
    interruptView = new Uint8Array(interruptBuffer);
    pyodide.setInterruptBuffer(interruptView);
  }

  return pyodide;
}

function withCapturedOutput(runtime, fn) {
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
        // Best-effort; if Pyodide doesn't support overriding streams in this build,
        // we still run the script but return empty output.
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
    // Restore default behavior between executions to avoid output from one run
    // being captured by the next. We intentionally route to console here.
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

function getWasmMemoryBytes(runtime) {
  const mod = runtime?._module;
  const buf = mod?.wasmMemory?.buffer ?? mod?.HEAP8?.buffer;
  return buf?.byteLength ?? null;
}

function applyNetworkSandbox(permissions) {
  const mode = permissions?.network ?? "none";

  if (mode === "none") {
    try {
      self.fetch = async () => {
        throw new Error("Network access is not permitted");
      };
    } catch {
      // Best-effort: some hosts may expose non-writable globals.
    }

    try {
      self.WebSocket = class BlockedWebSocket {
        constructor() {
          throw new Error("Network access is not permitted");
        }
      };
    } catch {
      // ignore
    }
    return;
  }

  if (mode === "allowlist") {
    const allowlist = new Set(permissions?.networkAllowlist ?? []);
    try {
      self.fetch = async (input, init) => {
        const url = typeof input === "string" ? input : input?.url;
        const hostname = new URL(url, self.location?.href ?? "https://localhost").hostname;
        if (!allowlist.has(hostname)) {
          throw new Error(`Network access to ${hostname} is not permitted`);
        }
        if (typeof originalFetch !== "function") {
          throw new Error("Network access is not permitted (fetch is unavailable)");
        }
        return originalFetch(input, init);
      };
    } catch {
      // ignore
    }

    try {
      self.WebSocket = class AllowlistWebSocket {
        constructor(url, protocols) {
          const hostname = new URL(url, self.location?.href ?? "https://localhost").hostname;
          if (!allowlist.has(hostname)) {
            throw new Error(`Network access to ${hostname} is not permitted`);
          }
          if (typeof originalWebSocket !== "function") {
            throw new Error("Network access is not permitted (WebSocket is unavailable)");
          }
          return new originalWebSocket(url, protocols);
        }
      };
    } catch {
      // ignore
    }
    return;
  }

  // full access
  try {
    self.fetch = originalFetch;
  } catch {
    // ignore
  }
  try {
    self.WebSocket = originalWebSocket;
  } catch {
    // ignore
  }
}

function coercePyProxy(value) {
  if (!value || (typeof value !== "object" && typeof value !== "function")) return value;
  if (typeof value.toJs !== "function") return value;

  // When called from Python, Pyodide may pass dict/list values as PyProxy
  // instances. Those are not structured-cloneable, so convert to plain JS before
  // sending across the worker boundary.
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

function rpcCallSync(method, params) {
  if (typeof SharedArrayBuffer === "undefined") {
    throw new Error(
      "SharedArrayBuffer is required for the worker-backed Pyodide formula bridge. " +
        "Enable crossOriginIsolated (COOP/COEP) or run Pyodide on the main thread.",
    );
  }

  const normalizedParams = normalizeRpcParams(params);

  const responseBuffer = new SharedArrayBuffer(8 + rpcBufferBytes);
  const header = new Int32Array(responseBuffer, 0, 2);
  const payload = new Uint8Array(responseBuffer, 8);

  header[0] = 0;
  header[1] = 0;

  self.postMessage({ type: "rpc", method, params: normalizedParams, responseBuffer });

  const waitResult = Atomics.wait(header, 0, 0, rpcTimeoutMs);
  if (waitResult === "timed-out") {
    throw new Error(`Spreadsheet RPC "${method}" timed out after ${rpcTimeoutMs}ms`);
  }

  const length = Atomics.load(header, 1);
  // Some browsers disallow decoding directly from SharedArrayBuffer-backed views.
  // Copy into an ArrayBuffer-backed Uint8Array before decoding.
  const text = decoder.decode(payload.slice(0, length));
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
    create_sheet: (name, index) => rpcCallSync("create_sheet", index === undefined ? { name } : { name, index }),
    get_sheet_name: (sheet_id) => rpcCallSync("get_sheet_name", { sheet_id }),
    rename_sheet: (sheet_id, name) => rpcCallSync("rename_sheet", { sheet_id, name }),

    get_selection: () => rpcCallSync("get_selection", null),
    set_selection: (selection) => rpcCallSync("set_selection", { selection }),

    get_range_values: (range) => rpcCallSync("get_range_values", { range }),
    set_range_values: (range, values) => rpcCallSync("set_range_values", { range, values }),
    set_cell_value: (range, value) => rpcCallSync("set_cell_value", { range, value }),
    get_cell_formula: (range) => rpcCallSync("get_cell_formula", { range }),
    set_cell_formula: (range, formula) => rpcCallSync("set_cell_formula", { range, formula }),
    clear_range: (range) => rpcCallSync("clear_range", { range }),

    set_range_format: (range, format) => rpcCallSync("set_range_format", { range, format }),
    get_range_format: (range) => rpcCallSync("get_range_format", { range }),
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
# Keep the formula API path at the front of sys.path without accumulating duplicates.
_formula_root = ${JSON.stringify(rootDir)}
while _formula_root in sys.path:
    sys.path.remove(_formula_root)
sys.path.insert(0, _formula_root)
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
  }, timeoutMs);

  try {
    return await runtime.runPythonAsync(code);
  } finally {
    clearTimeout(timer);
    interruptView[0] = 0;
  }
}

self.onmessage = (event) => {
  void (async () => {
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
      executeQueue = executeQueue
        .then(async () => {
          const requestId = msg.requestId;
          let stdout = "";
          let stderr = "";
          try {
            const runtime = await loadPyodideOnce({ indexURL: msg.indexURL });

            const permissions = msg.permissions ?? {};
            applyNetworkSandbox(permissions);
            await applyPythonSandbox(runtime, permissions);

            const maxMemoryBytes = msg.maxMemoryBytes;
            const beforeMem = getWasmMemoryBytes(runtime);

            const { value: result, stdout: capturedStdout, stderr: capturedStderr } = await withCapturedOutput(
              runtime,
              () => runWithTimeout(runtime, msg.code, msg.timeoutMs),
            );
            stdout = capturedStdout;
            stderr = capturedStderr;

            const afterMem = getWasmMemoryBytes(runtime);
            if (
              Number.isFinite(maxMemoryBytes) &&
              maxMemoryBytes > 0 &&
              afterMem != null &&
              afterMem > maxMemoryBytes
            ) {
              throw new Error(`Pyodide memory limit exceeded: ${afterMem} bytes > ${maxMemoryBytes} bytes`);
            }

            // If memory grew substantially during the run, still return the result but
            // include some debugging metadata.
            self.postMessage({
              type: "result",
              requestId,
              success: true,
              result,
              stdout,
              stderr,
              memory: { before: beforeMem, after: afterMem },
            });
          } catch (err) {
            stdout = err?.stdout ?? stdout;
            stderr = err?.stderr ?? stderr;
            self.postMessage({
              type: "result",
              requestId,
              success: false,
              error: err?.message ?? String(err),
              stdout,
              stderr,
            });
          }
        })
        .catch((err) => {
          // Keep the execution queue alive even if a postMessage/handler bug throws.
          try {
            console.error("Unhandled Pyodide worker execution error:", err);
          } catch {
            // ignore
          }
        });
    }
  })().catch((err) => {
    // `onmessage` handlers ignore returned promises; ensure we always terminate the
    // promise chain to avoid unhandled rejections.
    try {
      console.error("Unhandled Pyodide worker message error:", err);
    } catch {
      // ignore
    }
  });
};
