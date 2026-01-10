/**
 * Browser runtime (Pyodide) – implemented as a Worker-backed runtime that loads
 * Pyodide and executes scripts off the UI thread.
 *
 * This package intentionally keeps the runtime surface stable while the
 * underlying spreadsheet engine bridge solidifies.
 */
import { dispatchRpc } from "./rpc.js";

function defaultPermissions() {
  return { filesystem: "none", network: "none" };
}

function writeSharedRpcResponse(sharedBuffer, { result, error }) {
  const header = new Int32Array(sharedBuffer, 0, 2);
  const payload = new Uint8Array(sharedBuffer, 8);

  const encoder = new TextEncoder();
  let bytes = encoder.encode(JSON.stringify({ result, error }));

  if (bytes.length > payload.length) {
    bytes = encoder.encode(JSON.stringify({ result: null, error: "RPC response too large for shared buffer" }));
  }

  payload.fill(0);
  payload.set(bytes);
  Atomics.store(header, 1, bytes.length);
  Atomics.store(header, 0, 1);
  Atomics.notify(header, 0);
}

export class PyodideRuntime {
  constructor(options = {}) {
    this.workerUrl = options.workerUrl ?? new URL("./pyodide-worker.js", import.meta.url);
    this.timeoutMs = options.timeoutMs ?? 5_000;
    this.maxMemoryBytes = options.maxMemoryBytes ?? 256 * 1024 * 1024;
    this.permissions = options.permissions ?? defaultPermissions();
    this.api = options.api;
    this.formulaFiles = options.formulaFiles;
  }

  /**
   * Initialize the worker and load Pyodide.
   *
   * Note: this is a minimal scaffold; the full integration will register a
   * `formula_bridge` JS module that forwards spreadsheet operations to the core
   * engine.
   */
  async initialize(options = {}) {
    if (typeof Worker === "undefined") {
      throw new Error("PyodideRuntime requires Web Worker support");
    }

    this.api = options.api ?? this.api;
    this.formulaFiles = options.formulaFiles ?? this.formulaFiles;
    this.permissions = options.permissions ?? this.permissions;

    if (!this.formulaFiles) {
      throw new Error("PyodideRuntime.initialize requires { formulaFiles } to install the in-repo formula API");
    }

    // Pyodide's upstream loader is typically pulled in via `importScripts()`,
    // which requires a classic worker.
    this.worker = new Worker(this.workerUrl, { type: "classic" });

    this._onRpcMessage = (event) => {
      const msg = event.data;
      if (!msg || msg.type !== "rpc") return;

      if (typeof SharedArrayBuffer === "undefined") {
        return;
      }

      const sharedBuffer = msg.responseBuffer;
      if (!(sharedBuffer instanceof SharedArrayBuffer)) {
        return;
      }

      const method = msg.method;
      const params = msg.params;

      (async () => {
        try {
          if (!this.api) {
            throw new Error("PyodideRuntime has no spreadsheet API configured");
          }
          const result = await dispatchRpc(this.api, method, params);
          writeSharedRpcResponse(sharedBuffer, { result, error: null });
        } catch (err) {
          writeSharedRpcResponse(sharedBuffer, {
            result: null,
            error: err instanceof Error ? err.message : String(err),
          });
        }
      })();
    };

    this.worker.addEventListener("message", this._onRpcMessage);

    await new Promise((resolve, reject) => {
      const onMessage = (event) => {
        if (event.data?.type === "ready") {
          this.worker.removeEventListener("message", onMessage);
          if (event.data?.error) {
            reject(new Error(event.data.error));
          } else {
            resolve();
          }
        }
      };
      const onError = (err) => {
        this.worker.removeEventListener("message", onMessage);
        reject(err);
      };
      this.worker.addEventListener("message", onMessage);
      this.worker.addEventListener("error", onError, { once: true });
      this.worker.postMessage({
        type: "init",
        maxMemoryBytes: this.maxMemoryBytes,
        permissions: this.permissions,
        formulaFiles: this.formulaFiles,
      });
    });
  }

  /**
   * Execute a Python script inside the Pyodide worker.
   *
   * Spreadsheet operations are expected to be bridged by injecting a JS module
   * (e.g. `formula_bridge`) into Pyodide.
   */
  async execute(code, { timeoutMs, maxMemoryBytes, permissions } = {}) {
    if (!this.worker) {
      throw new Error("PyodideRuntime not initialized; call initialize() first");
    }

    const effectiveTimeout = timeoutMs ?? this.timeoutMs;
    const requestId = globalThis.crypto?.randomUUID?.() ?? String(Date.now());
    const effectiveMaxMemory = maxMemoryBytes ?? this.maxMemoryBytes;
    const effectivePermissions = permissions ?? this.permissions;

    return await new Promise((resolve, reject) => {
      // The worker enforces the timeout as well (via Pyodide interrupts). This
      // timer is a last-resort failsafe in case the worker stops responding.
      const timer =
        Number.isFinite(effectiveTimeout) && effectiveTimeout > 0
          ? setTimeout(() => {
              this.worker.terminate();
              reject(new Error("Pyodide script timed out"));
            }, effectiveTimeout + 250)
          : null;

      const onMessage = (event) => {
        const msg = event.data;
        if (msg?.type !== "result" || msg?.requestId !== requestId) return;
        this.worker.removeEventListener("message", onMessage);
        if (timer) clearTimeout(timer);
        if (msg.success) resolve(msg);
        else reject(new Error(msg.error || "Pyodide script failed"));
      };

      this.worker.addEventListener("message", onMessage);
      this.worker.postMessage({
        type: "execute",
        requestId,
        code,
        timeoutMs: effectiveTimeout,
        maxMemoryBytes: effectiveMaxMemory,
        permissions: effectivePermissions,
      });
    });
  }
}
