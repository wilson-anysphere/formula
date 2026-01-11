/**
 * Browser runtime (Pyodide).
 *
 * Prefer running Pyodide in a Worker (off the UI thread) when
 * `crossOriginIsolated` + `SharedArrayBuffer` are available. Fall back to a
 * main-thread Pyodide instance otherwise so Python scripting can still work in
 * non-COOP/COEP contexts (with degraded UX: the UI may freeze while Python runs).
 *
 * This package intentionally keeps the runtime surface stable while the
 * underlying spreadsheet engine bridge solidifies.
 */
import { dispatchRpc } from "./rpc.js";
import { formulaFiles as bundledFormulaFiles } from "./formula-files.generated.js";
import {
  applyNetworkSandbox as applyMainThreadNetworkSandbox,
  applyPythonSandbox,
  bootstrapFormulaBridge,
  getWasmMemoryBytes,
  installFormulaFiles,
  loadPyodideMainThread,
  registerFormulaBridge,
  setFormulaBridgeApi,
  resolveIndexURL,
  runWithTimeout,
  withCapturedOutput,
} from "./pyodide-main-thread.js";

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
    this.indexURL = options.indexURL;
    this.mode = options.mode ?? "auto";
    this.rpcTimeoutMs = options.rpcTimeoutMs;
    this.rpcBufferBytes = options.rpcBufferBytes;
    this.timeoutMs = options.timeoutMs ?? 5_000;
    this.maxMemoryBytes = options.maxMemoryBytes ?? 256 * 1024 * 1024;
    this.permissions = options.permissions ?? defaultPermissions();
    this.api = options.api;
    this.formulaFiles = options.formulaFiles ?? bundledFormulaFiles;
    this.onOutput = options.onOutput ?? null;
    this.worker = null;
    this._onRpcMessage = null;
    this.pyodide = null;
    this._interruptView = null;
    this._mainThreadReady = false;
    this._executeQueue = Promise.resolve();
    this.backendMode = null;
    this.initialized = false;
  }

  destroy() {
    if (this.worker) {
      try {
        if (this._onRpcMessage) {
          this.worker.removeEventListener("message", this._onRpcMessage);
        }
        this.worker.terminate();
      } catch {
        // ignore
      } finally {
        this.worker = null;
        this._onRpcMessage = null;
      }
    }

    // Main-thread Pyodide cannot be fully unloaded, but dropping references lets
    // callers re-initialize and allows GC to reclaim JS-side state.
    this.pyodide = null;
    this._interruptView = null;
    this._mainThreadReady = false;
    this._executeQueue = Promise.resolve();
    this.backendMode = null;
    this.initialized = false;
  }

  /**
   * Return the backend mode this runtime will use given the current `mode`
   * setting and environment.
   *
   * This is safe to call before `initialize()` and is used by UI code to surface
   * degraded-mode warnings.
   */
  getBackendMode() {
    return this.backendMode ?? resolveBackendMode(this.mode);
  }

  _assertWorkerBackendAvailable() {
    if (typeof Worker === "undefined") {
      throw new Error("PyodideRuntime requires Web Worker support");
    }
    if (typeof SharedArrayBuffer === "undefined" || globalThis.crossOriginIsolated !== true) {
      throw new Error(
        "PyodideRuntime worker mode requires SharedArrayBuffer in a cross-origin isolated context (COOP/COEP)",
      );
    }
  }

  async _initializeWorker() {
    this._assertWorkerBackendAvailable();

    if (this.worker) {
      try {
        if (this._onRpcMessage) {
          this.worker.removeEventListener("message", this._onRpcMessage);
        }
        this.worker.terminate();
      } catch {
        // ignore
      } finally {
        this.worker = null;
        this._onRpcMessage = null;
      }
    }

    // Pyodide's upstream loader is typically pulled in via `importScripts()`,
    // which requires a classic worker.
    this.worker = new Worker(this.workerUrl, { type: "classic" });

    this._onRpcMessage = (event) => {
      const msg = event.data;
      if (!msg || typeof msg.type !== "string") return;

      if (msg.type === "output") {
        if (typeof this.onOutput === "function") {
          try {
            this.onOutput({
              requestId: typeof msg.requestId === "string" ? msg.requestId : null,
              stream: msg.stream === "stderr" ? "stderr" : "stdout",
              text: typeof msg.text === "string" ? msg.text : String(msg.text ?? ""),
            });
          } catch {
            // ignore output handler errors
          }
        }
        return;
      }

      if (msg.type !== "rpc") return;

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
            this.destroy();
            reject(new Error(event.data.error));
          } else {
            resolve();
          }
        }
      };
      const onError = (err) => {
        this.worker.removeEventListener("message", onMessage);
        this.destroy();
        reject(err);
      };
      this.worker.addEventListener("message", onMessage);
      this.worker.addEventListener("error", onError, { once: true });
      this.worker.postMessage({
        type: "init",
        indexURL: this.indexURL,
        maxMemoryBytes: this.maxMemoryBytes,
        permissions: this.permissions,
        formulaFiles: this.formulaFiles,
        rpcTimeoutMs: this.rpcTimeoutMs,
        rpcBufferBytes: this.rpcBufferBytes,
      });
    });
  }

  async _initializeMainThread() {
    const resolvedIndexURL = resolveIndexURL(this.indexURL);
    this.indexURL = resolvedIndexURL;

    this.pyodide = await loadPyodideMainThread({ indexURL: resolvedIndexURL });

    if (
      !this._interruptView &&
      typeof SharedArrayBuffer !== "undefined" &&
      typeof this.pyodide.setInterruptBuffer === "function"
    ) {
      const interruptBuffer = new SharedArrayBuffer(1);
      this._interruptView = new Uint8Array(interruptBuffer);
      this.pyodide.setInterruptBuffer(this._interruptView);
    }

    setFormulaBridgeApi(this.pyodide, this.api);

    if (!this._mainThreadReady) {
      const rootDir = installFormulaFiles(this.pyodide, this.formulaFiles);
      registerFormulaBridge(this.pyodide);
      await bootstrapFormulaBridge(this.pyodide, rootDir);
      this._mainThreadReady = true;
    }

    await applyPythonSandbox(this.pyodide, this.permissions ?? {});
  }

  /**
   * Initialize the runtime and load Pyodide.
   *
   * Modes:
   * - `auto` (default): use a Worker backend when `SharedArrayBuffer` is
   *   available and `crossOriginIsolated` is true, otherwise fall back to
   *   main-thread Pyodide.
   * - `worker`: force Worker mode (requires COOP/COEP + SharedArrayBuffer).
   * - `mainThread`: force main-thread Pyodide (UI thread will block).
   */
  async initialize(options = {}) {
    const requestedMode = options.mode ?? this.mode ?? "auto";
    this.mode = requestedMode;
    this.api = options.api ?? this.api;
    this.formulaFiles = options.formulaFiles ?? this.formulaFiles;
    this.permissions = options.permissions ?? this.permissions;
    this.indexURL = options.indexURL ?? this.indexURL;
    this.rpcTimeoutMs = options.rpcTimeoutMs ?? this.rpcTimeoutMs;
    this.rpcBufferBytes = options.rpcBufferBytes ?? this.rpcBufferBytes;
    this.onOutput = options.onOutput ?? this.onOutput;

    if (!this.formulaFiles) {
      throw new Error("PyodideRuntime.initialize requires { formulaFiles } to install the in-repo formula API");
    }

    const selectedMode = resolveBackendMode(requestedMode);

    if (this.initialized && this.backendMode === selectedMode) {
      // Already initialized. The worker-side RPC handler and main-thread bridge
      // both read `this.api` at call time, so callers can swap the bridge by
      // updating `api` + `activeSheetId` without reloading Pyodide.
      return;
    }

    if (this.initialized) {
      this.destroy();
    }

    this.backendMode = selectedMode;
    try {
      if (selectedMode === "worker") {
        await this._initializeWorker();
      } else {
        await this._initializeMainThread();
      }

      this.initialized = true;
    } catch (err) {
      // Ensure we don't leave partially-initialized state around (e.g. failed
      // script load, worker creation error).
      this.destroy();
      throw err;
    }
  }

  /**
   * Execute a Python script.
   */
  async execute(code, { timeoutMs, maxMemoryBytes, permissions, requestId } = {}) {
    if (!this.initialized || !this.backendMode) {
      throw new Error("PyodideRuntime not initialized; call initialize() first");
    }

    const effectiveTimeout = timeoutMs ?? this.timeoutMs;
    const effectiveRequestId = requestId ?? globalThis.crypto?.randomUUID?.() ?? String(Date.now());
    const effectiveMaxMemory = maxMemoryBytes ?? this.maxMemoryBytes;
    const effectivePermissions = permissions ?? this.permissions;

    if (this.backendMode === "worker") {
      if (!this.worker) {
        throw new Error("PyodideRuntime worker backend not initialized; call initialize() first");
      }

      return await new Promise((resolve, reject) => {
        const worker = this.worker;
        // The worker enforces the timeout as well (via Pyodide interrupts). This
        // timer is a last-resort failsafe in case the worker stops responding.
        const timer =
          Number.isFinite(effectiveTimeout) && effectiveTimeout > 0
            ? setTimeout(() => {
                this.destroy();
                reject(new Error("Pyodide script timed out"));
              }, effectiveTimeout + 250)
            : null;

        const onMessage = (event) => {
          const msg = event.data;
          if (msg?.type !== "result" || msg?.requestId !== effectiveRequestId) return;
          worker.removeEventListener("message", onMessage);
          worker.removeEventListener("error", onError);
          if (timer) clearTimeout(timer);
          if (msg.success) {
            resolve(msg);
            return;
          }

          // If the worker exceeded memory limits, it's safer to reset the runtime.
          if (typeof msg.error === "string" && msg.error.includes("Pyodide memory limit exceeded")) {
            this.destroy();
          }

          const err = new Error(msg.error || "Pyodide script failed");
          if (typeof msg.stdout === "string" && msg.stdout.length > 0) {
            err.stdout = msg.stdout;
          }
          if (typeof msg.stderr === "string" && msg.stderr.length > 0) {
            err.stderr = msg.stderr;
          }
          reject(err);
        };

        const onError = (err) => {
          worker.removeEventListener("message", onMessage);
          if (timer) clearTimeout(timer);
          this.destroy();
          reject(err);
        };

        worker.addEventListener("message", onMessage);
        worker.addEventListener("error", onError, { once: true });
        worker.postMessage({
          type: "execute",
          requestId: effectiveRequestId,
          code,
          indexURL: this.indexURL,
          timeoutMs: effectiveTimeout,
          maxMemoryBytes: effectiveMaxMemory,
          permissions: effectivePermissions,
        });
      });
    }

    if (!this.pyodide) {
      throw new Error("PyodideRuntime mainThread backend not initialized; call initialize() first");
    }

    const run = async () => {
      let stdout = "";
      let stderr = "";
      setFormulaBridgeApi(this.pyodide, this.api);
      const beforeMem = getWasmMemoryBytes(this.pyodide);
      const restoreNetworkSandbox = applyMainThreadNetworkSandbox(effectivePermissions);

      try {
        await applyPythonSandbox(this.pyodide, effectivePermissions);

        const { value: result, stdout: capturedStdout, stderr: capturedStderr } = await withCapturedOutput(this.pyodide, () =>
          runWithTimeout(this.pyodide, code, effectiveTimeout, this._interruptView),
        );

        stdout = capturedStdout;
        stderr = capturedStderr;

        const afterMem = getWasmMemoryBytes(this.pyodide);
        if (Number.isFinite(effectiveMaxMemory) && effectiveMaxMemory > 0 && afterMem != null && afterMem > effectiveMaxMemory) {
          throw new Error(`Pyodide memory limit exceeded: ${afterMem} bytes > ${effectiveMaxMemory} bytes`);
        }

        return {
          type: "result",
          requestId: effectiveRequestId,
          success: true,
          result,
          stdout,
          stderr,
          memory: { before: beforeMem, after: afterMem },
        };
      } catch (err) {
        stdout = err?.stdout ?? stdout;
        stderr = err?.stderr ?? stderr;

        // If we exceeded memory limits, drop references so callers can re-init.
        if (err instanceof Error && err.message.includes("Pyodide memory limit exceeded")) {
          this.destroy();
        }

        const wrapped = new Error(err?.message ?? String(err));
        if (typeof stdout === "string" && stdout.length > 0) wrapped.stdout = stdout;
        if (typeof stderr === "string" && stderr.length > 0) wrapped.stderr = stderr;
        throw wrapped;
      } finally {
        try {
          restoreNetworkSandbox();
        } catch {
          // ignore
        }
      }
    };

    // Ensure main-thread executions don't overlap (stdout/stderr, sandbox state).
    this._executeQueue = this._executeQueue.then(run, run);
    return await this._executeQueue;
  }
}

function resolveBackendMode(mode) {
  const requested = mode ?? "auto";
  if (requested === "worker") return "worker";
  if (requested === "mainThread") return "mainThread";

  const canUseWorker =
    typeof Worker !== "undefined" &&
    typeof SharedArrayBuffer !== "undefined" &&
    globalThis.crossOriginIsolated === true;
  return canUseWorker ? "worker" : "mainThread";
}
