/**
 * Browser runtime (Pyodide) – implemented as a Worker-backed runtime that loads
 * Pyodide and executes scripts off the UI thread.
 *
 * This package intentionally keeps the runtime surface stable while the
 * underlying spreadsheet engine bridge solidifies.
 */
export class PyodideRuntime {
  constructor(options = {}) {
    this.workerUrl = options.workerUrl ?? new URL("./pyodide-worker.js", import.meta.url);
    this.timeoutMs = options.timeoutMs ?? 5_000;
    this.maxMemoryBytes = options.maxMemoryBytes ?? 256 * 1024 * 1024;
  }

  /**
   * Initialize the worker and load Pyodide.
   *
   * Note: this is a minimal scaffold; the full integration will register a
   * `formula_bridge` JS module that forwards spreadsheet operations to the core
   * engine.
   */
  async initialize() {
    if (typeof Worker === "undefined") {
      throw new Error("PyodideRuntime requires Web Worker support");
    }

    this.worker = new Worker(this.workerUrl, { type: "module" });
    await new Promise((resolve, reject) => {
      const onMessage = (event) => {
        if (event.data?.type === "ready") {
          this.worker.removeEventListener("message", onMessage);
          resolve();
        }
      };
      const onError = (err) => {
        this.worker.removeEventListener("message", onMessage);
        reject(err);
      };
      this.worker.addEventListener("message", onMessage);
      this.worker.addEventListener("error", onError, { once: true });
      this.worker.postMessage({ type: "init", maxMemoryBytes: this.maxMemoryBytes });
    });
  }

  /**
   * Execute a Python script inside the Pyodide worker.
   *
   * Spreadsheet operations are expected to be bridged by injecting a JS module
   * (e.g. `formula_bridge`) into Pyodide.
   */
  async execute(code, { timeoutMs } = {}) {
    if (!this.worker) {
      throw new Error("PyodideRuntime not initialized; call initialize() first");
    }

    const effectiveTimeout = timeoutMs ?? this.timeoutMs;
    const requestId = crypto.randomUUID?.() ?? String(Date.now());

    return await new Promise((resolve, reject) => {
      const timer =
        Number.isFinite(effectiveTimeout) && effectiveTimeout > 0
          ? setTimeout(() => {
              this.worker.terminate();
              reject(new Error("Pyodide script timed out"));
            }, effectiveTimeout)
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
      this.worker.postMessage({ type: "execute", requestId, code });
    });
  }
}

