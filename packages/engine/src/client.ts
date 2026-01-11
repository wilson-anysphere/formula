import type { CellChange, CellData, CellScalar, RpcOptions } from "./protocol";
import { defaultWasmBinaryUrl, defaultWasmModuleUrl } from "./wasm";
import { EngineWorker } from "./worker/EngineWorker";

export interface EngineClient {
  /**
   * Force initialization of the underlying worker/WASM engine.
   *
   * Note: the engine may still lazy-load WASM on first request.
   */
  init(): Promise<void>;
  newWorkbook(): Promise<void>;
  loadWorkbookFromJson(json: string): Promise<void>;
  /**
   * Load a workbook by passing raw `.xlsx` bytes directly to the WASM engine.
   */
  loadWorkbookFromXlsxBytes(bytes: Uint8Array, options?: RpcOptions): Promise<void>;
  toJson(): Promise<string>;
  getCell(address: string, sheet?: string, options?: RpcOptions): Promise<CellData>;
  getRange(range: string, sheet?: string, options?: RpcOptions): Promise<CellData[][]>;
  /**
   * Set a single cell, batched across the current microtask to minimize RPC
   * overhead.
   */
  setCell(address: string, value: CellScalar, sheet?: string): Promise<void>;
  /**
   * Set multiple cells in a single RPC call.
   *
   * Useful when applying large delta batches (paste, imports) without creating
   * per-cell promises.
   */
  setCells(
    updates: Array<{ address: string; value: CellScalar; sheet?: string }>,
    options?: RpcOptions
  ): Promise<void>;
  setRange(range: string, values: CellScalar[][], sheet?: string, options?: RpcOptions): Promise<void>;
  recalculate(sheet?: string, options?: RpcOptions): Promise<CellChange[]>;
  terminate(): void;
}

export function createEngineClient(options?: { wasmModuleUrl?: string; wasmBinaryUrl?: string }): EngineClient {
  if (typeof Worker === "undefined") {
    throw new Error("createEngineClient() requires a Worker-capable environment");
  }

  const wasmModuleUrl = options?.wasmModuleUrl ?? defaultWasmModuleUrl();
  const wasmBinaryUrl = options?.wasmBinaryUrl ?? defaultWasmBinaryUrl();

  // Vite supports Worker construction via `new URL(..., import.meta.url)` and will
  // bundle the Worker entrypoint correctly for both dev and production builds.
  //
  // In React 18 StrictMode (dev-only), effects intentionally run twice
  // (setup → cleanup → setup). Callers may `terminate()` between setups, so we
  // support tearing down the worker and reconnecting on-demand.
  let worker: Worker | null = null;
  let engine: EngineWorker | null = null;
  let enginePromise: Promise<EngineWorker> | null = null;
  let generation = 0;

  const ensureWorker = () => {
    if (worker) return worker;
    worker = new Worker(new URL("./engine.worker.ts", import.meta.url), {
      type: "module"
    });
    return worker;
  };

  const connect = () => {
    if (enginePromise) {
      return enginePromise;
    }

    const connectGeneration = ++generation;
    const activeWorker = ensureWorker();

    enginePromise = EngineWorker.connect({
      worker: activeWorker,
      wasmModuleUrl,
      wasmBinaryUrl
    });

    void enginePromise
      .then((connected) => {
        // If the caller terminated/restarted while we were connecting, immediately
        // dispose the stale connection.
        if (connectGeneration !== generation) {
          connected.terminate();
          return;
        }
        engine = connected;
      })
      .catch(() => {
        // Allow retries on the next call.
        if (connectGeneration !== generation) {
          return;
        }
        enginePromise = null;
        engine = null;
        worker?.terminate();
        worker = null;
      });

    return enginePromise;
  };

  const withEngine = async <T>(fn: (engine: EngineWorker) => Promise<T>): Promise<T> => {
    const connected = await connect();
    return await fn(connected);
  };

  return {
    init: async () => {
      await connect();
    },
    newWorkbook: async () => await withEngine((connected) => connected.newWorkbook()),
    loadWorkbookFromJson: async (json) => await withEngine((connected) => connected.loadWorkbookFromJson(json)),
    loadWorkbookFromXlsxBytes: async (bytes, rpcOptions) =>
      await withEngine((connected) => connected.loadWorkbookFromXlsxBytes(bytes, rpcOptions)),
    toJson: async () => await withEngine((connected) => connected.toJson()),
    getCell: async (address, sheet, rpcOptions) =>
      await withEngine((connected) => connected.getCell(address, sheet, rpcOptions)),
    getRange: async (range, sheet, rpcOptions) =>
      await withEngine((connected) => connected.getRange(range, sheet, rpcOptions)),
    setCell: async (address, value, sheet) => await withEngine((connected) => connected.setCell(address, value, sheet)),
    setCells: async (updates, rpcOptions) => await withEngine((connected) => connected.setCells(updates, rpcOptions)),
    setRange: async (range, values, sheet, rpcOptions) =>
      await withEngine((connected) => connected.setRange(range, values, sheet, rpcOptions)),
    recalculate: async (sheet, rpcOptions) => await withEngine((connected) => connected.recalculate(sheet, rpcOptions)),
    terminate: () => {
      generation++;
      enginePromise = null;
      engine?.terminate();
      engine = null;
      worker?.terminate();
      worker = null;
    }
  };
}
