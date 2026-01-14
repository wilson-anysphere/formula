/// <reference lib="webworker" />

import type {
  CellScalar,
  FormulaParseOptions,
  InitMessage,
  RpcCancel,
  RpcRequest,
  WorkerInboundMessage,
  WorkerOutboundMessage,
} from "./protocol.ts";

type WasmWorkbookInstance = {
  getCell(address: string, sheet?: string): unknown;
  getCellRich?: (address: string, sheet?: string) => unknown;
  goalSeek?: (request: unknown) => unknown;
  getPivotSchema?: (sheet: string, sourceRangeA1: string, sampleSize?: number) => unknown;
  calculatePivot?: (
    sheet: string,
    sourceRangeA1: string,
    destinationTopLeftA1: string,
    config: unknown
  ) => unknown;
  setCell(address: string, value: CellScalar, sheet?: string): void;
  setCellRich?: (address: string, value: unknown, sheet?: string) => void;
  setCells?: (updates: Array<{ address: string; value: CellScalar; sheet?: string }>) => void;
  setLocale?: (localeId: string) => boolean;
  getRange(range: string, sheet?: string): unknown;
  setRange(range: string, values: CellScalar[][], sheet?: string): void;
  recalculate(sheet?: string): unknown;
  applyOperation?: (op: unknown) => unknown;
  setSheetDimensions?: (sheet: string, rows: number, cols: number) => void;
  getSheetDimensions?: (sheet: string) => { rows: number; cols: number };
  setWorkbookFileMetadata?: (directory: string | null, filename: string | null) => void;
  setCellStyleId?: (sheet: string, address: string, styleId: number) => void;
  setColWidth?: (sheet: string, col: number, width: number | null) => void;
  setColHidden?: (sheet: string, col: number, hidden: boolean) => void;
  internStyle?: (style: unknown) => number;
  toJson(): string;
};

type WasmModule = {
  default?: (module_or_path?: unknown) => Promise<void> | void;
  lexFormula: (formula: string, options?: FormulaParseOptions) => unknown;
  parseFormulaPartial: (formula: string, cursor?: number, options?: FormulaParseOptions) => unknown;
  canonicalizeFormula?: (formula: string, localeId: string, referenceStyle?: "A1" | "R1C1") => string;
  localizeFormula?: (formula: string, localeId: string, referenceStyle?: "A1" | "R1C1") => string;
  rewriteFormulasForCopyDelta?: (requests: unknown) => unknown;
  WasmWorkbook: {
    new (): WasmWorkbookInstance;
    fromJson(json: string): WasmWorkbookInstance;
    fromXlsxBytes?: (bytes: Uint8Array) => WasmWorkbookInstance;
  };
  lexFormulaPartial?: (formula: string, options?: unknown) => unknown;
};

let port: MessagePort | null = null;
let wasmModuleUrl: string | null = null;
let wasmBinaryUrl: string | null = null;
let workbook: WasmWorkbookInstance | null = null;

function normalizeCellScalar(value: unknown): CellScalar {
  // wasm-bindgen maps `Option<T>` to `T | undefined` in JS. Our public protocol uses `null`
  // for empty cells, so normalize `undefined` at the worker boundary.
  if (value === undefined) return null;
  if (value === null) return null;
  if (typeof value === "string" || typeof value === "number" || typeof value === "boolean") return value;
  return null;
}

function normalizeCellData(value: unknown): unknown {
  if (!value || typeof value !== "object") return value;
  const cell = value as any;
  if (!("input" in cell) && !("value" in cell)) return value;
  return { ...cell, input: normalizeCellScalar(cell.input), value: normalizeCellScalar(cell.value) };
}

function normalizeRangeData(value: unknown): unknown {
  if (!Array.isArray(value)) return value;
  return value.map((row) => (Array.isArray(row) ? row.map((cell) => normalizeCellData(cell)) : row));
}

function normalizeCellChanges(value: unknown): unknown {
  if (!Array.isArray(value)) return value;
  return value.map((change) => {
    if (!change || typeof change !== "object") return change;
    const obj = change as any;
    if (!("value" in obj)) return change;
    return { ...obj, value: normalizeCellScalar(obj.value) };
  });
}

function normalizePivotCalculation(value: unknown): unknown {
  if (!value || typeof value !== "object") return value;
  const obj = value as any;
  if (!Array.isArray(obj.writes)) return value;
  return { ...obj, writes: normalizeCellChanges(obj.writes) };
}

function cloneToPlainData(value: unknown): unknown {
  // wasm-bindgen APIs can return objects with prototypes that are structured-clone
  // safe but not ideal for RPC consumers. Normalize editor-tooling results into
  // plain objects/arrays at the worker boundary.
  if (value === undefined) return null;
  try {
    return JSON.parse(JSON.stringify(value));
  } catch {
    return value;
  }
}

let cancelledRequests = new Set<number>();
const pendingRequestIds = new Set<number>();

// Cancels can arrive for request IDs that will never be sent (e.g. abort signal
// already fired before the request message is posted). Track those separately
// in a bounded structure so they don't leak forever.
const preCancelledRequestIds = new Set<number>();
const preCancelledRequestQueue: number[] = [];
const MAX_PRE_CANCELLED_REQUEST_IDS = 4096;

// Cancellation messages can arrive after the worker has already responded (e.g.
// main thread aborts while a response is in-flight). Track a bounded set of
// recently completed request IDs so late cancellations can be ignored without
// growing `cancelledRequests` forever.
const completedRequestIds = new Set<number>();
const completedRequestQueue: number[] = [];
const MAX_COMPLETED_REQUEST_IDS = 1024;

function markRequestCompleted(id: number): void {
  pendingRequestIds.delete(id);
  cancelledRequests.delete(id);
  preCancelledRequestIds.delete(id);

  completedRequestIds.add(id);
  completedRequestQueue.push(id);
  if (completedRequestQueue.length > MAX_COMPLETED_REQUEST_IDS) {
    const oldest = completedRequestQueue.shift();
    if (oldest != null) {
      completedRequestIds.delete(oldest);
    }
  }
}

function trackCancellation(id: number): void {
  if (completedRequestIds.has(id)) {
    return;
  }

  if (pendingRequestIds.has(id)) {
    cancelledRequests.add(id);
    return;
  }

  if (preCancelledRequestIds.has(id)) {
    return;
  }

  preCancelledRequestIds.add(id);
  preCancelledRequestQueue.push(id);
  if (preCancelledRequestQueue.length > MAX_PRE_CANCELLED_REQUEST_IDS) {
    const oldest = preCancelledRequestQueue.shift();
    if (oldest != null) {
      preCancelledRequestIds.delete(oldest);
    }
  }
}

function freeWorkbook(instance: WasmWorkbookInstance | null): void {
  // wasm-bindgen classes expose an eager `free()` API. Prefer it so `newWorkbook`
  // / `loadFromJson` don't rely on GC timing to release WASM allocations.
  try {
    (instance as any)?.free?.();
  } catch {
    // Ignore failures; worst case the object is left for GC/finalization.
  }
}

let wasmModulePromise: Promise<WasmModule> | null = null;
let wasmModulePromiseUrl: string | null = null;

async function loadWasmModule(moduleUrl: string): Promise<WasmModule> {
  // Vite will try to pre-bundle dynamic imports unless explicitly told not to.
  // The `@vite-ignore` hint prevents Vite from trying to pre-bundle arbitrary URLs.
  // eslint-disable-next-line @typescript-eslint/ban-ts-comment
  // @ts-ignore - `moduleUrl` is runtime-defined (Vite dev server / asset URL).
  const mod = (await import(/* @vite-ignore */ moduleUrl)) as WasmModule;
  const init = mod.default;
  if (init) {
    if (wasmBinaryUrl) {
      // wasm-bindgen >=0.2.105 prefers an object parameter, but older versions
      // accepted `module_or_path` directly. Try the modern form first to avoid
      // a noisy console warning, then fall back for compatibility.
      try {
        await init({ module_or_path: wasmBinaryUrl });
      } catch {
        await init(wasmBinaryUrl);
      }
    } else {
      await init();
    }
  }
  return mod;
}

function getWasmModule(moduleUrl: string): Promise<WasmModule> {
  if (wasmModulePromise && wasmModulePromiseUrl === moduleUrl) {
    return wasmModulePromise;
  }

  wasmModulePromiseUrl = moduleUrl;
  wasmModulePromise = loadWasmModule(moduleUrl).catch((err) => {
    // If initialization fails (e.g. transient network error during dev), allow
    // future requests to retry by clearing the cached promise.
    wasmModulePromise = null;
    wasmModulePromiseUrl = null;
    throw err;
  });
  return wasmModulePromise;
}

async function ensureWorkbook(moduleUrl: string): Promise<WasmWorkbookInstance> {
  const mod = await getWasmModule(moduleUrl);
  if (!workbook) {
    workbook = new mod.WasmWorkbook();
  }
  return workbook;
}

function postMessageToMain(msg: WorkerOutboundMessage): void {
  port?.postMessage(msg);
}

function isCancelled(id: number): boolean {
  return cancelledRequests.has(id);
}

function consumeCancellation(id: number): boolean {
  if (!cancelledRequests.has(id)) {
    return false;
  }
  cancelledRequests.delete(id);
  return true;
}

async function handleRequest(message: WorkerInboundMessage): Promise<void> {
  if (message.type === "cancel") {
    trackCancellation((message as RpcCancel).id);
    return;
  }

  const req = message as RpcRequest;
  const id = req.id;

  if (!wasmModuleUrl) {
    postMessageToMain({
      type: "response",
      id,
      ok: false,
      error: "worker not initialized",
    });
    markRequestCompleted(id);
    return;
  }

  if (consumeCancellation(id)) {
    markRequestCompleted(id);
    return;
  }

  try {
    const mod = await getWasmModule(wasmModuleUrl);

    if (consumeCancellation(id)) {
      markRequestCompleted(id);
      return;
    }

    const params = req.params as any;
    let result: unknown;

    switch (req.method) {
      case "ping":
        result = "pong";
        break;
      case "lexFormula":
        {
          const lexFormula = mod.lexFormula;
          if (typeof lexFormula !== "function") {
            throw new Error("lexFormula: wasm module does not export lexFormula()");
          }
          result = cloneToPlainData(lexFormula(params.formula, params.options));
        }
        break;
      case "canonicalizeFormula":
        {
          const canonicalizeFormula = mod.canonicalizeFormula;
          if (typeof canonicalizeFormula !== "function") {
            throw new Error("canonicalizeFormula: wasm module does not export canonicalizeFormula()");
          }
          result = canonicalizeFormula(params.formula, params.localeId, params.referenceStyle);
        }
        break;
      case "localizeFormula":
        {
          const localizeFormula = mod.localizeFormula;
          if (typeof localizeFormula !== "function") {
            throw new Error("localizeFormula: wasm module does not export localizeFormula()");
          }
          result = localizeFormula(params.formula, params.localeId, params.referenceStyle);
        }
        break;
      case "lexFormulaPartial":
        {
          const lexFormulaPartial = mod.lexFormulaPartial;
          if (typeof lexFormulaPartial !== "function") {
            throw new Error("lexFormulaPartial: wasm module does not export lexFormulaPartial()");
          }
          result = cloneToPlainData(lexFormulaPartial(params.formula, params.options));
        }
        break;
      case "parseFormulaPartial":
        {
          const parseFormulaPartial = mod.parseFormulaPartial;
          if (typeof parseFormulaPartial !== "function") {
            throw new Error("parseFormulaPartial: wasm module does not export parseFormulaPartial()");
          }
          result = cloneToPlainData(parseFormulaPartial(params.formula, params.cursor, params.options));
        }
        break;
      case "rewriteFormulasForCopyDelta":
        {
          const rewrite = mod.rewriteFormulasForCopyDelta;
          if (typeof rewrite !== "function") {
            throw new Error("rewriteFormulasForCopyDelta: wasm module does not export rewriteFormulasForCopyDelta()");
          }
          // This RPC can return large arrays (e.g. paste/fill), so avoid JSON clone overhead.
          result = rewrite(params.requests);
        }
        break;
      case "newWorkbook":
        {
          const next = new mod.WasmWorkbook();
          freeWorkbook(workbook);
          workbook = next;
        }
        result = null;
        break;
      case "loadFromJson":
        {
          const next = mod.WasmWorkbook.fromJson(params.json);
          freeWorkbook(workbook);
          workbook = next;
        }
        result = null;
        break;
      case "loadFromXlsxBytes":
        {
          const rawBytes = params.bytes as unknown;
          let bytes: Uint8Array;
          if (rawBytes instanceof Uint8Array) {
            bytes = rawBytes;
          } else if (rawBytes instanceof ArrayBuffer) {
            bytes = new Uint8Array(rawBytes);
          } else if (ArrayBuffer.isView(rawBytes) && rawBytes.buffer instanceof ArrayBuffer) {
            bytes = new Uint8Array(rawBytes.buffer, rawBytes.byteOffset, rawBytes.byteLength);
          } else {
            throw new Error(
              "loadFromXlsxBytes: expected params.bytes to be a Uint8Array/ArrayBuffer/ArrayBufferView"
            );
          }

          const fromXlsxBytes = mod.WasmWorkbook.fromXlsxBytes;
          if (typeof fromXlsxBytes !== "function") {
            throw new Error("loadFromXlsxBytes: WasmWorkbook.fromXlsxBytes is not available in this WASM build");
          }

          const next = fromXlsxBytes(bytes);
          freeWorkbook(workbook);
          workbook = next;
        }
        result = null;
        break;
      default:
        {
          const wb = await ensureWorkbook(wasmModuleUrl);

          if (consumeCancellation(id)) {
            markRequestCompleted(id);
            return;
          }

          switch (req.method) {
            case "toJson":
              result = wb.toJson();
              break;
            case "getCell":
              result = normalizeCellData(wb.getCell(params.address, params.sheet));
              break;
            case "getCellRich":
              if (typeof (wb as any).getCellRich !== "function") {
                throw new Error("getCellRich: WasmWorkbook.getCellRich is not available in this WASM build");
              }
              result = cloneToPlainData((wb as any).getCellRich(params.address, params.sheet));
              break;
            case "getRange":
              result = normalizeRangeData(wb.getRange(params.range, params.sheet));
              break;
            case "setSheetDimensions":
              if (typeof (wb as any).setSheetDimensions !== "function") {
                throw new Error("setSheetDimensions: not available in this WASM build");
              }
              (wb as any).setSheetDimensions(params.sheet, params.rows, params.cols);
              result = null;
              break;
            case "getSheetDimensions":
              if (typeof (wb as any).getSheetDimensions !== "function") {
                throw new Error("getSheetDimensions: not available in this WASM build");
              }
              result = (wb as any).getSheetDimensions(params.sheet);
              break;
            case "setCells":
              if (typeof (wb as any).setCells === "function") {
                (wb as any).setCells(params.updates);
              } else {
                for (const update of params.updates as Array<any>) {
                  wb.setCell(update.address, update.value, update.sheet);
                }
              }
              result = null;
              break;
            case "setCellRich":
              if (typeof (wb as any).setCellRich !== "function") {
                throw new Error("setCellRich: WasmWorkbook.setCellRich is not available in this WASM build");
              }
              (wb as any).setCellRich(params.address, params.value, params.sheet);
              result = null;
              break;
            case "setRange":
              wb.setRange(params.range, params.values, params.sheet);
              result = null;
              break;
            case "setLocale":
              if (typeof (wb as any).setLocale === "function") {
                result = (wb as any).setLocale(params.localeId);
              } else {
                result = false;
              }
              break;
            case "setWorkbookFileMetadata":
              if (typeof (wb as any).setWorkbookFileMetadata !== "function") {
                throw new Error(
                  "setWorkbookFileMetadata: WasmWorkbook.setWorkbookFileMetadata is not available in this WASM build"
                );
              }
              (wb as any).setWorkbookFileMetadata(params.directory ?? null, params.filename ?? null);
              result = null;
              break;
            case "setCellStyleId":
              if (typeof (wb as any).setCellStyleId !== "function") {
                throw new Error("setCellStyleId: WasmWorkbook.setCellStyleId is not available in this WASM build");
              }
              (wb as any).setCellStyleId(params.sheet, params.address, params.styleId);
              result = null;
              break;
            case "setColWidth":
              if (typeof (wb as any).setColWidth !== "function") {
                throw new Error("setColWidth: WasmWorkbook.setColWidth is not available in this WASM build");
              }
              (wb as any).setColWidth(params.sheet, params.col, params.width ?? null);
              result = null;
              break;
            case "setColHidden":
              if (typeof (wb as any).setColHidden !== "function") {
                throw new Error("setColHidden: WasmWorkbook.setColHidden is not available in this WASM build");
              }
              (wb as any).setColHidden(params.sheet, params.col, Boolean(params.hidden));
              result = null;
              break;
            case "internStyle":
              if (typeof (wb as any).internStyle !== "function") {
                throw new Error("internStyle: WasmWorkbook.internStyle is not available in this WASM build");
              }
              result = (wb as any).internStyle(params.style);
              break;
            case "recalculate":
              result = normalizeCellChanges(wb.recalculate(params.sheet));
              break;
            case "applyOperation":
              if (typeof (wb as any).applyOperation === "function") {
                result = cloneToPlainData((wb as any).applyOperation(params.op));
              } else {
                throw new Error("applyOperation: WasmWorkbook.applyOperation is not available in this WASM build");
              }
              break;
            case "goalSeek":
              if (typeof (wb as any).goalSeek !== "function") {
                throw new Error("goalSeek: WasmWorkbook.goalSeek is not available in this WASM build");
              }
              result = cloneToPlainData((wb as any).goalSeek(params));
              break;
            case "getPivotSchema":
              if (typeof (wb as any).getPivotSchema !== "function") {
                throw new Error("getPivotSchema: WasmWorkbook.getPivotSchema is not available in this WASM build");
              }
              result = cloneToPlainData((wb as any).getPivotSchema(params.sheet, params.sourceRangeA1, params.sampleSize));
              break;
            case "calculatePivot":
              if (typeof (wb as any).calculatePivot !== "function") {
                throw new Error("calculatePivot: WasmWorkbook.calculatePivot is not available in this WASM build");
              }
              result = normalizePivotCalculation(
                cloneToPlainData(
                  (wb as any).calculatePivot(
                    params.sheet,
                    params.sourceRangeA1,
                    params.destinationTopLeftA1,
                    params.config
                  )
                )
              );
              break;
            default:
              throw new Error(`unknown method: ${req.method}`);
          }
        }
    }

    if (isCancelled(id)) {
      // Cancellation might arrive after the request starts; we still perform the work
      // but suppress the response.
      cancelledRequests.delete(id);
      markRequestCompleted(id);
      return;
    }

    postMessageToMain({ type: "response", id, ok: true, result });
    markRequestCompleted(id);
  } catch (err) {
    if (isCancelled(id)) {
      cancelledRequests.delete(id);
      markRequestCompleted(id);
      return;
    }

    postMessageToMain({
      type: "response",
      id,
      ok: false,
      error: err instanceof Error ? err.message : String(err),
    });
    markRequestCompleted(id);
  }
}

function isWorkerInboundMessage(data: unknown): data is WorkerInboundMessage {
  if (!data || typeof data !== "object" || !("type" in data)) {
    return false;
  }

  const type = (data as any).type;
  if (type !== "request" && type !== "cancel") {
    return false;
  }

  if (typeof (data as any).id !== "number") {
    return false;
  }

  if (type === "request") {
    return typeof (data as any).method === "string";
  }

  return true;
}

let requestQueue: Promise<void> = Promise.resolve();

self.addEventListener("message", (event: MessageEvent<unknown>) => {
  const data = event.data;

  const msg = data as InitMessage;
  if (!msg || typeof msg !== "object" || (msg as any).type !== "init") {
    return;
  }

  port = msg.port;
  wasmModuleUrl = msg.wasmModuleUrl;
  wasmBinaryUrl = msg.wasmBinaryUrl ?? null;

  port.addEventListener("message", (inner: MessageEvent<unknown>) => {
    const inbound = inner.data;
    if (!isWorkerInboundMessage(inbound)) {
      return;
    }

    if (inbound.type === "cancel") {
      // Handle cancels immediately so in-flight requests can observe cancellation.
      trackCancellation(inbound.id);
      return;
    }

    pendingRequestIds.add(inbound.id);
    if (preCancelledRequestIds.has(inbound.id)) {
      preCancelledRequestIds.delete(inbound.id);
      cancelledRequests.add(inbound.id);
    }

    // Serialize request processing to avoid interleaving workbook mutations.
    requestQueue = requestQueue
      .then(() => handleRequest(inbound))
      .catch(() => {
        // `handleRequest` should catch and respond to all errors, but if something
        // escapes we don't want to wedge the queue.
      });
  });
  port.start?.();

  postMessageToMain({ type: "ready" });
});
