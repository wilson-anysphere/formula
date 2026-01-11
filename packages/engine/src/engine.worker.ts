/// <reference lib="webworker" />

import type {
  CellScalar,
  InitMessage,
  RpcCancel,
  RpcRequest,
  WorkerInboundMessage,
  WorkerOutboundMessage,
} from "./protocol";

type WasmWorkbookInstance = {
  getCell(address: string, sheet?: string): unknown;
  setCell(address: string, value: CellScalar, sheet?: string): void;
  getRange(range: string, sheet?: string): unknown;
  setRange(range: string, values: CellScalar[][], sheet?: string): void;
  recalculate(sheet?: string): unknown;
  toJson(): string;
};

type WasmModule = {
  default?: (module_or_path?: unknown) => Promise<void> | void;
  WasmWorkbook: {
    new (): WasmWorkbookInstance;
    fromJson(json: string): WasmWorkbookInstance;
    fromXlsxBytes?: (bytes: Uint8Array) => WasmWorkbookInstance;
  };
};

let port: MessagePort | null = null;
let wasmModuleUrl: string | null = null;
let wasmBinaryUrl: string | null = null;
let workbook: WasmWorkbookInstance | null = null;

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
            throw new Error("loadFromXlsxBytes: expected params.bytes to be a Uint8Array/ArrayBuffer/ArrayBufferView");
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
              result = wb.getCell(params.address, params.sheet);
              break;
            case "getRange":
              result = wb.getRange(params.range, params.sheet);
              break;
            case "setCells":
              for (const update of params.updates as Array<any>) {
                wb.setCell(update.address, update.value, update.sheet);
              }
              result = null;
              break;
            case "setRange":
              wb.setRange(params.range, params.values, params.sheet);
              result = null;
              break;
            case "recalculate":
              result = wb.recalculate(params.sheet);
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
