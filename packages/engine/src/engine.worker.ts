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
  };
};

let port: MessagePort | null = null;
let wasmModuleUrl: string | null = null;
let wasmBinaryUrl: string | null = null;
let workbook: WasmWorkbookInstance | null = null;

let cancelledRequests = new Set<number>();

// Cancellation messages can arrive after the worker has already responded (e.g.
// main thread aborts while a response is in-flight). Track a bounded set of
// recently completed request IDs so late cancellations can be ignored without
// growing `cancelledRequests` forever.
const completedRequestIds = new Set<number>();
const completedRequestQueue: number[] = [];
const MAX_COMPLETED_REQUEST_IDS = 1024;

function markRequestCompleted(id: number): void {
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
  cancelledRequests.add(id);
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
  wasmModulePromise = loadWasmModule(moduleUrl);
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
      return;
    }

    const wb = await ensureWorkbook(wasmModuleUrl);

    if (consumeCancellation(id)) {
      return;
    }

    const params = req.params as any;
    let result: unknown;

    switch (req.method) {
      case "ping":
        result = "pong";
        break;
      case "newWorkbook":
        workbook = new mod.WasmWorkbook();
        result = null;
        break;
      case "loadFromJson":
        workbook = mod.WasmWorkbook.fromJson(params.json);
        result = null;
        break;
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
    void handleRequest(inner.data as WorkerInboundMessage);
  });
  port.start?.();

  postMessageToMain({ type: "ready" });
});
