/// <reference lib="webworker" />

import type {
  CellScalar,
  InitMessage,
  RpcCancel,
  RpcRequest,
  WorkerInboundMessage,
  WorkerOutboundMessage,
} from "./protocol";

type EngineRequest = {
  id: number;
  method: "init" | "ping";
  params?: unknown;
};

type EngineResponse =
  | { id: number; ok: true; result: unknown }
  | { id: number; ok: false; error: string };

let initialized = false;

async function handleEngine(method: EngineRequest["method"], _params?: unknown) {
  switch (method) {
    case "init": {
      initialized = true;
      return;
    }
    case "ping": {
      return initialized ? "pong" : "not-initialized";
    }
  }
}

async function respondToEngineRequest(request: EngineRequest): Promise<void> {
  const { id, method, params } = request;

  try {
    const result = await handleEngine(method, params);
    const response: EngineResponse = { id, ok: true, result };
    self.postMessage(response);
  } catch (error) {
    const response: EngineResponse = {
      id,
      ok: false,
      error: error instanceof Error ? error.message : String(error),
    };
    self.postMessage(response);
  }
}

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

let wasmModulePromise: Promise<WasmModule> | null = null;
let wasmModulePromiseUrl: string | null = null;

async function loadWasmModule(moduleUrl: string): Promise<WasmModule> {
  // Vite will try to pre-bundle dynamic imports unless explicitly told not to.
  // The `@vite-ignore` hint prevents Vite from trying to pre-bundle arbitrary URLs.
  // eslint-disable-next-line @typescript-eslint/ban-ts-comment
  // @ts-ignore - `moduleUrl` is runtime-defined (Vite dev server / asset URL).
  const mod = (await import(/* @vite-ignore */ moduleUrl)) as WasmModule;
  await mod.default?.(wasmBinaryUrl ?? undefined);
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
    cancelledRequests.add((message as RpcCancel).id);
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
    return;
  }

  if (consumeCancellation(id)) {
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
      return;
    }

    postMessageToMain({ type: "response", id, ok: true, result });
  } catch (err) {
    if (isCancelled(id)) {
      cancelledRequests.delete(id);
      return;
    }

    postMessageToMain({
      type: "response",
      id,
      ok: false,
      error: err instanceof Error ? err.message : String(err),
    });
  }
}

function isEngineRequest(data: unknown): data is EngineRequest {
  return (
    !!data &&
    typeof data === "object" &&
    "id" in data &&
    typeof (data as any).id === "number" &&
    "method" in data &&
    ((data as any).method === "init" || (data as any).method === "ping")
  );
}

self.addEventListener("message", (event: MessageEvent<unknown>) => {
  const data = event.data;

  if (isEngineRequest(data)) {
    void respondToEngineRequest(data);
    return;
  }

  const msg = data as InitMessage;
  if (!msg || typeof msg !== "object" || (msg as any).type !== "init") {
    return;
  }

  port = msg.port;
  wasmModuleUrl = msg.wasmModuleUrl;
  wasmBinaryUrl = msg.wasmBinaryUrl ?? null;

  port.start?.();
  port.addEventListener("message", (inner: MessageEvent<unknown>) => {
    void handleRequest(inner.data as WorkerInboundMessage);
  });

  postMessageToMain({ type: "ready" });
});
