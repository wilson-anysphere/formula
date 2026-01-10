/// <reference lib="webworker" />

import type {
  CellScalar,
  InitMessage,
  RpcCancel,
  RpcRequest,
  WorkerInboundMessage,
  WorkerOutboundMessage
} from "./protocol";

type WasmModule = {
  default?: () => Promise<void> | void;
  WasmWorkbook: {
    new (): {
      getCell(address: string, sheet?: string): unknown;
      setCell(address: string, value: CellScalar, sheet?: string): void;
      getRange(range: string, sheet?: string): unknown;
      setRange(range: string, values: CellScalar[][], sheet?: string): void;
      recalculate(sheet?: string): unknown;
      toJson(): string;
    };
    fromJson(json: string): any;
  };
};

let port: MessagePort | null = null;
let workbook: InstanceType<WasmModule["WasmWorkbook"]> | null = null;
let cancelled = new Set<number>();
let wasmModulePromise: Promise<WasmModule> | null = null;

async function loadWasmModule(moduleUrl: string): Promise<WasmModule> {
  // In Vite/Tauri dev the module URL may be a fully-qualified URL.
  // The `@vite-ignore` hint prevents Vite from trying to pre-bundle it.
  // eslint-disable-next-line @typescript-eslint/ban-ts-comment
  // @ts-ignore
  const mod = (await import(/* @vite-ignore */ moduleUrl)) as WasmModule;
  await mod.default?.();
  return mod;
}

function postMessageToMain(msg: WorkerOutboundMessage): void {
  port?.postMessage(msg);
}

async function ensureWorkbook(moduleUrl: string): Promise<void> {
  if (!wasmModulePromise) {
    wasmModulePromise = loadWasmModule(moduleUrl);
  }
  const mod = await wasmModulePromise;
  if (!workbook) {
    workbook = new mod.WasmWorkbook();
  }
}

async function handleRequest(
  moduleUrl: string,
  message: WorkerInboundMessage
): Promise<void> {
  if (message.type === "cancel") {
    cancelled.add((message as RpcCancel).id);
    return;
  }

  const req = message as RpcRequest;
  try {
    await ensureWorkbook(moduleUrl);

    if (cancelled.has(req.id)) {
      cancelled.delete(req.id);
      return;
    }

    if (!workbook) {
      throw new Error("workbook not initialized");
    }

    const params = req.params as any;
    let result: unknown;

    switch (req.method) {
      case "newWorkbook":
        workbook = new (await wasmModulePromise!).WasmWorkbook();
        result = null;
        break;
      case "loadFromJson": {
        const mod = await wasmModulePromise!;
        workbook = mod.WasmWorkbook.fromJson(params.json);
        result = null;
        break;
      }
      case "toJson":
        result = workbook.toJson();
        break;
      case "getCell":
        result = workbook.getCell(params.address, params.sheet);
        break;
      case "setCells":
        for (const update of params.updates as Array<any>) {
          workbook.setCell(update.address, update.value, update.sheet);
        }
        result = null;
        break;
      case "getRange":
        result = workbook.getRange(params.range, params.sheet);
        break;
      case "setRange":
        workbook.setRange(params.range, params.values, params.sheet);
        result = null;
        break;
      case "recalculate":
        result = workbook.recalculate(params.sheet);
        break;
      default:
        throw new Error(`unknown method: ${req.method}`);
    }

    postMessageToMain({ type: "response", id: req.id, ok: true, result });
  } catch (err) {
    postMessageToMain({
      type: "response",
      id: (message as RpcRequest).id,
      ok: false,
      error: err instanceof Error ? err.message : String(err)
    });
  }
}

self.addEventListener("message", (event: MessageEvent<unknown>) => {
  const msg = event.data as InitMessage;
  if (!msg || typeof msg !== "object" || (msg as any).type !== "init") {
    return;
  }

  port = msg.port;
  port.start?.();
  port.addEventListener("message", (inner: MessageEvent<unknown>) => {
    void handleRequest(msg.wasmModuleUrl, inner.data as WorkerInboundMessage);
  });

  postMessageToMain({ type: "ready" });
});
