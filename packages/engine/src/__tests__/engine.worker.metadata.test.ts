import { afterAll, describe, expect, it } from "vitest";

import type { InitMessage, RpcRequest, RpcResponseErr, RpcResponseOk, WorkerOutboundMessage } from "../protocol.ts";

class MockWorkerGlobal {
  private readonly listeners = new Set<(event: MessageEvent<unknown>) => void>();

  addEventListener(_type: "message", listener: (event: MessageEvent<unknown>) => void): void {
    this.listeners.add(listener);
  }

  dispatchMessage(data: unknown): void {
    const event = { data } as MessageEvent<unknown>;
    for (const listener of this.listeners) {
      listener(event);
    }
  }
}

class MockMessagePort {
  onmessage: ((event: MessageEvent<unknown>) => void) | null = null;
  private listeners = new Set<(event: MessageEvent<unknown>) => void>();
  private other: MockMessagePort | null = null;

  connect(other: MockMessagePort) {
    this.other = other;
  }

  postMessage(message: unknown): void {
    queueMicrotask(() => {
      this.other?.dispatchMessage(message);
    });
  }

  start(): void {}

  close(): void {
    this.listeners.clear();
    this.onmessage = null;
    this.other = null;
  }

  addEventListener(_type: "message", listener: (event: MessageEvent<unknown>) => void): void {
    this.listeners.add(listener);
  }

  removeEventListener(_type: "message", listener: (event: MessageEvent<unknown>) => void): void {
    this.listeners.delete(listener);
  }

  private dispatchMessage(data: unknown): void {
    const event = { data } as MessageEvent<unknown>;
    this.onmessage?.(event);
    for (const listener of this.listeners) {
      listener(event);
    }
  }
}

function createMockChannel(): { port1: MockMessagePort; port2: MessagePort } {
  const port1 = new MockMessagePort();
  const port2 = new MockMessagePort();
  port1.connect(port2);
  port2.connect(port1);
  return { port1, port2: port2 as unknown as MessagePort };
}

async function waitForMessage(
  port: MockMessagePort,
  predicate: (msg: WorkerOutboundMessage) => boolean
): Promise<WorkerOutboundMessage> {
  return await new Promise((resolve) => {
    const handler = (event: MessageEvent<unknown>) => {
      const msg = event.data as WorkerOutboundMessage;
      if (msg && typeof msg === "object" && predicate(msg)) {
        port.removeEventListener("message", handler);
        resolve(msg);
      }
    };
    port.addEventListener("message", handler);
  });
}

async function setupWorker(options: { wasmModuleUrl: string }) {
  await loadWorkerModule();

  const channel = createMockChannel();
  const init: InitMessage = {
    type: "init",
    port: channel.port2,
    wasmModuleUrl: options.wasmModuleUrl
  };
  workerGlobal.dispatchMessage(init);

  await waitForMessage(channel.port1, (msg) => msg.type === "ready");

  return {
    port: channel.port1,
    dispose: () => channel.port1.close()
  };
}

async function sendRequest(
  port: MockMessagePort,
  req: RpcRequest
): Promise<RpcResponseOk | RpcResponseErr> {
  const responsePromise = waitForMessage(port, (msg) => msg.type === "response" && msg.id === req.id) as Promise<
    RpcResponseOk | RpcResponseErr
  >;
  port.postMessage(req);
  return await responsePromise;
}

describe("engine.worker workbook metadata RPCs", () => {
  afterAll(() => {
    (globalThis as any).self = previousSelf;
  });

  it("dispatches workbook metadata setters to the underlying wasm workbook", async () => {
    (globalThis as any).__ENGINE_WORKER_TEST_CALLS__ = [];
    const wasmModuleUrl = new URL("./fixtures/mockWasmWorkbookMetadata.mjs", import.meta.url).href;
    const { port, dispose } = await setupWorker({ wasmModuleUrl });

    try {
      await sendRequest(port, { type: "request", id: 0, method: "newWorkbook", params: {} });
      let resp = await sendRequest(port, {
        type: "request",
        id: 1,
        method: "setWorkbookFileMetadata",
        params: { directory: "/tmp", filename: "book.xlsx" }
      });
      expect(resp.ok).toBe(true);

      resp = await sendRequest(port, {
        type: "request",
        id: 2,
        method: "setCellStyleId",
        params: { sheet: "Sheet1", address: "A1", styleId: 7 }
      });
      expect(resp.ok).toBe(true);

      resp = await sendRequest(port, {
        type: "request",
        id: 3,
        method: "setColWidth",
        params: { sheet: "Sheet1", col: 2, width: 120 }
      });
      expect(resp.ok).toBe(true);

      resp = await sendRequest(port, {
        type: "request",
        id: 4,
        method: "setColHidden",
        params: { sheet: "Sheet1", col: 2, hidden: true }
      });
      expect(resp.ok).toBe(true);

      resp = await sendRequest(port, {
        type: "request",
        id: 5,
        method: "internStyle",
        params: { style: { font: { bold: true } } }
      });
      expect(resp.ok).toBe(true);
      expect((resp as RpcResponseOk).result).toBe(42);

      expect((globalThis as any).__ENGINE_WORKER_TEST_CALLS__).toEqual([
        ["setWorkbookFileMetadata", "/tmp", "book.xlsx"],
        ["setCellStyleId", "Sheet1", "A1", 7],
        ["setColWidth", "Sheet1", 2, 120],
        ["setColHidden", "Sheet1", 2, true],
        ["internStyle", { font: { bold: true } }]
      ]);
    } finally {
      dispose();
      delete (globalThis as any).__ENGINE_WORKER_TEST_CALLS__;
    }
  });

  it("returns a clear error when the wasm workbook does not support a metadata method", async () => {
    const wasmModuleUrl = new URL("./fixtures/mockWasmWorkbookNoMetadata.mjs", import.meta.url).href;
    const { port, dispose } = await setupWorker({ wasmModuleUrl });

    try {
      await sendRequest(port, { type: "request", id: 0, method: "newWorkbook", params: {} });
      const resp = await sendRequest(port, {
        type: "request",
        id: 1,
        method: "setColHidden",
        params: { sheet: "Sheet1", col: 0, hidden: true }
      });
      expect(resp.ok).toBe(false);
      expect((resp as RpcResponseErr).error).toMatch(/setColHidden: WasmWorkbook\.setColHidden is not available/i);
    } finally {
      dispose();
    }
  });
});

const previousSelf = (globalThis as any).self;
const workerGlobal = new MockWorkerGlobal();
// `engine.worker.ts` expects a WebWorker-like `self`.
(globalThis as any).self = workerGlobal;

let workerModulePromise: Promise<unknown> | null = null;
function loadWorkerModule(): Promise<unknown> {
  if (!workerModulePromise) {
    workerModulePromise = import("../engine.worker.ts");
  }
  return workerModulePromise;
}
