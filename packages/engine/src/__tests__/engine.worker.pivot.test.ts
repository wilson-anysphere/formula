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
  private listeners = new Map<string, Set<(event: MessageEvent<unknown>) => void>>();
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

  addEventListener(type: string, listener: (event: MessageEvent<unknown>) => void): void {
    const key = String(type ?? "");
    let set = this.listeners.get(key);
    if (!set) {
      set = new Set();
      this.listeners.set(key, set);
    }
    set.add(listener);
  }

  removeEventListener(type: string, listener: (event: MessageEvent<unknown>) => void): void {
    const key = String(type ?? "");
    this.listeners.get(key)?.delete(listener);
  }

  private dispatchMessage(data: unknown): void {
    const event = { data } as MessageEvent<unknown>;
    this.onmessage?.(event);
    for (const listener of this.listeners.get("message") ?? []) {
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
  predicate: (msg: WorkerOutboundMessage) => boolean,
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

async function sendRequest(port: MockMessagePort, req: RpcRequest): Promise<RpcResponseOk | RpcResponseErr> {
  const responsePromise = waitForMessage(port, (msg) => msg.type === "response" && msg.id === req.id) as Promise<
    RpcResponseOk | RpcResponseErr
  >;
  port.postMessage(req);
  return await responsePromise;
}

async function setupWorker(options: { wasmModuleUrl: string }) {
  await loadWorkerModule();

  const channel = createMockChannel();
  const init: InitMessage = {
    type: "init",
    port: channel.port2,
    wasmModuleUrl: options.wasmModuleUrl,
  };
  workerGlobal.dispatchMessage(init);

  await waitForMessage(channel.port1, (msg) => msg.type === "ready");
  return { port: channel.port1, dispose: () => channel.port1.close() };
}

describe("engine.worker pivot RPC normalization", () => {
  afterAll(() => {
    (globalThis as any).self = previousSelf;
  });

  it("normalizes calculatePivot writes by converting undefined values to null", async () => {
    (globalThis as any).__ENGINE_WORKER_TEST_CALLS__ = [];
    const wasmModuleUrl = new URL("./fixtures/mockWasmWorkbookPivot.mjs", import.meta.url).href;
    const { port, dispose } = await setupWorker({ wasmModuleUrl });

    try {
      await sendRequest(port, { type: "request", id: 0, method: "newWorkbook", params: {} });

      const resp = await sendRequest(port, {
        type: "request",
        id: 1,
        method: "calculatePivot",
        params: {
          sheet: "Sheet1",
          sourceRangeA1: "A1:B2",
          destinationTopLeftA1: "D1",
          // The mocked wasm workbook ignores the config, but the worker always forwards it.
          config: {},
        },
      });

      expect(resp.ok).toBe(true);
      const payload = (resp as RpcResponseOk).result as any;
      expect(payload).toHaveProperty("writes");
      const d1 = (payload.writes as any[]).find((c) => c.sheet === "Sheet1" && c.address === "D1");
      expect(d1).toBeTruthy();
      expect(d1.value).toBeNull();

      expect((globalThis as any).__ENGINE_WORKER_TEST_CALLS__).toEqual([["calculatePivot", "Sheet1"]]);
    } finally {
      dispose();
      delete (globalThis as any).__ENGINE_WORKER_TEST_CALLS__;
    }
  });

  it("trims sheet names for pivot RPCs", async () => {
    (globalThis as any).__ENGINE_WORKER_TEST_CALLS__ = [];
    const wasmModuleUrl = new URL("./fixtures/mockWasmWorkbookPivot.mjs", import.meta.url).href;
    const { port, dispose } = await setupWorker({ wasmModuleUrl });

    try {
      await sendRequest(port, { type: "request", id: 0, method: "newWorkbook", params: {} });

      const resp = await sendRequest(port, {
        type: "request",
        id: 1,
        method: "calculatePivot",
        params: {
          sheet: "  Sheet1  ",
          sourceRangeA1: "A1:B2",
          destinationTopLeftA1: "D1",
          config: {},
        },
      });

      expect(resp.ok).toBe(true);
      expect((globalThis as any).__ENGINE_WORKER_TEST_CALLS__).toEqual([["calculatePivot", "Sheet1"]]);
    } finally {
      dispose();
      delete (globalThis as any).__ENGINE_WORKER_TEST_CALLS__;
    }
  });

  it("trims and defaults sheet names for getPivotSchema/getPivotFieldItems/getPivotFieldItemsPaged", async () => {
    (globalThis as any).__ENGINE_WORKER_TEST_CALLS__ = [];
    const wasmModuleUrl = new URL("./fixtures/mockWasmWorkbookPivot.mjs", import.meta.url).href;
    const { port, dispose } = await setupWorker({ wasmModuleUrl });

    try {
      await sendRequest(port, { type: "request", id: 0, method: "newWorkbook", params: {} });

      let resp = await sendRequest(port, {
        type: "request",
        id: 1,
        method: "getPivotSchema",
        params: { sheet: "  Sheet1  ", sourceRangeA1: "A1:B2", sampleSize: 10 },
      });
      expect(resp.ok).toBe(true);

      resp = await sendRequest(port, {
        type: "request",
        id: 2,
        method: "getPivotFieldItems",
        params: { sheet: "   ", sourceRangeA1: "A1:B2", field: "Category" },
      });
      expect(resp.ok).toBe(true);

      resp = await sendRequest(port, {
        type: "request",
        id: 3,
        method: "getPivotFieldItemsPaged",
        params: { sheet: "  Sheet2  ", sourceRangeA1: "A1:B2", field: "Category", offset: 0, limit: 5 },
      });
      expect(resp.ok).toBe(true);

      expect((globalThis as any).__ENGINE_WORKER_TEST_CALLS__).toEqual([
        ["getPivotSchema", "Sheet1", "A1:B2", 10],
        ["getPivotFieldItems", "Sheet1", "A1:B2", "Category"],
        ["getPivotFieldItemsPaged", "Sheet2", "A1:B2", "Category", 0, 5],
      ]);
    } finally {
      dispose();
      delete (globalThis as any).__ENGINE_WORKER_TEST_CALLS__;
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
