import { afterAll, describe, expect, it } from "vitest";

import type { InitMessage, RpcRequest, WorkerOutboundMessage, WorkbookSheetInfoDto } from "../protocol.ts";

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

async function sendRequest(port: MockMessagePort, req: RpcRequest): Promise<any> {
  const responsePromise = waitForMessage(port, (msg) => msg.type === "response" && msg.id === req.id);
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

describe("engine.worker getWorkbookInfo fallback respects sheetOrder", () => {
  afterAll(() => {
    (globalThis as any).self = previousSelf;
  });

  it("orders sheets using sheetOrder when wasm does not export getWorkbookInfo()", async () => {
    const wasmModuleUrl = new URL("./fixtures/mockWasmWorkbookJsonSheetOrder.mjs", import.meta.url).href;
    const { port, dispose } = await setupWorker({ wasmModuleUrl });

    try {
      await sendRequest(port, { type: "request", id: 0, method: "newWorkbook", params: {} });

      const resp = await sendRequest(port, { type: "request", id: 1, method: "getWorkbookInfo", params: {} });

      expect(resp.ok).toBe(true);
      const sheets: WorkbookSheetInfoDto[] = (resp as any).result?.sheets ?? [];
      expect(sheets.map((s) => s.id)).toEqual(["Sheet2", "Sheet1", "Empty"]);

      // Use an `as const` entry tuple so TS infers a stable Map value type (instead of `{}`).
      const byId = new Map<string, WorkbookSheetInfoDto>(sheets.map((s) => [s.id, s] as const));
      expect(byId.get("Sheet1")?.visibility).toBe("hidden");
      expect(byId.get("Sheet1")?.tabColor).toEqual({ rgb: "FFFF0000" });
      expect(byId.get("Sheet2")?.visibility).toBe("veryHidden");
      expect(byId.get("Empty")?.tabColor).toEqual({ theme: 1, tint: 0.5 });
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
