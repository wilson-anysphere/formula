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
    wasmModuleUrl: options.wasmModuleUrl,
  };
  workerGlobal.dispatchMessage(init);

  await waitForMessage(channel.port1, (msg) => msg.type === "ready");

  return {
    port: channel.port1,
    dispose: () => channel.port1.close(),
  };
}

async function sendRequest(port: MockMessagePort, req: RpcRequest): Promise<RpcResponseOk | RpcResponseErr> {
  const responsePromise = waitForMessage(port, (msg) => msg.type === "response" && msg.id === req.id) as Promise<
    RpcResponseOk | RpcResponseErr
  >;
  port.postMessage(req);
  return await responsePromise;
}

describe("engine.worker locale metadata RPCs", () => {
  afterAll(() => {
    (globalThis as any).self = previousSelf;
  });

  it("dispatches supportedLocaleIds and getLocaleInfo without requiring a workbook", async () => {
    const wasmModuleUrl = new URL("./fixtures/mockWasmLocaleInfo.mjs", import.meta.url).href;
    const { port, dispose } = await setupWorker({ wasmModuleUrl });

    try {
      let resp = await sendRequest(port, { type: "request", id: 1, method: "supportedLocaleIds", params: {} });
      expect(resp.ok).toBe(true);
      expect((resp as RpcResponseOk).result).toEqual(["de-DE", "en-US", "es-ES", "fr-FR", "ja-JP"]);

      resp = await sendRequest(port, { type: "request", id: 2, method: "getLocaleInfo", params: { localeId: "de-DE" } });
      expect(resp.ok).toBe(true);
      expect((resp as RpcResponseOk).result).toEqual({
        localeId: "de-DE",
        decimalSeparator: ",",
        argSeparator: ";",
        arrayRowSeparator: ";",
        arrayColSeparator: "\\",
        thousandsSeparator: ".",
        isRtl: false,
        booleanTrue: "WAHR",
        booleanFalse: "FALSCH",
      });
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
