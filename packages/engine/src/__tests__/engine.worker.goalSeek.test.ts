import { afterAll, describe, expect, it } from "vitest";

import type { InitMessage, RpcRequest, RpcResponseErr, RpcResponseOk, WorkerOutboundMessage } from "../protocol.ts";

class MockWorkerGlobal {
  private readonly listeners = new Map<string, Set<(event: MessageEvent<unknown>) => void>>();

  addEventListener(type: string, listener: (event: MessageEvent<unknown>) => void): void {
    const key = String(type ?? "");
    let set = this.listeners.get(key);
    if (!set) {
      set = new Set();
      this.listeners.set(key, set);
    }
    set.add(listener);
  }

  dispatchMessage(data: unknown): void {
    const event = { data } as MessageEvent<unknown>;
    for (const listener of this.listeners.get("message") ?? []) {
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

describe("engine.worker goalSeek RPC normalization", () => {
  afterAll(() => {
    (globalThis as any).self = previousSelf;
  });

  it("normalizes the new wasm goalSeek response shape and converts undefined change values to null", async () => {
    const wasmModuleUrl = new URL("./fixtures/mockWasmWorkbookGoalSeek.mjs", import.meta.url).href;
    const { port, dispose } = await setupWorker({ wasmModuleUrl });

    try {
      await sendRequest(port, { type: "request", id: 0, method: "newWorkbook", params: {} });

      const resp = await sendRequest(port, {
        type: "request",
        id: 1,
        method: "goalSeek",
        params: {
          sheet: "Sheet1",
          targetCell: "B1",
          targetValue: 25,
          changingCell: "A1",
          derivativeStep: 0.1,
        },
      });

      expect(resp.ok).toBe(true);
      const payload = (resp as RpcResponseOk).result as any;
      expect(payload).toHaveProperty("result");
      expect(payload).toHaveProperty("changes");
      expect(payload.result.status).toBe("Converged");
      expect(payload.result.solution).toBe(5);

      const c1 = (payload.changes as any[]).find((c) => c.sheet === "Sheet1" && c.address === "C1");
      expect(c1).toBeTruthy();
      // The mock wasm module returns `undefined` for `C1`; the worker should normalize this to `null`.
      expect(c1.value).toBeNull();
    } finally {
      dispose();
    }
  });

  it("wraps legacy flat goalSeek responses into the new { result, changes } shape", async () => {
    const wasmModuleUrl = new URL("./fixtures/mockWasmWorkbookGoalSeek.mjs", import.meta.url).href;
    const { port, dispose } = await setupWorker({ wasmModuleUrl });

    try {
      await sendRequest(port, { type: "request", id: 0, method: "newWorkbook", params: {} });

      const resp = await sendRequest(port, {
        type: "request",
        id: 1,
        method: "goalSeek",
        params: {
          sheet: "Sheet1",
          targetCell: "B1",
          targetValue: 25,
          changingCell: "A1",
        },
      });

      expect(resp.ok).toBe(true);
      const payload = (resp as RpcResponseOk).result as any;
      expect(payload.result.status).toBe("Converged");
      expect(payload.result.solution).toBe(5);
      // Legacy flat payloads may omit `finalOutput`; the worker should reconstruct it as
      // `targetValue + finalError` (matches `formula_engine::what_if::goal_seek` semantics).
      expect(payload.result.finalOutput).toBe(25);
      // Legacy responses did not include changes.
      expect(payload.changes).toEqual([]);
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
