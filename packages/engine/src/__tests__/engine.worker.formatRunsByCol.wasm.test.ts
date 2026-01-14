import { afterAll, describe, expect, it } from "vitest";

import type { InitMessage, RpcRequest, WorkerOutboundMessage } from "../protocol.ts";

const skipWasmBuild = process.env.FORMULA_SKIP_WASM_BUILD === "1" || process.env.FORMULA_SKIP_WASM_BUILD === "true";
const describeWasm = skipWasmBuild ? describe.skip : describe;

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

describeWasm("engine.worker range-run formatting integration (wasm)", () => {
  afterAll(() => {
    (globalThis as any).self = previousSelf;
  });

  it('updates CELL("protect") results after setFormatRunsByCol', async () => {
    const wasmModuleUrl = new URL("./fixtures/formulaWasmNodeWrapper.mjs", import.meta.url).href;
    const { port, dispose } = await setupWorker({ wasmModuleUrl });

    try {
      await sendRequest(port, { type: "request", id: 0, method: "newWorkbook", params: {} });

      // Style that explicitly unlocks cells.
      const styleResp = await sendRequest(port, {
        type: "request",
        id: 1,
        method: "internStyle",
        params: { style: { protection: { locked: false, hidden: false } } },
      });
      expect(styleResp.ok).toBe(true);
      const styleId = styleResp.result as number;
      expect(Number.isFinite(styleId)).toBe(true);

      // Two reference cells so we can check the run boundary (A1 included, A2 excluded).
      await sendRequest(port, {
        type: "request",
        id: 2,
        method: "setCells",
        params: {
          updates: [
            { address: "A1", value: "x", sheet: "Sheet1" },
            { address: "A2", value: "y", sheet: "Sheet1" },
            { address: "B1", value: '=CELL("protect",A1)', sheet: "Sheet1" },
            { address: "B2", value: '=CELL("protect",A2)', sheet: "Sheet1" },
          ],
        },
      });

      await sendRequest(port, { type: "request", id: 3, method: "recalculate", params: {} });

      let resp = await sendRequest(port, {
        type: "request",
        id: 4,
        method: "getCell",
        params: { address: "B1", sheet: "Sheet1" },
      });
      expect(resp.ok).toBe(true);
      // Excel defaults to locked cells.
      expect(resp.result.value).toBe(1);

      resp = await sendRequest(port, {
        type: "request",
        id: 5,
        method: "getCell",
        params: { address: "B2", sheet: "Sheet1" },
      });
      expect(resp.ok).toBe(true);
      expect(resp.result.value).toBe(1);

      // Apply a range-run only for row 0 (A1), not row 1 (A2).
      const applyResp = await sendRequest(port, {
        type: "request",
        id: 6,
        method: "setFormatRunsByCol",
        params: { sheet: "Sheet1", col: 0, runs: [{ startRow: 0, endRowExclusive: 1, styleId }] },
      });
      expect(applyResp.ok).toBe(true);

      await sendRequest(port, { type: "request", id: 7, method: "recalculate", params: {} });

      resp = await sendRequest(port, {
        type: "request",
        id: 8,
        method: "getCell",
        params: { address: "B1", sheet: "Sheet1" },
      });
      expect(resp.ok).toBe(true);
      expect(resp.result.value).toBe(0);

      resp = await sendRequest(port, {
        type: "request",
        id: 9,
        method: "getCell",
        params: { address: "B2", sheet: "Sheet1" },
      });
      expect(resp.ok).toBe(true);
      // Row 1 is outside the run and should remain locked.
      expect(resp.result.value).toBe(1);

      // Clearing the runs should restore the default formatting.
      const clearResp = await sendRequest(port, {
        type: "request",
        id: 10,
        method: "setFormatRunsByCol",
        params: { sheet: "Sheet1", col: 0, runs: [] },
      });
      expect(clearResp.ok).toBe(true);

      await sendRequest(port, { type: "request", id: 11, method: "recalculate", params: {} });

      resp = await sendRequest(port, {
        type: "request",
        id: 12,
        method: "getCell",
        params: { address: "B1", sheet: "Sheet1" },
      });
      expect(resp.ok).toBe(true);
      expect(resp.result.value).toBe(1);

      // Validate wasm-side run parsing rejects invalid ranges.
      const invalidResp = await sendRequest(port, {
        type: "request",
        id: 13,
        method: "setFormatRunsByCol",
        params: { sheet: "Sheet1", col: 0, runs: [{ startRow: 1, endRowExclusive: 1, styleId }] },
      });
      expect(invalidResp.ok).toBe(false);
      expect(String(invalidResp.error)).toMatch(/endRowExclusive/i);
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
