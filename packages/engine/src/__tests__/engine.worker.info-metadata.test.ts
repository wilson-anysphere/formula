import { afterAll, describe, expect, it } from "vitest";

import type { InitMessage, RpcRequest, WorkerOutboundMessage } from "../protocol.ts";

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

describe("engine.worker INFO() metadata integration", () => {
  afterAll(() => {
    (globalThis as any).self = previousSelf;
  });

  it("propagates setEngineInfo + setSheetOrigin via worker RPC", async () => {
    const wasmModuleUrl = new URL("./fixtures/mockWasmWorkbookEvalInfo.mjs", import.meta.url).href;
    const { port, dispose } = await setupWorker({ wasmModuleUrl });

    try {
      await sendRequest(port, { type: "request", id: 0, method: "newWorkbook", params: {} });

      // Seed formulas that depend on INFO metadata.
      await sendRequest(port, {
        type: "request",
        id: 1,
        method: "setCells",
        params: {
          updates: [
            { address: "A1", value: '=INFO("system")', sheet: "Sheet1" },
            { address: "A2", value: '=INFO("osversion")', sheet: "Sheet1" },
            { address: "A3", value: '=INFO("release")', sheet: "Sheet1" },
            { address: "A4", value: '=INFO("version")', sheet: "Sheet1" },
            { address: "A5", value: '=INFO("memavail")', sheet: "Sheet1" },
            { address: "A6", value: '=INFO("totmem")', sheet: "Sheet1" },
            { address: "A7", value: '=INFO("directory")', sheet: "Sheet1" },
            { address: "A8", value: '=INFO("origin")', sheet: "Sheet1" },
            { address: "A1", value: '=INFO("origin")', sheet: "Sheet2" },
          ],
        },
      });

      await sendRequest(port, {
        type: "request",
        id: 2,
        method: "setEngineInfo",
        params: {
          info: {
            system: "mac",
            osversion: "14.0",
            release: "sonoma",
            version: "1.2.3",
            memavail: 123.5,
            totmem: 456.25,
            directory: "/tmp",
          },
        },
      });

      // Host-provided sheet origin state drives `INFO("origin")`.
      await sendRequest(port, {
        type: "request",
        id: 4,
        method: "setSheetOrigin",
        params: { sheet: "  Sheet1  ", origin: "B2" },
      });

      await sendRequest(port, { type: "request", id: 5, method: "recalculate", params: {} });

      const read = async (id: number, sheet: string, address: string) => {
        const resp = await sendRequest(port, { type: "request", id, method: "getCell", params: { address, sheet } });
        expect(resp.ok).toBe(true);
        return resp.result.value as unknown;
      };

      expect(await read(6, "Sheet1", "A1")).toBe("mac");
      expect(await read(7, "Sheet1", "A2")).toBe("14.0");
      expect(await read(8, "Sheet1", "A3")).toBe("sonoma");
      expect(await read(9, "Sheet1", "A4")).toBe("1.2.3");
      expect(await read(10, "Sheet1", "A5")).toBe(123.5);
      expect(await read(11, "Sheet1", "A6")).toBe(456.25);
      // Excel-compatible directory results include a trailing path separator.
      expect(await read(12, "Sheet1", "A7")).toBe("/tmp/");
      expect(await read(13, "Sheet1", "A8")).toBe("$B$2");

      // Sheet2 uses the default origin.
      expect(await read(14, "Sheet2", "A1")).toBe("$A$1");
    } finally {
      dispose();
    }
  });

  it("normalizes empty strings as unset and validates memavail/totmem finiteness", async () => {
    const wasmModuleUrl = new URL("./fixtures/mockWasmWorkbookEvalInfo.mjs", import.meta.url).href;
    const { port, dispose } = await setupWorker({ wasmModuleUrl });

    try {
      await sendRequest(port, { type: "request", id: 0, method: "newWorkbook", params: {} });

      await sendRequest(port, {
        type: "request",
        id: 1,
        method: "setCells",
        params: {
          updates: [
            { address: "A1", value: '=INFO("system")', sheet: "Sheet1" },
            { address: "A2", value: '=INFO("osversion")', sheet: "Sheet1" },
            { address: "A3", value: '=INFO("origin")', sheet: "Sheet1" },
            { address: "A4", value: '=INFO("memavail")', sheet: "Sheet1" },
          ],
        },
      });

      await sendRequest(port, {
        type: "request",
        id: 2,
        method: "setEngineInfo",
        params: { info: { system: "mac", osversion: "14.0", memavail: 100 } },
      });
      await sendRequest(port, {
        type: "request",
        id: 3,
        method: "setSheetOrigin",
        params: { sheet: "  Sheet1  ", origin: "B2" },
      });
      await sendRequest(port, { type: "request", id: 5, method: "recalculate", params: {} });

      const read = async (id: number, address: string) => {
        const resp = await sendRequest(port, { type: "request", id, method: "getCell", params: { address, sheet: "Sheet1" } });
        expect(resp.ok).toBe(true);
        return resp.result.value as unknown;
      };

      expect(await read(6, "A1")).toBe("mac");
      expect(await read(7, "A2")).toBe("14.0");
      expect(await read(8, "A3")).toBe("$B$2");
      expect(await read(9, "A4")).toBe(100);

      // Clearing the origin should fall back to `$A$1`.
      await sendRequest(port, {
        type: "request",
        id: 10,
        method: "setSheetOrigin",
        params: { sheet: "  Sheet1  ", origin: "" },
      });
      await sendRequest(port, { type: "request", id: 11, method: "recalculate", params: {} });
      expect(await read(12, "A3")).toBe("$A$1");

      // Clearing string metadata should revert to Excel defaults / #N/A.
      await sendRequest(port, {
        type: "request",
        id: 13,
        method: "setEngineInfo",
        params: { info: { system: "", osversion: "" } },
      });
      await sendRequest(port, { type: "request", id: 14, method: "recalculate", params: {} });
      expect(await read(15, "A1")).toBe("pcdos");
      expect(await read(16, "A2")).toBe("#N/A");

      // memavail must be finite numbers.
      const bad = await sendRequest(port, {
        type: "request",
        id: 17,
        method: "setEngineInfo",
        params: { info: { memavail: Infinity } },
      });
      expect(bad.ok).toBe(false);
      expect(String((bad as any).error)).toContain("memavail");

      // Failed update should not clobber existing metadata.
      await sendRequest(port, { type: "request", id: 18, method: "recalculate", params: {} });
      expect(await read(19, "A4")).toBe(100);

      // Clearing memavail should restore #N/A.
      await sendRequest(port, {
        type: "request",
        id: 20,
        method: "setEngineInfo",
        params: { info: { memavail: null } },
      });
      await sendRequest(port, { type: "request", id: 21, method: "recalculate", params: {} });
      expect(await read(22, "A4")).toBe("#N/A");
    } finally {
      dispose();
    }
  });

  it('supports legacy setInfoOrigin / setInfoOriginForSheet fallbacks for INFO("origin")', async () => {
    const wasmModuleUrl = new URL("./fixtures/mockWasmWorkbookEvalInfo.mjs", import.meta.url).href;
    const { port, dispose } = await setupWorker({ wasmModuleUrl });

    try {
      await sendRequest(port, { type: "request", id: 0, method: "newWorkbook", params: {} });

      await sendRequest(port, {
        type: "request",
        id: 1,
        method: "setCells",
        params: {
          updates: [
            { address: "A1", value: '=INFO("origin")', sheet: "Sheet1" },
            { address: "A1", value: '=INFO("origin")', sheet: "Sheet2" },
          ],
        },
      });

      const read = async (id: number, sheet: string) => {
        const resp = await sendRequest(port, { type: "request", id, method: "getCell", params: { address: "A1", sheet } });
        expect(resp.ok).toBe(true);
        return resp.result.value as unknown;
      };

      await sendRequest(port, { type: "request", id: 2, method: "recalculate", params: {} });
      expect(await read(3, "Sheet1")).toBe("$A$1");
      expect(await read(4, "Sheet2")).toBe("$A$1");

      // Workbook-level legacy origin (A1 string) should be normalized to absolute form.
      await sendRequest(port, { type: "request", id: 5, method: "setInfoOrigin", params: { origin: "B2" } });
      await sendRequest(port, { type: "request", id: 6, method: "recalculate", params: {} });
      expect(await read(7, "Sheet1")).toBe("$B$2");
      expect(await read(8, "Sheet2")).toBe("$B$2");

      // Per-sheet legacy override should take precedence.
      await sendRequest(port, {
        type: "request",
        id: 9,
        method: "setInfoOriginForSheet",
        params: { sheet: "Sheet1", origin: "C5" },
      });
      await sendRequest(port, { type: "request", id: 10, method: "recalculate", params: {} });
      expect(await read(11, "Sheet1")).toBe("$C$5");
      expect(await read(12, "Sheet2")).toBe("$B$2");

      // Non-A1 legacy origin strings should be returned verbatim for backward compatibility.
      await sendRequest(port, { type: "request", id: 13, method: "setInfoOrigin", params: { origin: "workbook-origin" } });
      await sendRequest(port, { type: "request", id: 14, method: "recalculate", params: {} });
      expect(await read(15, "Sheet1")).toBe("$C$5");
      expect(await read(16, "Sheet2")).toBe("workbook-origin");

      // Clearing the per-sheet override should fall back to the workbook-level legacy string.
      await sendRequest(port, {
        type: "request",
        id: 17,
        method: "setInfoOriginForSheet",
        params: { sheet: "Sheet1", origin: "" },
      });
      await sendRequest(port, { type: "request", id: 18, method: "recalculate", params: {} });
      expect(await read(19, "Sheet1")).toBe("workbook-origin");
      expect(await read(20, "Sheet2")).toBe("workbook-origin");

      // Clearing the workbook-level origin should restore the `$A$1` default.
      await sendRequest(port, { type: "request", id: 21, method: "setInfoOrigin", params: { origin: "" } });
      await sendRequest(port, { type: "request", id: 22, method: "recalculate", params: {} });
      expect(await read(23, "Sheet1")).toBe("$A$1");
      expect(await read(24, "Sheet2")).toBe("$A$1");
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
