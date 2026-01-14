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

describe("engine.worker workbook file metadata integration", () => {
  afterAll(() => {
    (globalThis as any).self = previousSelf;
  });

  it('formats CELL("filename") without a directory when only filename is known', async () => {
    const wasmModuleUrl = new URL("./fixtures/mockWasmWorkbookEvalFilename.mjs", import.meta.url).href;
    const { port, dispose } = await setupWorker({ wasmModuleUrl });

    try {
      await sendRequest(port, { type: "request", id: 0, method: "newWorkbook", params: {} });

      await sendRequest(port, {
        type: "request",
        id: 1,
        method: "setCells",
        params: {
          updates: [
            { address: "A1", value: '=CELL("filename")', sheet: "Sheet1" },
            { address: "A2", value: '=INFO("directory")', sheet: "Sheet1" },
            { address: "A1", value: '=CELL("filename")', sheet: "Sheet2" },
            { address: "A3", value: '=CELL("filename",Sheet2!A1)', sheet: "Sheet1" },
          ],
        },
      });

      // Simulate a web-style save where only the workbook name is known.
      await sendRequest(port, {
        type: "request",
        id: 2,
        method: "setWorkbookFileMetadata",
        params: { directory: null, filename: "book.xlsx" },
      });
      await sendRequest(port, { type: "request", id: 3, method: "recalculate", params: {} });

      let resp = await sendRequest(port, { type: "request", id: 4, method: "getCell", params: { address: "A1", sheet: "Sheet1" } });
      expect(resp.ok).toBe(true);
      expect(resp.result.value).toBe("[book.xlsx]Sheet1");

      resp = await sendRequest(port, { type: "request", id: 5, method: "getCell", params: { address: "A2", sheet: "Sheet1" } });
      expect(resp.ok).toBe(true);
      // INFO("directory") requires a non-empty directory string.
      expect(resp.result.value).toBe("#N/A");

      resp = await sendRequest(port, { type: "request", id: 6, method: "getCell", params: { address: "A1", sheet: "Sheet2" } });
      expect(resp.ok).toBe(true);
      expect(resp.result.value).toBe("[book.xlsx]Sheet2");

      // `CELL("filename", reference)` should use the reference's sheet name component.
      resp = await sendRequest(port, { type: "request", id: 7, method: "getCell", params: { address: "A3", sheet: "Sheet1" } });
      expect(resp.ok).toBe(true);
      expect(resp.result.value).toBe("[book.xlsx]Sheet2");
    } finally {
      dispose();
    }
  });

  it('prefers EngineInfo.directory over workbook file metadata for INFO("directory")', async () => {
    const wasmModuleUrl = new URL("./fixtures/mockWasmWorkbookEvalFilename.mjs", import.meta.url).href;
    const { port, dispose } = await setupWorker({ wasmModuleUrl });

    try {
      await sendRequest(port, { type: "request", id: 0, method: "newWorkbook", params: {} });

      await sendRequest(port, {
        type: "request",
        id: 1,
        method: "setCells",
        params: {
          updates: [
            { address: "A1", value: '=INFO("directory")', sheet: "Sheet1" },
            { address: "A2", value: '=CELL("filename")', sheet: "Sheet1" },
          ],
        },
      });

      await sendRequest(port, {
        type: "request",
        id: 2,
        method: "setWorkbookFileMetadata",
        params: { directory: "/workbook", filename: "book.xlsx" },
      });
      await sendRequest(port, { type: "request", id: 3, method: "recalculate", params: {} });

      let resp = await sendRequest(port, { type: "request", id: 4, method: "getCell", params: { address: "A1", sheet: "Sheet1" } });
      expect(resp.ok).toBe(true);
      expect(resp.result.value).toBe("/workbook/");

      resp = await sendRequest(port, { type: "request", id: 5, method: "getCell", params: { address: "A2", sheet: "Sheet1" } });
      expect(resp.ok).toBe(true);
      expect(resp.result.value).toBe("/workbook/[book.xlsx]Sheet1");

      // Host-provided INFO("directory") override should win over workbook file metadata.
      await sendRequest(port, {
        type: "request",
        id: 6,
        method: "setEngineInfo",
        params: { info: { directory: "/host" } },
      });
      await sendRequest(port, { type: "request", id: 7, method: "recalculate", params: {} });

      resp = await sendRequest(port, { type: "request", id: 8, method: "getCell", params: { address: "A1", sheet: "Sheet1" } });
      expect(resp.ok).toBe(true);
      expect(resp.result.value).toBe("/host/");

      // CELL("filename") continues to use workbook file metadata (not the host override).
      resp = await sendRequest(port, { type: "request", id: 9, method: "getCell", params: { address: "A2", sheet: "Sheet1" } });
      expect(resp.ok).toBe(true);
      expect(resp.result.value).toBe("/workbook/[book.xlsx]Sheet1");

      // Clearing the override restores workbook-metadata behavior.
      await sendRequest(port, {
        type: "request",
        id: 10,
        method: "setEngineInfo",
        params: { info: { directory: "" } },
      });
      await sendRequest(port, { type: "request", id: 11, method: "recalculate", params: {} });
      resp = await sendRequest(port, { type: "request", id: 12, method: "getCell", params: { address: "A1", sheet: "Sheet1" } });
      expect(resp.ok).toBe(true);
      expect(resp.result.value).toBe("/workbook/");
    } finally {
      dispose();
    }
  });

  it('uses the reference sheet name in CELL("filename", reference) without quoting', async () => {
    const wasmModuleUrl = new URL("./fixtures/mockWasmWorkbookEvalFilename.mjs", import.meta.url).href;
    const { port, dispose } = await setupWorker({ wasmModuleUrl });

    try {
      await sendRequest(port, { type: "request", id: 0, method: "newWorkbook", params: {} });

      await sendRequest(port, {
        type: "request",
        id: 1,
        method: "setCells",
        params: {
          updates: [{ address: "A1", value: '=CELL("filename",\'Other Sheet\'!A1)', sheet: "Sheet1" }],
        },
      });

      await sendRequest(port, {
        type: "request",
        id: 2,
        method: "setWorkbookFileMetadata",
        params: { directory: "/tmp", filename: "book.xlsx" },
      });
      await sendRequest(port, { type: "request", id: 3, method: "recalculate", params: {} });

      const resp = await sendRequest(port, { type: "request", id: 4, method: "getCell", params: { address: "A1", sheet: "Sheet1" } });
      expect(resp.ok).toBe(true);
      // The output should include the referenced sheet display name without Excel quoting.
      expect(resp.result.value).toBe("/tmp/[book.xlsx]Other Sheet");
    } finally {
      dispose();
    }
  });

  it('infers a trailing path separator for Windows-style directories', async () => {
    const wasmModuleUrl = new URL("./fixtures/mockWasmWorkbookEvalFilename.mjs", import.meta.url).href;
    const { port, dispose } = await setupWorker({ wasmModuleUrl });

    try {
      await sendRequest(port, { type: "request", id: 0, method: "newWorkbook", params: {} });

      await sendRequest(port, {
        type: "request",
        id: 1,
        method: "setCells",
        params: {
          updates: [
            { address: "A1", value: '=CELL("filename")', sheet: "Sheet1" },
            { address: "A2", value: '=INFO("directory")', sheet: "Sheet1" },
          ],
        },
      });

      // Use a Windows-style directory without a trailing separator; the engine should infer `\\`.
      await sendRequest(port, {
        type: "request",
        id: 2,
        method: "setWorkbookFileMetadata",
        params: { directory: "C:\\tmp", filename: "book.xlsx" },
      });
      await sendRequest(port, { type: "request", id: 3, method: "recalculate", params: {} });

      let resp = await sendRequest(port, { type: "request", id: 4, method: "getCell", params: { address: "A1", sheet: "Sheet1" } });
      expect(resp.ok).toBe(true);
      expect(resp.result.value).toBe("C:\\tmp\\[book.xlsx]Sheet1");

      resp = await sendRequest(port, { type: "request", id: 5, method: "getCell", params: { address: "A2", sheet: "Sheet1" } });
      expect(resp.ok).toBe(true);
      expect(resp.result.value).toBe("C:\\tmp\\");
    } finally {
      dispose();
    }
  });

  it('updates CELL("filename") results after setWorkbookFileMetadata', async () => {
    const wasmModuleUrl = new URL("./fixtures/mockWasmWorkbookEvalFilename.mjs", import.meta.url).href;
    const { port, dispose } = await setupWorker({ wasmModuleUrl });

    try {
      await sendRequest(port, { type: "request", id: 0, method: "newWorkbook", params: {} });

      // Seed a simple formula that depends on workbook metadata.
      await sendRequest(port, {
        type: "request",
        id: 1,
        method: "setCells",
        params: {
          updates: [
            { address: "A1", value: '=CELL("filename")', sheet: "Sheet1" },
            { address: "A2", value: '=INFO("directory")', sheet: "Sheet1" },
            { address: "A1", value: '=CELL("filename")', sheet: "Sheet2" },
            { address: "A3", value: '=CELL("filename",Sheet2!A1)', sheet: "Sheet1" },
          ],
        },
      });

      await sendRequest(port, { type: "request", id: 2, method: "recalculate", params: {} });

      let resp = await sendRequest(port, { type: "request", id: 3, method: "getCell", params: { address: "A1", sheet: "Sheet1" } });
      expect(resp.ok).toBe(true);
      expect(resp.result.value).toBe("");

      resp = await sendRequest(port, { type: "request", id: 4, method: "getCell", params: { address: "A2", sheet: "Sheet1" } });
      expect(resp.ok).toBe(true);
      expect(resp.result.value).toBe("#N/A");

      resp = await sendRequest(port, { type: "request", id: 5, method: "getCell", params: { address: "A1", sheet: "Sheet2" } });
      expect(resp.ok).toBe(true);
      expect(resp.result.value).toBe("");

      resp = await sendRequest(port, { type: "request", id: 6, method: "getCell", params: { address: "A3", sheet: "Sheet1" } });
      expect(resp.ok).toBe(true);
      expect(resp.result.value).toBe("");

      // Simulate Save As.
      await sendRequest(port, {
        type: "request",
        id: 7,
        method: "setWorkbookFileMetadata",
        params: { directory: "/tmp", filename: "book.xlsx" },
      });

      await sendRequest(port, { type: "request", id: 8, method: "recalculate", params: {} });

      resp = await sendRequest(port, { type: "request", id: 9, method: "getCell", params: { address: "A1", sheet: "Sheet1" } });
      expect(resp.ok).toBe(true);
      expect(resp.result.value).toBe("/tmp/[book.xlsx]Sheet1");

      resp = await sendRequest(port, { type: "request", id: 10, method: "getCell", params: { address: "A2", sheet: "Sheet1" } });
      expect(resp.ok).toBe(true);
      expect(resp.result.value).toBe("/tmp/");

      resp = await sendRequest(port, { type: "request", id: 11, method: "getCell", params: { address: "A1", sheet: "Sheet2" } });
      expect(resp.ok).toBe(true);
      expect(resp.result.value).toBe("/tmp/[book.xlsx]Sheet2");

      // `CELL("filename", reference)` should use the reference's sheet name component.
      resp = await sendRequest(port, { type: "request", id: 12, method: "getCell", params: { address: "A3", sheet: "Sheet1" } });
      expect(resp.ok).toBe(true);
      expect(resp.result.value).toBe("/tmp/[book.xlsx]Sheet2");

      // Simulate creating a new, unsaved workbook (metadata cleared).
      await sendRequest(port, {
        type: "request",
        id: 13,
        method: "setWorkbookFileMetadata",
        params: { directory: null, filename: null },
      });
      await sendRequest(port, { type: "request", id: 14, method: "recalculate", params: {} });

      resp = await sendRequest(port, { type: "request", id: 15, method: "getCell", params: { address: "A1", sheet: "Sheet1" } });
      expect(resp.ok).toBe(true);
      expect(resp.result.value).toBe("");

      resp = await sendRequest(port, { type: "request", id: 16, method: "getCell", params: { address: "A2", sheet: "Sheet1" } });
      expect(resp.ok).toBe(true);
      expect(resp.result.value).toBe("#N/A");

      resp = await sendRequest(port, { type: "request", id: 17, method: "getCell", params: { address: "A1", sheet: "Sheet2" } });
      expect(resp.ok).toBe(true);
      expect(resp.result.value).toBe("");

      resp = await sendRequest(port, { type: "request", id: 18, method: "getCell", params: { address: "A3", sheet: "Sheet1" } });
      expect(resp.ok).toBe(true);
      expect(resp.result.value).toBe("");
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
