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
        method: "setRowStyleId",
        params: { sheet: "Sheet1", row: 5, styleId: 9 }
      });
      expect(resp.ok).toBe(true);

      resp = await sendRequest(port, {
        type: "request",
        id: 4,
        method: "setRowStyleId",
        // Clear semantics: `null` should be treated as "reset" (worker forwards `0` to wasm).
        params: { sheet: "Sheet1", row: 6, styleId: null }
      });
      expect(resp.ok).toBe(true);

      resp = await sendRequest(port, {
        type: "request",
        id: 5,
        method: "setColStyleId",
        params: { sheet: "Sheet1", col: 2, styleId: 11 }
      });
      expect(resp.ok).toBe(true);

      resp = await sendRequest(port, {
        type: "request",
        id: 6,
        method: "setColStyleId",
        // Clear semantics: `null` should be treated as "reset" (worker forwards `0` to wasm).
        params: { sheet: "Sheet1", col: 3, styleId: null }
      });
      expect(resp.ok).toBe(true);

      resp = await sendRequest(port, {
        type: "request",
        id: 7,
        method: "setSheetDefaultStyleId",
        params: { sheet: "Sheet1", styleId: 13 }
      });
      expect(resp.ok).toBe(true);

      resp = await sendRequest(port, {
        type: "request",
        id: 8,
        method: "setSheetDefaultStyleId",
        // Clear semantics: `null` should be treated as "reset" (worker forwards `0` to wasm).
        params: { sheet: "Sheet1", styleId: null }
      });
      expect(resp.ok).toBe(true);

      resp = await sendRequest(port, {
        type: "request",
        id: 9,
        method: "setFormatRunsByCol",
        // Omit sheet; worker should default to "Sheet1".
        params: { col: 2, runs: [{ startRow: 0, endRowExclusive: 10, styleId: 17 }] }
      });
      expect(resp.ok).toBe(true);

      resp = await sendRequest(port, {
        type: "request",
        id: 10,
        method: "setColWidth",
        params: { sheet: "Sheet1", col: 2, width: 120 }
      });
      expect(resp.ok).toBe(true);

      resp = await sendRequest(port, {
        type: "request",
        id: 11,
        method: "setColHidden",
        params: { sheet: "Sheet1", col: 2, hidden: true }
      });
      expect(resp.ok).toBe(true);

      resp = await sendRequest(port, {
        type: "request",
        id: 12,
        method: "internStyle",
        params: { style: { font: { bold: true } } }
      });
      expect(resp.ok).toBe(true);
      expect((resp as RpcResponseOk).result).toBe(42);

      resp = await sendRequest(port, {
        type: "request",
        id: 13,
        method: "setRowStyleId",
        params: { row: 0, styleId: null }
      });
      expect(resp.ok).toBe(true);

      resp = await sendRequest(port, {
        type: "request",
        id: 14,
        method: "setColStyleId",
        params: { col: 0, styleId: null }
      });
      expect(resp.ok).toBe(true);

      resp = await sendRequest(port, {
        type: "request",
        id: 15,
        method: "setSheetDefaultStyleId",
        params: { styleId: null }
      });
      expect(resp.ok).toBe(true);

      expect((globalThis as any).__ENGINE_WORKER_TEST_CALLS__).toEqual([
        ["setWorkbookFileMetadata", "/tmp", "book.xlsx"],
        ["setCellStyleId", "Sheet1", "A1", 7],
        ["setRowStyleId", "Sheet1", 5, 9],
        ["setRowStyleId", "Sheet1", 6, 0],
        ["setColStyleId", "Sheet1", 2, 11],
        ["setColStyleId", "Sheet1", 3, 0],
        ["setSheetDefaultStyleId", "Sheet1", 13],
        ["setSheetDefaultStyleId", "Sheet1", 0],
        ["setFormatRunsByCol", "Sheet1", 2, [{ startRow: 0, endRowExclusive: 10, styleId: 17 }]],
        ["setColWidth", "Sheet1", 2, 120],
        ["setColHidden", "Sheet1", 2, true],
        ["internStyle", { font: { bold: true } }],
        // The worker defaults the sheet name to "Sheet1" when the caller omits it, and
        // normalizes `null`/`undefined` style ids to `0` for backward compatibility.
        ["setRowStyleId", "Sheet1", 0, 0],
        ["setColStyleId", "Sheet1", 0, 0],
        ["setSheetDefaultStyleId", "Sheet1", 0]
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

  it("falls back to the legacy sheet-last setCellStyleId signature when needed", async () => {
    (globalThis as any).__ENGINE_WORKER_TEST_CALLS__ = [];
    const wasmModuleUrl = new URL("./fixtures/mockWasmWorkbookMetadataLegacyCellStyle.mjs", import.meta.url).href;
    const { port, dispose } = await setupWorker({ wasmModuleUrl });

    try {
      await sendRequest(port, { type: "request", id: 0, method: "newWorkbook", params: {} });
      const resp = await sendRequest(port, {
        type: "request",
        id: 1,
        method: "setCellStyleId",
        params: { sheet: "Sheet1", address: "A1", styleId: 7 }
      });
      expect(resp.ok).toBe(true);

      expect((globalThis as any).__ENGINE_WORKER_TEST_CALLS__).toEqual([["setCellStyleId", "A1", 7, "Sheet1"]]);
    } finally {
      dispose();
      delete (globalThis as any).__ENGINE_WORKER_TEST_CALLS__;
    }
  });

  it("defaults blank sheet names to Sheet1 for sheet-scoped metadata RPCs", async () => {
    (globalThis as any).__ENGINE_WORKER_TEST_CALLS__ = [];
    const wasmModuleUrl = new URL("./fixtures/mockWasmWorkbookMetadata.mjs", import.meta.url).href;
    const { port, dispose } = await setupWorker({ wasmModuleUrl });

    try {
      await sendRequest(port, { type: "request", id: 0, method: "newWorkbook", params: {} });

      let resp = await sendRequest(port, {
        type: "request",
        id: 1,
        method: "setCellStyleId",
        params: { sheet: "", address: "A1", styleId: 7 }
      });
      expect(resp.ok).toBe(true);

      resp = await sendRequest(port, {
        type: "request",
        id: 2,
        method: "setRowStyleId",
        params: { sheet: "   ", row: 5, styleId: 9 }
      });
      expect(resp.ok).toBe(true);

      resp = await sendRequest(port, {
        type: "request",
        id: 3,
        method: "setColStyleId",
        params: { sheet: "", col: 2, styleId: 11 }
      });
      expect(resp.ok).toBe(true);

      resp = await sendRequest(port, {
        type: "request",
        id: 4,
        method: "setSheetDefaultStyleId",
        params: { sheet: " ", styleId: 13 }
      });
      expect(resp.ok).toBe(true);

      resp = await sendRequest(port, {
        type: "request",
        id: 5,
        method: "setColWidth",
        params: { sheet: "", col: 2, width: 120 }
      });
      expect(resp.ok).toBe(true);

      resp = await sendRequest(port, {
        type: "request",
        id: 6,
        method: "setColWidthChars",
        params: { sheet: "", col: 3, widthChars: 8.5 }
      });
      expect(resp.ok).toBe(true);

      resp = await sendRequest(port, {
        type: "request",
        id: 7,
        method: "setColHidden",
        params: { sheet: "", col: 2, hidden: true }
      });
      expect(resp.ok).toBe(true);

      resp = await sendRequest(port, {
        type: "request",
        id: 8,
        method: "setSheetDimensions",
        params: { sheet: "", rows: 10, cols: 20 }
      });
      expect(resp.ok).toBe(true);

      resp = await sendRequest(port, {
        type: "request",
        id: 9,
        method: "getSheetDimensions",
        params: { sheet: "" }
      });
      expect(resp.ok).toBe(true);
      expect((resp as RpcResponseOk).result).toEqual({ rows: 100, cols: 200 });

      resp = await sendRequest(port, {
        type: "request",
        id: 10,
        method: "setFormatRunsByCol",
        params: { sheet: "", col: 2, runs: [{ startRow: 0, endRowExclusive: 1, styleId: 17 }] }
      });
      expect(resp.ok).toBe(true);

      expect((globalThis as any).__ENGINE_WORKER_TEST_CALLS__).toEqual([
        ["setCellStyleId", "Sheet1", "A1", 7],
        ["setRowStyleId", "Sheet1", 5, 9],
        ["setColStyleId", "Sheet1", 2, 11],
        ["setSheetDefaultStyleId", "Sheet1", 13],
        ["setColWidth", "Sheet1", 2, 120],
        ["setColWidthChars", "Sheet1", 3, 8.5],
        ["setColHidden", "Sheet1", 2, true],
        ["setSheetDimensions", "Sheet1", 10, 20],
        ["getSheetDimensions", "Sheet1"],
        ["setFormatRunsByCol", "Sheet1", 2, [{ startRow: 0, endRowExclusive: 1, styleId: 17 }]]
      ]);
    } finally {
      dispose();
      delete (globalThis as any).__ENGINE_WORKER_TEST_CALLS__;
    }
  });

  it("trims sheet ids for setSheetDisplayName", async () => {
    (globalThis as any).__ENGINE_WORKER_TEST_CALLS__ = [];
    const wasmModuleUrl = new URL("./fixtures/mockWasmWorkbookMetadata.mjs", import.meta.url).href;
    const { port, dispose } = await setupWorker({ wasmModuleUrl });

    try {
      await sendRequest(port, { type: "request", id: 0, method: "newWorkbook", params: {} });

      let resp = await sendRequest(port, {
        type: "request",
        id: 1,
        method: "setSheetDisplayName",
        params: { sheetId: "  Sheet2  ", name: "Budget" }
      });
      expect(resp.ok).toBe(true);

      resp = await sendRequest(port, {
        type: "request",
        id: 2,
        method: "setSheetDisplayName",
        params: { sheetId: "   ", name: "Main" }
      });
      expect(resp.ok).toBe(true);

      expect((globalThis as any).__ENGINE_WORKER_TEST_CALLS__).toEqual([
        ["setSheetDisplayName", "Sheet2", "Budget"],
        ["setSheetDisplayName", "Sheet1", "Main"]
      ]);
    } finally {
      dispose();
      delete (globalThis as any).__ENGINE_WORKER_TEST_CALLS__;
    }
  });

  it("trims sheet names for renameSheet", async () => {
    (globalThis as any).__ENGINE_WORKER_TEST_CALLS__ = [];
    const wasmModuleUrl = new URL("./fixtures/mockWasmWorkbookMetadata.mjs", import.meta.url).href;
    const { port, dispose } = await setupWorker({ wasmModuleUrl });

    try {
      await sendRequest(port, { type: "request", id: 0, method: "newWorkbook", params: {} });

      let resp = await sendRequest(port, {
        type: "request",
        id: 1,
        method: "renameSheet",
        params: { oldName: "  Sheet1  ", newName: "  Budget  " }
      });
      expect(resp.ok).toBe(true);
      expect((resp as RpcResponseOk).result).toBe(true);

      resp = await sendRequest(port, {
        type: "request",
        id: 2,
        method: "renameSheet",
        params: { oldName: "   ", newName: "Budget" }
      });
      expect(resp.ok).toBe(true);
      expect((resp as RpcResponseOk).result).toBe(false);

      expect((globalThis as any).__ENGINE_WORKER_TEST_CALLS__).toEqual([["renameSheet", "Sheet1", "Budget"]]);
    } finally {
      dispose();
      delete (globalThis as any).__ENGINE_WORKER_TEST_CALLS__;
    }
  });

  it("defaults blank sheet names to Sheet1 for setInfoOriginForSheet", async () => {
    (globalThis as any).__ENGINE_WORKER_TEST_CALLS__ = [];
    const wasmModuleUrl = new URL("./fixtures/mockWasmWorkbookMetadata.mjs", import.meta.url).href;
    const { port, dispose } = await setupWorker({ wasmModuleUrl });

    try {
      await sendRequest(port, { type: "request", id: 0, method: "newWorkbook", params: {} });

      let resp = await sendRequest(port, {
        type: "request",
        id: 1,
        method: "setInfoOriginForSheet",
        params: { sheet: "  Sheet2  ", origin: "B2" }
      });
      expect(resp.ok).toBe(true);

      resp = await sendRequest(port, {
        type: "request",
        id: 2,
        method: "setInfoOriginForSheet",
        params: { sheet: "   ", origin: null }
      });
      expect(resp.ok).toBe(true);

      expect((globalThis as any).__ENGINE_WORKER_TEST_CALLS__).toEqual([
        ["setInfoOriginForSheet", "Sheet2", "B2"],
        ["setInfoOriginForSheet", "Sheet1", null]
      ]);
    } finally {
      dispose();
      delete (globalThis as any).__ENGINE_WORKER_TEST_CALLS__;
    }
  });

  it("treats blank sheet names as missing for sheet-optional cell edit RPCs", async () => {
    (globalThis as any).__ENGINE_WORKER_TEST_CALLS__ = [];
    const wasmModuleUrl = new URL("./fixtures/mockWasmWorkbookMetadata.mjs", import.meta.url).href;
    const { port, dispose } = await setupWorker({ wasmModuleUrl });

    try {
      await sendRequest(port, { type: "request", id: 0, method: "newWorkbook", params: {} });

      let resp = await sendRequest(port, {
        type: "request",
        id: 1,
        method: "setCells",
        params: {
          updates: [
            { sheet: "", address: "A1", value: 1 },
            { sheet: "Sheet2", address: "A2", value: 2 },
            { sheet: "   ", address: "A3", value: 3 }
          ]
        }
      });
      expect(resp.ok).toBe(true);

      resp = await sendRequest(port, {
        type: "request",
        id: 2,
        method: "setRange",
        params: { sheet: "", range: "A1:A1", values: [[1]] }
      });
      expect(resp.ok).toBe(true);

      resp = await sendRequest(port, {
        type: "request",
        id: 3,
        method: "setCellRich",
        params: { sheet: " ", address: "B1", value: null }
      });
      expect(resp.ok).toBe(true);

      resp = await sendRequest(port, {
        type: "request",
        id: 4,
        method: "getCell",
        params: { sheet: "", address: "A1" }
      });
      expect(resp.ok).toBe(true);
      expect((resp as RpcResponseOk).result).toEqual({ sheet: "Sheet1", address: "A1", input: null, value: null });

      expect((globalThis as any).__ENGINE_WORKER_TEST_CALLS__).toEqual([
        [
          "setCells",
          [
            { address: "A1", value: 1 },
            { address: "A2", value: 2, sheet: "Sheet2" },
            { address: "A3", value: 3 }
          ]
        ],
        ["setRange", "A1:A1", [[1]], undefined],
        ["setCellRich", "B1", null, undefined]
      ]);
    } finally {
      dispose();
      delete (globalThis as any).__ENGINE_WORKER_TEST_CALLS__;
    }
  });

  it("defaults blank sheet names to Sheet1 for applyOperation", async () => {
    (globalThis as any).__ENGINE_WORKER_TEST_CALLS__ = [];
    const wasmModuleUrl = new URL("./fixtures/mockWasmWorkbookMetadata.mjs", import.meta.url).href;
    const { port, dispose } = await setupWorker({ wasmModuleUrl });

    try {
      await sendRequest(port, { type: "request", id: 0, method: "newWorkbook", params: {} });

      const op = { type: "InsertRows", sheet: "   ", row: 0, count: 1 } as const;
      const resp = await sendRequest(port, { type: "request", id: 1, method: "applyOperation", params: { op } });
      expect(resp.ok).toBe(true);

      expect((globalThis as any).__ENGINE_WORKER_TEST_CALLS__).toEqual([
        ["applyOperation", { type: "InsertRows", sheet: "Sheet1", row: 0, count: 1 }]
      ]);
    } finally {
      dispose();
      delete (globalThis as any).__ENGINE_WORKER_TEST_CALLS__;
    }
  });

  it("trims whitespace in sheet names for sheet-optional cell edit RPCs", async () => {
    (globalThis as any).__ENGINE_WORKER_TEST_CALLS__ = [];
    const wasmModuleUrl = new URL("./fixtures/mockWasmWorkbookMetadata.mjs", import.meta.url).href;
    const { port, dispose } = await setupWorker({ wasmModuleUrl });

    try {
      await sendRequest(port, { type: "request", id: 0, method: "newWorkbook", params: {} });

      let resp = await sendRequest(port, {
        type: "request",
        id: 1,
        method: "setCells",
        params: { updates: [{ sheet: "  Sheet2  ", address: "A1", value: 1 }] }
      });
      expect(resp.ok).toBe(true);

      resp = await sendRequest(port, {
        type: "request",
        id: 2,
        method: "setRange",
        params: { sheet: "  Sheet2  ", range: "A1:A1", values: [[2]] }
      });
      expect(resp.ok).toBe(true);

      resp = await sendRequest(port, {
        type: "request",
        id: 3,
        method: "setCellRich",
        params: { sheet: "  Sheet2  ", address: "B1", value: null }
      });
      expect(resp.ok).toBe(true);

      resp = await sendRequest(port, {
        type: "request",
        id: 4,
        method: "getCell",
        params: { sheet: "  Sheet2  ", address: "A1" }
      });
      expect(resp.ok).toBe(true);
      expect((resp as RpcResponseOk).result).toEqual({ sheet: "Sheet2", address: "A1", input: null, value: null });

      expect((globalThis as any).__ENGINE_WORKER_TEST_CALLS__).toEqual([
        ["setCells", [{ address: "A1", value: 1, sheet: "Sheet2" }]],
        ["setRange", "A1:A1", [[2]], "Sheet2"],
        ["setCellRich", "B1", null, "Sheet2"]
      ]);
    } finally {
      dispose();
      delete (globalThis as any).__ENGINE_WORKER_TEST_CALLS__;
    }
  });

  it("trims whitespace in sheet names for applyOperation", async () => {
    (globalThis as any).__ENGINE_WORKER_TEST_CALLS__ = [];
    const wasmModuleUrl = new URL("./fixtures/mockWasmWorkbookMetadata.mjs", import.meta.url).href;
    const { port, dispose } = await setupWorker({ wasmModuleUrl });

    try {
      await sendRequest(port, { type: "request", id: 0, method: "newWorkbook", params: {} });

      const op = { type: "InsertRows", sheet: "  Sheet2  ", row: 0, count: 1 } as const;
      const resp = await sendRequest(port, { type: "request", id: 1, method: "applyOperation", params: { op } });
      expect(resp.ok).toBe(true);

      expect((globalThis as any).__ENGINE_WORKER_TEST_CALLS__).toEqual([
        ["applyOperation", { type: "InsertRows", sheet: "Sheet2", row: 0, count: 1 }]
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
