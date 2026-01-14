import { describe, expect, it, vi } from "vitest";

import { EngineWorker, type MessageChannelLike, type WorkerLike } from "../worker/EngineWorker.ts";
import type {
  CellValueRich,
  InitMessage,
  RpcRequest,
  WorkerInboundMessage,
  WorkerOutboundMessage
} from "../protocol.ts";

class MockMessagePort {
  onmessage: ((event: MessageEvent<unknown>) => void) | null = null;
  public readonly sent: Array<{ message: unknown; transfer?: Transferable[] }> = [];
  private listeners = new Set<(event: MessageEvent<unknown>) => void>();
  private other: MockMessagePort | null = null;

  connect(other: MockMessagePort) {
    this.other = other;
  }

  postMessage(message: unknown, transfer?: Transferable[]): void {
    this.sent.push({ message, transfer });
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

function createMockChannel(): MessageChannelLike {
  const port1 = new MockMessagePort();
  const port2 = new MockMessagePort();
  port1.connect(port2);
  port2.connect(port1);
  return { port1, port2: port2 as unknown as MessagePort };
}

class MockWorker implements WorkerLike {
  public serverPort: MockMessagePort | null = null;
  public received: WorkerInboundMessage[] = [];
  public terminated = false;

  postMessage(message: unknown): void {
    const init = message as InitMessage;
    if (!init || typeof init !== "object" || (init as any).type !== "init") {
      return;
    }

    this.serverPort = init.port as unknown as MockMessagePort;
    this.serverPort.addEventListener("message", (event) => {
      const msg = event.data as WorkerInboundMessage;
      this.received.push(msg);

      if (msg.type === "request") {
        const req = msg as RpcRequest;
        if (req.method === "setCells") {
          const response: WorkerOutboundMessage = {
            type: "response",
            id: req.id,
            ok: true,
            result: null
          };
          this.serverPort?.postMessage(response);
          return;
        }
        if (req.method === "setLocale") {
          const response: WorkerOutboundMessage = {
            type: "response",
            id: req.id,
            ok: true,
            result: true
          };
          this.serverPort?.postMessage(response);
          return;
        }
        if (req.method === "setSheetDimensions") {
          const response: WorkerOutboundMessage = {
            type: "response",
            id: req.id,
            ok: true,
            result: null
          };
          this.serverPort?.postMessage(response);
          return;
        }
        if (req.method === "getSheetDimensions") {
          const response: WorkerOutboundMessage = {
            type: "response",
            id: req.id,
            ok: true,
            result: { rows: 2_100_000, cols: 16_384 }
          };
          this.serverPort?.postMessage(response);
          return;
        }
        if (req.method === "internStyle") {
          const response: WorkerOutboundMessage = {
            type: "response",
            id: req.id,
            ok: true,
            result: 123
          };
          this.serverPort?.postMessage(response);
          return;
        }
        if (
          req.method === "setWorkbookFileMetadata" ||
          req.method === "setCellStyleId" ||
          req.method === "setColWidth" ||
          req.method === "setColHidden"
        ) {
          const response: WorkerOutboundMessage = {
            type: "response",
            id: req.id,
            ok: true,
            result: null
          };
          this.serverPort?.postMessage(response);
          return;
        }

        // Default: echo response.
        const response: WorkerOutboundMessage = {
          type: "response",
          id: req.id,
          ok: true,
          result: req.params
        };
        this.serverPort?.postMessage(response);
      }
    });

    const ready: WorkerOutboundMessage = { type: "ready" };
    this.serverPort.postMessage(ready);
  }

  terminate(): void {
    this.terminated = true;
    this.serverPort?.close();
  }
}

describe("EngineWorker RPC", () => {
  it("batches consecutive setCell calls into a single setCells request", async () => {
    const worker = new MockWorker();
    const engine = await EngineWorker.connect({
      worker,
      wasmModuleUrl: "mock://wasm",
      channelFactory: createMockChannel
    });

    await Promise.all([engine.setCell("A1", 1), engine.setCell("A2", 2)]);

    const requests = worker.received.filter(
      (msg): msg is RpcRequest => msg.type === "request" && (msg as RpcRequest).method === "setCells"
    );

    expect(requests).toHaveLength(1);
    expect((requests[0].params as any).updates).toEqual([
      { address: "A1", value: 1, sheet: undefined },
      { address: "A2", value: 2, sheet: undefined }
    ]);
  });

  it("sends lexFormulaPartial requests with formula + options params", async () => {
    const worker = new MockWorker();
    const engine = await EngineWorker.connect({
      worker,
      wasmModuleUrl: "mock://wasm",
      channelFactory: createMockChannel
    });

    await engine.lexFormulaPartial("=\"hello", { localeId: "de-DE", referenceStyle: "R1C1" });

    const requests = worker.received.filter(
      (msg): msg is RpcRequest => msg.type === "request" && (msg as RpcRequest).method === "lexFormulaPartial"
    );

    expect(requests).toHaveLength(1);
    expect(requests[0].params).toEqual({
      formula: "=\"hello",
      options: { localeId: "de-DE", referenceStyle: "R1C1" }
    });
  });

  it("does not silently drop malformed parse options passed as the 3rd arg to parseFormulaPartial", async () => {
    const worker = new MockWorker();
    const engine = await EngineWorker.connect({
      worker,
      wasmModuleUrl: "mock://wasm",
      channelFactory: createMockChannel
    });

    await expect(engine.parseFormulaPartial("=1+2", undefined, { localeID: "de-DE" } as any)).rejects.toThrow(
      /options must be \{ localeId\?: string, referenceStyle\?:/
    );

    const requests = worker.received.filter(
      (msg): msg is RpcRequest => msg.type === "request" && (msg as RpcRequest).method === "parseFormulaPartial"
    );
    expect(requests).toHaveLength(0);
  });

  it("transfers xlsx ArrayBuffer when loading workbooks from bytes", async () => {
    const worker = new MockWorker();
    const channel = createMockChannel();
    const engine = await EngineWorker.connect({
      worker,
      wasmModuleUrl: "mock://wasm",
      channelFactory: () => channel
    });

    const bytes = new Uint8Array([1, 2, 3, 4]);
    await engine.loadWorkbookFromXlsxBytes(bytes);

    const requests = worker.received.filter(
      (msg): msg is RpcRequest => msg.type === "request" && (msg as RpcRequest).method === "loadFromXlsxBytes"
    );

    expect(requests).toHaveLength(1);

    const clientPort = channel.port1 as unknown as MockMessagePort;
    const call = clientPort.sent.find((entry) => {
      const msg = entry.message as any;
      return msg?.type === "request" && msg?.method === "loadFromXlsxBytes";
    });

    expect(call?.transfer).toHaveLength(1);
    expect(call?.transfer?.[0]).toBe(bytes.buffer);
  });

  it("transfers only the view range when loading xlsx bytes from a subarray", async () => {
    const worker = new MockWorker();
    const channel = createMockChannel();
    const engine = await EngineWorker.connect({
      worker,
      wasmModuleUrl: "mock://wasm",
      channelFactory: () => channel
    });

    const backing = new Uint8Array([0, 0, 1, 2, 3, 4, 0, 0]);
    const view = backing.subarray(2, 6);
    await engine.loadWorkbookFromXlsxBytes(view);

    const clientPort = channel.port1 as unknown as MockMessagePort;
    const call = clientPort.sent.find((entry) => {
      const msg = entry.message as any;
      return msg?.type === "request" && msg?.method === "loadFromXlsxBytes";
    });

    expect(call?.transfer).toHaveLength(1);
    const transferred = call?.transfer?.[0] as ArrayBuffer | undefined;
    expect(transferred).toBeInstanceOf(ArrayBuffer);
    expect(transferred?.byteLength).toBe(view.byteLength);

    const paramsBytes = (call?.message as any)?.params?.bytes as Uint8Array | undefined;
    expect(paramsBytes).toBeInstanceOf(Uint8Array);
    expect(paramsBytes?.byteLength).toBe(view.byteLength);
    expect(paramsBytes?.buffer).toBe(transferred);
  });

  it("flushes pending setCell batches before loading workbook from xlsx bytes", async () => {
    const worker = new MockWorker();
    const engine = await EngineWorker.connect({
      worker,
      wasmModuleUrl: "mock://wasm",
      channelFactory: createMockChannel
    });

    // Simulate a caller firing-and-forgetting setCell; loadWorkbookFromXlsxBytes must
    // flush the pending microtask-batched setCells before replacing the workbook.
    void engine.setCell("A1", 1);

    await engine.loadWorkbookFromXlsxBytes(new Uint8Array([1, 2, 3, 4]));

    const methods = worker.received
      .filter((msg): msg is RpcRequest => msg.type === "request")
      .map((msg) => msg.method);

    expect(methods).toEqual(["setCells", "loadFromXlsxBytes"]);
  });

  it("flushes pending setCell batches before setSheetDimensions", async () => {
    const worker = new MockWorker();
    const engine = await EngineWorker.connect({
      worker,
      wasmModuleUrl: "mock://wasm",
      channelFactory: createMockChannel
    });

    // Simulate a caller firing-and-forgetting setCell; setSheetDimensions should
    // flush the pending microtask-batched setCells first.
    void engine.setCell("A1", 1);

    await engine.setSheetDimensions("Sheet1", 2_100_000, 16_384);

    const methods = worker.received
      .filter((msg): msg is RpcRequest => msg.type === "request")
      .map((msg) => msg.method);

    expect(methods).toEqual(["setCells", "setSheetDimensions"]);
  });

  it("supports getSheetDimensions RPC requests", async () => {
    const worker = new MockWorker();
    const engine = await EngineWorker.connect({
      worker,
      wasmModuleUrl: "mock://wasm",
      channelFactory: createMockChannel
    });

    const dims = await engine.getSheetDimensions("Sheet1");
    expect(dims).toEqual({ rows: 2_100_000, cols: 16_384 });

    const requests = worker.received.filter(
      (msg): msg is RpcRequest => msg.type === "request" && (msg as RpcRequest).method === "getSheetDimensions"
    );
    expect(requests).toHaveLength(1);
    expect(requests[0].params).toEqual({ sheet: "Sheet1" });
  });

  it("sends getCellRich RPC requests with the expected params", async () => {
    const worker = new MockWorker();
    const engine = await EngineWorker.connect({
      worker,
      wasmModuleUrl: "mock://wasm",
      channelFactory: createMockChannel
    });

    await engine.getCellRich("A1", "Sheet1");

    const requests = worker.received.filter(
      (msg): msg is RpcRequest => msg.type === "request" && (msg as RpcRequest).method === "getCellRich"
    );
    expect(requests).toHaveLength(1);
    expect(requests[0].params).toEqual({ address: "A1", sheet: "Sheet1" });
  });

  it("flushes pending setCell batches before setCellRich", async () => {
    const worker = new MockWorker();
    const engine = await EngineWorker.connect({
      worker,
      wasmModuleUrl: "mock://wasm",
      channelFactory: createMockChannel
    });

    // Fire-and-forget setCell leaves a pending microtask flush.
    void engine.setCell("A1", 1);

    const value: CellValueRich = { type: "entity", value: { displayValue: "Acme" } };
    await engine.setCellRich("A2", value, "Sheet1");

    const methods = worker.received
      .filter((msg): msg is RpcRequest => msg.type === "request")
      .map((msg) => msg.method);

    expect(methods).toEqual(["setCells", "setCellRich"]);
  });

  it("sends setCellRich RPC requests with the expected params", async () => {
    const worker = new MockWorker();
    const engine = await EngineWorker.connect({
      worker,
      wasmModuleUrl: "mock://wasm",
      channelFactory: createMockChannel
    });

    const value: CellValueRich = {
      type: "entity",
      value: { displayValue: "Acme", properties: { Price: { type: "number", value: 12.5 } } }
    };
    await engine.setCellRich("A1", value, "Sheet1");

    const requests = worker.received.filter(
      (msg): msg is RpcRequest => msg.type === "request" && (msg as RpcRequest).method === "setCellRich"
    );
    expect(requests).toHaveLength(1);
    expect(requests[0].params).toEqual({ address: "A1", value, sheet: "Sheet1" });
  });

  it("supports request cancellation via AbortSignal", async () => {
    const worker = new MockWorker();
    const engine = await EngineWorker.connect({
      worker,
      wasmModuleUrl: "mock://wasm",
      channelFactory: createMockChannel
    });

    const controller = new AbortController();
    const promise = engine.getCell("A1", undefined, { signal: controller.signal });
    controller.abort();

    await expect(promise).rejects.toThrow(/aborted/i);

    const cancelMessages = worker.received.filter((msg) => msg.type === "cancel");
    expect(cancelMessages).toHaveLength(1);
  });

  it("times out requests and sends cancel", async () => {
    vi.useFakeTimers();
    const worker = new MockWorker();
    worker.postMessage = (message: unknown) => {
      const init = message as InitMessage;
      worker.serverPort = init.port as unknown as MockMessagePort;
      worker.serverPort.addEventListener("message", (event) => {
        const msg = event.data as WorkerInboundMessage;
        worker.received.push(msg);
        // Intentionally don't respond.
      });
      worker.serverPort.postMessage({ type: "ready" } as WorkerOutboundMessage);
    };

    const engine = await EngineWorker.connect({
      worker,
      wasmModuleUrl: "mock://wasm",
      channelFactory: createMockChannel
    });

    const promise = engine.getCell("A1", undefined, { timeoutMs: 50 });
    const expectation = expect(promise).rejects.toThrow(/timed out/i);
    await vi.advanceTimersByTimeAsync(50);
    await expectation;

    const cancelMessages = worker.received.filter((msg) => msg.type === "cancel");
    expect(cancelMessages).toHaveLength(1);

    vi.useRealTimers();
  });

  it("sends lexFormula RPC requests with the expected params", async () => {
    const worker = new MockWorker();
    const engine = await EngineWorker.connect({
      worker,
      wasmModuleUrl: "mock://wasm",
      channelFactory: createMockChannel
    });

    await engine.lexFormula("=1+2");

    const requests = worker.received.filter(
      (msg): msg is RpcRequest => msg.type === "request" && (msg as RpcRequest).method === "lexFormula"
    );
    expect(requests).toHaveLength(1);
    expect(requests[0].params).toEqual({ formula: "=1+2", options: undefined });
  });

  it("forwards lexFormula options in the RPC params", async () => {
    const worker = new MockWorker();
    const engine = await EngineWorker.connect({
      worker,
      wasmModuleUrl: "mock://wasm",
      channelFactory: createMockChannel
    });

    await engine.lexFormula("=1+2", { localeId: "de-DE", referenceStyle: "R1C1" });

    const requests = worker.received.filter(
      (msg): msg is RpcRequest => msg.type === "request" && (msg as RpcRequest).method === "lexFormula"
    );
    expect(requests).toHaveLength(1);
    expect(requests[0].params).toEqual({
      formula: "=1+2",
      options: { localeId: "de-DE", referenceStyle: "R1C1" }
    });
  });

  it("supports lexFormula overload with rpcOptions as the second argument", async () => {
    const worker = new MockWorker();
    const engine = await EngineWorker.connect({
      worker,
      wasmModuleUrl: "mock://wasm",
      channelFactory: createMockChannel
    });

    await engine.lexFormula("=1+2", { timeoutMs: 1_000 });

    const requests = worker.received.filter(
      (msg): msg is RpcRequest => msg.type === "request" && (msg as RpcRequest).method === "lexFormula"
    );
    expect(requests).toHaveLength(1);
    expect(requests[0].params).toEqual({
      formula: "=1+2",
      options: undefined
    });
  });

  it("does not flush pending setCell batches when calling lexFormula", async () => {
    const worker = new MockWorker();
    const engine = await EngineWorker.connect({
      worker,
      wasmModuleUrl: "mock://wasm",
      channelFactory: createMockChannel
    });

    // Intentionally do not await: this leaves a scheduled microtask flush pending.
    const setPromise = engine.setCell("A1", 1);

    // `lexFormula` is editor tooling and should not force-flush workbook edits.
    await engine.lexFormula("=1+2");

    // Clean up to avoid leaking an in-flight flush promise.
    await setPromise;

    const requests = worker.received.filter((msg): msg is RpcRequest => msg.type === "request");
    const lexIndex = requests.findIndex((req) => req.method === "lexFormula");
    const setCellsIndex = requests.findIndex((req) => req.method === "setCells");

    expect(lexIndex).toBeGreaterThanOrEqual(0);
    expect(setCellsIndex).toBeGreaterThanOrEqual(0);
    expect(lexIndex).toBeLessThan(setCellsIndex);
  });

  it("supports lexFormulaPartial overload with rpcOptions as the second argument", async () => {
    const worker = new MockWorker();
    const engine = await EngineWorker.connect({
      worker,
      wasmModuleUrl: "mock://wasm",
      channelFactory: createMockChannel
    });

    await engine.lexFormulaPartial("=SUM(1,", { timeoutMs: 1_000 });

    const requests = worker.received.filter(
      (msg): msg is RpcRequest => msg.type === "request" && (msg as RpcRequest).method === "lexFormulaPartial"
    );
    expect(requests).toHaveLength(1);
    expect(requests[0].params).toEqual({
      formula: "=SUM(1,",
      options: undefined
    });
  });

  it("sends parseFormulaPartial RPC requests with the expected params", async () => {
    const worker = new MockWorker();
    const engine = await EngineWorker.connect({
      worker,
      wasmModuleUrl: "mock://wasm",
      channelFactory: createMockChannel
    });

    await engine.parseFormulaPartial("=SUM(1,", 6);

    const requests = worker.received.filter(
      (msg): msg is RpcRequest => msg.type === "request" && (msg as RpcRequest).method === "parseFormulaPartial"
    );
    expect(requests).toHaveLength(1);
    expect(requests[0].params).toEqual({ formula: "=SUM(1,", cursor: 6, options: undefined });
  });

  it("supports parseFormulaPartial overload with options as the second argument", async () => {
    const worker = new MockWorker();
    const engine = await EngineWorker.connect({
      worker,
      wasmModuleUrl: "mock://wasm",
      channelFactory: createMockChannel
    });

    await engine.parseFormulaPartial("=SUM(1,", { localeId: "de-DE", referenceStyle: "R1C1" });

    const requests = worker.received.filter(
      (msg): msg is RpcRequest => msg.type === "request" && (msg as RpcRequest).method === "parseFormulaPartial"
    );
    expect(requests).toHaveLength(1);
    expect(requests[0].params).toEqual({
      formula: "=SUM(1,",
      cursor: undefined,
      options: { localeId: "de-DE", referenceStyle: "R1C1" }
    });
  });

  it("forwards parseFormulaPartial options in the RPC params", async () => {
    const worker = new MockWorker();
    const engine = await EngineWorker.connect({
      worker,
      wasmModuleUrl: "mock://wasm",
      channelFactory: createMockChannel
    });

    await engine.parseFormulaPartial("=SUM(1,", 6, { localeId: "de-DE", referenceStyle: "R1C1" });

    const requests = worker.received.filter(
      (msg): msg is RpcRequest => msg.type === "request" && (msg as RpcRequest).method === "parseFormulaPartial"
    );
    expect(requests).toHaveLength(1);
    expect(requests[0].params).toEqual({
      formula: "=SUM(1,",
      cursor: 6,
      options: { localeId: "de-DE", referenceStyle: "R1C1" }
    });
  });

  it("does not flush pending setCell batches when calling parseFormulaPartial", async () => {
    const worker = new MockWorker();
    const engine = await EngineWorker.connect({
      worker,
      wasmModuleUrl: "mock://wasm",
      channelFactory: createMockChannel
    });

    // Intentionally do not await: this leaves a scheduled microtask flush pending.
    const setPromise = engine.setCell("A1", 1);

    // `parseFormulaPartial` is editor tooling and should not force-flush workbook edits.
    await engine.parseFormulaPartial("=SUM(1,", 6);

    // Clean up to avoid leaking an in-flight flush promise.
    await setPromise;

    const requests = worker.received.filter((msg): msg is RpcRequest => msg.type === "request");
    const parseIndex = requests.findIndex((req) => req.method === "parseFormulaPartial");
    const setCellsIndex = requests.findIndex((req) => req.method === "setCells");

    expect(parseIndex).toBeGreaterThanOrEqual(0);
    expect(setCellsIndex).toBeGreaterThanOrEqual(0);
    expect(parseIndex).toBeLessThan(setCellsIndex);
  });

  it("forwards setLocale calls with the correct RPC method name", async () => {
    const worker = new MockWorker();
    const engine = await EngineWorker.connect({
      worker,
      wasmModuleUrl: "mock://wasm",
      channelFactory: createMockChannel
    });

    const ok = await engine.setLocale("de-DE");
    expect(ok).toBe(true);

    const requests = worker.received.filter(
      (msg): msg is RpcRequest => msg.type === "request" && (msg as RpcRequest).method === "setLocale"
    );
    expect(requests).toHaveLength(1);
    expect(requests[0].params).toEqual({ localeId: "de-DE" });
  });

  it("forwards applyOperation calls with the correct RPC method name", async () => {
    const worker = new MockWorker();
    const engine = await EngineWorker.connect({
      worker,
      wasmModuleUrl: "mock://wasm",
      channelFactory: createMockChannel
    });

    const op = { type: "InsertRows", sheet: "Sheet1", row: 0, count: 1 } as const;
    await engine.applyOperation(op);

    const requests = worker.received.filter(
      (msg): msg is RpcRequest => msg.type === "request" && (msg as RpcRequest).method === "applyOperation"
    );
    expect(requests).toHaveLength(1);
    expect(requests[0].params).toEqual({ op });
  });

  it("sends setWorkbookFileMetadata RPC requests with the expected params", async () => {
    const worker = new MockWorker();
    const engine = await EngineWorker.connect({
      worker,
      wasmModuleUrl: "mock://wasm",
      channelFactory: createMockChannel
    });

    await engine.setWorkbookFileMetadata("/tmp", "book.xlsx");

    const requests = worker.received.filter(
      (msg): msg is RpcRequest => msg.type === "request" && (msg as RpcRequest).method === "setWorkbookFileMetadata"
    );
    expect(requests).toHaveLength(1);
    expect(requests[0].params).toEqual({ directory: "/tmp", filename: "book.xlsx" });
  });

  it("flushes pending setCell batches before setColWidth", async () => {
    const worker = new MockWorker();
    const engine = await EngineWorker.connect({
      worker,
      wasmModuleUrl: "mock://wasm",
      channelFactory: createMockChannel
    });

    // Fire-and-forget setCell leaves a pending microtask flush.
    void engine.setCell("A1", 1);

    await engine.setColWidth(3, 120, "Sheet1");

    const methods = worker.received
      .filter((msg): msg is RpcRequest => msg.type === "request")
      .map((msg) => msg.method);

    expect(methods).toEqual(["setCells", "setColWidth"]);
  });

  it("sends setCellStyleId RPC requests with the expected params", async () => {
    const worker = new MockWorker();
    const engine = await EngineWorker.connect({
      worker,
      wasmModuleUrl: "mock://wasm",
      channelFactory: createMockChannel
    });

    await engine.setCellStyleId("A1", 7, "Sheet1");

    const requests = worker.received.filter(
      (msg): msg is RpcRequest => msg.type === "request" && (msg as RpcRequest).method === "setCellStyleId"
    );
    expect(requests).toHaveLength(1);
    expect(requests[0].params).toEqual({ sheet: "Sheet1", address: "A1", styleId: 7 });
  });

  it("internStyle returns the style id from the worker", async () => {
    const worker = new MockWorker();
    const engine = await EngineWorker.connect({
      worker,
      wasmModuleUrl: "mock://wasm",
      channelFactory: createMockChannel
    });

    const styleId = await engine.internStyle({ font: { bold: true } });
    expect(styleId).toBe(123);

    const requests = worker.received.filter(
      (msg): msg is RpcRequest => msg.type === "request" && (msg as RpcRequest).method === "internStyle"
    );
    expect(requests).toHaveLength(1);
    expect(requests[0].params).toEqual({ style: { font: { bold: true } } });
  });
});
