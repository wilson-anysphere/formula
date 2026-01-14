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
  public closed = false;
  private listeners = new Map<string, Set<(event: MessageEvent<unknown>) => void>>();
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
    this.closed = true;
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

  getListenerCount(): number {
    let count = 0;
    for (const set of this.listeners.values()) {
      count += set.size;
    }
    return count;
  }

  private dispatchMessage(data: unknown): void {
    const event = { data } as MessageEvent<unknown>;
    this.onmessage?.(event);
    for (const listener of this.listeners.get("message") ?? []) {
      listener(event);
    }
  }

  dispatchMessageError(): void {
    const event = {} as MessageEvent<unknown>;
    for (const listener of this.listeners.get("messageerror") ?? []) {
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
        if (req.method === "ping") {
          const response: WorkerOutboundMessage = {
            type: "response",
            id: req.id,
            ok: true,
            result: "pong"
          };
          this.serverPort?.postMessage(response);
          return;
        }
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
        if (req.method === "supportedLocaleIds") {
          const response: WorkerOutboundMessage = {
            type: "response",
            id: req.id,
            ok: true,
            result: ["de-DE", "en-US"]
          };
          this.serverPort?.postMessage(response);
          return;
        }
        if (req.method === "getLocaleInfo") {
          const response: WorkerOutboundMessage = {
            type: "response",
            id: req.id,
            ok: true,
            result: {
              localeId: "de-DE",
              decimalSeparator: ",",
              argSeparator: ";",
              arrayRowSeparator: ";",
              arrayColSeparator: "\\",
              thousandsSeparator: ".",
              isRtl: false,
              booleanTrue: "WAHR",
              booleanFalse: "FALSCH"
            }
          };
          this.serverPort?.postMessage(response);
          return;
        }
        if (req.method === "getCalcSettings") {
          const response: WorkerOutboundMessage = {
            type: "response",
            id: req.id,
            ok: true,
            result: {
              calculationMode: "manual",
              calculateBeforeSave: true,
              fullPrecision: true,
              fullCalcOnLoad: false,
              iterative: { enabled: false, maxIterations: 100, maxChange: 0.001 }
            }
          };
          this.serverPort?.postMessage(response);
          return;
        }
        if (req.method === "setCalcSettings") {
          const response: WorkerOutboundMessage = {
            type: "response",
            id: req.id,
            ok: true,
            result: null
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

class HangingWorker implements WorkerLike {
  public serverPort: MockMessagePort | null = null;
  public terminated = false;
  public postMessageCalls = 0;

  postMessage(message: unknown): void {
    this.postMessageCalls += 1;
    const init = message as InitMessage;
    if (!init || typeof init !== "object" || (init as any).type !== "init") {
      return;
    }

    this.serverPort = init.port as unknown as MockMessagePort;
    // Intentionally never post the initial `{ type: "ready" }` handshake.
  }

  terminate(): void {
    this.terminated = true;
    this.serverPort?.close();
    this.serverPort = null;
  }
}

class ErrorSetCellsWorker implements WorkerLike {
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

      if (msg.type !== "request") {
        return;
      }

      const req = msg as RpcRequest;
      if (req.method === "setCells") {
        const response: WorkerOutboundMessage = {
          type: "response",
          id: req.id,
          ok: false,
          error: "setCells failed"
        };
        this.serverPort?.postMessage(response);
        return;
      }

      const response: WorkerOutboundMessage = {
        type: "response",
        id: req.id,
        ok: true,
        result: null
      };
      this.serverPort?.postMessage(response);
    });

    this.serverPort.postMessage({ type: "ready" } as WorkerOutboundMessage);
  }

  terminate(): void {
    this.terminated = true;
    this.serverPort?.close();
    this.serverPort = null;
  }
}

class ErrorSheetOriginWorker implements WorkerLike {
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

      if (msg.type !== "request") {
        return;
      }

      const req = msg as RpcRequest;
      if (req.method === "setSheetOrigin") {
        const response: WorkerOutboundMessage = {
          type: "response",
          id: req.id,
          ok: false,
          error: "setSheetOrigin failed"
        };
        this.serverPort?.postMessage(response);
        return;
      }

      const response: WorkerOutboundMessage = { type: "response", id: req.id, ok: true, result: null };
      this.serverPort?.postMessage(response);
    });

    this.serverPort.postMessage({ type: "ready" } as WorkerOutboundMessage);
  }

  terminate(): void {
    this.terminated = true;
    this.serverPort?.close();
    this.serverPort = null;
  }
}

describe("EngineWorker RPC", () => {
  it("supports ping RPC requests", async () => {
    const worker = new MockWorker();
    const engine = await EngineWorker.connect({
      worker,
      wasmModuleUrl: "mock://wasm",
      channelFactory: createMockChannel
    });

    const result = await engine.ping();
    expect(result).toBe("pong");

    const requests = worker.received.filter((msg): msg is RpcRequest => msg.type === "request");
    expect(requests[0]?.method).toBe("ping");
    expect(requests[0]?.params).toEqual({});
  });

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

  it("does not emit unhandled rejections when fire-and-forgetting a failing setCell flush", async () => {
    const worker = new ErrorSetCellsWorker();
    const engine = await EngineWorker.connect({
      worker,
      wasmModuleUrl: "mock://wasm",
      channelFactory: createMockChannel
    });

    // Intentionally fire-and-forget: the scheduled flush will reject because setCells fails.
    void engine.setCell("A1", 1);

    // Allow the microtask flush + response to run. Without internal handling, this would surface
    // as an unhandled rejection and fail the test run.
    await new Promise((resolve) => setTimeout(resolve, 0));

    engine.terminate();
  });

  it("does not emit unhandled rejections when fire-and-forgetting a failing setSheetOrigin", async () => {
    const worker = new ErrorSheetOriginWorker();
    const engine = await EngineWorker.connect({
      worker,
      wasmModuleUrl: "mock://wasm",
      channelFactory: createMockChannel
    });

    void engine.setSheetOrigin("Sheet1", "A1");
    await new Promise((resolve) => setTimeout(resolve, 0));

    engine.terminate();
  });

  it("rejects pending RPCs when the underlying worker emits an error after connect", async () => {
    class ErrorAfterReadyWorker implements WorkerLike {
      terminated = false;
      postMessageCalls = 0;
      private readonly errorListeners = new Set<(event: any) => void>();
      private port: MockMessagePort | null = null;

      addEventListener(type: string, listener: (event: any) => void): void {
        if (type === "error") this.errorListeners.add(listener);
      }

      removeEventListener(type: string, listener: (event: any) => void): void {
        if (type === "error") this.errorListeners.delete(listener);
      }

      postMessage(message: unknown): void {
        this.postMessageCalls += 1;
        const init = message as InitMessage;
        if (!init || typeof init !== "object" || (init as any).type !== "init") return;
        this.port = init.port as unknown as MockMessagePort;

        // Do not respond to requests; we'll trigger a fatal worker error instead.
        this.port.addEventListener("message", () => {});
        this.port.postMessage({ type: "ready" } as WorkerOutboundMessage);
      }

      emitError(message: string): void {
        for (const listener of this.errorListeners) {
          listener({ message });
        }
      }

      terminate(): void {
        this.terminated = true;
        try {
          this.port?.close();
        } catch {
          // ignore
        }
        this.port = null;
      }
    }

    const worker = new ErrorAfterReadyWorker();
    const engine = await EngineWorker.connect({
      worker,
      wasmModuleUrl: "mock://wasm",
      channelFactory: createMockChannel
    });

    const pingPromise = engine.ping();
    worker.emitError("crash");

    await expect(pingPromise).rejects.toThrow(/worker error/i);
    expect(worker.terminated).toBe(true);
  });

  it("rejects pending RPCs when the MessagePort emits messageerror", async () => {
    class NoResponseWorker implements WorkerLike {
      terminated = false;
      private port: MockMessagePort | null = null;

      postMessage(message: unknown): void {
        const init = message as InitMessage;
        if (!init || typeof init !== "object" || (init as any).type !== "init") return;
        this.port = init.port as unknown as MockMessagePort;
        this.port.addEventListener("message", () => {});
        this.port.postMessage({ type: "ready" } as WorkerOutboundMessage);
      }

      terminate(): void {
        this.terminated = true;
        try {
          this.port?.close();
        } catch {
          // ignore
        }
        this.port = null;
      }
    }

    const worker = new NoResponseWorker();
    const channel = createMockChannel();
    const engine = await EngineWorker.connect({
      worker,
      wasmModuleUrl: "mock://wasm",
      channelFactory: () => channel,
    });

    const pingPromise = engine.ping();
    const expectation = expect(pingPromise).rejects.toThrow(/messageerror/i);

    (channel.port1 as unknown as MockMessagePort).dispatchMessageError();

    await expectation;
    expect(worker.terminated).toBe(true);
  });

  it("rejects new RPCs after terminate() instead of hanging", async () => {
    const worker = new MockWorker();
    const engine = await EngineWorker.connect({
      worker,
      wasmModuleUrl: "mock://wasm",
      channelFactory: createMockChannel
    });

    engine.terminate();

    await expect(
      Promise.race([
        engine.ping().then(
          () => "resolved",
          (err) => err
        ),
        new Promise<"timeout">((resolve) => setTimeout(() => resolve("timeout"), 50))
      ])
    ).resolves.not.toBe("timeout");

    await expect(engine.ping()).rejects.toThrow(/terminated/i);
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

  it("transfers xlsx ArrayBuffer when loading encrypted workbooks from bytes", async () => {
    const worker = new MockWorker();
    const channel = createMockChannel();
    const engine = await EngineWorker.connect({
      worker,
      wasmModuleUrl: "mock://wasm",
      channelFactory: () => channel
    });

    const bytes = new Uint8Array([1, 2, 3, 4]);
    await engine.loadWorkbookFromEncryptedXlsxBytes(bytes, "secret");

    const requests = worker.received.filter(
      (msg): msg is RpcRequest => msg.type === "request" && (msg as RpcRequest).method === "loadFromEncryptedXlsxBytes"
    );
    expect(requests).toHaveLength(1);

    const clientPort = channel.port1 as unknown as MockMessagePort;
    const call = clientPort.sent.find((entry) => {
      const msg = entry.message as any;
      return msg?.type === "request" && msg?.method === "loadFromEncryptedXlsxBytes";
    });

    expect(call?.transfer).toHaveLength(1);
    expect(call?.transfer?.[0]).toBe(bytes.buffer);
    expect((call?.message as any)?.params?.password).toBe("secret");
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

  it("transfers only the view range when loading xlsx bytes from a Buffer slice", async () => {
    const worker = new MockWorker();
    const channel = createMockChannel();
    const engine = await EngineWorker.connect({
      worker,
      wasmModuleUrl: "mock://wasm",
      channelFactory: () => channel
    });

    const backing = Buffer.from([0, 0, 1, 2, 3, 4, 0, 0]);
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

  it("transfers only the view range when loading encrypted xlsx bytes from a Buffer slice", async () => {
    const worker = new MockWorker();
    const channel = createMockChannel();
    const engine = await EngineWorker.connect({
      worker,
      wasmModuleUrl: "mock://wasm",
      channelFactory: () => channel
    });

    const backing = Buffer.from([0, 0, 1, 2, 3, 4, 0, 0]);
    const view = backing.subarray(2, 6);
    await engine.loadWorkbookFromEncryptedXlsxBytes(view, "secret");

    const clientPort = channel.port1 as unknown as MockMessagePort;
    const call = clientPort.sent.find((entry) => {
      const msg = entry.message as any;
      return msg?.type === "request" && msg?.method === "loadFromEncryptedXlsxBytes";
    });

    expect(call?.transfer).toHaveLength(1);
    const transferred = call?.transfer?.[0] as ArrayBuffer | undefined;
    expect(transferred).toBeInstanceOf(ArrayBuffer);
    expect(transferred?.byteLength).toBe(view.byteLength);

    const paramsBytes = (call?.message as any)?.params?.bytes as Uint8Array | undefined;
    expect(paramsBytes).toBeInstanceOf(Uint8Array);
    expect(paramsBytes?.byteLength).toBe(view.byteLength);
    expect(paramsBytes?.buffer).toBe(transferred);
    expect((call?.message as any)?.params?.password).toBe("secret");
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

  it("flushes pending setCell batches before loading workbook from encrypted xlsx bytes", async () => {
    const worker = new MockWorker();
    const engine = await EngineWorker.connect({
      worker,
      wasmModuleUrl: "mock://wasm",
      channelFactory: createMockChannel
    });

    // Simulate a caller firing-and-forgetting setCell; loadWorkbookFromEncryptedXlsxBytes must
    // flush the pending microtask-batched setCells before replacing the workbook.
    void engine.setCell("A1", 1);

    await engine.loadWorkbookFromEncryptedXlsxBytes(new Uint8Array([1, 2, 3, 4]), "secret");

    const methods = worker.received
      .filter((msg): msg is RpcRequest => msg.type === "request")
      .map((msg) => msg.method);

    expect(methods).toEqual(["setCells", "loadFromEncryptedXlsxBytes"]);
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

  it("routes row/col/sheet style layer updates to the worker without throwing", async () => {
    const worker = new MockWorker();
    const engine = await EngineWorker.connect({
      worker,
      wasmModuleUrl: "mock://wasm",
      channelFactory: createMockChannel
    });

    await engine.setRowStyleId("Sheet1", 5, 123);
    await engine.setColStyleId("Sheet1", 2, 0);
    await engine.setSheetDefaultStyleId("Sheet1", 42);

    const requests = worker.received.filter((msg): msg is RpcRequest => msg.type === "request");
    const methods = requests.map((r) => r.method);
    expect(methods).toEqual(["setRowStyleId", "setColStyleId", "setSheetDefaultStyleId"]);
    expect(requests[0].params).toEqual({ sheet: "Sheet1", row: 5, styleId: 123 });
    expect(requests[1].params).toEqual({ sheet: "Sheet1", col: 2, styleId: 0 });
    expect(requests[2].params).toEqual({ sheet: "Sheet1", styleId: 42 });
  });

  it("supports sheet-first row/col/sheet style signatures (null clears)", async () => {
    const worker = new MockWorker();
    const engine = await EngineWorker.connect({
      worker,
      wasmModuleUrl: "mock://wasm",
      channelFactory: createMockChannel
    });

    await engine.setRowStyleId("Sheet1", 5, 123);
    await engine.setColStyleId("Sheet1", 2, null);
    await engine.setSheetDefaultStyleId("Sheet1", 42);

    const requests = worker.received.filter((msg): msg is RpcRequest => msg.type === "request");
    const methods = requests.map((r) => r.method);
    expect(methods).toEqual(["setRowStyleId", "setColStyleId", "setSheetDefaultStyleId"]);
    expect(requests[0].params).toEqual({ sheet: "Sheet1", row: 5, styleId: 123 });
    expect(requests[1].params).toEqual({ sheet: "Sheet1", col: 2, styleId: null });
    expect(requests[2].params).toEqual({ sheet: "Sheet1", styleId: 42 });
  });

  it("accepts legacy row/col/sheet style layer call signatures (styleId=0 clears)", async () => {
    const worker = new MockWorker();
    const engine = await EngineWorker.connect({
      worker,
      wasmModuleUrl: "mock://wasm",
      channelFactory: createMockChannel
    });

    await engine.setRowStyleId(5, 123, "Sheet1");
    await engine.setRowStyleId(6, 0, "Sheet1");
    await engine.setColStyleId(2, 456, "Sheet1");
    await engine.setColStyleId(3, 0, "Sheet1");
    await engine.setSheetDefaultStyleId(42, "Sheet1");
    await engine.setSheetDefaultStyleId(0, "Sheet1");

    const requests = worker.received.filter((msg): msg is RpcRequest => msg.type === "request");
    expect(requests.map((r) => r.method)).toEqual([
      "setRowStyleId",
      "setRowStyleId",
      "setColStyleId",
      "setColStyleId",
      "setSheetDefaultStyleId",
      "setSheetDefaultStyleId"
    ]);

    expect(requests[0].params).toEqual({ sheet: "Sheet1", row: 5, styleId: 123 });
    expect(requests[1].params).toEqual({ sheet: "Sheet1", row: 6, styleId: 0 });
    expect(requests[2].params).toEqual({ sheet: "Sheet1", col: 2, styleId: 456 });
    expect(requests[3].params).toEqual({ sheet: "Sheet1", col: 3, styleId: 0 });
    expect(requests[4].params).toEqual({ sheet: "Sheet1", styleId: 42 });
    expect(requests[5].params).toEqual({ sheet: "Sheet1", styleId: 0 });
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

  it("sends getRangeCompact RPC requests with the expected params", async () => {
    const worker = new MockWorker();
    const engine = await EngineWorker.connect({
      worker,
      wasmModuleUrl: "mock://wasm",
      channelFactory: createMockChannel
    });

    await engine.getRangeCompact("A1:B2", "Sheet1");

    const requests = worker.received.filter(
      (msg): msg is RpcRequest => msg.type === "request" && (msg as RpcRequest).method === "getRangeCompact"
    );
    expect(requests).toHaveLength(1);
    expect(requests[0].params).toEqual({ range: "A1:B2", sheet: "Sheet1" });
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

  it("supports supportedLocaleIds and getLocaleInfo module-level RPCs", async () => {
    const worker = new MockWorker();
    const engine = await EngineWorker.connect({
      worker,
      wasmModuleUrl: "mock://wasm",
      channelFactory: createMockChannel
    });

    const ids = await engine.supportedLocaleIds();
    expect(ids).toEqual(["de-DE", "en-US"]);

    const info = await engine.getLocaleInfo("de-DE");
    expect(info).toEqual({
      localeId: "de-DE",
      decimalSeparator: ",",
      argSeparator: ";",
      arrayRowSeparator: ";",
      arrayColSeparator: "\\",
      thousandsSeparator: ".",
      isRtl: false,
      booleanTrue: "WAHR",
      booleanFalse: "FALSCH"
    });

    const supportedReqs = worker.received.filter(
      (msg): msg is RpcRequest => msg.type === "request" && (msg as RpcRequest).method === "supportedLocaleIds"
    );
    expect(supportedReqs).toHaveLength(1);
    expect(supportedReqs[0].params).toEqual({});

    const infoReqs = worker.received.filter(
      (msg): msg is RpcRequest => msg.type === "request" && (msg as RpcRequest).method === "getLocaleInfo"
    );
    expect(infoReqs).toHaveLength(1);
    expect(infoReqs[0].params).toEqual({ localeId: "de-DE" });
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

  it("supports getCalcSettings / setCalcSettings RPCs", async () => {
    const worker = new MockWorker();
    const engine = await EngineWorker.connect({
      worker,
      wasmModuleUrl: "mock://wasm",
      channelFactory: createMockChannel
    });

    const settings = await engine.getCalcSettings();
    expect(settings).toEqual({
      calculationMode: "manual",
      calculateBeforeSave: true,
      fullPrecision: true,
      fullCalcOnLoad: false,
      iterative: { enabled: false, maxIterations: 100, maxChange: 0.001 }
    });

    await engine.setCalcSettings({
      calculationMode: "automatic",
      calculateBeforeSave: false,
      fullPrecision: true,
      fullCalcOnLoad: true,
      iterative: { enabled: true, maxIterations: 10, maxChange: 0.0001 }
    });

    const requests = worker.received.filter(
      (msg): msg is RpcRequest => msg.type === "request" && (msg as RpcRequest).method === "setCalcSettings"
    );
    expect(requests).toHaveLength(1);
    expect(requests[0].params).toEqual({
      settings: {
        calculationMode: "automatic",
        calculateBeforeSave: false,
        fullPrecision: true,
        fullCalcOnLoad: true,
        iterative: { enabled: true, maxIterations: 10, maxChange: 0.0001 }
      }
    });
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

  it("flushes pending setCell batches before setColWidthChars", async () => {
    const worker = new MockWorker();
    const engine = await EngineWorker.connect({
      worker,
      wasmModuleUrl: "mock://wasm",
      channelFactory: createMockChannel
    });

    // Fire-and-forget setCell leaves a pending microtask flush.
    void engine.setCell("A1", 1);

    await engine.setColWidthChars("Sheet1", 3, 8.43);

    const methods = worker.received
      .filter((msg): msg is RpcRequest => msg.type === "request")
      .map((msg) => msg.method);

    expect(methods).toEqual(["setCells", "setColWidthChars"]);
  });

  it("sends setColWidthChars RPC requests with the expected params", async () => {
    const worker = new MockWorker();
    const engine = await EngineWorker.connect({
      worker,
      wasmModuleUrl: "mock://wasm",
      channelFactory: createMockChannel
    });

    await engine.setColWidthChars("Sheet1", 3, 8.43);

    const requests = worker.received.filter(
      (msg): msg is RpcRequest => msg.type === "request" && (msg as RpcRequest).method === "setColWidthChars"
    );
    expect(requests).toHaveLength(1);
    expect(requests[0].params).toEqual({ sheet: "Sheet1", col: 3, widthChars: 8.43 });
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

  it("rejects connect immediately when the AbortSignal is already aborted (no init post)", async () => {
    const worker = new HangingWorker();
    const controller = new AbortController();
    controller.abort();

    await expect(
      EngineWorker.connect({
        worker,
        wasmModuleUrl: "mock://wasm",
        channelFactory: createMockChannel,
        signal: controller.signal
      })
    ).rejects.toThrow(/aborted/i);

    // Guard against leaking a message port by posting init even after abort.
    expect(worker.postMessageCalls).toBe(0);
    expect(worker.serverPort).toBeNull();
    expect(worker.terminated).toBe(true);
  });

  it("rejects connect and terminates the worker when aborted before the ready handshake", async () => {
    const worker = new HangingWorker();
    const controller = new AbortController();

    const promise = EngineWorker.connect({
      worker,
      wasmModuleUrl: "mock://wasm",
      channelFactory: createMockChannel,
      signal: controller.signal
    });

    // Ensure the init message was posted (we only abort after connect begins).
    expect(worker.postMessageCalls).toBe(1);
    expect(worker.serverPort).not.toBeNull();

    controller.abort();

    await expect(promise).rejects.toThrow(/aborted/i);
    expect(worker.terminated).toBe(true);
    expect(worker.serverPort).toBeNull();
  });

  it("rejects connect when the ready handshake does not arrive before timeoutMs", async () => {
    vi.useFakeTimers();
    const worker = new HangingWorker();
    const promise = EngineWorker.connect({
      worker,
      wasmModuleUrl: "mock://wasm",
      channelFactory: createMockChannel,
      timeoutMs: 50,
    });

    const expectation = expect(promise).rejects.toThrow(/timed out/i);
    await vi.advanceTimersByTimeAsync(50);
    await expectation;
    expect(worker.terminated).toBe(true);

    vi.useRealTimers();
  });

  it("rejects connect when the worker emits an error before the ready handshake", async () => {
    class ErroringWorker implements WorkerLike {
      terminated = false;
      postMessageCalls = 0;
      private readonly errorListeners = new Set<(event: any) => void>();
      private port: MockMessagePort | null = null;

      addEventListener(type: string, listener: (event: any) => void): void {
        if (type === "error") this.errorListeners.add(listener);
      }

      removeEventListener(type: string, listener: (event: any) => void): void {
        if (type === "error") this.errorListeners.delete(listener);
      }

      postMessage(message: unknown): void {
        this.postMessageCalls += 1;
        const init = message as InitMessage;
        if (!init || typeof init !== "object" || (init as any).type !== "init") {
          return;
        }
        this.port = init.port as unknown as MockMessagePort;
        queueMicrotask(() => {
          for (const listener of this.errorListeners) {
            listener({ message: "worker crashed" });
          }
        });
      }

      terminate(): void {
        this.terminated = true;
        try {
          this.port?.close();
        } catch {
          // ignore
        }
        this.port = null;
      }
    }

    const worker = new ErroringWorker();
    const channel = createMockChannel();

    await expect(
      EngineWorker.connect({
        worker,
        wasmModuleUrl: "mock://wasm",
        channelFactory: () => channel,
      })
    ).rejects.toThrow(/worker error/i);

    expect(worker.postMessageCalls).toBe(1);
    expect(worker.terminated).toBe(true);
    expect((channel.port1 as unknown as MockMessagePort).closed).toBe(true);
    expect((channel.port2 as unknown as MockMessagePort).closed).toBe(true);
  });

  it("cleans up ports + abort listener when Worker.postMessage throws", async () => {
    class ThrowingWorker implements WorkerLike {
      terminated = false;
      postMessageCalls = 0;
      postMessage(): void {
        this.postMessageCalls += 1;
        throw new Error("boom");
      }
      terminate(): void {
        this.terminated = true;
      }
    }

    class TrackingAbortSignal {
      aborted = false;
      listeners = new Set<() => void>();
      addEventListener(_type: string, listener: () => void): void {
        this.listeners.add(listener);
      }
      removeEventListener(_type: string, listener: () => void): void {
        this.listeners.delete(listener);
      }
    }

    const worker = new ThrowingWorker();
    const channel = createMockChannel();
    const signal = new TrackingAbortSignal();

    await expect(
      EngineWorker.connect({
        worker,
        wasmModuleUrl: "mock://wasm",
        channelFactory: () => channel,
        signal: signal as any
      })
    ).rejects.toThrow(/boom/);

    expect(worker.postMessageCalls).toBe(1);
    expect(worker.terminated).toBe(true);
    expect(signal.listeners.size).toBe(0);

    expect((channel.port1 as unknown as MockMessagePort).closed).toBe(true);
    expect((channel.port2 as unknown as MockMessagePort).closed).toBe(true);
  });

  it("does not throw if terminate() races with already-torn-down worker/ports", async () => {
    const worker = new MockWorker();
    const channel = createMockChannel();
    const port1 = channel.port1 as unknown as MockMessagePort;

    // Simulate buggy/hostile teardown implementations that throw even after performing cleanup.
    const originalClose = port1.close.bind(port1);
    (port1 as any).close = () => {
      originalClose();
      throw new Error("port close boom");
    };

    const engine = await EngineWorker.connect({
      worker,
      wasmModuleUrl: "mock://wasm",
      channelFactory: () => channel,
    });

    const originalTerminate = worker.terminate.bind(worker);
    (worker as any).terminate = () => {
      originalTerminate();
      throw new Error("worker terminate boom");
    };

    expect(() => engine.terminate()).not.toThrow();
  });

  it("removes the MessagePort message listener even when port.close() is unavailable", async () => {
    const worker = new MockWorker();
    const channel = createMockChannel();
    const port1 = channel.port1 as unknown as MockMessagePort;
    // Simulate an environment where MessagePortLike.close is not exposed.
    (port1 as any).close = undefined;

    const engine = await EngineWorker.connect({
      worker,
      wasmModuleUrl: "mock://wasm",
      channelFactory: () => channel,
    });

    expect(port1.getListenerCount()).toBeGreaterThan(0);

    engine.terminate();

    expect(port1.getListenerCount()).toBe(0);
  });
});
