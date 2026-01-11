import { describe, expect, it, vi } from "vitest";

import { EngineWorker, type MessageChannelLike, type WorkerLike } from "../worker/EngineWorker";
import type {
  InitMessage,
  RpcRequest,
  WorkerInboundMessage,
  WorkerOutboundMessage
} from "../protocol";

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

  it("sends loadFromXlsxBytes when loading workbook from xlsx bytes", async () => {
    const worker = new MockWorker();
    const engine = await EngineWorker.connect({
      worker,
      wasmModuleUrl: "mock://wasm",
      channelFactory: createMockChannel
    });

    const bytes = new Uint8Array([1, 2, 3, 4]);
    await engine.loadWorkbookFromXlsxBytes(bytes);

    const requests = worker.received.filter(
      (msg): msg is RpcRequest => msg.type === "request" && (msg as RpcRequest).method === "loadFromXlsxBytes"
    );

    expect(requests).toHaveLength(1);
    expect((requests[0].params as any).bytes).toEqual(bytes);
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
});
