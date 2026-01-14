import { afterAll, describe, expect, it } from "vitest";

import type {
  InitMessage,
  RpcCancel,
  RpcRequest,
  RpcResponseErr,
  RpcResponseOk,
  WorkerOutboundMessage
} from "../protocol.ts";

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
  closed = false;
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

  private dispatchMessage(data: unknown): void {
    const event = { data } as MessageEvent<unknown>;
    this.onmessage?.(event);
    for (const listener of this.listeners.get("message") ?? []) {
      listener(event);
    }
  }
}

function createMockChannel(): { port1: MockMessagePort; port2: MockMessagePort } {
  const port1 = new MockMessagePort();
  const port2 = new MockMessagePort();
  port1.connect(port2);
  port2.connect(port1);
  return { port1, port2 };
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

async function sendRequest(port: MockMessagePort, req: RpcRequest): Promise<RpcResponseOk | RpcResponseErr> {
  const responsePromise = waitForMessage(port, (msg) => msg.type === "response" && msg.id === req.id) as Promise<
    RpcResponseOk | RpcResponseErr
  >;
  port.postMessage(req);
  return await responsePromise;
}

async function waitFor(condition: () => boolean, timeoutMs = 1_000): Promise<void> {
  const start = Date.now();
  while (!condition()) {
    if (Date.now() - start > timeoutMs) {
      throw new Error("Timed out waiting for condition");
    }
    await new Promise((resolve) => setTimeout(resolve, 0));
  }
}

describe("engine.worker resilience", () => {
  afterAll(() => {
    (globalThis as any).self = previousSelf;
  });

  it("responds with an error when a response cannot be structured-cloned (DataCloneError)", async () => {
    await loadWorkerModule();

    const channel = createMockChannel();
    const originalPostMessage = channel.port2.postMessage.bind(channel.port2);
    let responseAttempts = 0;
    channel.port2.postMessage = (message: unknown) => {
      if ((message as any)?.type === "response") {
        responseAttempts += 1;
        if (responseAttempts === 1) {
          const err = new Error("DataCloneError");
          (err as any).name = "DataCloneError";
          throw err;
        }
      }
      originalPostMessage(message);
    };

    const wasmModuleUrl = new URL("./fixtures/mockWasmWorkbookMetadata.mjs", import.meta.url).href;
    workerGlobal.dispatchMessage({
      type: "init",
      port: channel.port2 as unknown as MessagePort,
      wasmModuleUrl,
    } satisfies InitMessage);

    await waitForMessage(channel.port1, (msg) => msg.type === "ready");

    const response1 = await sendRequest(channel.port1, { type: "request", id: 1, method: "ping", params: {} });
    expect((response1 as RpcResponseErr).ok).toBe(false);
    expect((response1 as RpcResponseErr).error).toMatch(/datacloneerror/i);

    const response2 = await sendRequest(channel.port1, { type: "request", id: 2, method: "ping", params: {} });
    expect((response2 as RpcResponseOk).ok).toBe(true);
    expect((response2 as RpcResponseOk).result).toBe("pong");

    channel.port1.close();
  });

  it("does not leak cancellation state when posting a response throws", async () => {
    await loadWorkerModule();

    const channel = createMockChannel();
    let responseAttempts = 0;
    const originalPostMessage = channel.port2.postMessage.bind(channel.port2);
    channel.port2.postMessage = (message: unknown) => {
      const msgType = (message as any)?.type;
      if (msgType === "response") {
        responseAttempts += 1;
        // First response attempt throws (simulating a closed port). Subsequent responses deliver normally.
        if (responseAttempts === 1) {
          throw new Error("postMessage boom");
        }
      }
      originalPostMessage(message);
    };

    const wasmModuleUrl = new URL("./fixtures/mockWasmWorkbookMetadata.mjs", import.meta.url).href;
    const init: InitMessage = {
      type: "init",
      port: channel.port2 as unknown as MessagePort,
      wasmModuleUrl,
    };
    workerGlobal.dispatchMessage(init);

    await waitForMessage(channel.port1, (msg) => msg.type === "ready");

    // First request: the worker will attempt to post a response and throw. We don't wait for the response
    // (it won't arrive), but we do wait until the worker attempted to post it so we know request handling
    // reached the response stage.
    channel.port1.postMessage({ type: "request", id: 1, method: "ping", params: {} });
    await waitFor(() => responseAttempts === 1);

    // Cancel after the request finishes. If the worker failed to mark the request as completed (e.g. it
    // threw before cleanup), this cancel could poison future requests with the same id by leaving the id
    // in `cancelledRequests`.
    const cancel: RpcCancel = { type: "cancel", id: 1 };
    channel.port1.postMessage(cancel);

    // Second request (id reused for test determinism): should still get a response now that `postMessage`
    // is working again.
    const response = await Promise.race([
      sendRequest(channel.port1, { type: "request", id: 1, method: "ping", params: {} }),
      new Promise<"timeout">((resolve) => setTimeout(() => resolve("timeout"), 250)),
    ]);

    expect(response).not.toBe("timeout");
    expect((response as RpcResponseOk).ok).toBe(true);
    expect((response as RpcResponseOk).result).toBe("pong");

    channel.port1.close();
  });

  it("clears cancellation state and closes the previous port on re-init", async () => {
    await loadWorkerModule();

    const wasmModuleUrl = new URL("./fixtures/mockWasmWorkbookMetadata.mjs", import.meta.url).href;

    const channel1 = createMockChannel();
    workerGlobal.dispatchMessage({
      type: "init",
      port: channel1.port2 as unknown as MessagePort,
      wasmModuleUrl,
    } satisfies InitMessage);
    await waitForMessage(channel1.port1, (msg) => msg.type === "ready");

    // Send a cancel for an id that will never be requested on this connection.
    channel1.port1.postMessage({ type: "cancel", id: 42 } satisfies RpcCancel);
    // Allow the cancel microtask to be processed so it actually registers in worker state.
    await new Promise((resolve) => setTimeout(resolve, 0));

    const channel2 = createMockChannel();
    workerGlobal.dispatchMessage({
      type: "init",
      port: channel2.port2 as unknown as MessagePort,
      wasmModuleUrl,
    } satisfies InitMessage);
    await waitForMessage(channel2.port1, (msg) => msg.type === "ready");

    // The new init should close the previous port so it doesn't leak or keep delivering messages.
    expect(channel1.port2.closed).toBe(true);

    // Reuse the cancelled id on the new connection. If cancellation state wasn't reset, this request
    // would be dropped without a response.
    const response = await Promise.race([
      sendRequest(channel2.port1, { type: "request", id: 42, method: "ping", params: {} }),
      new Promise<"timeout">((resolve) => setTimeout(() => resolve("timeout"), 250)),
    ]);

    expect(response).not.toBe("timeout");
    expect((response as RpcResponseOk).ok).toBe(true);
    expect((response as RpcResponseOk).result).toBe("pong");

    channel1.port1.close();
    channel2.port1.close();
  });

  it("does not deliver responses from an in-flight request to a new port after re-init", async () => {
    await loadWorkerModule();

    let resolveInit: (() => void) | null = null;
    (globalThis as any).__ENGINE_WORKER_DELAY_INIT_PROMISE__ = new Promise<void>((resolve) => {
      resolveInit = resolve;
    });

    const delayedWasmModuleUrl = new URL("./fixtures/mockWasmWorkbookDelayedInit.mjs", import.meta.url).href;
    const fastWasmModuleUrl = new URL("./fixtures/mockWasmWorkbookMetadata.mjs", import.meta.url).href;

    const channel1 = createMockChannel();
    workerGlobal.dispatchMessage({
      type: "init",
      port: channel1.port2 as unknown as MessagePort,
      wasmModuleUrl: delayedWasmModuleUrl,
    } satisfies InitMessage);
    await waitForMessage(channel1.port1, (msg) => msg.type === "ready");

    // Start an in-flight request that will block while the wasm module init is delayed.
    channel1.port1.postMessage({ type: "request", id: 1, method: "ping", params: {} } satisfies RpcRequest);

    const channel2 = createMockChannel();
    workerGlobal.dispatchMessage({
      type: "init",
      port: channel2.port2 as unknown as MessagePort,
      wasmModuleUrl: fastWasmModuleUrl,
    } satisfies InitMessage);
    await waitForMessage(channel2.port1, (msg) => msg.type === "ready");

    // New connection should close the old port so it doesn't leak handles.
    expect(channel1.port2.closed).toBe(true);

    // Ensure the new port is functional.
    const response2 = await sendRequest(channel2.port1, { type: "request", id: 2, method: "ping", params: {} });
    expect((response2 as RpcResponseOk).ok).toBe(true);

    // Release the delayed init; the old request may resume execution, but must not respond on the
    // new port (generation guards should drop it).
    resolveInit?.();
    await new Promise((resolve) => setTimeout(resolve, 0));

    const unexpected = await Promise.race([
      waitForMessage(channel2.port1, (msg) => msg.type === "response" && msg.id === 1),
      new Promise<"timeout">((resolve) => setTimeout(() => resolve("timeout"), 50)),
    ]);
    expect(unexpected).toBe("timeout");

    delete (globalThis as any).__ENGINE_WORKER_DELAY_INIT_PROMISE__;
    channel1.port1.close();
    channel2.port1.close();
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
