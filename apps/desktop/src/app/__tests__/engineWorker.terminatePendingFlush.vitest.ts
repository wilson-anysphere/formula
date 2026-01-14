import { describe, expect, it } from "vitest";

import { EngineWorker } from "@formula/engine";

class MockMessagePort {
  private listeners = new Map<string, Set<(event: MessageEvent<unknown>) => void>>();
  private other: MockMessagePort | null = null;
  private closed = false;

  connect(other: MockMessagePort) {
    this.other = other;
  }

  postMessage(message: unknown): void {
    if (this.closed) {
      throw new Error("MockMessagePort: postMessage after close()");
    }
    queueMicrotask(() => {
      this.other?.dispatchMessage(message);
    });
  }

  start(): void {}

  close(): void {
    this.closed = true;
    this.listeners.clear();
    this.other = null;
  }

  addEventListener(type: string, listener: (event: MessageEvent<unknown>) => void): void {
    if (this.closed) return;
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
    if (this.closed) return;
    const event = { data } as MessageEvent<unknown>;
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

class MockWorker {
  private port: MockMessagePort | null = null;

  postMessage(message: unknown): void {
    const init = message as any;
    if (!init || typeof init !== "object" || init.type !== "init") {
      return;
    }

    this.port = init.port as unknown as MockMessagePort;

    // A real worker immediately posts a ready handshake on the MessagePort.
    this.port.postMessage({ type: "ready" });
  }

  terminate(): void {
    this.port?.close();
    this.port = null;
  }
}

describe("EngineWorker terminate() flush semantics", () => {
  it("does not emit an unhandled rejection when terminate races with a pending setCell microtask flush", async () => {
    const unhandled: unknown[] = [];
    const handler = (reason: unknown) => {
      unhandled.push(reason);
    };
    process.on("unhandledRejection", handler);

    try {
      const worker = new MockWorker();
      const engine = await EngineWorker.connect({
        worker: worker as any,
        wasmModuleUrl: "mock://wasm",
        channelFactory: createMockChannel as any,
      });

      // Fire-and-forget (simulates common UI code paths).
      void engine.setCell("A1", 1);

      // Terminate immediately (simulates teardown / strict-mode cleanup).
      engine.terminate();

      // Allow the scheduled microtask flush (and Node's unhandledRejection bookkeeping) to run.
      await new Promise((resolve) => setTimeout(resolve, 0));

      expect(unhandled).toEqual([]);
      expect((engine as any).pendingCellUpdates).toEqual([]);
      expect((engine as any).flushPromise).toBeNull();
    } finally {
      process.off("unhandledRejection", handler);
    }
  });
});
