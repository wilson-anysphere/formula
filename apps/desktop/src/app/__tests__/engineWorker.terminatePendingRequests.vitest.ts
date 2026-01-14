import { describe, expect, it, vi } from "vitest";

import { EngineWorker } from "@formula/engine";

class TrackingAbortSignal {
  aborted = false;
  added: Array<{ type: string; listener: any }> = [];
  removed: Array<{ type: string; listener: any }> = [];

  addEventListener(type: string, listener: any): void {
    this.added.push({ type, listener });
  }

  removeEventListener(type: string, listener: any): void {
    this.removed.push({ type, listener });
  }
}

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

class MockWorkerNoResponse {
  private port: MockMessagePort | null = null;

  postMessage(message: unknown): void {
    const init = message as any;
    if (!init || typeof init !== "object" || init.type !== "init") {
      return;
    }

    this.port = init.port as unknown as MockMessagePort;
    this.port.addEventListener("message", () => {
      // Intentionally do not respond to requests (so they remain pending).
    });
    this.port.postMessage({ type: "ready" });
  }

  terminate(): void {
    this.port?.close();
    this.port = null;
  }
}

describe("EngineWorker terminate() cleanup", () => {
  it("clears timeouts + abort listeners for pending requests", async () => {
    vi.useFakeTimers();
    const signal = new TrackingAbortSignal();

    try {
      const worker = new MockWorkerNoResponse();
      const engine = await EngineWorker.connect({
        worker: worker as any,
        wasmModuleUrl: "mock://wasm",
        channelFactory: createMockChannel as any,
      });

      const promise = engine.getCell("A1", undefined, { timeoutMs: 10_000, signal: signal as any });

      // Allow the async getCell() to pass its `await this.flush()` and schedule the request.
      await Promise.resolve();

      expect(vi.getTimerCount()).toBe(1);
      expect(signal.added).toHaveLength(1);

      engine.terminate();

      await expect(promise).rejects.toThrow(/worker terminated/i);
      expect(vi.getTimerCount()).toBe(0);
      expect(signal.removed).toHaveLength(1);
    } finally {
      vi.useRealTimers();
    }
  });
});
