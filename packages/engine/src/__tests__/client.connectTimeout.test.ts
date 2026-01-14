import { describe, expect, it, vi } from "vitest";

import { createEngineClient } from "../client.ts";

type InitMessage = { type: "init"; port: MessagePort };

class TestWorker {
  static mode: "hang" | "ready" = "hang";

  private port: MessagePort | null = null;

  constructor(_url: any, _options?: any) {}

  postMessage(message: unknown): void {
    const init = message as InitMessage;
    if (!init || typeof init !== "object" || (init as any).type !== "init") return;
    this.port = init.port;

    if (TestWorker.mode === "ready") {
      queueMicrotask(() => {
        try {
          this.port?.postMessage({ type: "ready" });
        } catch {
          // ignore
        }
      });
    }
  }

  terminate(): void {
    try {
      this.port?.close();
    } catch {
      // ignore
    }
    this.port = null;
  }
}

describe("createEngineClient() connect timeout", () => {
  it("rejects init() when the worker never sends the ready handshake within connectTimeoutMs", async () => {
    vi.useFakeTimers();
    const originalWorker = (globalThis as any).Worker;
    (globalThis as any).Worker = TestWorker;

    try {
      TestWorker.mode = "hang";
      const engine = createEngineClient({
        wasmModuleUrl: "mock://wasm",
        wasmBinaryUrl: "mock://wasm_bg.wasm",
        connectTimeoutMs: 50,
      });

      const initPromise = engine.init();
      const expectation = expect(initPromise).rejects.toThrow(/timed out/i);
      await vi.advanceTimersByTimeAsync(50);
      await expectation;
    } finally {
      (globalThis as any).Worker = originalWorker;
      vi.useRealTimers();
    }
  });

  it("can re-init after a connect timeout (fresh Worker)", async () => {
    vi.useFakeTimers();
    const originalWorker = (globalThis as any).Worker;
    (globalThis as any).Worker = TestWorker;

    try {
      TestWorker.mode = "hang";
      const engine = createEngineClient({
        wasmModuleUrl: "mock://wasm",
        wasmBinaryUrl: "mock://wasm_bg.wasm",
        connectTimeoutMs: 50,
      });

      const initPromise = engine.init();
      const expectation = expect(initPromise).rejects.toThrow(/timed out/i);
      await vi.advanceTimersByTimeAsync(50);
      await expectation;

      TestWorker.mode = "ready";
      await expect(engine.init()).resolves.toBeUndefined();
    } finally {
      (globalThis as any).Worker = originalWorker;
      vi.useRealTimers();
    }
  });
});
