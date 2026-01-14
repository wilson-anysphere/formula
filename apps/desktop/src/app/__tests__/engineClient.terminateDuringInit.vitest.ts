import { describe, expect, it } from "vitest";

import { createEngineClient } from "@formula/engine";

type InitMessage = { type: "init"; port: MessagePort };

// Minimal Worker stub that can either hang (never post "ready") or respond immediately.
class TestWorker {
  static mode: "hang" | "ready" = "hang";

  private port: MessagePort | null = null;

  constructor(_url: any, _options?: any) {}

  postMessage(message: unknown): void {
    const init = message as InitMessage;
    if (!init || typeof init !== "object" || (init as any).type !== "init") return;
    this.port = init.port;

    if (TestWorker.mode === "ready") {
      // Simulate the worker handshake expected by EngineWorker.connect.
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

describe("createEngineClient() teardown", () => {
  it("rejects init() when terminate() is called before the worker 'ready' handshake", async () => {
    const originalWorker = (globalThis as any).Worker;
    (globalThis as any).Worker = TestWorker;

    try {
      TestWorker.mode = "hang";
      const engine = createEngineClient({ wasmModuleUrl: "mock://wasm", wasmBinaryUrl: "mock://wasm_bg.wasm" });

      const initPromise = engine.init();
      engine.terminate();

      // Fail fast if init never settles (regression guard against hung connect promises).
      const settled = await Promise.race([
        initPromise.then(
          () => "resolved",
          (err) => err
        ),
        new Promise<"timeout">((resolve) => setTimeout(() => resolve("timeout"), 100)),
      ]);

      expect(settled).not.toBe("resolved");
      expect(settled).not.toBe("timeout");
      expect(String((settled as any)?.message ?? settled)).toMatch(/terminated/i);
    } finally {
      (globalThis as any).Worker = originalWorker;
    }
  });

  it("can re-init after terminate() (new worker instance)", async () => {
    const originalWorker = (globalThis as any).Worker;
    (globalThis as any).Worker = TestWorker;

    try {
      TestWorker.mode = "hang";
      const engine = createEngineClient({ wasmModuleUrl: "mock://wasm", wasmBinaryUrl: "mock://wasm_bg.wasm" });
      const initPromise = engine.init();
      engine.terminate();
      await expect(initPromise).rejects.toThrow(/terminated/i);

      // Switch the stub worker to respond with "ready" for the next connection.
      TestWorker.mode = "ready";
      await expect(engine.init()).resolves.toBeUndefined();
    } finally {
      (globalThis as any).Worker = originalWorker;
    }
  });
});

