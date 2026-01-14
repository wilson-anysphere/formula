import { describe, expect, it } from "vitest";

import { formulaWasmNodeEntryUrl } from "../../../../scripts/build-formula-wasm-node.mjs";

import { EngineWorker, type MessageChannelLike, type WorkerLike } from "../worker/EngineWorker.ts";
import type {
  CellChange,
  GoalSeekResponse,
  InitMessage,
  RpcRequest,
  WorkerInboundMessage,
  WorkerOutboundMessage,
} from "../protocol.ts";

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

function createMockChannel(): MessageChannelLike {
  const port1 = new MockMessagePort();
  const port2 = new MockMessagePort();
  port1.connect(port2);
  port2.connect(port1);
  return { port1, port2: port2 as unknown as MessagePort };
}

async function loadFormulaWasm() {
  const entry = formulaWasmNodeEntryUrl();
  // wasm-pack `--target nodejs` outputs CommonJS. Under ESM dynamic import, the exports
  // are exposed on `default`.
  // eslint-disable-next-line @typescript-eslint/ban-ts-comment
  // @ts-ignore - `@vite-ignore` is required for runtime-defined file URLs.
  const mod = await import(/* @vite-ignore */ entry);
  return (mod as any).default ?? mod;
}

class WasmBackedWorker implements WorkerLike {
  private readonly wasm: any;
  private port: MockMessagePort | null = null;
  private workbook: any | null = null;

  constructor(wasm: any) {
    this.wasm = wasm;
  }

  postMessage(message: unknown): void {
    const init = message as InitMessage;
    if (!init || typeof init !== "object" || (init as any).type !== "init") {
      return;
    }

    this.port = init.port as unknown as MockMessagePort;
    this.workbook = new this.wasm.WasmWorkbook();

    this.port.addEventListener("message", (event) => {
      const msg = event.data as WorkerInboundMessage;
      if (!msg || typeof msg !== "object" || (msg as any).type !== "request") {
        return;
      }

      const req = msg as RpcRequest;
      const params = req.params as any;

      try {
        let result: unknown;
        switch (req.method) {
          case "newWorkbook":
            this.workbook = new this.wasm.WasmWorkbook();
            result = null;
            break;
          case "setCells":
            if (typeof this.workbook?.setCells === "function") {
              this.workbook?.setCells(params.updates);
            } else {
              for (const update of params.updates as Array<any>) {
                this.workbook?.setCell(update.address, update.value, update.sheet);
              }
            }
            result = null;
            break;
          case "goalSeek":
            if (typeof this.workbook?.goalSeek !== "function") {
              throw new Error("goalSeek: not supported by this wasm build");
            }
            result = this.workbook.goalSeek(params);
            break;
          default:
            throw new Error(`unsupported method: ${req.method}`);
        }

        const response: WorkerOutboundMessage = { type: "response", id: req.id, ok: true, result };
        this.port?.postMessage(response);
      } catch (err) {
        const response: WorkerOutboundMessage = {
          type: "response",
          id: req.id,
          ok: false,
          error: err instanceof Error ? err.message : String(err),
        };
        this.port?.postMessage(response);
      }
    });

    const ready: WorkerOutboundMessage = { type: "ready" };
    this.port.postMessage(ready);
  }

  terminate(): void {
    this.port?.close();
    this.port = null;
    this.workbook = null;
  }
}

const skipWasmBuild = process.env.FORMULA_SKIP_WASM_BUILD === "1" || process.env.FORMULA_SKIP_WASM_BUILD === "true";
const describeWasm = skipWasmBuild ? describe.skip : describe;

describeWasm("EngineWorker goalSeek (wasm)", () => {
  it("returns { result, changes } and converges on a solution", async () => {
    const wasm = await loadFormulaWasm();
    const worker = new WasmBackedWorker(wasm);
    const engine = await EngineWorker.connect({
      worker,
      wasmModuleUrl: "mock://wasm",
      channelFactory: createMockChannel,
    });

    try {
      await engine.newWorkbook();

      await engine.setCell("A1", 1, "Sheet1");
      await engine.setCell("B1", "=A1*A1", "Sheet1");

      const response: GoalSeekResponse = await engine.goalSeek({
        sheet: "Sheet1",
        targetCell: "B1",
        targetValue: 25,
        changingCell: "A1",
      });

      expect(response).toHaveProperty("result");
      expect(response).toHaveProperty("changes");

      expect(response.result.status).toBe("Converged");
      expect(Number.isFinite(response.result.solution)).toBe(true);
      expect(Math.abs(response.result.solution - 5)).toBeLessThan(0.01);
      expect(Math.abs(response.result.finalOutput - 25)).toBeLessThan(0.01);
      expect(Math.abs(response.result.finalError)).toBeLessThan(0.01);

      const changes: CellChange[] = response.changes;
      expect(changes.length).toBeGreaterThan(0);

      const a1 = changes.find((c) => c.sheet === "Sheet1" && c.address === "A1");
      expect(a1).toBeTruthy();
      expect(typeof a1?.value).toBe("number");
      expect(Math.abs((a1?.value as number) - 5)).toBeLessThan(0.01);

      const b1 = changes.find((c) => c.sheet === "Sheet1" && c.address === "B1");
      expect(b1).toBeTruthy();
      expect(typeof b1?.value).toBe("number");
      expect(Math.abs((b1?.value as number) - 25)).toBeLessThan(0.01);
    } finally {
      engine.terminate();
    }
  });
});
