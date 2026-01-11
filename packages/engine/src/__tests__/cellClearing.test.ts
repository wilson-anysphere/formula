import { describe, expect, it } from "vitest";

import path from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";

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

async function loadFormulaWasm() {
  const __filename = fileURLToPath(import.meta.url);
  const __dirname = path.dirname(__filename);
  const repoRoot = path.resolve(__dirname, "../../../..");

  const entry = pathToFileURL(
    path.join(repoRoot, "crates", "formula-wasm", "pkg-node", "formula_wasm.js")
  ).href;

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
          case "loadFromJson":
            this.workbook = this.wasm.WasmWorkbook.fromJson(params.json);
            result = null;
            break;
          case "toJson":
            result = this.workbook?.toJson();
            break;
          case "getCell":
            result = this.workbook?.getCell(params.address, params.sheet);
            break;
          case "setCells":
            for (const update of params.updates as Array<any>) {
              this.workbook?.setCell(update.address, update.value, update.sheet);
            }
            result = null;
            break;
          case "setRange":
            this.workbook?.setRange(params.range, params.values, params.sheet);
            result = null;
            break;
          case "recalculate":
            result = this.workbook?.recalculate(params.sheet);
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
          error: err instanceof Error ? err.message : String(err)
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

describe("EngineWorker null clear semantics", () => {
  it("treats setCell(..., null) as clearing the cell (sparse semantics)", async () => {
    const wasm = await loadFormulaWasm();
    const worker = new WasmBackedWorker(wasm);

    const engine = await EngineWorker.connect({
      worker,
      wasmModuleUrl: "mock://wasm",
      channelFactory: createMockChannel
    });

    try {
      await engine.newWorkbook();
      await Promise.all([engine.setCell("A1", 1), engine.setCell("A2", "=A1*2")]);
      await engine.recalculate();
      expect((await engine.getCell("A2")).value).toBe(2);

      await engine.setCell("A1", null);
      const changes = await engine.recalculate();

      expect(changes).toEqual([{ sheet: "Sheet1", address: "A2", value: 0 }]);

      const a1 = await engine.getCell("A1");
      expect(a1.input).toBeNull();
      expect(a1.value).toBeNull();

      expect((await engine.getCell("A2")).value).toBe(0);

      const exported = JSON.parse(await engine.toJson());
      expect(exported.sheets.Sheet1.cells).not.toHaveProperty("A1");
    } finally {
      engine.terminate();
    }
  });

  it("reports dynamic array spill outputs as recalc changes", async () => {
    const wasm = await loadFormulaWasm();
    const worker = new WasmBackedWorker(wasm);

    const engine = await EngineWorker.connect({
      worker,
      wasmModuleUrl: "mock://wasm",
      channelFactory: createMockChannel
    });

    try {
      await engine.newWorkbook();
      await engine.setCell("A1", "=SEQUENCE(1,2)");

      const changes = await engine.recalculate();
      expect(changes).toEqual([
        { sheet: "Sheet1", address: "A1", value: 1 },
        { sheet: "Sheet1", address: "B1", value: 2 }
      ]);

      expect((await engine.getCell("A1")).value).toBe(1);
      const b1 = await engine.getCell("B1");
      expect(b1.input).toBeNull();
      expect(b1.value).toBe(2);
    } finally {
      engine.terminate();
    }
  });

  it("clears spill output cells when a spill cell is overwritten", async () => {
    const wasm = await loadFormulaWasm();
    const worker = new WasmBackedWorker(wasm);

    const engine = await EngineWorker.connect({
      worker,
      wasmModuleUrl: "mock://wasm",
      channelFactory: createMockChannel
    });

    try {
      await engine.newWorkbook();
      await engine.setCell("A1", "=SEQUENCE(1,3)");
      await engine.recalculate();

      // Overwrite a spill output cell with a literal. The spill origin should become #SPILL!,
      // and any remaining spill outputs should be cleared back to blank (null).
      await engine.setCell("B1", 5);
      const changes = await engine.recalculate();

      expect(changes).toEqual([
        { sheet: "Sheet1", address: "A1", value: "#SPILL!" },
        { sheet: "Sheet1", address: "C1", value: null }
      ]);

      const b1 = await engine.getCell("B1");
      expect(b1.input).toBe(5);
      expect(b1.value).toBe(5);
      expect((await engine.getCell("C1")).value).toBeNull();
    } finally {
      engine.terminate();
    }
  });

  it("filters recalc changes by sheet name (case-insensitive)", async () => {
    const wasm = await loadFormulaWasm();
    const worker = new WasmBackedWorker(wasm);

    const engine = await EngineWorker.connect({
      worker,
      wasmModuleUrl: "mock://wasm",
      channelFactory: createMockChannel
    });

    try {
      await engine.newWorkbook();
      await engine.setCell("A1", 1, "Sheet1");
      await engine.setCell("A2", "=A1*2", "Sheet1");

      await engine.setCell("A1", 3, "Sheet2");
      await engine.setCell("A2", "=A1*2", "Sheet2");

      await engine.recalculate();
      expect((await engine.getCell("A2", "Sheet1")).value).toBe(2);
      expect((await engine.getCell("A2", "Sheet2")).value).toBe(6);

      await engine.setCell("A1", 4, "Sheet2");
      const changes = await engine.recalculate("sHeEt2");

      expect(changes).toEqual([{ sheet: "Sheet2", address: "A2", value: 8 }]);
      expect((await engine.getCell("A2", "Sheet2")).value).toBe(8);
      // Sheet1 should not have been included in the delta list.
      expect((await engine.getCell("A2", "Sheet1")).value).toBe(2);
    } finally {
      engine.terminate();
    }
  });

  it("sheet-scoped recalc still updates cross-sheet dependents in the engine", async () => {
    const wasm = await loadFormulaWasm();
    const worker = new WasmBackedWorker(wasm);

    const engine = await EngineWorker.connect({
      worker,
      wasmModuleUrl: "mock://wasm",
      channelFactory: createMockChannel
    });

    try {
      await engine.newWorkbook();
      await engine.setCell("A1", 1, "Sheet1");
      await engine.setCell("A1", "=Sheet1!A1*2", "Sheet2");
      await engine.recalculate();
      expect((await engine.getCell("A1", "Sheet2")).value).toBe(2);

      await engine.setCell("A1", 2, "Sheet1");
      const changes = await engine.recalculate("Sheet1");
      // The sheet-scoped delta list is filtered, but the engine still recalculates Sheet2.
      expect(changes).toEqual([]);
      expect((await engine.getCell("A1", "Sheet2")).value).toBe(4);
    } finally {
      engine.terminate();
    }
  });

  it("reports formula edits that clear a previously displayed value", async () => {
    const wasm = await loadFormulaWasm();
    const worker = new WasmBackedWorker(wasm);

    const engine = await EngineWorker.connect({
      worker,
      wasmModuleUrl: "mock://wasm",
      channelFactory: createMockChannel
    });

    try {
      await engine.newWorkbook();
      await engine.setCell("A1", "=1");
      await engine.recalculate();
      expect((await engine.getCell("A1")).value).toBe(1);

      // Edit formula to reference a blank cell; the new result is blank so the recalc delta must
      // still report {A1: null} to clear the old cached value.
      await engine.setCell("A1", "=A2");
      const changes = await engine.recalculate();
      expect(changes).toEqual([{ sheet: "Sheet1", address: "A1", value: null }]);
      expect((await engine.getCell("A1")).value).toBeNull();
    } finally {
      engine.terminate();
    }
  });

  it("treats explicit null cells in JSON as absent and omits them on export", async () => {
    const wasm = await loadFormulaWasm();
    const worker = new WasmBackedWorker(wasm);

    const engine = await EngineWorker.connect({
      worker,
      wasmModuleUrl: "mock://wasm",
      channelFactory: createMockChannel
    });

    try {
      await engine.loadWorkbookFromJson(
        JSON.stringify({
          sheets: {
            Sheet1: {
              cells: {
                A1: null,
                A2: "=A1*2"
              }
            }
          }
        })
      );

      await engine.recalculate();
      expect((await engine.getCell("A1")).input).toBeNull();
      expect((await engine.getCell("A2")).value).toBe(0);

      const exported = JSON.parse(await engine.toJson());
      expect(exported.sheets.Sheet1.cells).not.toHaveProperty("A1");
      expect(exported.sheets.Sheet1.cells).toHaveProperty("A2");
    } finally {
      engine.terminate();
    }
  });

  it("clears null entries passed to setRange and updates dependents", async () => {
    const wasm = await loadFormulaWasm();
    const worker = new WasmBackedWorker(wasm);

    const engine = await EngineWorker.connect({
      worker,
      wasmModuleUrl: "mock://wasm",
      channelFactory: createMockChannel
    });

    try {
      await engine.newWorkbook();
      await engine.setRange("A1:B1", [[1, 2]]);
      await engine.setCell("C1", "=A1+B1");
      await engine.recalculate();
      expect((await engine.getCell("C1")).value).toBe(3);

      await engine.setRange("A1", [[null]]);
      const changes = await engine.recalculate();
      expect(changes).toEqual(expect.arrayContaining([{ sheet: "Sheet1", address: "C1", value: 2 }]));

      const a1 = await engine.getCell("A1");
      expect(a1.input).toBeNull();
      expect(a1.value).toBeNull();

      const exported = JSON.parse(await engine.toJson());
      expect(exported.sheets.Sheet1.cells).not.toHaveProperty("A1");
    } finally {
      engine.terminate();
    }
  });
});
