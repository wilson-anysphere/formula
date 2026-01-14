import { describe, expect, it } from "vitest";

import { formulaWasmNodeEntryUrl } from "../../../../scripts/build-formula-wasm-node.mjs";

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
  // The Node-compatible wasm-bindgen build lives under `crates/formula-wasm/pkg-node/`.
  // It's a generated (gitignored) directory. Vitest's global setup (`scripts/vitest.global-setup.mjs`)
  // builds/refreshes it for CI runs. If a test run skips that setup (e.g. by setting
  // `FORMULA_SKIP_WASM_BUILD=1` or `FORMULA_SKIP_WASM_BUILD=true`), fail fast with a helpful error instead of trying to
  // run a slow build inside a Vitest worker thread (which can trigger RPC timeouts).
  const entry = formulaWasmNodeEntryUrl();

  // wasm-pack `--target nodejs` outputs CommonJS. Under ESM dynamic import, the exports
  // are exposed on `default`.
  // eslint-disable-next-line @typescript-eslint/ban-ts-comment
  // @ts-ignore - `@vite-ignore` is required for runtime-defined file URLs.
  try {
    const mod = await import(/* @vite-ignore */ entry);
    return (mod as any).default ?? mod;
  } catch (err) {
    throw new Error(
      `Failed to import formula-wasm Node build (${entry}). ` +
        `Run \`node scripts/build-formula-wasm-node.mjs\` (or rerun vitest without FORMULA_SKIP_WASM_BUILD).\n\n` +
        `Original error: ${err instanceof Error ? err.message : String(err)}`,
    );
  }
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
          case "getCellRich":
            result = this.workbook?.getCellRich(params.address, params.sheet);
            break;
          case "setCellRich":
            this.workbook?.setCellRich(params.address, params.value, params.sheet);
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
          case "setRange":
            this.workbook?.setRange(params.range, params.values, params.sheet);
            result = null;
            break;
          case "recalculate":
            result = this.workbook?.recalculate(params.sheet);
            break;
          case "renameSheet":
            result = Boolean(this.workbook?.renameSheet?.(params.oldName, params.newName));
            break;
          case "setWorkbookFileMetadata":
            this.workbook?.setWorkbookFileMetadata?.(params.directory, params.filename);
            result = null;
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

const skipWasmBuild = process.env.FORMULA_SKIP_WASM_BUILD === "1" || process.env.FORMULA_SKIP_WASM_BUILD === "true";
const describeWasm = skipWasmBuild ? describe.skip : describe;

describeWasm("EngineWorker null clear semantics", () => {
  it("normalizes formula input text using formula-model display semantics", async () => {
    const wasm = await loadFormulaWasm();
    const worker = new WasmBackedWorker(wasm);

    const engine = await EngineWorker.connect({
      worker,
      wasmModuleUrl: "mock://wasm",
      channelFactory: createMockChannel
    });

    try {
      await engine.newWorkbook();

      await engine.setCell("A1", "  =  SUM(A1:A2)  ");
      const cell = await engine.getCell("A1");
      expect(cell.input).toBe("=SUM(A1:A2)");
    } finally {
      engine.terminate();
    }
  });

  it("resolves sheet names using Unicode NFKC + case-insensitive compare (no duplicate sheet aliases)", async () => {
    const wasm = await loadFormulaWasm();
    const worker = new WasmBackedWorker(wasm);

    const engine = await EngineWorker.connect({
      worker,
      wasmModuleUrl: "mock://wasm",
      channelFactory: createMockChannel
    });

    try {
      await engine.newWorkbook();

      // Angstrom sign (U+212B) normalizes to Å (U+00C5) under NFKC. The engine should treat
      // both as equivalent sheet names for lookups, without creating duplicate wrapper entries.
      expect(await engine.renameSheet("Sheet1", "Å")).toBe(true);
      await engine.setCell("A1", 1, "Å");

      const exported = JSON.parse(await engine.toJson());
      // `toJson()` should include sheetOrder so sheet tab order round-trips (3D references / SHEET()).
      expect(exported.sheetOrder).toEqual(["Å"]);
      expect(Object.keys(exported.sheets)).toEqual(["Å"]);
      expect(exported.sheets["Å"].cells.A1).toBe(1);
    } finally {
      engine.terminate();
    }
  });

  it("updates CELL(\"address\") and CELL(\"filename\") outputs after a sheet rename", async () => {
    const wasm = await loadFormulaWasm();
    const worker = new WasmBackedWorker(wasm);

    const engine = await EngineWorker.connect({
      worker,
      wasmModuleUrl: "mock://wasm",
      channelFactory: createMockChannel
    });

    try {
      await engine.newWorkbook();
      await engine.setWorkbookFileMetadata(null, "Book1.xlsx");

      await engine.setCell("A1", '=CELL("address",Sheet1!A1)', "Sheet2");
      await engine.setCell("A2", '=CELL("filename",Sheet1!A1)', "Sheet2");
      await engine.recalculate();

      const beforeAddress = await engine.getCell("A1", "Sheet2");
      expect(beforeAddress.value).toBe("Sheet1!$A$1");

      const beforeFilename = await engine.getCell("A2", "Sheet2");
      expect(beforeFilename.value).toBe("[Book1.xlsx]Sheet1");

      expect(await engine.renameSheet("Sheet1", "Budget")).toBe(true);
      await engine.recalculate();

      const afterAddress = await engine.getCell("A1", "Sheet2");
      expect(afterAddress.value).toBe("Budget!$A$1");

      const afterFilename = await engine.getCell("A2", "Sheet2");
      expect(afterFilename.value).toBe("[Book1.xlsx]Budget");

      // Renaming to a name that requires quoting should update CELL("address") outputs to include
      // the quoted sheet name while leaving CELL("filename") unquoted (Excel semantics).
      expect(await engine.renameSheet("Budget", "Budget 2026")).toBe(true);
      await engine.recalculate();

      const afterQuotedAddress = await engine.getCell("A1", "Sheet2");
      expect(afterQuotedAddress.value).toBe("'Budget 2026'!$A$1");

      const afterQuotedFilename = await engine.getCell("A2", "Sheet2");
      expect(afterQuotedFilename.value).toBe("[Book1.xlsx]Budget 2026");

      // Ensure stored formula inputs are also rewritten (toJson/getCell.input should match).
      const exported = JSON.parse(await engine.toJson());
      expect(exported.sheets?.Sheet2?.cells?.A1).toContain("'Budget 2026'!A1");
      expect(exported.sheets?.Sheet2?.cells?.A2).toContain("'Budget 2026'!A1");
    } finally {
      engine.terminate();
    }
  });

  it("supports rich values (entity) with field access formulas", async () => {
    const wasm = await loadFormulaWasm();
    const worker = new WasmBackedWorker(wasm);
    const engine = await EngineWorker.connect({
      worker,
      wasmModuleUrl: "mock://wasm",
      channelFactory: createMockChannel
    });

    try {
      await engine.newWorkbook();

      const entity = {
        type: "entity",
        value: {
          entityType: "stock",
          entityId: "AAPL",
          displayValue: "Apple Inc.",
          properties: {
            Price: { type: "number", value: 12.5 }
          }
        }
      } as const;

      await engine.setCellRich?.("A1", entity, "Sheet1");
      await engine.setCell("B1", "=A1.Price", "Sheet1");
      await engine.recalculate("Sheet1");

      const b1 = await engine.getCell("B1", "Sheet1");
      expect(b1.value).toBe(12.5);

      const a1 = await engine.getCell("A1", "Sheet1");
      expect(a1.input).toBeNull();
      expect(a1.value).toBe("Apple Inc.");

      const a1Rich = await engine.getCellRich?.("A1", "Sheet1");
      expect(a1Rich).toEqual({
        sheet: "Sheet1",
        address: "A1",
        input: entity,
        value: entity
      });
    } finally {
      engine.terminate();
    }
  });

  it("roundtrips rich_text input through getCellRich while degrading scalar getCell", async () => {
    const wasm = await loadFormulaWasm();
    const worker = new WasmBackedWorker(wasm);
    const engine = await EngineWorker.connect({
      worker,
      wasmModuleUrl: "mock://wasm",
      channelFactory: createMockChannel
    });

    try {
      await engine.newWorkbook();

      const richText: CellValueRich = {
        type: "rich_text",
        value: {
          text: "Hello",
          runs: [
            {
              start: 0,
              end: 5,
              style: { bold: true }
            }
          ]
        }
      };

      await engine.setCellRich?.("A1", richText, "Sheet1");

      const a1 = await engine.getCell("A1", "Sheet1");
      expect(a1.input).toBe("Hello");
      expect(a1.value).toBe("Hello");

      const a1Rich = await engine.getCellRich?.("A1", "Sheet1");
      // RichTextRunStyle may round-trip with additional explicit null fields for unset style props.
      // Assert the essential structure without requiring an exact key-for-key match.
      expect(a1Rich?.input).toMatchObject(richText);
      expect(a1Rich?.value).toEqual({ type: "string", value: "Hello" });
    } finally {
      engine.terminate();
    }
  });

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

  it("returns no recalc changes when nothing is dirty", async () => {
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

      const changes = await engine.recalculate();
      expect(changes).toEqual([{ sheet: "Sheet1", address: "A2", value: 2 }]);

      const changes2 = await engine.recalculate();
      expect(changes2).toEqual([]);
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

  it("surfaces bare LAMBDA values as #CALC!", async () => {
    const wasm = await loadFormulaWasm();
    const worker = new WasmBackedWorker(wasm);

    const engine = await EngineWorker.connect({
      worker,
      wasmModuleUrl: "mock://wasm",
      channelFactory: createMockChannel
    });

    try {
      await engine.newWorkbook();
      await engine.setCell("A1", "=LAMBDA(x,x)");

      const changes = await engine.recalculate();
      expect(changes).toEqual([{ sheet: "Sheet1", address: "A1", value: "#CALC!" }]);

      const a1 = await engine.getCell("A1");
      expect(a1.value).toBe("#CALC!");
    } finally {
      engine.terminate();
    }
  });

  it("treats a bare '=' input as literal text (not a formula)", async () => {
    const wasm = await loadFormulaWasm();
    const worker = new WasmBackedWorker(wasm);

    const engine = await EngineWorker.connect({
      worker,
      wasmModuleUrl: "mock://wasm",
      channelFactory: createMockChannel
    });

    try {
      await engine.newWorkbook();
      await engine.setCell("A1", "=");
      const changes = await engine.recalculate();
      expect(changes).toEqual([]);

      const a1 = await engine.getCell("A1");
      expect(a1.input).toBe("=");
      expect(a1.value).toBe("=");

      const exported = JSON.parse(await engine.toJson());
      expect(exported.sheets.Sheet1.cells).toHaveProperty("A1", "=");
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

  it("clears spill output cells when the spill origin is edited", async () => {
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
      await engine.recalculate();
      expect((await engine.getCell("B1")).value).toBe(2);

      // Edit the spill origin formula so it no longer spills; the previous spill output should be
      // cleared (blank/null) in the next recalc change list.
      await engine.setCell("A1", "=1");
      const changes = await engine.recalculate();
      expect(changes).toEqual([
        { sheet: "Sheet1", address: "A1", value: 1 },
        { sheet: "Sheet1", address: "B1", value: null }
      ]);
      expect((await engine.getCell("B1")).value).toBeNull();
    } finally {
      engine.terminate();
    }
  });

  it("clears spill output cells when a spill shrinks", async () => {
    const wasm = await loadFormulaWasm();
    const worker = new WasmBackedWorker(wasm);

    const engine = await EngineWorker.connect({
      worker,
      wasmModuleUrl: "mock://wasm",
      channelFactory: createMockChannel
    });

    try {
      await engine.newWorkbook();
      await engine.setCell("A2", 2);
      await engine.setCell("A1", "=SEQUENCE(1,A2)");
      await engine.recalculate();
      expect((await engine.getCell("B1")).value).toBe(2);

      // Shrink the spill width from 2 cells to 1; B1 should be returned as a delta back to blank.
      await engine.setCell("A2", 1);
      const changes = await engine.recalculate();
      expect(changes).toEqual([{ sheet: "Sheet1", address: "B1", value: null }]);
      expect((await engine.getCell("B1")).value).toBeNull();
    } finally {
      engine.terminate();
    }
  });

  it("returns recalc changes in deterministic (sheet, row, col) order", async () => {
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

      await engine.setCell("A1", 10, "Sheet2");
      await engine.setCell("A2", "=A1*2", "Sheet2");

      await engine.recalculate();

      // Dirty both sheets before a single recalc tick. Recalculate should return formula deltas
      // sorted by (sheet, row, col): Sheet1 before Sheet2.
      await engine.setCell("A1", 2, "Sheet1");
      await engine.setCell("A1", 11, "Sheet2");

      const changes = await engine.recalculate();
      expect(changes).toEqual([
        { sheet: "Sheet1", address: "A2", value: 4 },
        { sheet: "Sheet2", address: "A2", value: 22 }
      ]);
    } finally {
      engine.terminate();
    }
  });

  it("returns recalc changes in row-major order within a sheet", async () => {
    const wasm = await loadFormulaWasm();
    const worker = new WasmBackedWorker(wasm);

    const engine = await EngineWorker.connect({
      worker,
      wasmModuleUrl: "mock://wasm",
      channelFactory: createMockChannel
    });

    try {
      await engine.newWorkbook();
      await engine.setCell("A1", 1);
      await engine.setCell("B1", "=A1+1");
      await engine.setCell("A2", "=A1*2");

      await engine.recalculate();

      await engine.setCell("A1", 2);
      const changes = await engine.recalculate();
      expect(changes).toEqual([
        { sheet: "Sheet1", address: "B1", value: 3 },
        { sheet: "Sheet1", address: "A2", value: 4 }
      ]);
    } finally {
      engine.terminate();
    }
  });

  it("does not error when recalculate is called with an unknown sheet name", async () => {
    const wasm = await loadFormulaWasm();
    const worker = new WasmBackedWorker(wasm);

    const engine = await EngineWorker.connect({
      worker,
      wasmModuleUrl: "mock://wasm",
      channelFactory: createMockChannel
    });

    try {
      await engine.newWorkbook();
      // The WASM workbook API accepts an optional sheet argument but does not scope/filter
      // recalculation by sheet. Unknown sheets should be ignored (not treated as an error).
      await expect(engine.recalculate("MissingSheet")).resolves.toEqual([]);
    } finally {
      engine.terminate();
    }
  });

  it("does not filter recalc changes by sheet name (case-insensitive)", async () => {
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

      // Dirty both sheets, then request a sheet-scoped recalc. The returned changes should still
      // include all sheets so the client-side cache remains coherent.
      await engine.setCell("A1", 2, "Sheet1");
      await engine.setCell("A1", 4, "Sheet2");
      const changes = await engine.recalculate("sHeEt2");

      expect(changes).toEqual([
        { sheet: "Sheet1", address: "A2", value: 4 },
        { sheet: "Sheet2", address: "A2", value: 8 }
      ]);
      expect((await engine.getCell("A2", "Sheet2")).value).toBe(8);
      expect((await engine.getCell("A2", "Sheet1")).value).toBe(4);
    } finally {
      engine.terminate();
    }
  });

  it("sheet-scoped recalc still updates cross-sheet dependents and returns their deltas", async () => {
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
      // The engine should still surface the dependent sheet's delta (no filtering by sheet arg).
      expect(changes).toEqual([{ sheet: "Sheet2", address: "A1", value: 4 }]);
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

  it("applies localeId when loading workbook JSON (localized formulas parse correctly)", async () => {
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
          localeId: "de-DE",
          sheets: {
            Sheet1: {
              cells: {
                // de-DE: localized function name + semicolon arg separator + decimal comma.
                A1: "=SUMME(1,5;2,5)"
              }
            }
          }
        })
      );

      await engine.recalculate();
      expect((await engine.getCell("A1")).value).toBe(4);

      const exported = JSON.parse(await engine.toJson());
      expect(exported.localeId).toBe("de-DE");
      // Engine persists canonical formulas internally.
      expect(exported.sheets.Sheet1.cells.A1).toBe("=SUM(1.5,2.5)");
    } finally {
      engine.terminate();
    }
  });

  it("roundtrips toJson()/fromJson() for comma-decimal locales by disambiguating canonical formulas", async () => {
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
          localeId: "de-DE",
          sheets: {
            Sheet1: {
              cells: {
                // de-DE: semicolon argument separator.
                A1: "=LOG(8;2)"
              }
            }
          }
        })
      );

      await engine.recalculate();
      expect((await engine.getCell("A1")).value).toBe(3);

      const exportedStr = await engine.toJson();
      const exported = JSON.parse(exportedStr);
      expect(exported.localeId).toBe("de-DE");
      expect(exported.formulaLanguage).toBe("canonical");
      expect(exported.sheets.Sheet1.cells.A1).toBe("=LOG(8,2)");

      // Ensure `fromJson(toJson(x))` is stable for comma-decimal locales like de-DE.
      await engine.loadWorkbookFromJson(exportedStr);
      await engine.recalculate();
      expect((await engine.getCell("A1")).value).toBe(3);
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
