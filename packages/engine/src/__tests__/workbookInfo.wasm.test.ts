import { describe, expect, it } from "vitest";
import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";

import { formulaWasmNodeEntryUrl } from "../../../../scripts/build-formula-wasm-node.mjs";

import { EngineWorker, type MessageChannelLike, type WorkerLike } from "../worker/EngineWorker.ts";
import type {
  CellValueRich,
  InitMessage,
  RpcRequest,
  WorkerInboundMessage,
  WorkerOutboundMessage,
  WorkbookInfoDto,
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
  // wasm-pack `--target nodejs` outputs CommonJS. Under ESM dynamic import, the exports are
  // exposed on `default`.
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
          case "setCellRich":
            this.workbook?.setCellRich(params.address, params.value, params.sheet);
            result = null;
            break;
          case "setSheetDimensions":
            this.workbook?.setSheetDimensions(params.sheet, params.rows, params.cols);
            result = null;
            break;
          case "getWorkbookInfo":
            if (typeof this.workbook?.getWorkbookInfo !== "function") {
              throw new Error("getWorkbookInfo: not supported by this wasm build");
            }
            result = this.workbook.getWorkbookInfo();
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

describeWasm("EngineWorker getWorkbookInfo (wasm)", () => {
  it("includes optional sheet visibility + tabColor metadata when available", async () => {
    const wasm = await loadFormulaWasm();
    const fixturePath = fileURLToPath(
      new URL("../../../../crates/formula-xlsx/tests/fixtures/sheet-metadata.xlsx", import.meta.url)
    );
    const bytes = new Uint8Array(readFileSync(fixturePath));

    const workbook = wasm.WasmWorkbook.fromXlsxBytes(bytes);
    const info = workbook.getWorkbookInfo() as WorkbookInfoDto;

    const byId = new Map(info.sheets.map((sheet) => [sheet.id, sheet]));
    expect(Array.from(byId.keys()).sort()).toEqual(["Hidden", "VeryHidden", "Visible"]);

    const visible = byId.get("Visible")!;
    expect(visible.visibility).toBeUndefined();
    expect(visible.tabColor).toEqual({ rgb: "FFFF0000" });

    const hidden = byId.get("Hidden")!;
    expect(hidden.visibility).toBe("hidden");
    expect(hidden.tabColor).toBeUndefined();

    const veryHidden = byId.get("VeryHidden")!;
    expect(veryHidden.visibility).toBe("veryHidden");
    expect(veryHidden.tabColor).toBeUndefined();
  });

  it("round-trips sheet visibility + tabColor through toJson/fromJson", async () => {
    const wasm = await loadFormulaWasm();
    const fixturePath = fileURLToPath(
      new URL("../../../../crates/formula-xlsx/tests/fixtures/sheet-metadata.xlsx", import.meta.url)
    );
    const bytes = new Uint8Array(readFileSync(fixturePath));

    const workbook = wasm.WasmWorkbook.fromXlsxBytes(bytes);
    const json = workbook.toJson();
    const parsed = JSON.parse(json) as any;

    expect(parsed?.sheets?.Visible?.tabColor).toEqual({ rgb: "FFFF0000" });
    expect(parsed?.sheets?.Visible?.visibility).toBeUndefined();
    expect(parsed?.sheets?.Hidden?.visibility).toBe("hidden");
    expect(parsed?.sheets?.VeryHidden?.visibility).toBe("veryHidden");

    const roundtripped = wasm.WasmWorkbook.fromJson(json);
    const info = roundtripped.getWorkbookInfo() as WorkbookInfoDto;

    const byId = new Map(info.sheets.map((sheet) => [sheet.id, sheet]));
    expect(byId.get("Visible")?.tabColor).toEqual({ rgb: "FFFF0000" });
    expect(byId.get("Hidden")?.visibility).toBe("hidden");
    expect(byId.get("VeryHidden")?.visibility).toBe("veryHidden");
  });

  it("returns sheet list, dimensions, and best-effort used ranges (including rich inputs)", async () => {
    const wasm = await loadFormulaWasm();
    const worker = new WasmBackedWorker(wasm);
    const engine = await EngineWorker.connect({
      worker,
      wasmModuleUrl: "mock://wasm",
      channelFactory: createMockChannel,
    });

    try {
      await engine.newWorkbook();

      // Create an empty sheet entry without storing any cells.
      await engine.setSheetDimensions("Empty", 1_048_576, 16_384);

      // Non-default dimensions should be reported in the metadata payload.
      await engine.setSheetDimensions("Sheet1", 100, 50);

      await engine.setCell("A1", 1, "Sheet1");

      const entity: CellValueRich = {
        type: "entity",
        value: {
          entityType: "stock",
          entityId: "AAPL",
          displayValue: "Apple Inc.",
          properties: {
            Price: { type: "number", value: 12.5 },
          },
        },
      };
      // Store a rich (non-scalar) input outside the scalar map to ensure used range scanning
      // covers `sheets_rich`.
      await engine.setCellRich?.("C3", entity, "Sheet1");

      // Create a second sheet with a scalar input.
      await engine.setCell("D4", 42, "Sheet2");

      const info = (await engine.getWorkbookInfo()) as WorkbookInfoDto;

      expect(info.path).toBeNull();
      expect(info.origin_path).toBeNull();

      const byId = new Map(info.sheets.map((sheet) => [sheet.id, sheet]));
      expect(Array.from(byId.keys()).sort()).toEqual(["Empty", "Sheet1", "Sheet2"]);

      const sheet1 = byId.get("Sheet1")!;
      expect(sheet1.name).toBe("Sheet1");
      expect(sheet1.rowCount).toBe(100);
      expect(sheet1.colCount).toBe(50);
      expect(sheet1.usedRange).toEqual({ start_row: 0, start_col: 0, end_row: 2, end_col: 2 });

      const sheet2 = byId.get("Sheet2")!;
      expect(sheet2.name).toBe("Sheet2");
      // Defaults should be omitted.
      expect(sheet2.rowCount).toBeUndefined();
      expect(sheet2.colCount).toBeUndefined();
      expect(sheet2.usedRange).toEqual({ start_row: 3, start_col: 3, end_row: 3, end_col: 3 });

      const empty = byId.get("Empty")!;
      expect(empty.name).toBe("Empty");
      expect(empty.rowCount).toBeUndefined();
      expect(empty.colCount).toBeUndefined();
      expect(empty.usedRange).toBeUndefined();
    } finally {
      engine.terminate();
    }
  });
});
