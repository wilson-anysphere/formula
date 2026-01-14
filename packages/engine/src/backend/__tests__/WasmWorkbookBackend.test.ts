import { describe, expect, it, vi } from "vitest";

import type { EngineClient } from "../../client.ts";
import { colToName, fromA1, toA1, toA1Range } from "../a1.ts";
import { normalizeFormulaText } from "../formula.ts";
import { WasmWorkbookBackend } from "../WasmWorkbookBackend.ts";

describe("A1 helpers", () => {
  it("converts 0-based columns to Excel column names", () => {
    expect(colToName(0)).toBe("A");
    expect(colToName(25)).toBe("Z");
    expect(colToName(26)).toBe("AA");
    expect(colToName(27)).toBe("AB");
    expect(colToName(51)).toBe("AZ");
    expect(colToName(52)).toBe("BA");
    expect(colToName(701)).toBe("ZZ");
    expect(colToName(702)).toBe("AAA");
  });

  it("converts 0-based row/col coords to A1 addresses", () => {
    expect(toA1(0, 0)).toBe("A1");
    expect(toA1(1, 0)).toBe("A2");
    expect(toA1(0, 1)).toBe("B1");
    expect(toA1(9, 25)).toBe("Z10");
  });

  it("formats an A1 range (collapsing to a single cell when needed)", () => {
    expect(toA1Range(0, 0, 0, 0)).toBe("A1");
    expect(toA1Range(0, 0, 1, 1)).toBe("A1:B2");
  });

  it("parses A1 addresses into 0-based row/col coords", () => {
    expect(fromA1("A1")).toEqual({ row0: 0, col0: 0 });
    expect(fromA1("B2")).toEqual({ row0: 1, col0: 1 });
    expect(fromA1("$AA$10")).toEqual({ row0: 9, col0: 26 });
  });
});

describe("formula normalization", () => {
  it("ensures formulas start with '=' and strips leading whitespace", () => {
    expect(normalizeFormulaText("A1*2")).toBe("=A1*2");
    expect(normalizeFormulaText(" =A1*2")).toBe("=A1*2");
    expect(normalizeFormulaText("=A1*2")).toBe("=A1*2");
  });
});

describe("WasmWorkbookBackend", () => {
  it("translates setRange row/col rectangles into engine A1 range calls (with formula normalization)", async () => {
    const engine: EngineClient = {
      init: vi.fn(async () => {}),
      newWorkbook: vi.fn(async () => {}),
      loadWorkbookFromJson: vi.fn(async () => {}),
      loadWorkbookFromXlsxBytes: vi.fn(async () => {}),
      toJson: vi.fn(async () => "{}"),
      getCell: vi.fn(async () => ({ sheet: "Sheet1", address: "A1", input: null, value: null })),
      getRange: vi.fn(async () => []),
      setCell: vi.fn(async () => {}),
      setCells: vi.fn(async () => {}),
      setRange: vi.fn(async () => {}),
      setWorkbookFileMetadata: vi.fn(async () => {}),
      setCellStyleId: vi.fn(async () => {}),
      setColWidth: vi.fn(async () => {}),
      setColHidden: vi.fn(async () => {}),
      internStyle: vi.fn(async () => 0),
      setLocale: vi.fn(async () => true),
      recalculate: vi.fn(async () => []),
      setSheetDimensions: vi.fn(async () => {}),
      getSheetDimensions: vi.fn(async () => ({ rows: 1_048_576, cols: 16_384 })),
      applyOperation: vi.fn(async () => ({ changedCells: [], movedRanges: [], formulaRewrites: [] })),
      rewriteFormulasForCopyDelta: vi.fn(async () => []),
      lexFormula: vi.fn(async () => []),
      lexFormulaPartial: vi.fn(async () => ({ tokens: [], error: null })),
      parseFormulaPartial: vi.fn(async () => ({ ast: null, error: null, context: { function: null } })),
      terminate: vi.fn(),
    };

    const backend = new WasmWorkbookBackend(engine);

    await backend.setRange({
      sheetId: "Sheet1",
      startRow: 0,
      startCol: 0,
      endRow: 1,
      endCol: 1,
      values: [
        [
          { value: 1, formula: null },
          { value: 123, formula: " A1*2" },
        ],
        [
          { value: true, formula: null },
          { value: { text: "Hello", runs: [] }, formula: null },
        ],
      ],
    });

    expect(engine.setRange).toHaveBeenCalledTimes(1);
    expect(engine.setRange).toHaveBeenCalledWith(
      "A1:B2",
      [
        [1, "=A1*2"],
        [true, "Hello"],
      ],
      "Sheet1",
    );

    expect(engine.recalculate).toHaveBeenCalledTimes(1);
    expect(engine.recalculate).toHaveBeenCalledWith("Sheet1");
  });

  it("loads a workbook from XLSX bytes (and clears used range tracking)", async () => {
    const calls: string[] = [];
    let resolveLoad: (() => void) | undefined;
    const loadPromise = new Promise<void>((resolve) => {
      resolveLoad = resolve;
    });

    const engine: EngineClient = {
      init: vi.fn(async () => {}),
      newWorkbook: vi.fn(async () => {}),
      loadWorkbookFromJson: vi.fn(async () => {}),
      loadWorkbookFromXlsxBytes: vi.fn(async () => {
        calls.push("loadWorkbookFromXlsxBytes");
        return await loadPromise;
      }),
      toJson: vi.fn(async () => "{}"),
      getCell: vi.fn(async () => ({ sheet: "Sheet1", address: "A1", input: null, value: null })),
      getRange: vi.fn(async () => []),
      setCell: vi.fn(async () => {}),
      setCells: vi.fn(async () => {}),
      setRange: vi.fn(async () => {}),
      setWorkbookFileMetadata: vi.fn(async () => {}),
      setCellStyleId: vi.fn(async () => {}),
      setColWidth: vi.fn(async () => {}),
      setColHidden: vi.fn(async () => {}),
      internStyle: vi.fn(async () => 0),
      setLocale: vi.fn(async () => true),
      recalculate: vi.fn(async () => {
        calls.push("recalculate");
        return [];
      }),
      setSheetDimensions: vi.fn(async () => {}),
      getSheetDimensions: vi.fn(async () => ({ rows: 1_048_576, cols: 16_384 })),
      applyOperation: vi.fn(async () => ({ changedCells: [], movedRanges: [], formulaRewrites: [] })),
      rewriteFormulasForCopyDelta: vi.fn(async () => []),
      lexFormula: vi.fn(async () => []),
      lexFormulaPartial: vi.fn(async () => ({ tokens: [], error: null })),
      parseFormulaPartial: vi.fn(async () => ({ ast: null, error: null, context: { function: null } })),
      terminate: vi.fn(),
    };

    const backend = new WasmWorkbookBackend(engine);

    await backend.setCell({ sheetId: "Sheet1", row: 2, col: 3, value: 123, formula: null });
    expect(await backend.getSheetUsedRange("Sheet1")).toEqual({ start_row: 2, end_row: 2, start_col: 3, end_col: 3 });

    calls.length = 0;

    const bytes = new Uint8Array([1, 2, 3]);
    const openPromise = backend.openWorkbookFromBytes(bytes);

    expect(engine.loadWorkbookFromXlsxBytes).toHaveBeenCalledTimes(1);
    expect(engine.loadWorkbookFromXlsxBytes).toHaveBeenCalledWith(bytes);
    expect(engine.recalculate).toHaveBeenCalledTimes(1); // from setCell only
    expect(calls).toEqual(["loadWorkbookFromXlsxBytes"]);

    resolveLoad?.();
    const info = await openPromise;

    expect(engine.recalculate).toHaveBeenCalledTimes(2);
    expect(calls).toEqual(["loadWorkbookFromXlsxBytes", "recalculate"]);
    expect(info).toEqual({
      path: null,
      origin_path: null,
      sheets: [{ id: "Sheet1", name: "Sheet1" }],
    });
    expect(backend.getWorkbookInfo()).toEqual(info);
    expect(await backend.getSheetUsedRange("Sheet1")).toBeNull();
  });

  it("loads workbooks from raw xlsx bytes, triggers a full recalc, and seeds used ranges", async () => {
    const bytes = new Uint8Array([1, 2, 3]);
    const workbookJson = JSON.stringify({
      sheets: {
        Sheet1: {
          cells: {
            A1: 1,
            B2: 2,
            C3: "=A1+B2",
          },
        },
        Sheet2: {
          cells: {
            D4: "Hello",
          },
        },
        Empty: { cells: {} },
      },
    });

    const engine: EngineClient = {
      init: vi.fn(async () => {}),
      newWorkbook: vi.fn(async () => {}),
      loadWorkbookFromJson: vi.fn(async () => {}),
      loadWorkbookFromXlsxBytes: vi.fn(async () => {}),
      toJson: vi.fn(async () => workbookJson),
      getCell: vi.fn(async () => ({ sheet: "Sheet1", address: "A1", input: null, value: null })),
      getRange: vi.fn(async () => []),
      setCell: vi.fn(async () => {}),
      setCells: vi.fn(async () => {}),
      setRange: vi.fn(async () => {}),
      setWorkbookFileMetadata: vi.fn(async () => {}),
      setCellStyleId: vi.fn(async () => {}),
      setColWidth: vi.fn(async () => {}),
      setColHidden: vi.fn(async () => {}),
      internStyle: vi.fn(async () => 0),
      setLocale: vi.fn(async () => true),
      recalculate: vi.fn(async () => []),
      setSheetDimensions: vi.fn(async () => {}),
      getSheetDimensions: vi.fn(async () => ({ rows: 1_048_576, cols: 16_384 })),
      applyOperation: vi.fn(async () => ({ changedCells: [], movedRanges: [], formulaRewrites: [] })),
      rewriteFormulasForCopyDelta: vi.fn(async () => []),
      lexFormula: vi.fn(async () => []),
      lexFormulaPartial: vi.fn(async () => ({ tokens: [], error: null })),
      parseFormulaPartial: vi.fn(async () => ({ ast: null, error: null, context: { function: null } })),
      terminate: vi.fn(),
    };

    const backend = new WasmWorkbookBackend(engine);
    const info = await backend.openWorkbookFromBytes(bytes);

    expect(engine.loadWorkbookFromXlsxBytes).toHaveBeenCalledTimes(1);
    expect(engine.loadWorkbookFromXlsxBytes).toHaveBeenCalledWith(bytes);

    expect(engine.recalculate).toHaveBeenCalledTimes(1);
    expect(engine.recalculate).toHaveBeenCalledWith();

    expect(info).toEqual({
      path: null,
      origin_path: null,
      sheets: [
        { id: "Sheet1", name: "Sheet1" },
        { id: "Sheet2", name: "Sheet2" },
        { id: "Empty", name: "Empty" },
      ],
    });

    expect(await backend.getSheetUsedRange("Sheet1")).toEqual({
      start_row: 0,
      start_col: 0,
      end_row: 2,
      end_col: 2,
    });

    expect(await backend.getSheetUsedRange("Sheet2")).toEqual({
      start_row: 3,
      start_col: 3,
      end_row: 3,
      end_col: 3,
    });

    expect(await backend.getSheetUsedRange("Empty")).toBeNull();
  });
});
