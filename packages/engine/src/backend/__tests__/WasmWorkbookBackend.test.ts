import { describe, expect, it, vi } from "vitest";

import type { EngineClient } from "../../client.ts";
import type { CalcSettings, CellDataCompact } from "../../protocol.ts";
import { colToName, fromA1, toA1, toA1Range } from "../a1.ts";
import { normalizeFormulaText } from "../formula.ts";
import { WasmWorkbookBackend } from "../WasmWorkbookBackend.ts";

const defaultCalcSettings: CalcSettings = {
  calculationMode: "manual",
  calculateBeforeSave: true,
  fullPrecision: true,
  fullCalcOnLoad: false,
  iterative: { enabled: false, maxIterations: 100, maxChange: 0.001 },
};

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
  function createMockEngine(overrides: Partial<EngineClient>): EngineClient {
    return {
      init: overrides.init ?? vi.fn(async () => {}),
      newWorkbook: overrides.newWorkbook ?? vi.fn(async () => {}),
      loadWorkbookFromJson: overrides.loadWorkbookFromJson ?? vi.fn(async () => {}),
      loadWorkbookFromXlsxBytes: overrides.loadWorkbookFromXlsxBytes ?? vi.fn(async () => {}),
      toJson: overrides.toJson ?? vi.fn(async () => "{}"),
      getWorkbookInfo: overrides.getWorkbookInfo,
      getCell:
        overrides.getCell ?? vi.fn(async () => ({ sheet: "Sheet1", address: "A1", input: null, value: null })),
      getCellRich: overrides.getCellRich,
      getRange: overrides.getRange ?? vi.fn(async () => []),
      getRangeCompact: overrides.getRangeCompact,
      setCell: overrides.setCell ?? vi.fn(async () => {}),
      setCellRich: overrides.setCellRich,
      setCells: overrides.setCells ?? vi.fn(async () => {}),
      setRange: overrides.setRange ?? vi.fn(async () => {}),
      setWorkbookFileMetadata: overrides.setWorkbookFileMetadata ?? vi.fn(async () => {}),
      setCellStyleId: overrides.setCellStyleId ?? vi.fn(async () => {}),
      setRowStyleId: overrides.setRowStyleId,
      setColStyleId: overrides.setColStyleId,
      setSheetDefaultStyleId: overrides.setSheetDefaultStyleId,
      setColWidth: overrides.setColWidth ?? vi.fn(async () => {}),
      setColHidden: overrides.setColHidden ?? vi.fn(async () => {}),
      setSheetDefaultColWidth:
        overrides.setSheetDefaultColWidth ?? vi.fn(async (_sheet: string, _widthChars: number | null) => {}),
      internStyle: overrides.internStyle ?? vi.fn(async () => 0),
      setLocale: overrides.setLocale ?? vi.fn(async () => true),
      getCalcSettings: overrides.getCalcSettings ?? vi.fn(async () => defaultCalcSettings),
      setCalcSettings: overrides.setCalcSettings ?? vi.fn(async () => {}),
      setEngineInfo: overrides.setEngineInfo ?? vi.fn(async () => {}),
      setInfoOrigin: overrides.setInfoOrigin ?? vi.fn(async () => {}),
      setInfoOriginForSheet: overrides.setInfoOriginForSheet ?? vi.fn(async () => {}),
      setColFormatRuns: overrides.setColFormatRuns ?? vi.fn(async () => {}),
      recalculate: overrides.recalculate ?? vi.fn(async () => []),
      setSheetDimensions: overrides.setSheetDimensions ?? vi.fn(async () => {}),
      getSheetDimensions: overrides.getSheetDimensions ?? vi.fn(async () => ({ rows: 1_048_576, cols: 16_384 })),
      renameSheet: overrides.renameSheet ?? vi.fn(async () => true),
      setSheetOrigin: overrides.setSheetOrigin ?? vi.fn(async () => {}),
      setColWidthChars: overrides.setColWidthChars ?? vi.fn(async () => {}),
      applyOperation: overrides.applyOperation ?? vi.fn(async () => ({ changedCells: [], movedRanges: [], formulaRewrites: [] })),
      rewriteFormulasForCopyDelta: overrides.rewriteFormulasForCopyDelta ?? vi.fn(async () => []),
      getPivotSchema: overrides.getPivotSchema,
      getPivotFieldItems: overrides.getPivotFieldItems,
      getPivotFieldItemsPaged: overrides.getPivotFieldItemsPaged,
      calculatePivot: overrides.calculatePivot,
      goalSeek: overrides.goalSeek,
      canonicalizeFormula: overrides.canonicalizeFormula,
      localizeFormula: overrides.localizeFormula,
      lexFormula: overrides.lexFormula ?? vi.fn(async () => []),
      lexFormulaPartial: overrides.lexFormulaPartial ?? vi.fn(async () => ({ tokens: [], error: null })),
      parseFormulaPartial:
        overrides.parseFormulaPartial ?? vi.fn(async () => ({ ast: null, error: null, context: { function: null } })),
      terminate: overrides.terminate ?? vi.fn(),
    };
  }

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
      setSheetDefaultColWidth: vi.fn(async (_sheet: string, _widthChars: number | null) => {}),
      internStyle: vi.fn(async () => 0),
      setLocale: vi.fn(async () => true),
      getCalcSettings: vi.fn(async () => defaultCalcSettings),
      setCalcSettings: vi.fn(async () => {}),
      setEngineInfo: vi.fn(async () => {}),
      setInfoOrigin: vi.fn(async () => {}),
      setInfoOriginForSheet: vi.fn(async () => {}),
      setColFormatRuns: vi.fn(async () => {}),
      recalculate: vi.fn(async () => []),
      setSheetDimensions: vi.fn(async () => {}),
      getSheetDimensions: vi.fn(async () => ({ rows: 1_048_576, cols: 16_384 })),
      renameSheet: vi.fn(async () => true),
      setSheetDisplayName: vi.fn(async () => {}),
      setColWidthChars: vi.fn(async () => {}),
      setSheetOrigin: vi.fn(async () => {}),
      setRowStyleId: vi.fn(async () => {}),
      setColStyleId: vi.fn(async () => {}),
      setSheetDefaultStyleId: vi.fn(async () => {}),
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

  it("prefers getRangeCompact when available and returns the same RangeData as legacy getRange", async () => {
    const legacy = [
      [
        { sheet: "Sheet1", address: "A1", input: 1, value: 1 },
        // Include whitespace to ensure backend normalization logic is stable.
        { sheet: "Sheet1", address: "B1", input: "  =A1*2  ", value: 2 },
      ],
      [
        { sheet: "Sheet1", address: "A2", input: null, value: null },
        { sheet: "Sheet1", address: "B2", input: "Hello", value: "Hello" },
      ],
    ];

    const compact: CellDataCompact[][] = [
      [
        [1, 1],
        ["  =A1*2  ", 2],
      ],
      [
        [null, null],
        ["Hello", "Hello"],
      ],
    ] satisfies CellDataCompact[][];

    const engineCompact = createMockEngine({
      getRangeCompact: vi.fn(async () => compact),
      getRange: vi.fn(async () => legacy),
    });
    const backendCompact = new WasmWorkbookBackend(engineCompact);

    const engineLegacy = createMockEngine({
      getRange: vi.fn(async () => legacy),
    });
    const backendLegacy = new WasmWorkbookBackend(engineLegacy);

    const params = { sheetId: "Sheet1", startRow: 0, startCol: 0, endRow: 1, endCol: 1 };
    const gotCompact = await backendCompact.getRange(params);
    const gotLegacy = await backendLegacy.getRange(params);

    expect(gotCompact).toEqual(gotLegacy);
    expect(gotCompact).toEqual({
      start_row: 0,
      start_col: 0,
      values: [
        [
          { value: 1, formula: null, display_value: "1" },
          { value: 2, formula: "=A1*2", display_value: "2" },
        ],
        [
          { value: null, formula: null, display_value: "" },
          { value: "Hello", formula: null, display_value: "Hello" },
        ],
      ],
    });

    expect(engineCompact.getRangeCompact!).toHaveBeenCalledTimes(1);
    expect(engineCompact.getRange).toHaveBeenCalledTimes(0);
    expect(engineLegacy.getRange).toHaveBeenCalledTimes(1);
  });

  it("falls back to getRange when getRangeCompact is not supported", async () => {
    const legacy = [[{ sheet: "Sheet1", address: "A1", input: 1, value: 1 }]];

    const engine = createMockEngine({
      getRangeCompact: vi.fn(async () => {
        throw new Error("getRangeCompact: WasmWorkbook.getRangeCompact is not available in this WASM build");
      }),
      getRange: vi.fn(async () => legacy),
    });

    const backend = new WasmWorkbookBackend(engine);
    const params = { sheetId: "Sheet1", startRow: 0, startCol: 0, endRow: 0, endCol: 0 };
    const got = await backend.getRange(params);
    const got2 = await backend.getRange(params);

    expect(got).toEqual({
      start_row: 0,
      start_col: 0,
      values: [[{ value: 1, formula: null, display_value: "1" }]],
    });
    expect(got2).toEqual(got);

    // Once we've observed a missing compact API error, the backend should stop trying
    // (avoids paying for exceptions on every range read).
    expect(engine.getRangeCompact!).toHaveBeenCalledTimes(1);
    expect(engine.getRange).toHaveBeenCalledTimes(2);
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
      setSheetDefaultColWidth: vi.fn(async (_sheet: string, _widthChars: number | null) => {}),
      internStyle: vi.fn(async () => 0),
      setLocale: vi.fn(async () => true),
      getCalcSettings: vi.fn(async () => defaultCalcSettings),
      setCalcSettings: vi.fn(async () => {}),
      setEngineInfo: vi.fn(async () => {}),
      setInfoOrigin: vi.fn(async () => {}),
      setInfoOriginForSheet: vi.fn(async () => {}),
      setColFormatRuns: vi.fn(async () => {}),
      recalculate: vi.fn(async () => {
        calls.push("recalculate");
        return [];
      }),
      setSheetDimensions: vi.fn(async () => {}),
      getSheetDimensions: vi.fn(async () => ({ rows: 1_048_576, cols: 16_384 })),
      renameSheet: vi.fn(async () => true),
      setSheetDisplayName: vi.fn(async () => {}),
      setColWidthChars: vi.fn(async () => {}),
      setSheetOrigin: vi.fn(async () => {}),
      setRowStyleId: vi.fn(async () => {}),
      setColStyleId: vi.fn(async () => {}),
      setSheetDefaultStyleId: vi.fn(async () => {}),
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

  it("loads workbooks from raw xlsx bytes using getWorkbookInfo() when available (seeding used ranges without toJson)", async () => {
    const bytes = new Uint8Array([1, 2, 3]);

    const meta = {
      path: null,
      origin_path: null,
      sheets: [
        {
          id: "Sheet1",
          name: "Sheet1",
          visibility: "hidden",
          tabColor: { rgb: "FF0000" },
          usedRange: { start_row: 0, start_col: 0, end_row: 2, end_col: 2 },
        },
        {
          id: "Sheet2",
          name: "Sheet2",
          visibility: "veryHidden",
          usedRange: null,
        },
        {
          id: "Empty",
          name: "Empty",
          tabColor: { theme: 1, tint: 0.5 },
        },
      ],
    } as const;

    const engine: EngineClient = {
      init: vi.fn(async () => {}),
      newWorkbook: vi.fn(async () => {}),
      loadWorkbookFromJson: vi.fn(async () => {}),
      loadWorkbookFromXlsxBytes: vi.fn(async () => {}),
      getWorkbookInfo: vi.fn(async () => meta as any),
      toJson: vi.fn(async () => {
        throw new Error("toJson should not be called when getWorkbookInfo is available");
      }),
      getCell: vi.fn(async () => ({ sheet: "Sheet1", address: "A1", input: null, value: null })),
      getRange: vi.fn(async () => []),
      setCell: vi.fn(async () => {}),
      setCells: vi.fn(async () => {}),
      setRange: vi.fn(async () => {}),
      setWorkbookFileMetadata: vi.fn(async () => {}),
      setCellStyleId: vi.fn(async () => {}),
      setColWidth: vi.fn(async () => {}),
      setColHidden: vi.fn(async () => {}),
      setSheetDefaultColWidth: vi.fn(async (_sheet: string, _widthChars: number | null) => {}),
      internStyle: vi.fn(async () => 0),
      setLocale: vi.fn(async () => true),
      getCalcSettings: vi.fn(async () => defaultCalcSettings),
      setCalcSettings: vi.fn(async () => {}),
      setEngineInfo: vi.fn(async () => {}),
      setInfoOrigin: vi.fn(async () => {}),
      setInfoOriginForSheet: vi.fn(async () => {}),
      setColFormatRuns: vi.fn(async () => {}),
      recalculate: vi.fn(async () => []),
      setSheetDimensions: vi.fn(async () => {}),
      getSheetDimensions: vi.fn(async () => ({ rows: 1_048_576, cols: 16_384 })),
      renameSheet: vi.fn(async () => true),
      setColWidthChars: vi.fn(async () => {}),
      setSheetOrigin: vi.fn(async () => {}),
      setRowStyleId: vi.fn(async () => {}),
      setColStyleId: vi.fn(async () => {}),
      setSheetDefaultStyleId: vi.fn(async () => {}),
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

    expect(engine.getWorkbookInfo).toHaveBeenCalledTimes(1);
    expect(engine.toJson).toHaveBeenCalledTimes(0);

    expect(info).toEqual({
      path: null,
      origin_path: null,
      sheets: [
        { id: "Sheet1", name: "Sheet1", visibility: "hidden", tabColor: { rgb: "FF0000" } },
        { id: "Sheet2", name: "Sheet2", visibility: "veryHidden" },
        { id: "Empty", name: "Empty", tabColor: { theme: 1, tint: 0.5 } },
      ],
    });

    expect(await backend.getSheetUsedRange("Sheet1")).toEqual({
      start_row: 0,
      start_col: 0,
      end_row: 2,
      end_col: 2,
    });
    expect(await backend.getSheetUsedRange("Sheet2")).toBeNull();
    expect(await backend.getSheetUsedRange("Empty")).toBeNull();
  });

  it("loads workbooks from raw xlsx bytes, triggers a full recalc, and seeds used ranges", async () => {
    const bytes = new Uint8Array([1, 2, 3]);
    const workbookJson = JSON.stringify({
      sheets: {
        Sheet1: {
          visibility: "hidden",
          tabColor: { rgb: "FFFF0000" },
          cells: {
            A1: 1,
            B2: 2,
            C3: "=A1+B2",
          },
        },
        Sheet2: {
          visibility: "veryHidden",
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
      getWorkbookInfo: vi.fn(async () => {
        throw new Error("getWorkbookInfo: not supported");
      }),
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
      setSheetDefaultColWidth: vi.fn(async (_sheet: string, _widthChars: number | null) => {}),
      internStyle: vi.fn(async () => 0),
      setLocale: vi.fn(async () => true),
      getCalcSettings: vi.fn(async () => defaultCalcSettings),
      setCalcSettings: vi.fn(async () => {}),
      setEngineInfo: vi.fn(async () => {}),
      setInfoOrigin: vi.fn(async () => {}),
      setInfoOriginForSheet: vi.fn(async () => {}),
      setColFormatRuns: vi.fn(async () => {}),
      recalculate: vi.fn(async () => []),
      setSheetDimensions: vi.fn(async () => {}),
      getSheetDimensions: vi.fn(async () => ({ rows: 1_048_576, cols: 16_384 })),
      renameSheet: vi.fn(async () => true),
      setSheetDisplayName: vi.fn(async () => {}),
      setColWidthChars: vi.fn(async () => {}),
      setSheetOrigin: vi.fn(async () => {}),
      setRowStyleId: vi.fn(async () => {}),
      setColStyleId: vi.fn(async () => {}),
      setSheetDefaultStyleId: vi.fn(async () => {}),
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
        { id: "Sheet1", name: "Sheet1", visibility: "hidden", tabColor: { rgb: "FFFF0000" } },
        { id: "Sheet2", name: "Sheet2", visibility: "veryHidden" },
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

  it("prefers sheetOrder from toJson() when deriving sheet tab order in the legacy metadata path", async () => {
    const bytes = new Uint8Array([1, 2, 3]);
    const workbookJson = JSON.stringify({
      sheetOrder: ["Sheet2", "Sheet1", "Empty"],
      // Deliberately encode sheets in a different order to ensure we don't rely on Object.keys().
      sheets: {
        Empty: { cells: {} },
        Sheet1: { cells: { A1: 1 } },
        Sheet2: { cells: { B2: 2 } },
      },
    });

    const engine: EngineClient = {
      init: vi.fn(async () => {}),
      newWorkbook: vi.fn(async () => {}),
      loadWorkbookFromJson: vi.fn(async () => {}),
      loadWorkbookFromXlsxBytes: vi.fn(async () => {}),
      getWorkbookInfo: vi.fn(async () => {
        throw new Error("getWorkbookInfo: not supported");
      }),
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
      setSheetDefaultColWidth: vi.fn(async (_sheet: string, _widthChars: number | null) => {}),
      internStyle: vi.fn(async () => 0),
      setLocale: vi.fn(async () => true),
      getCalcSettings: vi.fn(async () => defaultCalcSettings),
      setCalcSettings: vi.fn(async () => {}),
      setEngineInfo: vi.fn(async () => {}),
      setInfoOrigin: vi.fn(async () => {}),
      setInfoOriginForSheet: vi.fn(async () => {}),
      setColFormatRuns: vi.fn(async () => {}),
      recalculate: vi.fn(async () => []),
      setSheetDimensions: vi.fn(async () => {}),
      getSheetDimensions: vi.fn(async () => ({ rows: 1_048_576, cols: 16_384 })),
      renameSheet: vi.fn(async () => true),
      setColWidthChars: vi.fn(async () => {}),
      setSheetOrigin: vi.fn(async () => {}),
      setRowStyleId: vi.fn(async () => {}),
      setColStyleId: vi.fn(async () => {}),
      setSheetDefaultStyleId: vi.fn(async () => {}),
      applyOperation: vi.fn(async () => ({ changedCells: [], movedRanges: [], formulaRewrites: [] })),
      rewriteFormulasForCopyDelta: vi.fn(async () => []),
      lexFormula: vi.fn(async () => []),
      lexFormulaPartial: vi.fn(async () => ({ tokens: [], error: null })),
      parseFormulaPartial: vi.fn(async () => ({ ast: null, error: null, context: { function: null } })),
      terminate: vi.fn(),
    };

    const backend = new WasmWorkbookBackend(engine);
    const info = await backend.openWorkbookFromBytes(bytes);

    expect(info).toEqual({
      path: null,
      origin_path: null,
      sheets: [
        { id: "Sheet2", name: "Sheet2" },
        { id: "Sheet1", name: "Sheet1" },
        { id: "Empty", name: "Empty" },
      ],
    });
  });
});
