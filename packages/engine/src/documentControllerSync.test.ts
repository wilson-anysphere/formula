import { describe, expect, it, vi } from "vitest";

// DocumentController is authored in JS. We keep a minimal `.d.ts` next to it so this
// TS test can import the runtime implementation under `strict` typechecking.
import { DocumentController } from "../../../apps/desktop/src/document/documentController.js";
import type { CellChange } from "./protocol.ts";
import {
  engineApplyDeltas,
  engineApplyDocumentChange,
  engineHydrateFromDocument,
  exportDocumentToEngineWorkbookJson,
  type EngineSyncTarget,
} from "./documentControllerSync.ts";

describe("DocumentController → engine workbook JSON exporter", () => {
  it("exports scalar values, rich text as plain text, and normalizes formulas", () => {
    const doc = new DocumentController();

    doc.setCellValue("Sheet1", "A1", 1);
    doc.setCellFormula("Sheet1", "A2", "A1*2"); // note: no leading '='
    doc.setCellValue("Sheet1", "B1", { text: "Hello", runs: [{ start: 0, end: 5, style: { bold: true } }] });

    // Formatting-only cell: should not be emitted to the engine JSON.
    doc.setRangeFormat("Sheet1", "C1", { font: { italic: true } });

    const json = exportDocumentToEngineWorkbookJson(doc);

    expect(json).toEqual({
      sheetOrder: ["Sheet1"],
      sheets: {
        Sheet1: {
          cells: {
            A1: 1,
            A2: "=A1*2",
            B1: "Hello",
          },
        },
      },
    });
  });

  it("exports stable sheet ids even when they differ from display names", () => {
    const doc = new DocumentController();
    doc.setCellValue("Sheet1", "A1", 1);

    doc.addSheet({ sheetId: "sheet_2", name: "Budget" });
    doc.setCellValue("sheet_2", "A1", 2);

    expect(exportDocumentToEngineWorkbookJson(doc)).toEqual({
      sheetOrder: ["Sheet1", "sheet_2"],
      sheets: {
        Sheet1: { cells: { A1: 1 } },
        sheet_2: { cells: { A1: 2 } },
      },
    });
  });

  it("preserves DocumentController sheet tab order (sheetOrder) rather than sheetMeta insertion order", () => {
    const doc = new DocumentController();
    doc.setCellValue("Sheet1", "A1", 1);

    doc.addSheet({ sheetId: "sheet_2", name: "Budget" });
    doc.addSheet({ sheetId: "sheet_3", name: "Costs" });

    // Reorder the sheets so the model order differs from the sheetMeta insertion order.
    doc.reorderSheets(["sheet_3", "Sheet1", "sheet_2"]);

    const json = exportDocumentToEngineWorkbookJson(doc);
    expect(json.sheetOrder).toEqual(["sheet_3", "Sheet1", "sheet_2"]);
  });

  it("includes localeId in exported workbook JSON when provided", () => {
    const doc = new DocumentController();
    doc.setCellFormula("Sheet1", "A1", "1+1");

    const json = exportDocumentToEngineWorkbookJson(doc, { localeId: "de-DE" });
    expect(json.localeId).toBe("de-DE");
    expect(json.sheetOrder).toEqual(["Sheet1"]);
    expect(json.sheets.Sheet1.cells.A1).toBe("=1+1");
  });

  it("hydrates an engine from exported JSON in a single load+recalc step", async () => {
    const doc = new DocumentController();
    doc.setCellValue("Sheet1", "A1", 1);
    doc.setCellFormula("Sheet1", "A2", "A1*2");

    const calls: string[] = [];
    const engine = {
      loadWorkbookFromJson: vi.fn(async (_serialized: string) => {
        calls.push("loadWorkbookFromJson");
      }),
      setCell: vi.fn(async () => {}),
      recalculate: vi.fn(async () => {
        calls.push("recalculate");
        return [];
      }),
    };

    await engineHydrateFromDocument(engine, doc);

    expect(engine.loadWorkbookFromJson).toHaveBeenCalledTimes(1);
    const serialized = engine.loadWorkbookFromJson.mock.calls[0]?.[0];
    expect(JSON.parse(String(serialized))).toEqual(exportDocumentToEngineWorkbookJson(doc));

    expect(engine.recalculate).toHaveBeenCalledTimes(1);
    expect(calls).toEqual(["loadWorkbookFromJson", "recalculate"]);
  });

  it("embeds localeId in engine hydration JSON when provided", async () => {
    const doc = new DocumentController();
    doc.setCellFormula("Sheet1", "A1", "1+1");

    const engine = {
      loadWorkbookFromJson: vi.fn(async (_serialized: string) => {}),
      setCell: vi.fn(async () => {}),
      recalculate: vi.fn(async () => []),
    };

    await engineHydrateFromDocument(engine, doc, { localeId: "de-DE" });

    const serialized = engine.loadWorkbookFromJson.mock.calls[0]?.[0];
    const parsed = JSON.parse(String(serialized));
    expect(parsed.localeId).toBe("de-DE");
  });

  it("hydrates sheet display names when the engine supports setSheetDisplayName", async () => {
    const doc = new DocumentController();
    doc.addSheet({ sheetId: "sheet_2", name: "Budget" });

    const engine = {
      loadWorkbookFromJson: vi.fn(async () => {}),
      setCell: vi.fn(async () => {}),
      recalculate: vi.fn(async () => []),
      setSheetDisplayName: vi.fn(async () => {}),
    };

    await engineHydrateFromDocument(engine, doc);
    expect(engine.setSheetDisplayName).toHaveBeenCalledWith("sheet_2", "Budget");
  });

  it("applies incremental deltas without emitting format-only updates when style sync hooks are absent", async () => {
    const engine = {
      loadWorkbookFromJson: vi.fn(async () => {}),
      setCell: vi.fn(async () => {}),
      setCells: vi.fn(async () => {}),
      recalculate: vi.fn(async () => []),
    };

    await engineApplyDeltas(engine, [
      {
        sheetId: "Sheet1",
        row: 0,
        col: 0,
        before: { value: null, formula: null, styleId: 0 },
        after: { value: 1, formula: null, styleId: 0 },
      },
      // Formatting-only delta: should be ignored.
      {
        sheetId: "Sheet1",
        row: 0,
        col: 1,
        before: { value: null, formula: null, styleId: 0 },
        after: { value: null, formula: null, styleId: 42 },
      },
    ]);

    expect(engine.setCells).toHaveBeenCalledTimes(1);
    expect(engine.setCells).toHaveBeenCalledWith([{ address: "A1", value: 1, sheet: "Sheet1" }]);
    expect(engine.recalculate).toHaveBeenCalledTimes(1);
  });

  it("supports mapping DocumentController sheet ids to engine sheet names when applying deltas", async () => {
    const engine = {
      loadWorkbookFromJson: vi.fn(async () => {}),
      setCell: vi.fn(async () => {}),
      setCells: vi.fn(async () => {}),
      recalculate: vi.fn(async () => []),
    };

    await engineApplyDeltas(
      engine,
      [
        {
          sheetId: "sheet_2",
          row: 0,
          col: 0,
          before: { value: null, formula: null, styleId: 0 },
          after: { value: 1, formula: null, styleId: 0 },
        },
      ],
      { sheetIdToSheet: (sheetId) => (sheetId === "sheet_2" ? "Budget" : sheetId) },
    );

    expect(engine.setCells).toHaveBeenCalledTimes(1);
    expect(engine.setCells).toHaveBeenCalledWith([{ address: "A1", value: 1, sheet: "Budget" }]);
  });

  it("propagates cleared cells as null inputs (sparse clears)", async () => {
    const engine = {
      loadWorkbookFromJson: vi.fn(async () => {}),
      setCell: vi.fn(async () => {}),
      setCells: vi.fn(async () => {}),
      recalculate: vi.fn(async () => []),
    };

    await engineApplyDeltas(engine, [
      {
        sheetId: "Sheet1",
        row: 0,
        col: 0,
        before: { value: 1, formula: null, styleId: 0 },
        after: { value: null, formula: null, styleId: 0 },
      },
    ]);

    expect(engine.setCells).toHaveBeenCalledTimes(1);
    expect(engine.setCells).toHaveBeenCalledWith([{ address: "A1", value: null, sheet: "Sheet1" }]);
    expect(engine.recalculate).toHaveBeenCalledTimes(1);
  });
});

describe("engine sync helpers", () => {
  class FakeEngine implements EngineSyncTarget {
    readonly loadedJson: string[] = [];
    readonly setCalls: Array<{ address: string; value: unknown; sheet?: string }> = [];
    readonly sheetDisplayNameCalls: Array<{ sheetId: string; name: string }> = [];
    readonly internStyleCalls: unknown[] = [];
    readonly setStyleCalls: Array<{ address: string; styleId: number; sheet?: string }> = [];
    readonly sheetDefaultStyleCalls: Array<{ sheet: string; styleId: number | null }> = [];
    readonly rowStyleCalls: Array<{ sheet: string; row: number; styleId: number | null }> = [];
    readonly colStyleCalls: Array<{ sheet: string; col: number; styleId: number | null }> = [];
    readonly formatRunsByColCalls: Array<{
      sheet: string;
      col: number;
      runs: Array<{ startRow: number; endRowExclusive: number; styleId: number }>;
    }> = [];
    readonly colFormatRunsCalls: Array<{
      sheet: string;
      col: number;
      runs: Array<{ startRow: number; endRowExclusive: number; styleId: number }>;
    }> = [];
    readonly colWidthCalls: Array<{ sheet: string; col: number; widthChars: number | null }> = [];
    readonly renameSheetCalls: Array<{ oldName: string; newName: string }> = [];
    readonly recalcCalls: Array<string | undefined> = [];
    constructor(private readonly recalcResult: CellChange[]) {}

    async loadWorkbookFromJson(json: string): Promise<void> {
      this.loadedJson.push(json);
    }

    async setCell(address: string, value: any, sheet?: string): Promise<void> {
      this.setCalls.push({ address, value, sheet });
    }

    async setSheetDisplayName(sheetId: string, name: string): Promise<void> {
      this.sheetDisplayNameCalls.push({ sheetId, name });
    }

    async internStyle(style: unknown): Promise<number> {
      this.internStyleCalls.push(style);
      return this.internStyleCalls.length;
    }

    async setCellStyleId(sheet: string, address: string, styleId: number): Promise<void> {
      this.setStyleCalls.push({ address, styleId, sheet });
    }

    async setSheetDefaultStyleId(sheet: string, styleId: number | null): Promise<void> {
      this.sheetDefaultStyleCalls.push({ sheet, styleId });
    }

    async setRowStyleId(sheet: string, row: number, styleId: number | null): Promise<void> {
      this.rowStyleCalls.push({ sheet, row, styleId });
    }

    async setColStyleId(sheet: string, col: number, styleId: number | null): Promise<void> {
      this.colStyleCalls.push({ sheet, col, styleId });
    }

    async setFormatRunsByCol(
      sheet: string,
      col: number,
      runs: Array<{ startRow: number; endRowExclusive: number; styleId: number }>,
    ): Promise<void> {
      this.formatRunsByColCalls.push({ sheet, col, runs });
    }

    async setColFormatRuns(
      sheet: string,
      col: number,
      runs: Array<{ startRow: number; endRowExclusive: number; styleId: number }>,
    ): Promise<void> {
      this.colFormatRunsCalls.push({ sheet, col, runs });
    }

    async setColWidth(sheet: string, col: number, widthChars: number | null): Promise<void> {
      this.colWidthCalls.push({ sheet, col, widthChars });
    }

    async renameSheet(oldName: string, newName: string): Promise<boolean> {
      this.renameSheetCalls.push({ oldName, newName });
      return true;
    }

    async recalculate(sheet?: string): Promise<CellChange[]> {
      this.recalcCalls.push(sheet);
      return this.recalcResult;
    }
  }

  it("engineHydrateFromDocument returns the engine's recalc changes", async () => {
    const doc = new DocumentController();
    doc.setCellValue("Sheet1", "A1", 1);
    doc.setCellFormula("Sheet1", "A2", "A1*2");

    const expected: CellChange[] = [{ sheet: "Sheet1", address: "A2", value: 2 }];
    const engine = new FakeEngine(expected);

    const changes = await engineHydrateFromDocument(engine, doc);

    expect(engine.loadedJson).toHaveLength(1);
    expect(engine.recalcCalls).toEqual([undefined]);
    expect(changes).toEqual(expected);
  });

  it("engineApplyDeltas skips formatting-only edits when style sync hooks are absent (no recalc)", async () => {
    const engine = new FakeEngine([{ sheet: "Sheet1", address: "A1", value: 1 }]);

    // Disable style sync hooks for this test.
    (engine as Partial<EngineSyncTarget> & { internStyle?: unknown; setCellStyleId?: unknown }).internStyle = undefined;
    (engine as Partial<EngineSyncTarget> & { internStyle?: unknown; setCellStyleId?: unknown }).setCellStyleId = undefined;

    const changes = await engineApplyDeltas(engine, [
      {
        sheetId: "Sheet1",
        row: 0,
        col: 0,
        before: { value: 1, formula: null, styleId: 0 },
        after: { value: 1, formula: null, styleId: 123 },
      },
    ]);

    expect(changes).toEqual([]);
    expect(engine.setCalls).toEqual([]);
    expect(engine.recalcCalls).toEqual([]);
  });

  it("engineHydrateFromDocument syncs styleIds when style sync hooks are available", async () => {
    const doc = new DocumentController();
    doc.setRangeFormat("Sheet1", "A1", { font: { italic: true } });
    const styleId = doc.getCell("Sheet1", "A1").styleId;
    expect(styleId).not.toBe(0);

    const engine = new FakeEngine([]);

    await engineHydrateFromDocument(engine, doc);

    expect(engine.internStyleCalls).toEqual([doc.styleTable.get(styleId)]);
    expect(engine.setStyleCalls).toEqual([{ address: "A1", styleId: 1, sheet: "Sheet1" }]);
  });

  it("engineHydrateFromDocument syncs sheet/row/col default style layers when the optional hooks are available", async () => {
    const doc = new DocumentController();
    doc.addSheet({ sheetId: "sheet_2", name: "Budget" });

    doc.setRowFormat("sheet_2", 0, { font: { bold: true } });
    doc.setColFormat("sheet_2", 1, { font: { bold: true } });
    doc.setSheetFormat("sheet_2", { font: { bold: true } });

    const sheetModel = (doc as any)?.model?.sheets?.get?.("sheet_2");
    const rowDocStyleId = sheetModel?.rowStyleIds?.get?.(0) as number | undefined;
    const colDocStyleId = sheetModel?.colStyleIds?.get?.(1) as number | undefined;
    const sheetDocStyleId = sheetModel?.defaultStyleId as number | undefined;
    expect(rowDocStyleId).toBeTruthy();
    expect(rowDocStyleId).toBe(colDocStyleId);
    expect(rowDocStyleId).toBe(sheetDocStyleId);

    const engine = new FakeEngine([]);
    await engineHydrateFromDocument(engine, doc);

    // Engine is addressed by stable sheet ids, and interns the shared style once.
    expect(engine.internStyleCalls).toEqual([doc.styleTable.get(rowDocStyleId!)]);
    expect(engine.sheetDefaultStyleCalls).toEqual([{ sheet: "sheet_2", styleId: 1 }]);
    expect(engine.rowStyleCalls).toEqual([{ sheet: "sheet_2", row: 0, styleId: 1 }]);
    expect(engine.colStyleCalls).toEqual([{ sheet: "sheet_2", col: 1, styleId: 1 }]);
  });

  it("engineHydrateFromDocument syncs compressed range-run formatting when the engine supports it", async () => {
    const doc = new DocumentController();

    // Force range-run formatting (area > RANGE_RUN_FORMAT_THRESHOLD).
    doc.setRangeFormat("Sheet1", "A1:Z2000", { font: { italic: true } });

    const sheet = (doc as any)?.model?.sheets?.get?.("Sheet1");
    const col0Runs = sheet?.formatRunsByCol?.get?.(0) ?? [];
    expect(col0Runs.length).toBeGreaterThan(0);

    const docStyleId = col0Runs[0]?.styleId ?? 0;
    expect(docStyleId).not.toBe(0);

    const engine = new FakeEngine([]);
    await engineHydrateFromDocument(engine, doc);

    expect(engine.internStyleCalls).toEqual(expect.arrayContaining([doc.styleTable.get(docStyleId)]));

    const col0 = engine.formatRunsByColCalls.find((call) => call.sheet === "Sheet1" && call.col === 0);
    expect(col0).toBeTruthy();
    expect(col0?.runs).toEqual(
      expect.arrayContaining([{ startRow: 0, endRowExclusive: 2000, styleId: 1 }]),
    );
    // Prefer the modern API when available.
    expect(engine.colFormatRunsCalls).toEqual([]);
  });

  it("engineHydrateFromDocument syncs sheet view column widths into engine metadata", async () => {
    const doc = new DocumentController();
    doc.setColWidth("Sheet1", 0, 120);

    const engine = new FakeEngine([]);
    await engineHydrateFromDocument(engine, doc);

    // 120px -> Excel character width (default Calibri 11 metrics).
    expect(engine.colWidthCalls).toEqual([{ sheet: "Sheet1", col: 0, widthChars: 16.43 }]);
  });

  it("engineApplyDeltas propagates formatting-only deltas via internStyle + setCellStyleId", async () => {
    const doc = new DocumentController();
    // Intern a style into the document's style table without attaching it to a cell so
    // we can assert the delta path triggers `internStyle`.
    const docStyleId = doc.styleTable.intern({ font: { bold: true } });

    const engine = new FakeEngine([]);
    await engineHydrateFromDocument(engine, doc);

    await engineApplyDeltas(
      engine,
      [
        {
          sheetId: "Sheet1",
          row: 0,
          col: 0,
          before: { value: null, formula: null, styleId: 0 },
          after: { value: null, formula: null, styleId: docStyleId },
        },
        // Same style id should not be re-interned.
        {
          sheetId: "Sheet1",
          row: 0,
          col: 1,
          before: { value: null, formula: null, styleId: 0 },
          after: { value: null, formula: null, styleId: docStyleId },
        },
      ],
      { recalculate: false },
    );

    expect(engine.internStyleCalls).toEqual([doc.styleTable.get(docStyleId)]);
    expect(engine.setStyleCalls).toEqual([
      { address: "A1", styleId: 1, sheet: "Sheet1" },
      { address: "B1", styleId: 1, sheet: "Sheet1" },
    ]);
  });

  it("engineApplyDeltas propagates value changes and returns recalc changes", async () => {
    const expected: CellChange[] = [{ sheet: "Sheet1", address: "B1", value: 4 }];
    const engine = new FakeEngine(expected);

    const changes = await engineApplyDeltas(engine, [
      {
        sheetId: "Sheet1",
        row: 0,
        col: 0,
        before: { value: 1, formula: null, styleId: 0 },
        after: { value: 2, formula: null, styleId: 0 },
      },
    ]);

    expect(engine.setCalls).toEqual([{ address: "A1", value: 2, sheet: "Sheet1" }]);
    expect(engine.recalcCalls).toEqual([undefined]);
    expect(changes).toEqual(expected);
  });

  it("engineApplyDeltas propagates clears (null) via setCell and returns recalc changes", async () => {
    const expected: CellChange[] = [{ sheet: "Sheet1", address: "B1", value: 0 }];
    const engine = new FakeEngine(expected);

    const changes = await engineApplyDeltas(engine, [
      {
        sheetId: "Sheet1",
        row: 0,
        col: 0,
        before: { value: 1, formula: null, styleId: 0 },
        after: { value: null, formula: null, styleId: 0 },
      },
    ]);

    expect(engine.setCalls).toEqual([{ address: "A1", value: null, sheet: "Sheet1" }]);
    expect(engine.recalcCalls).toEqual([undefined]);
    expect(changes).toEqual(expected);
  });

  it("engineApplyDeltas normalizes formulas to start with '='", async () => {
    const engine = new FakeEngine([]);

    await engineApplyDeltas(engine, [
      {
        sheetId: "Sheet1",
        row: 0,
        col: 1,
        before: { value: null, formula: null, styleId: 0 },
        after: { value: null, formula: "A1*2", styleId: 0 },
      },
    ]);

    expect(engine.setCalls).toEqual([{ address: "B1", value: "=A1*2", sheet: "Sheet1" }]);
  });

  it("engineApplyDocumentChange propagates row/col/sheet style deltas via the optional formatting API", async () => {
    const doc = new DocumentController();

    const internStyle = vi.fn((_: unknown) => 100);
    const setRowStyleId = vi.fn();
    const setColStyleId = vi.fn();
    const setSheetDefaultStyleId = vi.fn();

    const engine: EngineSyncTarget = {
      loadWorkbookFromJson: vi.fn(async () => {}),
      setCell: vi.fn(async () => {}),
      recalculate: vi.fn(async () => []),
      internStyle,
      setRowStyleId,
      setColStyleId,
      setSheetDefaultStyleId,
    };

    const pending: Promise<unknown>[] = [];
    doc.on("change", (payload: any) => {
      pending.push(
        engineApplyDocumentChange(engine, payload, {
          recalculate: false,
          getStyleById: (styleId) => doc.styleTable.get(styleId),
        }),
      );
    });

    // Apply the same patch three times; DocumentController should reuse the same styleId, and the
    // sync helper should intern it into the engine once and then reuse the cached mapping.
    doc.setRowFormat("Sheet1", 0, { font: { bold: true } });
    doc.setColFormat("Sheet1", 0, { font: { bold: true } });
    doc.setSheetFormat("Sheet1", { font: { bold: true } });

    await Promise.all(pending);

    expect(internStyle).toHaveBeenCalledTimes(1);
    expect(setRowStyleId).toHaveBeenCalledWith("Sheet1", 0, 100);
    expect(setColStyleId).toHaveBeenCalledWith("Sheet1", 0, 100);
    expect(setSheetDefaultStyleId).toHaveBeenCalledWith("Sheet1", 100);
  });

  it("forces a recalc for cell styleId-only deltas even when DocumentController emits recalc=false", async () => {
    const expected: CellChange[] = [{ sheet: "Sheet1", address: "A1", value: 123 }];
    const engine = new FakeEngine(expected);

    const styleObj = { font: { bold: true } };
    const payload = {
      recalc: false,
      deltas: [
        {
          sheetId: "Sheet1",
          row: 0,
          col: 0,
          before: { value: null, formula: null, styleId: 0 },
          after: { value: null, formula: null, styleId: 1 },
        },
      ],
    };

    const changes = await engineApplyDocumentChange(engine, payload, {
      getStyleById: (styleId) => (styleId === 1 ? styleObj : null),
    });

    expect(engine.setCalls).toEqual([]);
    expect(engine.internStyleCalls).toEqual([styleObj]);
    expect(engine.setStyleCalls).toEqual([{ address: "A1", styleId: 1, sheet: "Sheet1" }]);
    expect(engine.recalcCalls).toEqual([undefined]);
    expect(changes).toEqual(expected);
  });

  it("does not force a recalc for cell styleId-only deltas when style sync hooks are absent", async () => {
    const engine = new FakeEngine([{ sheet: "Sheet1", address: "A1", value: 123 }]);

    // Disable style sync hooks for this test.
    (engine as Partial<EngineSyncTarget> & { internStyle?: unknown; setCellStyleId?: unknown }).internStyle = undefined;
    (engine as Partial<EngineSyncTarget> & { internStyle?: unknown; setCellStyleId?: unknown }).setCellStyleId = undefined;

    const payload = {
      recalc: false,
      deltas: [
        {
          sheetId: "Sheet1",
          row: 0,
          col: 0,
          before: { value: null, formula: null, styleId: 0 },
          after: { value: null, formula: null, styleId: 1 },
        },
      ],
    };

    const changes = await engineApplyDocumentChange(engine, payload, {
      getStyleById: () => ({ font: { bold: true } }),
    });

    expect(changes).toEqual([]);
    expect(engine.recalcCalls).toEqual([]);
  });

  it("forces a recalc for row/col/sheet style layer deltas even when DocumentController emits recalc=false", async () => {
    const styleObj = { font: { italic: true } };
    const recalcResult: CellChange[] = [{ sheet: "Sheet1", address: "B2", value: 456 }];

    const internStyle = vi.fn((_: unknown) => 100);
    const setRowStyleId = vi.fn(async () => {});
    const setColStyleId = vi.fn(async () => {});
    const setSheetDefaultStyleId = vi.fn(async () => {});
    const recalculate = vi.fn(async () => recalcResult);

    const engine: EngineSyncTarget = {
      loadWorkbookFromJson: vi.fn(async () => {}),
      setCell: vi.fn(async () => {}),
      recalculate,
      internStyle,
      setRowStyleId,
      setColStyleId,
      setSheetDefaultStyleId,
    };

    const payload = {
      recalc: false,
      rowStyleDeltas: [{ sheetId: "Sheet1", row: 0, afterStyleId: 7 }],
      colStyleDeltas: [{ sheetId: "Sheet1", col: 1, afterStyleId: 7 }],
      sheetStyleDeltas: [{ sheetId: "Sheet1", afterStyleId: 7 }],
    };

    const changes = await engineApplyDocumentChange(engine, payload, {
      getStyleById: (styleId) => (styleId === 7 ? styleObj : null),
    });

    expect(internStyle).toHaveBeenCalledTimes(1);
    expect(internStyle).toHaveBeenCalledWith(styleObj);
    expect(setRowStyleId).toHaveBeenCalledWith("Sheet1", 0, 100);
    expect(setColStyleId).toHaveBeenCalledWith("Sheet1", 1, 100);
    expect(setSheetDefaultStyleId).toHaveBeenCalledWith("Sheet1", 100);

    expect(recalculate).toHaveBeenCalledTimes(1);
    expect(changes).toEqual(recalcResult);
  });

  it("maps row/col/sheet style layer deltas through sheetIdToSheet when provided", async () => {
    const styleObj = { font: { italic: true } };
    const recalcResult: CellChange[] = [{ sheet: "Budget", address: "B2", value: 456 }];
    const sheetIdToSheet = (sheetId: string) => (sheetId === "sheet_2" ? "  Budget  " : sheetId);

    const internStyle = vi.fn((_: unknown) => 100);
    const setRowStyleId = vi.fn(async () => {});
    const setColStyleId = vi.fn(async () => {});
    const setSheetDefaultStyleId = vi.fn(async () => {});
    const recalculate = vi.fn(async () => recalcResult);

    const engine: EngineSyncTarget = {
      loadWorkbookFromJson: vi.fn(async () => {}),
      setCell: vi.fn(async () => {}),
      recalculate,
      internStyle,
      setRowStyleId,
      setColStyleId,
      setSheetDefaultStyleId,
    };

    const payload = {
      recalc: false,
      rowStyleDeltas: [{ sheetId: "sheet_2", row: 0, afterStyleId: 7 }],
      colStyleDeltas: [{ sheetId: "sheet_2", col: 1, afterStyleId: 7 }],
      sheetStyleDeltas: [{ sheetId: "sheet_2", afterStyleId: 7 }],
    };

    const changes = await engineApplyDocumentChange(engine, payload, {
      getStyleById: (styleId) => (styleId === 7 ? styleObj : null),
      sheetIdToSheet,
    });

    expect(internStyle).toHaveBeenCalledTimes(1);
    expect(internStyle).toHaveBeenCalledWith(styleObj);
    expect(setRowStyleId).toHaveBeenCalledWith("Budget", 0, 100);
    expect(setColStyleId).toHaveBeenCalledWith("Budget", 1, 100);
    expect(setSheetDefaultStyleId).toHaveBeenCalledWith("Budget", 100);

    expect(recalculate).toHaveBeenCalledTimes(1);
    expect(changes).toEqual(recalcResult);
  });

  it("clears row/col/sheet style layers by passing null (without interning styles) and still forces a recalc", async () => {
    const recalcResult: CellChange[] = [{ sheet: "Sheet1", address: "C3", value: 123 }];

    const internStyle = vi.fn(() => 999);
    const setRowStyleId = vi.fn(async () => {});
    const setColStyleId = vi.fn(async () => {});
    const setSheetDefaultStyleId = vi.fn(async () => {});
    const recalculate = vi.fn(async () => recalcResult);

    const engine: EngineSyncTarget = {
      loadWorkbookFromJson: vi.fn(async () => {}),
      setCell: vi.fn(async () => {}),
      recalculate,
      // Provide internStyle to assert it is not required for clearing (styleId=0).
      internStyle,
      setRowStyleId,
      setColStyleId,
      setSheetDefaultStyleId,
    };

    const payload = {
      recalc: false,
      rowStyleDeltas: [{ sheetId: "Sheet1", row: 0, afterStyleId: 0 }],
      colStyleDeltas: [{ sheetId: "Sheet1", col: 1, afterStyleId: 0 }],
      sheetStyleDeltas: [{ sheetId: "Sheet1", afterStyleId: 0 }],
    };

    const changes = await engineApplyDocumentChange(engine, payload);

    expect(internStyle).toHaveBeenCalledTimes(0);
    expect(setRowStyleId).toHaveBeenCalledWith("Sheet1", 0, null);
    expect(setColStyleId).toHaveBeenCalledWith("Sheet1", 1, null);
    expect(setSheetDefaultStyleId).toHaveBeenCalledWith("Sheet1", null);

    expect(recalculate).toHaveBeenCalledTimes(1);
    expect(changes).toEqual(recalcResult);
  });

  it("engineApplyDocumentChange syncs sheet view column width overrides into the engine (CELL width metadata)", async () => {
    const doc = new DocumentController();
    let payload: any = null;
    const unsubscribe = doc.on("change", (p: any) => {
      payload = p;
    });

    doc.setColWidth("Sheet1", 0, 120);
    unsubscribe();

    expect(Array.isArray(payload?.sheetViewDeltas)).toBe(true);
    expect(payload.sheetViewDeltas).toHaveLength(1);

    const engine = new FakeEngine([]);
    await engineApplyDocumentChange(engine, payload);

    // 120px -> Excel character width (default Calibri 11 metrics).
    expect(engine.colWidthCalls).toEqual([{ sheet: "Sheet1", col: 0, widthChars: 16.43 }]);
    // Column resizes should trigger a recalc even though DocumentController emits `recalc: false`.
    expect(engine.recalcCalls).toEqual([undefined]);
  });

  it("does not force a recalc for sheet view column widths when width hooks are absent", async () => {
    const doc = new DocumentController();
    let payload: any = null;
    const unsubscribe = doc.on("change", (p: any) => {
      payload = p;
    });

    doc.setColWidth("Sheet1", 0, 120);
    unsubscribe();

    const engine = new FakeEngine([]);
    (engine as Partial<EngineSyncTarget> & { setColWidth?: unknown }).setColWidth = undefined;

    const changes = await engineApplyDocumentChange(engine, payload);

    expect(changes).toEqual([]);
    expect(engine.colWidthCalls).toEqual([]);
    expect(engine.recalcCalls).toEqual([]);
  });

  it("engineApplyDocumentChange triggers a recalc tick for sheet renames (sheet meta deltas)", async () => {
    const doc = new DocumentController();
    let payload: any = null;
    const unsubscribe = doc.on("change", (p: any) => {
      payload = p;
    });

    doc.renameSheet("Sheet1", "Renamed");
    unsubscribe();

    expect(Array.isArray(payload?.sheetMetaDeltas)).toBe(true);
    expect(payload.sheetMetaDeltas).toHaveLength(1);
    expect(payload.recalc).toBe(false);

    const engine = new FakeEngine([]);
    await engineApplyDocumentChange(engine, payload);

    expect(engine.sheetDisplayNameCalls).toEqual([{ sheetId: "Sheet1", name: "Renamed" }]);
    expect(engine.recalcCalls).toEqual([undefined]);
  });

  it("syncs sheet renames via renameSheet and forces a recalc (CELL filename/address metadata)", async () => {
    const doc = new DocumentController();
    let payload: any = null;
    const unsubscribe = doc.on("change", (p: any) => {
      payload = p;
    });

    doc.renameSheet("Sheet1", "Budget");
    unsubscribe();

    expect(payload?.recalc).toBe(false);
    expect(Array.isArray(payload?.sheetMetaDeltas)).toBe(true);

    const calls: string[] = [];
    const renameSheet = vi.fn(async (_oldName: string, _newName: string) => {
      calls.push(`renameSheet:${_oldName}->${_newName}`);
      return true;
    });

    const engine: EngineSyncTarget = {
      loadWorkbookFromJson: vi.fn(async () => {}),
      setCell: vi.fn(async () => {}),
      recalculate: vi.fn(async () => {
        calls.push("recalculate");
        return [];
      }),
      renameSheet,
    };

    await engineApplyDocumentChange(engine, payload);

    expect(renameSheet).toHaveBeenCalledTimes(1);
    expect(renameSheet).toHaveBeenCalledWith("Sheet1", "Budget");
    expect(calls).toEqual(["renameSheet:Sheet1->Budget", "recalculate"]);
  });

  it("engineApplyDocumentChange triggers a recalc tick for formatting-only cell deltas", async () => {
    const doc = new DocumentController();
    let payload: any = null;
    const unsubscribe = doc.on("change", (p: any) => {
      payload = p;
    });

    doc.setRangeFormat("Sheet1", "C1", { font: { italic: true } });
    unsubscribe();
    expect(payload?.recalc).toBe(false);
    const docStyleId = doc.getCell("Sheet1", "C1").styleId;
    expect(docStyleId).not.toBe(0);

    const engine = new FakeEngine([]);
    await engineApplyDocumentChange(engine, payload, {
      getStyleById: (styleId) => doc.styleTable.get(styleId),
    });

    // Formatting-only changes should still advance the engine's recalculation tick so
    // metadata/volatile functions update.
    expect(engine.internStyleCalls).toEqual([doc.styleTable.get(docStyleId)]);
    expect(engine.setStyleCalls).toEqual([{ address: "C1", styleId: 1, sheet: "Sheet1" }]);
    expect(engine.recalcCalls).toEqual([undefined]);
  });
  it("engineApplyDocumentChange syncs compressed range-run formatting deltas into the engine", async () => {
    const doc = new DocumentController();
    let payload: any = null;
    const unsubscribe = doc.on("change", (p: any) => {
      payload = p;
    });

    // Force range-run formatting (area > RANGE_RUN_FORMAT_THRESHOLD).
    doc.setRangeFormat("Sheet1", "A1:Z2000", { font: { italic: true } });
    unsubscribe();

    expect(payload?.recalc).toBe(false);
    expect(Array.isArray(payload?.rangeRunDeltas)).toBe(true);
    expect(payload.rangeRunDeltas.length).toBeGreaterThan(0);

    const firstDelta = payload.rangeRunDeltas[0];
    const docStyleId = firstDelta?.afterRuns?.[0]?.styleId ?? 0;
    expect(docStyleId).not.toBe(0);

    const engine = new FakeEngine([]);
    await engineApplyDocumentChange(engine, payload, {
      getStyleById: (styleId) => doc.styleTable.get(styleId),
    });

    // Should intern the style used by the run and emit at least one run update for column A.
    expect(engine.internStyleCalls).toEqual(expect.arrayContaining([doc.styleTable.get(docStyleId)]));
    const col0 = engine.formatRunsByColCalls.find((call) => call.sheet === "Sheet1" && call.col === 0);
    expect(col0).toBeTruthy();
    expect(col0?.runs).toEqual(expect.arrayContaining([{ startRow: 0, endRowExclusive: 2000, styleId: 1 }]));
    // Prefer the modern API when available.
    expect(engine.colFormatRunsCalls).toEqual([]);

    // Range-run formatting edits are metadata-only; they should still advance the engine's recalc tick.
    expect(engine.recalcCalls).toEqual([undefined]);
  });

  it("engineApplyDocumentChange falls back to setColFormatRuns when setFormatRunsByCol is unavailable", async () => {
    const doc = new DocumentController();
    let payload: any = null;
    const unsubscribe = doc.on("change", (p: any) => {
      payload = p;
    });
    doc.setRangeFormat("Sheet1", "A1:Z2000", { font: { italic: true } });
    unsubscribe();

    const engine = new FakeEngine([]);
    // Shadow the prototype method so `typeof engine.setFormatRunsByCol` is not "function".
    (engine as any).setFormatRunsByCol = undefined;

    await engineApplyDocumentChange(engine, payload, {
      getStyleById: (styleId) => doc.styleTable.get(styleId),
    });

    expect(engine.formatRunsByColCalls).toEqual([]);
    const col0 = engine.colFormatRunsCalls.find((call) => call.sheet === "Sheet1" && call.col === 0);
    expect(col0).toBeTruthy();
    expect(col0?.runs).toEqual(expect.arrayContaining([{ startRow: 0, endRowExclusive: 2000, styleId: 1 }]));
  });

  it("engineApplyDocumentChange clears range-run formatting even when style sync hooks are unavailable", async () => {
    const doc = new DocumentController();

    // Seed a range-run formatting layer.
    doc.setRangeFormat("Sheet1", "A1:Z2000", { font: { italic: true } });

    // Capture the clearing payload.
    let payload: any = null;
    const unsubscribe = doc.on("change", (p: any) => {
      payload = p;
    });
    doc.setRangeFormat("Sheet1", "A1:Z2000", null);
    unsubscribe();

    expect(Array.isArray(payload?.rangeRunDeltas)).toBe(true);
    expect(payload.rangeRunDeltas.length).toBeGreaterThan(0);
    expect(payload.rangeRunDeltas[0]?.afterRuns).toEqual([]);

    const engine = new FakeEngine([]);
    // Disable style sync hooks to simulate a legacy engine that can't resolve style objects.
    (engine as any).internStyle = undefined;
    (engine as any).setCellStyleId = undefined;

    await engineApplyDocumentChange(engine, payload);

    const col0 = engine.formatRunsByColCalls.find((call) => call.sheet === "Sheet1" && call.col === 0);
    expect(col0).toBeTruthy();
    expect(col0?.runs).toEqual([]);
    // Clearing should still force a recalc tick so volatile metadata functions observe the change.
    expect(engine.recalcCalls).toEqual([undefined]);
  });

  it("engineApplyDocumentChange clears range-run formatting via setColFormatRuns fallback when setFormatRunsByCol is unavailable", async () => {
    const doc = new DocumentController();

    doc.setRangeFormat("Sheet1", "A1:Z2000", { font: { italic: true } });

    let payload: any = null;
    const unsubscribe = doc.on("change", (p: any) => {
      payload = p;
    });
    doc.setRangeFormat("Sheet1", "A1:Z2000", null);
    unsubscribe();

    const engine = new FakeEngine([]);
    (engine as any).setFormatRunsByCol = undefined;
    (engine as any).internStyle = undefined;
    (engine as any).setCellStyleId = undefined;

    await engineApplyDocumentChange(engine, payload);

    expect(engine.formatRunsByColCalls).toEqual([]);
    const col0 = engine.colFormatRunsCalls.find((call) => call.sheet === "Sheet1" && call.col === 0);
    expect(col0).toBeTruthy();
    expect(col0?.runs).toEqual([]);
    expect(engine.recalcCalls).toEqual([undefined]);
  });

  it("engineHydrateFromDocument falls back to setColFormatRuns when setFormatRunsByCol is unavailable", async () => {
    const doc = new DocumentController();
    doc.setRangeFormat("Sheet1", "A1:Z2000", { font: { italic: true } });

    const engine = new FakeEngine([]);
    (engine as any).setFormatRunsByCol = undefined;

    await engineHydrateFromDocument(engine, doc);

    expect(engine.formatRunsByColCalls).toEqual([]);
    const col0 = engine.colFormatRunsCalls.find((call) => call.sheet === "Sheet1" && call.col === 0);
    expect(col0).toBeTruthy();
    expect(col0?.runs).toEqual(expect.arrayContaining([{ startRow: 0, endRowExclusive: 2000, styleId: 1 }]));
  });
});
