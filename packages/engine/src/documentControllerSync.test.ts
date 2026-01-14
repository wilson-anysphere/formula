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
    readonly internStyleCalls: unknown[] = [];
    readonly setStyleCalls: Array<{ address: string; styleId: number; sheet?: string }> = [];
    readonly colWidthCalls: Array<{ sheet: string; col: number; widthChars: number | null }> = [];
    readonly recalcCalls: Array<string | undefined> = [];
    constructor(private readonly recalcResult: CellChange[]) {}

    async loadWorkbookFromJson(json: string): Promise<void> {
      this.loadedJson.push(json);
    }

    async setCell(address: string, value: any, sheet?: string): Promise<void> {
      this.setCalls.push({ address, value, sheet });
    }

    async internStyle(style: unknown): Promise<number> {
      this.internStyleCalls.push(style);
      return this.internStyleCalls.length;
    }

    async setCellStyleId(sheet: string, address: string, styleId: number): Promise<void> {
      this.setStyleCalls.push({ address, styleId, sheet });
    }

    async setColWidth(sheet: string, col: number, widthChars: number | null): Promise<void> {
      this.colWidthCalls.push({ sheet, col, widthChars });
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

  it("engineHydrateFromDocument syncs sheet view column widths into engine metadata", async () => {
    const doc = new DocumentController();
    doc.setColWidth("Sheet1", 0, 120);

    const engine = new FakeEngine([]);
    await engineHydrateFromDocument(engine, doc);

    // 120px -> Excel character width (1/256 precision): floor(((120-5)/7)*256)/256
    expect(engine.colWidthCalls).toEqual([{ sheet: "Sheet1", col: 0, widthChars: 16.42578125 }]);
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
    expect(setRowStyleId).toHaveBeenCalledWith(0, 100, "Sheet1");
    expect(setColStyleId).toHaveBeenCalledWith(0, 100, "Sheet1");
    expect(setSheetDefaultStyleId).toHaveBeenCalledWith(100, "Sheet1");
  });

  it("syncs sheet view column width overrides into the engine (CELL width metadata)", async () => {
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

    // 120px -> Excel character width (1/256 precision): floor(((120-5)/7)*256)/256
    expect(engine.colWidthCalls).toEqual([{ sheet: "Sheet1", col: 0, widthChars: 16.42578125 }]);
    // Column resizes should trigger a recalc even though DocumentController emits `recalc: false`.
    expect(engine.recalcCalls).toEqual([undefined]);
  });
});
