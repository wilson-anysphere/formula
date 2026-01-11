import { describe, expect, it, vi } from "vitest";

// DocumentController is authored in JS. We keep a minimal `.d.ts` next to it so this
// TS test can import the runtime implementation under `strict` typechecking.
import { DocumentController } from "../../../apps/desktop/src/document/documentController.js";
import { engineApplyDeltas, engineHydrateFromDocument, exportDocumentToEngineWorkbookJson } from "./documentControllerSync";

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
    const [serialized] = engine.loadWorkbookFromJson.mock.calls[0] ?? [];
    expect(JSON.parse(String(serialized))).toEqual(exportDocumentToEngineWorkbookJson(doc));

    expect(engine.recalculate).toHaveBeenCalledTimes(1);
    expect(calls).toEqual(["loadWorkbookFromJson", "recalculate"]);
  });

  it("applies incremental deltas without emitting format-only updates", async () => {
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
});
