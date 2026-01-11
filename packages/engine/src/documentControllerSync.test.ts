import { describe, expect, it } from "vitest";

// DocumentController is authored in JS. We keep a minimal `.d.ts` next to it so this
// TS test can import the runtime implementation under `strict` typechecking.
import { DocumentController } from "../../../apps/desktop/src/document/documentController.js";
import { exportDocumentToEngineWorkbookJson } from "./documentControllerSync";

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
});
