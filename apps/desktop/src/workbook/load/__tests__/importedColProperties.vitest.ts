import { describe, expect, it } from "vitest";

import { DocumentController } from "../../../document/documentController.js";
import { docColWidthsFromImportedColProperties, hiddenColsFromImportedColProperties } from "../importedColProperties.js";

describe("importedColProperties", () => {
  it("converts imported OOXML column widths into DocumentController colWidths (px)", () => {
    const payload = {
      schemaVersion: 1,
      colProperties: {
        // A: width 12 chars (visible)
        "0": { width: 12, hidden: false },
        // C: default width (visible)
        "2": { width: 8.43, hidden: false },
      },
    };

    const colWidths = docColWidthsFromImportedColProperties(payload);
    expect(colWidths).toEqual({
      // 12 chars -> 89px (per Excel default conversion rules).
      "0": 89,
      // 8.43 chars -> 64px (Excel default).
      "2": 64,
    });
  });

  it("extracts hidden columns from imported column metadata", () => {
    const payload = {
      schemaVersion: 1,
      colProperties: {
        "0": { width: 12, hidden: true },
        "1": { hidden: true },
        "2": { width: 8.43, hidden: false },
      },
    };

    expect(hiddenColsFromImportedColProperties(payload)).toEqual([0, 1]);
  });

  it("hydrates DocumentController sheet view state with imported colWidths", () => {
    const payload = {
      schemaVersion: 1,
      colProperties: {
        "0": { width: 12, hidden: true },
      },
    };

    const colWidths = docColWidthsFromImportedColProperties(payload);
    expect(colWidths).toEqual({ "0": 89 });

    const snapshot = {
      schemaVersion: 1,
      sheets: [
        {
          id: "Sheet1",
          name: "Sheet1",
          visibility: "visible",
          frozenRows: 0,
          frozenCols: 0,
          cells: [],
          colWidths,
        },
      ],
    };

    const doc = new DocumentController();
    doc.applyState(new TextEncoder().encode(JSON.stringify(snapshot)));

    expect(doc.getSheetView("Sheet1").colWidths).toEqual({ "0": 89 });
  });
});

