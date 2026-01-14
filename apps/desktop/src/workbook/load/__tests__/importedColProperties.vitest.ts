import { describe, expect, it } from "vitest";

import { DocumentController } from "../../../document/documentController.js";
import {
  defaultColWidthCharsFromImportedColProperties,
  docColWidthsFromImportedColProperties,
  hiddenColsFromImportedColProperties,
  sheetColWidthsFromViewOrImportedColProperties,
} from "../importedColProperties.js";

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
      defaultColWidth: 20,
      colProperties: {
        "0": { width: 12, hidden: true },
        "1": { hidden: true },
        "2": { width: 8.43, hidden: false },
      },
    };

    expect(hiddenColsFromImportedColProperties(payload)).toEqual([0, 1]);
  });

  it("extracts sheet default column width from imported metadata", () => {
    const payload = {
      schemaVersion: 1,
      defaultColWidth: 20,
      colProperties: {},
    };
    expect(defaultColWidthCharsFromImportedColProperties(payload)).toBe(20);
    expect(defaultColWidthCharsFromImportedColProperties({ schemaVersion: 1, colProperties: {} })).toBeNull();
  });

  it("falls back to imported column widths when persisted sheet view colWidths is empty", () => {
    const importedPayload = {
      schemaVersion: 1,
      colProperties: {
        // Mirrors `fixtures/xlsx/basic/row-col-attrs.xlsx`: B column width 25 chars.
        "1": { width: 25, hidden: false },
      },
    };

    const view = { schemaVersion: 1, colWidths: {} };
    expect(sheetColWidthsFromViewOrImportedColProperties(view, importedPayload)).toEqual({ "1": 180 });
  });

  it("prefers persisted sheet view colWidths when present", () => {
    const importedPayload = {
      schemaVersion: 1,
      colProperties: {
        "1": { width: 25, hidden: false },
      },
    };

    const view = { schemaVersion: 1, colWidths: { "1": 999 } };
    expect(sheetColWidthsFromViewOrImportedColProperties(view, importedPayload)).toEqual({ "1": 999 });
  });

  it("merges persisted sheet view colWidths with imported widths for other columns", () => {
    const importedPayload = {
      schemaVersion: 1,
      colProperties: {
        "1": { width: 25, hidden: false },
        "2": { width: 8.43, hidden: false },
      },
    };

    const view = { schemaVersion: 1, colWidths: { "0": 150 } };
    expect(sheetColWidthsFromViewOrImportedColProperties(view, importedPayload)).toEqual({
      "0": 150,
      "1": 180,
      "2": 64,
    });
  });

  it("hydrates DocumentController sheet view state with imported colWidths", () => {
    const payload = {
      schemaVersion: 1,
      colProperties: {
        // Mirrors `fixtures/xlsx/basic/row-col-attrs.xlsx`: B column width 25 chars.
        "1": { width: 25, hidden: false },
        // C is hidden in the source workbook, but has no explicit width.
        "2": { hidden: true },
      },
    };

    const colWidths = docColWidthsFromImportedColProperties(payload);
    expect(colWidths).toEqual({ "1": 180 });

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

    expect(doc.getSheetView("Sheet1").colWidths).toEqual({ "1": 180 });
  });
});
