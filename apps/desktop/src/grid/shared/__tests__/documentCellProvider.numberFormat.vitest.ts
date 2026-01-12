import { describe, expect, it } from "vitest";

import { DocumentController } from "../../../document/documentController.js";
import { applyNumberFormatPreset } from "../../../formatting/toolbar.js";
import { dateToExcelSerial } from "../../../shared/valueParsing.js";
import { DocumentCellProvider } from "../documentCellProvider.js";

describe("DocumentCellProvider numberFormat display rendering", () => {
  it("renders currency + percent presets as formatted strings and keeps numeric alignment", () => {
    const doc = new DocumentController();
    const sheetId = "Sheet1";

    doc.setCellValue(sheetId, "A1", 1234.5);
    applyNumberFormatPreset(doc, sheetId, "A1", "currency");
    doc.setCellValue(sheetId, "A4", 1234.5);
    applyNumberFormatPreset(doc, sheetId, "A4", "currency");

    doc.setCellValue(sheetId, "A2", 0.5);
    applyNumberFormatPreset(doc, sheetId, "A2", "percent");

    doc.setCellValue(sheetId, "A3", dateToExcelSerial(new Date(Date.UTC(2024, 0, 2))));
    applyNumberFormatPreset(doc, sheetId, "A3", "date");

    const provider = new DocumentCellProvider({
      document: doc,
      getSheetId: () => sheetId,
      headerRows: 1,
      headerCols: 1,
      rowCount: 10,
      colCount: 10,
      showFormulas: () => false,
      getComputedValue: ({ row, col }) => {
        const cell = doc.getCell(sheetId, { row, col });
        return cell.formula ? null : cell.value;
      }
    });

    const a1 = provider.getCell(1, 1);
    expect(a1?.value).toBe("$1,234.50");
    expect(a1?.style?.textAlign).toBe("end");

    const a2 = provider.getCell(2, 1);
    expect(a2?.value).toBe("50%");
    expect(a2?.style?.textAlign).toBe("end");

    const a3 = provider.getCell(3, 1);
    expect(a3?.value).toBe("1/2/2024");
    expect(a3?.style?.textAlign).toBe("end");

    // Numeric alignment should reuse a shared style object when the underlying formatting
    // has no explicit alignment (avoid per-cell `{...style, textAlign:'end'}` allocations).
    const a4 = provider.getCell(4, 1);
    expect(a4?.value).toBe("$1,234.50");
    expect(a4?.style?.textAlign).toBe("end");
    // A1 and A4 share the same numberFormat preset (currency), so the aligned style should be reused.
    expect(a1?.style).toBe(a4?.style);
  });

  it("does not override an explicit horizontal alignment when applying number formats", () => {
    const doc = new DocumentController();
    const sheetId = "Sheet1";

    doc.setCellValue(sheetId, "A1", 1234.5);
    applyNumberFormatPreset(doc, sheetId, "A1", "currency");
    // Explicitly left-align the cell; the provider should preserve this even though the formatted
    // value is rendered as a string.
    doc.setRangeFormat(sheetId, "A1", { alignment: { horizontal: "left" } });

    const provider = new DocumentCellProvider({
      document: doc,
      getSheetId: () => sheetId,
      headerRows: 1,
      headerCols: 1,
      rowCount: 10,
      colCount: 10,
      showFormulas: () => false,
      getComputedValue: ({ row, col }) => {
        const cell = doc.getCell(sheetId, { row, col });
        return cell.formula ? null : cell.value;
      }
    });

    const cell = provider.getCell(1, 1);
    expect(cell?.value).toBe("$1,234.50");
    // Explicit left alignment wins.
    expect(cell?.style?.textAlign).toBe("start");
  });
});
