import { describe, expect, it, vi } from "vitest";

import { DocumentController } from "../document/documentController.js";
import { createSheetNameResolverFromIdToNameMap } from "../sheet/sheetNameResolver.js";
import { DocumentControllerWorkbookAdapter } from "./documentControllerWorkbookAdapter.js";

describe("DocumentControllerWorkbookAdapter (sheet name resolver)", () => {
  it("resolves display names to stable sheet ids (no phantom sheet creation after rename)", () => {
    const controller = new DocumentController();
    controller.setCellValue("Sheet2", "A1", 1);

    const sheetNames = new Map<string, string>([["Sheet2", "Budget"]]);
    const sheetNameResolver = createSheetNameResolverFromIdToNameMap(sheetNames);

    const workbook = new DocumentControllerWorkbookAdapter(controller as any, { sheetNameResolver });

    const sheet = workbook.getSheet("Budget");
    expect(sheet.name).toBe("Budget");

    sheet.setCellValue("A1", 99);

    expect(controller.getCell("Sheet2", "A1").value).toBe(99);
    expect(controller.getSheetIds()).toContain("Sheet2");
    expect(controller.getSheetIds()).not.toContain("Budget");

    expect(() => workbook.getSheet("DoesNotExist")).toThrow(/Unknown sheet/i);
    expect(controller.getSheetIds()).not.toContain("DoesNotExist");

    workbook.dispose();
  });

  it("throws when DocumentController skips formatting due to safety caps", () => {
    const warnSpy = vi.spyOn(console, "warn").mockImplementation(() => {});
    const controller = new DocumentController();
    const workbook = new DocumentControllerWorkbookAdapter(controller as any, { activeSheetName: "Sheet1" });
    const sheet = workbook.getSheet("Sheet1");
    try {
      // Full-width formatting over >50k rows is rejected by DocumentController's safety cap.
      expect(() => sheet.getRange("A1:XFD60000").setFormat({ bold: true })).toThrow(/setFormat skipped/i);
    } finally {
      warnSpy.mockRestore();
      workbook.dispose();
    }
  });
});
