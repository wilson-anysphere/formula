import { describe, expect, it } from "vitest";

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
});
