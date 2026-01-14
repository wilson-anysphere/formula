import { describe, expect, it } from "vitest";

import { DocumentController } from "../../../document/documentController.js";
import { hydrateSheetBackgroundImagesFromBackend } from "../hydrateSheetBackgroundImages";

describe("hydrateSheetBackgroundImagesFromBackend", () => {
  it("does not resurrect deleted sheets when the workbook sheet store still references a stale id", async () => {
    const doc = new DocumentController();

    // Ensure Sheet1 exists so deleting Sheet2 doesn't trip the last-sheet guard.
    doc.getCell("Sheet1", { row: 0, col: 0 });
    doc.setCellValue("Sheet2", { row: 0, col: 0 }, "two");
    expect(doc.getSheetIds()).toEqual(["Sheet1", "Sheet2"]);

    doc.deleteSheet("Sheet2");
    expect(doc.getSheetIds()).toEqual(["Sheet1"]);

    const workbookSheetStore = {
      resolveIdByName: (name: string) => String(name),
      listAll: () => [{ id: "Sheet1" }, { id: "Sheet2" }],
    } as any;

    const app = {
      getDocument: () => doc,
      setSheetBackgroundImageId: () => {
        throw new Error("setSheetBackgroundImageId should not be called in this test");
      },
    } as any;

    await hydrateSheetBackgroundImagesFromBackend({
      app,
      workbookSheetStore,
      backend: { listImportedSheetBackgroundImages: async () => [] },
    });

    expect(doc.getSheetIds()).toEqual(["Sheet1"]);
  });
});

