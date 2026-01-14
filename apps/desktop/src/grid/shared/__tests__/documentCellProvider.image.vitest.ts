import { describe, expect, it, vi } from "vitest";

import { DocumentController } from "../../../document/documentController.js";
import { DocumentCellProvider } from "../documentCellProvider.js";

describe("DocumentCellProvider image values", () => {
  it("populates cell.image from a formula-model {type:\"image\"} envelope (camelCase) and uses altText as value", () => {
    const doc = new DocumentController();
    doc.setCellValue("Sheet1", "A1", {
      type: "image",
      value: { imageId: " image1.png ", altText: " Logo ", width: 128, height: 64 }
    });

    const provider = new DocumentCellProvider({
      document: doc,
      getSheetId: () => "Sheet1",
      headerRows: 1,
      headerCols: 1,
      rowCount: 10,
      colCount: 10,
      showFormulas: () => false,
      getComputedValue: () => null
    });

    const cell = provider.getCell(1, 1);
    expect(cell?.image).toEqual({ imageId: "image1.png", altText: "Logo", width: 128, height: 64 });
    expect(cell?.value).toBe("Logo");
  });

  it("accepts snake_case aliases for image payload fields", () => {
    const doc = new DocumentController();
    doc.setCellValue("Sheet1", "A1", {
      type: "image",
      value: { image_id: "image2.png", alt_text: "Alt", width: 10, height: 20 }
    });

    const provider = new DocumentCellProvider({
      document: doc,
      getSheetId: () => "Sheet1",
      headerRows: 1,
      headerCols: 1,
      rowCount: 10,
      colCount: 10,
      showFormulas: () => false,
      getComputedValue: () => null
    });

    const cell = provider.getCell(1, 1);
    expect(cell?.image).toEqual({ imageId: "image2.png", altText: "Alt", width: 10, height: 20 });
    expect(cell?.value).toBe("Alt");
  });

  it("falls back to [Image] when no altText is present", () => {
    const doc = new DocumentController();
    doc.setCellValue("Sheet1", "A1", { type: "image", value: { imageId: "image3.png" } });

    const provider = new DocumentCellProvider({
      document: doc,
      getSheetId: () => "Sheet1",
      headerRows: 1,
      headerCols: 1,
      rowCount: 10,
      colCount: 10,
      showFormulas: () => false,
      getComputedValue: () => null
    });

    const cell = provider.getCell(1, 1);
    expect(cell?.image?.imageId).toBe("image3.png");
    expect(cell?.value).toBe("[Image]");
  });

  it("accepts legacy direct payload shapes (no {type,value} envelope)", () => {
    const doc = new DocumentController();
    doc.setCellValue("Sheet1", "A1", { imageId: "image4.png", alt: "Legacy" });

    const provider = new DocumentCellProvider({
      document: doc,
      getSheetId: () => "Sheet1",
      headerRows: 1,
      headerCols: 1,
      rowCount: 10,
      colCount: 10,
      showFormulas: () => false,
      getComputedValue: () => null
    });

    const cell = provider.getCell(1, 1);
    expect(cell?.image).toEqual({ imageId: "image4.png", altText: "Legacy", width: undefined, height: undefined });
    expect(cell?.value).toBe("Legacy");
  });

  it("can surface images from computed values for formula cells", () => {
    const doc = new DocumentController();
    doc.setCellFormula("Sheet1", "A1", "=IMAGE(\"image5.png\")");

    const getComputedValue = vi.fn((coord: { row: number; col: number }) => {
      if (coord.row === 0 && coord.col === 0) {
        return { type: "image", value: { imageId: "image5.png", altText: "Computed" } };
      }
      return null;
    });

    const provider = new DocumentCellProvider({
      document: doc,
      getSheetId: () => "Sheet1",
      headerRows: 1,
      headerCols: 1,
      rowCount: 10,
      colCount: 10,
      showFormulas: () => false,
      getComputedValue
    });

    const cell = provider.getCell(1, 1);
    expect(getComputedValue).toHaveBeenCalled();
    expect(cell?.image?.imageId).toBe("image5.png");
    expect(cell?.value).toBe("Computed");
  });

  it("prefers cached in-cell image values for formula cells when present", () => {
    const doc = new DocumentController();
    // `setCellFormula` forces value=null, so use `setCell` to simulate an imported snapshot that
    // includes a cached rich-value image payload alongside the original formula text.
    (doc as any).model.setCell("Sheet1", 0, 0, {
      value: { type: "image", value: { imageId: "image6.png", altText: "Cached" } },
      formula: "=IMAGE(\"https://example.com\")",
      styleId: 0,
    } as any);

    const getComputedValue = vi.fn(() => ({ type: "image", value: { imageId: "computed.png", altText: "Computed" } }));

    const provider = new DocumentCellProvider({
      document: doc,
      getSheetId: () => "Sheet1",
      headerRows: 1,
      headerCols: 1,
      rowCount: 10,
      colCount: 10,
      showFormulas: () => false,
      getComputedValue,
    });

    const cell = provider.getCell(1, 1);
    expect(getComputedValue).not.toHaveBeenCalled();
    expect(cell?.image?.imageId).toBe("image6.png");
    expect(cell?.value).toBe("Cached");
  });
});
