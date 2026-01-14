import { describe, expect, it, vi } from "vitest";

import { MAX_INSERT_IMAGE_BYTES } from "../../../drawings/insertImageLimits.js";
import { mergeEmbeddedCellImagesIntoSnapshot } from "../embeddedCellImages.js";

describe("mergeEmbeddedCellImagesIntoSnapshot", () => {
  it("deduplicates snapshot images and preserves existing cell metadata", () => {
    const sheet = {
      id: "Sheet1",
      name: "Sheet1",
      cells: [{ row: 0, col: 0, value: 123, formula: "=1+1", format: { font: { bold: true } } }],
    };

    const result = mergeEmbeddedCellImagesIntoSnapshot({
      sheets: [sheet],
      images: [],
      embeddedCellImages: [
        {
          worksheet_part: "xl/worksheets/sheet1.xml",
          sheet_name: "Sheet1",
          row: 0,
          col: 0,
          image_id: "image2.png",
          bytes_base64: "abc",
          mime_type: "image/png",
          alt_text: "alt",
        },
        {
          worksheet_part: "xl/worksheets/sheet1.xml",
          sheet_name: "Sheet1",
          row: 0,
          col: 1,
          image_id: "image2.png",
          bytes_base64: "abc",
          mime_type: "image/png",
          alt_text: null,
        },
      ],
      resolveSheetIdByName: (name) => (name === "Sheet1" ? "Sheet1" : null),
      sheetIdsInOrder: ["Sheet1"],
      maxRows: 10_000,
      maxCols: 200,
    });

    expect(result.images).toEqual([{ id: "image2.png", bytesBase64: "abc=", mimeType: "image/png" }]);
    expect(sheet.cells).toEqual([
      {
        row: 0,
        col: 0,
        value: { type: "image", value: { imageId: "image2.png", altText: "alt" } },
        formula: "=1+1",
        format: { font: { bold: true } },
      },
      {
        row: 0,
        col: 1,
        value: { type: "image", value: { imageId: "image2.png" } },
        formula: null,
        format: null,
      },
    ]);
  });

  it("falls back to worksheet_part when sheet_name is missing", () => {
    const sheet1 = { id: "id-1", name: "Sheet1", cells: [] as any[] };
    const sheet2 = { id: "id-2", name: "Sheet2", cells: [] as any[] };

    mergeEmbeddedCellImagesIntoSnapshot({
      sheets: [sheet1, sheet2],
      embeddedCellImages: [
        {
          worksheet_part: "xl/worksheets/sheet2.xml",
          sheet_name: null,
          row: 5,
          col: 7,
          image_id: "image1.png",
          bytes_base64: "def",
          mime_type: "image/png",
          alt_text: null,
        },
      ],
      resolveSheetIdByName: () => null,
      sheetIdsInOrder: ["id-1", "id-2"],
      maxRows: 10_000,
      maxCols: 200,
    });

    expect(sheet2.cells).toEqual([
      { row: 5, col: 7, value: { type: "image", value: { imageId: "image1.png" } }, formula: null, format: null },
    ]);
  });

  it("skips images outside configured bounds", () => {
    const warn = vi.fn();
    const sheet = { id: "Sheet1", name: "Sheet1", cells: [] as any[] };

    mergeEmbeddedCellImagesIntoSnapshot({
      sheets: [sheet],
      embeddedCellImages: [
        {
          worksheet_part: "xl/worksheets/sheet1.xml",
          sheet_name: "Sheet1",
          row: 10,
          col: 0,
          image_id: "image1.png",
          bytes_base64: "ghi",
          mime_type: "image/png",
          alt_text: null,
        },
      ],
      resolveSheetIdByName: () => "Sheet1",
      sheetIdsInOrder: ["Sheet1"],
      maxRows: 10,
      maxCols: 200,
      warn,
    });

    expect(sheet.cells).toEqual([]);
    expect(warn).toHaveBeenCalled();
  });

  it("skips embedded cell images whose bytes_base64 exceeds MAX_INSERT_IMAGE_BYTES", () => {
    const warn = vi.fn();
    const sheet = { id: "Sheet1", name: "Sheet1", cells: [] as any[] };

    const desiredBytes = MAX_INSERT_IMAGE_BYTES + 1;
    const minLen = Math.ceil((desiredBytes * 4) / 3);
    const paddedLen = minLen + ((4 - (minLen % 4)) % 4);
    const oversizedBase64 = "A".repeat(paddedLen);

    const result = mergeEmbeddedCellImagesIntoSnapshot({
      sheets: [sheet],
      images: [],
      embeddedCellImages: [
        {
          worksheet_part: "xl/worksheets/sheet1.xml",
          sheet_name: "Sheet1",
          row: 0,
          col: 0,
          image_id: "image-too-big.png",
          bytes_base64: oversizedBase64,
          mime_type: "image/png",
          alt_text: null,
        },
      ],
      resolveSheetIdByName: (name) => (name === "Sheet1" ? "Sheet1" : null),
      sheetIdsInOrder: ["Sheet1"],
      maxRows: 10_000,
      maxCols: 200,
      warn,
    });

    expect(result.images).toEqual([]);
    expect(sheet.cells).toEqual([]);
    expect(warn).toHaveBeenCalled();
  });
});
