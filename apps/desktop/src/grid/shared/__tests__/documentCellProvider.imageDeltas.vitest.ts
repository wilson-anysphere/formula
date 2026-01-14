/**
 * @vitest-environment jsdom
 */

import { describe, expect, it } from "vitest";

import { DocumentController } from "../../../document/documentController.js";
import { DocumentCellProvider } from "../documentCellProvider.js";

describe("DocumentCellProvider imageDeltas", () => {
  it("redraws on imageDeltas without evicting cell caches", () => {
    const doc = new DocumentController();
    doc.setCellValue("Sheet1", "A1", "X");

    const provider = new DocumentCellProvider({
      document: doc,
      getSheetId: () => "Sheet1",
      headerRows: 1,
      headerCols: 1,
      rowCount: 10,
      colCount: 10,
      showFormulas: () => false,
      getComputedValue: () => null,
    });

    // Populate the per-sheet cell cache.
    provider.getCell(1, 1);
    const before = provider.getCacheStats();
    expect(before.sheetCache.size).toBeGreaterThan(0);

    const updates: any[] = [];
    provider.subscribe((update) => {
      updates.push(update);
    });

    // Emit an image-only change payload (no cell/format/view deltas).
    (doc as any).applyExternalImageCacheDeltas(
      [
        {
          imageId: "img-1",
          before: null,
          after: { bytes: new Uint8Array([1, 2, 3]), mimeType: "image/png" },
        },
      ],
      { source: "collab" },
    );

    const after = provider.getCacheStats();
    expect(after.sheetCache.size).toBeGreaterThan(0);
    expect(updates.some((u) => u?.type === "invalidateAll")).toBe(true);
  });
});

