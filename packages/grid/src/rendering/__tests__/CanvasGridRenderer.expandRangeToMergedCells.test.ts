import { describe, expect, it } from "vitest";
import type { CellProvider, CellRange } from "../../model/CellProvider";
import { CanvasGridRenderer } from "../CanvasGridRenderer";

describe("CanvasGridRenderer.expandRangeToMergedCells", () => {
  it("does not scan O(merge height) when expanding a single-cell range inside a very tall merge (getMergedRangeAt)", () => {
    const tallMerge: CellRange = { startRow: 0, endRow: 1_000_000, startCol: 0, endCol: 2 };

    let mergedAtCalls = 0;
    const provider: CellProvider = {
      getCell: () => null,
      getMergedRangeAt: (row, col) => {
        mergedAtCalls += 1;
        if (
          row >= tallMerge.startRow &&
          row < tallMerge.endRow &&
          col >= tallMerge.startCol &&
          col < tallMerge.endCol
        ) {
          return tallMerge;
        }
        return null;
      }
    };

    const renderer = new CanvasGridRenderer({ provider, rowCount: tallMerge.endRow, colCount: 10 });

    const expanded = (renderer as any).expandRangeToMergedCells({
      startRow: 500_000,
      endRow: 500_001,
      startCol: 0,
      endCol: 1
    }) as CellRange;

    expect(expanded).toEqual(tallMerge);

    // Should be O(1) calls (perimeter scans skipping over the merge), not ~1,000,000.
    expect(mergedAtCalls).toBeLessThan(100);
  });
});

