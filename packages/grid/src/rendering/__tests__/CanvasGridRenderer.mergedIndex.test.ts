import { describe, expect, it } from "vitest";
import type { CellProvider, CellRange } from "../../model/CellProvider";
import { CanvasGridRenderer } from "../CanvasGridRenderer";

describe("CanvasGridRenderer merged index", () => {
  it("indexes only visible rows for extremely tall merged ranges", () => {
    const tallMerge: CellRange = { startRow: 0, endRow: 1_000_000, startCol: 0, endCol: 3 };

    const provider: CellProvider = {
      getCell: (row, col) => ({ row, col, value: null }),
      getMergedRangesInRange: (range) => {
        const intersects =
          range.startRow < tallMerge.endRow &&
          range.endRow > tallMerge.startRow &&
          range.startCol < tallMerge.endCol &&
          range.endCol > tallMerge.startCol;
        return intersects ? [tallMerge] : [];
      }
    };

    const renderer = new CanvasGridRenderer({
      provider,
      rowCount: tallMerge.endRow,
      colCount: 10
    });

    // Avoid attaching canvases: we only want to exercise merged index construction.
    renderer.scroll.setViewportSize(320, 200);

    // Scroll deep inside the merge so the anchor row is offscreen.
    renderer.scroll.setScroll(0, renderer.scroll.rows.positionOf(500_000));

    const viewport = renderer.scroll.getViewportState();
    const mergedIndex = (renderer as any).getMergedIndex(viewport) as any;

    // Ensure we keep the full merge range (not clipped to the viewport).
    expect(mergedIndex.getRanges()).toEqual([tallMerge]);

    // But only materialize row spans for the visible viewport rows.
    const visibleRowCount = viewport.main.rows.end - viewport.main.rows.start;
    expect(mergedIndex.getIndexedRowCount()).toBe(visibleRowCount + 1);
    expect(mergedIndex.getIndexedRowCount()).toBeLessThan(1_000);

    // Interior cells that are visible must still be treated as merged.
    const sampleRow = viewport.main.rows.start;
    expect(mergedIndex.rangeAt({ row: sampleRow, col: 1 })).toEqual(tallMerge);
  });
});
