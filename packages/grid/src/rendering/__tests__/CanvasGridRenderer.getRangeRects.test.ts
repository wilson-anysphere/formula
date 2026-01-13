import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { CanvasGridRenderer } from "../CanvasGridRenderer";
import type { CellProvider } from "../../model/CellProvider";

describe("CanvasGridRenderer.getRangeRects", () => {
  beforeEach(() => {
    // Avoid scheduling real animation frames in the unit test environment.
    vi.stubGlobal("requestAnimationFrame", () => 0);
  });

  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it("returns a single rect for a range fully within the main scrollable quadrant", () => {
    const provider: CellProvider = { getCell: () => null };
    const renderer = new CanvasGridRenderer({
      provider,
      rowCount: 100,
      colCount: 100,
      defaultRowHeight: 10,
      defaultColWidth: 10
    });

    renderer.scroll.setViewportSize(100, 100);
    renderer.scroll.setFrozen(1, 1);
    renderer.scroll.setScroll(0, 0);

    const rects = renderer.getRangeRects({ startRow: 2, endRow: 4, startCol: 2, endCol: 5 });

    expect(rects).toEqual([{ x: 20, y: 20, width: 30, height: 20 }]);
  });

  it("splits a range spanning the frozen column boundary into frozen + scrollable rects and clips to the viewport", () => {
    const provider: CellProvider = { getCell: () => null };
    const renderer = new CanvasGridRenderer({
      provider,
      rowCount: 100,
      colCount: 100,
      defaultRowHeight: 10,
      defaultColWidth: 10
    });

    renderer.scroll.setViewportSize(80, 60);
    renderer.scroll.setFrozen(0, 2);
    renderer.scroll.setScroll(0, 0);

    const rects = renderer.getRangeRects({ startRow: 1, endRow: 3, startCol: 1, endCol: 4 });

    expect(rects).toEqual([
      { x: 10, y: 10, width: 10, height: 20 },
      { x: 20, y: 10, width: 20, height: 20 }
    ]);

    const viewport = renderer.getViewportState();
    for (const rect of rects) {
      expect(rect.x).toBeGreaterThanOrEqual(0);
      expect(rect.y).toBeGreaterThanOrEqual(0);
      expect(rect.x + rect.width).toBeLessThanOrEqual(viewport.width);
      expect(rect.y + rect.height).toBeLessThanOrEqual(viewport.height);
    }
  });

  it("clips ranges that are partially offscreen due to scroll (no negative coords or sizes)", () => {
    const provider: CellProvider = { getCell: () => null };
    const renderer = new CanvasGridRenderer({
      provider,
      rowCount: 100,
      colCount: 100,
      defaultRowHeight: 10,
      defaultColWidth: 10
    });

    renderer.scroll.setViewportSize(50, 50);
    renderer.scroll.setFrozen(0, 0);
    renderer.scroll.setScroll(20, 20);

    const rects = renderer.getRangeRects({ startRow: 1, endRow: 6, startCol: 1, endCol: 6 });

    expect(rects).toEqual([{ x: 0, y: 0, width: 40, height: 40 }]);
    for (const rect of rects) {
      expect(rect.x).toBeGreaterThanOrEqual(0);
      expect(rect.y).toBeGreaterThanOrEqual(0);
      expect(rect.width).toBeGreaterThan(0);
      expect(rect.height).toBeGreaterThan(0);
    }
  });
});

