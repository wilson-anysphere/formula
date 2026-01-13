import { describe, expect, it } from "vitest";

import { buildHitTestIndex, hitTestDrawings } from "../hitTest";
import { pxToEmu, type GridGeometry, type Viewport } from "../overlay";
import type { DrawingObject } from "../types";

const CELL = 10;

const geom: GridGeometry = {
  cellOriginPx: (cell) => ({ x: cell.col * CELL, y: cell.row * CELL }),
  cellSizePx: () => ({ width: CELL, height: CELL }),
};

function oneCellObject(
  id: number,
  opts: { row: number; col: number; widthPx: number; heightPx: number; zOrder?: number },
): DrawingObject {
  return {
    id,
    kind: { type: "shape" },
    anchor: {
      type: "oneCell",
      from: { cell: { row: opts.row, col: opts.col }, offset: { xEmu: 0, yEmu: 0 } },
      size: { cx: pxToEmu(opts.widthPx), cy: pxToEmu(opts.heightPx) },
    },
    zOrder: opts.zOrder ?? 0,
  };
}

describe("hitTestDrawings frozen panes + header offsets", () => {
  it("hits objects in the frozen top-left pane without applying scroll", () => {
    const objects = [oneCellObject(1, { row: 0, col: 0, widthPx: 5, heightPx: 5 })];
    const index = buildHitTestIndex(objects, geom, { bucketSizePx: 64 });

    const viewport: Viewport = {
      scrollX: 50,
      scrollY: 100,
      width: 200,
      height: 200,
      dpr: 1,
      frozenRows: 1,
      frozenCols: 1,
      headerOffsetX: 30,
      headerOffsetY: 40,
      frozenWidthPx: 30 + CELL,
      frozenHeightPx: 40 + CELL,
    };

    // Click inside the object as rendered under the headers.
    const hit = hitTestDrawings(index, viewport, 30 + 2, 40 + 2, geom);
    expect(hit?.object.id).toBe(1);
    expect(hit?.bounds).toEqual({ x: 30, y: 40, width: 5, height: 5 });

    // Pointer events in the header gutter should not hit drawings.
    expect(hitTestDrawings(index, viewport, 10, 10, geom)).toBeNull();
  });

  it("does not hit frozen-pane objects from other quadrants (clipped portions are not hittable)", () => {
    const objects = [oneCellObject(1, { row: 0, col: 0, widthPx: 30, heightPx: 5 })];
    const index = buildHitTestIndex(objects, geom, { bucketSizePx: 64 });

    const viewport: Viewport = {
      scrollX: 0,
      scrollY: 0,
      width: 200,
      height: 200,
      dpr: 1,
      frozenRows: 1,
      frozenCols: 1,
      frozenWidthPx: CELL,
      frozenHeightPx: CELL,
    };

    // The object extends beyond the frozen column boundary (x > CELL), but it is clipped
    // to the top-left quadrant, so the top-right quadrant should not register a hit.
    expect(hitTestDrawings(index, viewport, CELL + 1, 1, geom)).toBeNull();
  });

  it("hits objects in the top-right quadrant (x scrolls, y is pinned)", () => {
    const objects = [oneCellObject(1, { row: 0, col: 2, widthPx: 5, heightPx: 5 })];
    const index = buildHitTestIndex(objects, geom, { bucketSizePx: 64 });

    const viewport: Viewport = {
      scrollX: 5,
      scrollY: 100,
      width: 200,
      height: 200,
      dpr: 1,
      frozenRows: 1,
      frozenCols: 1,
      frozenWidthPx: CELL,
      frozenHeightPx: CELL,
    };

    // Sheet rect x = 2 * CELL. In the top-right quadrant it scrolls by scrollX.
    const hit = hitTestDrawings(index, viewport, 2 * CELL - viewport.scrollX + 1, 1, geom);
    expect(hit?.object.id).toBe(1);
    expect(hit?.bounds.x).toBe(2 * CELL - viewport.scrollX);
    expect(hit?.bounds.y).toBe(0);
  });
});
