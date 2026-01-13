import { describe, expect, it } from "vitest";

import { buildHitTestIndex, hitTestDrawings } from "../hitTest";
import { pxToEmu, type GridGeometry, type Viewport } from "../overlay";
import type { DrawingObject } from "../types";

const CELL = 10;

const geom: GridGeometry = {
  cellOriginPx: (cell) => ({ x: cell.col * CELL, y: cell.row * CELL }),
  cellSizePx: () => ({ width: CELL, height: CELL }),
};

function oneCellShape(id: number, row: number, col: number, widthPx: number, heightPx: number): DrawingObject {
  return {
    id,
    kind: { type: "shape" },
    anchor: {
      type: "oneCell",
      from: { cell: { row, col }, offset: { xEmu: 0, yEmu: 0 } },
      size: { cx: pxToEmu(widthPx), cy: pxToEmu(heightPx) },
    },
    zOrder: 0,
  };
}

describe("drawings hit testing with frozen panes", () => {
  it("hits objects in the frozen top-left pane even when scrolled", () => {
    const objects = [oneCellShape(1, 0, 0, 5, 5)];
    const index = buildHitTestIndex(objects, geom);
    const viewport: Viewport = {
      scrollX: 50,
      scrollY: 100,
      width: 200,
      height: 200,
      dpr: 1,
      frozenRows: 1,
      frozenCols: 1,
      frozenWidthPx: CELL,
      frozenHeightPx: CELL,
    };

    const hit = hitTestDrawings(index, viewport, 2, 2, geom);
    expect(hit?.object.id).toBe(1);
    expect(hit?.bounds.x).toBe(0);
    expect(hit?.bounds.y).toBe(0);
  });

  it("does not hit across quadrants when an object extends beyond a frozen boundary", () => {
    const objects = [oneCellShape(1, 0, 0, 30, 30)];
    const index = buildHitTestIndex(objects, geom);
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

    // Cursor is in the top-right pane; the shape is anchored in top-left and should be clipped away.
    expect(hitTestDrawings(index, viewport, CELL + 5, 5, geom)).toBeNull();
  });

  it("hits objects in the top-right pane using scrollX but not scrollY", () => {
    const objects = [oneCellShape(1, 0, 2, 5, 5)];
    const index = buildHitTestIndex(objects, geom);
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

    // Object sheet x = 20; screen x = 20 - scrollX, y pinned at 0.
    const hit = hitTestDrawings(index, viewport, 2 * CELL - viewport.scrollX + 1, 2, geom);
    expect(hit?.object.id).toBe(1);
  });
});

