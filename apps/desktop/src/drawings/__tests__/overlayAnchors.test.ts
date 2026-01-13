import { describe, expect, it } from "vitest";

import { anchorToRectPx, pxToEmu, type GridGeometry } from "../overlay";
import type { Anchor } from "../types";

function createGeomWithA1Origin(origin: { x: number; y: number }): GridGeometry {
  return {
    cellOriginPx: (cell) => ({ x: origin.x + cell.col * 100, y: origin.y + cell.row * 20 }),
    cellSizePx: () => ({ width: 100, height: 20 }),
  };
}

describe("anchorToRectPx", () => {
  it("offsets absolute anchors by the A1 origin (row/col headers)", () => {
    const geom = createGeomWithA1Origin({ x: 40, y: 22 });

    const anchor: Anchor = {
      type: "absolute",
      pos: { xEmu: pxToEmu(5), yEmu: pxToEmu(7) },
      size: { cx: pxToEmu(20), cy: pxToEmu(10) },
    };

    expect(anchorToRectPx(anchor, geom)).toEqual({ x: 45, y: 29, width: 20, height: 10 });
  });

  it("aligns oneCell(A1, 0,0) with absolute(0,0)", () => {
    const geom = createGeomWithA1Origin({ x: 40, y: 22 });

    const abs: Anchor = {
      type: "absolute",
      pos: { xEmu: pxToEmu(0), yEmu: pxToEmu(0) },
      size: { cx: pxToEmu(20), cy: pxToEmu(10) },
    };

    const oneCell: Anchor = {
      type: "oneCell",
      from: { cell: { row: 0, col: 0 }, offset: { xEmu: pxToEmu(0), yEmu: pxToEmu(0) } },
      size: { cx: pxToEmu(20), cy: pxToEmu(10) },
    };

    expect(anchorToRectPx(oneCell, geom)).toEqual(anchorToRectPx(abs, geom));
  });

  it("scales absolute anchors by zoom", () => {
    const geom = createGeomWithA1Origin({ x: 40, y: 22 });

    const anchor: Anchor = {
      type: "absolute",
      pos: { xEmu: pxToEmu(5), yEmu: pxToEmu(7) },
      size: { cx: pxToEmu(20), cy: pxToEmu(10) },
    };

    expect(anchorToRectPx(anchor, geom, 2)).toEqual({ x: 50, y: 36, width: 40, height: 20 });
  });
});
