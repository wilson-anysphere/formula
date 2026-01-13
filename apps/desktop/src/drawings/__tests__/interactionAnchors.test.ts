import { describe, expect, it } from "vitest";

import { pxToEmu, type GridGeometry } from "../overlay";
import { resizeAnchor, shiftAnchor } from "../interaction";
import type { Anchor } from "../types";

function createGeom(colWidths: number[], rowHeights: number[]): GridGeometry {
  const colWidth = (col: number) => colWidths[col] ?? colWidths[colWidths.length - 1] ?? 50;
  const rowHeight = (row: number) => rowHeights[row] ?? rowHeights[rowHeights.length - 1] ?? 20;

  return {
    cellOriginPx: ({ row, col }) => {
      let x = 0;
      for (let c = 0; c < col; c++) x += colWidth(c);
      let y = 0;
      for (let r = 0; r < row; r++) y += rowHeight(r);
      return { x, y };
    },
    cellSizePx: ({ row, col }) => ({ width: colWidth(col), height: rowHeight(row) }),
  };
}

describe("drawing anchor normalization", () => {
  const geom = createGeom([40, 60, 50], [20, 30, 25]);

  it("normalizes oneCell movement across non-uniform column boundaries (both directions)", () => {
    const anchor: Anchor = {
      type: "oneCell",
      from: { cell: { row: 0, col: 1 }, offset: { xEmu: pxToEmu(10), yEmu: pxToEmu(5) } },
      size: { cx: pxToEmu(20), cy: pxToEmu(10) },
    };

    const movedLeft = shiftAnchor(anchor, -25, 0, geom);
    expect(movedLeft.type).toBe("oneCell");
    if (movedLeft.type !== "oneCell") throw new Error("unexpected anchor type");
    expect(movedLeft.from.cell.col).toBe(0);
    expect(movedLeft.from.offset.xEmu).toBeCloseTo(pxToEmu(25));

    const movedRight = shiftAnchor(anchor, 70, 0, geom);
    expect(movedRight.type).toBe("oneCell");
    if (movedRight.type !== "oneCell") throw new Error("unexpected anchor type");
    expect(movedRight.from.cell.col).toBe(2);
    expect(movedRight.from.offset.xEmu).toBeCloseTo(pxToEmu(20));
  });

  it("normalizes oneCell movement across non-uniform row boundaries (both directions)", () => {
    const anchor: Anchor = {
      type: "oneCell",
      from: { cell: { row: 1, col: 0 }, offset: { xEmu: pxToEmu(5), yEmu: pxToEmu(10) } },
      size: { cx: pxToEmu(20), cy: pxToEmu(10) },
    };

    const movedUp = shiftAnchor(anchor, 0, -25, geom);
    expect(movedUp.type).toBe("oneCell");
    if (movedUp.type !== "oneCell") throw new Error("unexpected anchor type");
    expect(movedUp.from.cell.row).toBe(0);
    // row 0 height is 20, so 10 - 25 = -15 => 20 - 15 = 5
    expect(movedUp.from.offset.yEmu).toBeCloseTo(pxToEmu(5));

    const movedDown = shiftAnchor(anchor, 0, 35, geom);
    expect(movedDown.type).toBe("oneCell");
    if (movedDown.type !== "oneCell") throw new Error("unexpected anchor type");
    // row 1 height is 30, so 10 + 35 = 45 => 45 - 30 = 15 in row 2
    expect(movedDown.from.cell.row).toBe(2);
    expect(movedDown.from.offset.yEmu).toBeCloseTo(pxToEmu(15));
  });

  it("advances twoCell endpoints when resizing crosses cell boundaries", () => {
    const anchor: Anchor = {
      type: "twoCell",
      from: { cell: { row: 0, col: 0 }, offset: { xEmu: pxToEmu(10), yEmu: pxToEmu(5) } },
      to: { cell: { row: 0, col: 0 }, offset: { xEmu: pxToEmu(30), yEmu: pxToEmu(15) } },
    };

    const resized = resizeAnchor(anchor, "se", 30, 0, geom);
    expect(resized.type).toBe("twoCell");
    if (resized.type !== "twoCell") throw new Error("unexpected anchor type");
    // to.x: 30 + 30 = 60; col0 width is 40 => col1 offset 20
    expect(resized.to.cell.col).toBe(1);
    expect(resized.to.offset.xEmu).toBeCloseTo(pxToEmu(20));

    const resizedFrom = resizeAnchor(
      {
        type: "twoCell",
        from: { cell: { row: 0, col: 1 }, offset: { xEmu: pxToEmu(5), yEmu: pxToEmu(5) } },
        to: { cell: { row: 0, col: 1 }, offset: { xEmu: pxToEmu(25), yEmu: pxToEmu(15) } },
      },
      "nw",
      -10,
      0,
      geom,
    );
    expect(resizedFrom.type).toBe("twoCell");
    if (resizedFrom.type !== "twoCell") throw new Error("unexpected anchor type");
    // from.x: 5 - 10 = -5; col0 width is 40 => col0 offset 35
    expect(resizedFrom.from.cell.col).toBe(0);
    expect(resizedFrom.from.offset.xEmu).toBeCloseTo(pxToEmu(35));
  });
});

