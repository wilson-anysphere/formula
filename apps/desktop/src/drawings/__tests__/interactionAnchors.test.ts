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

  it("resizes absolute anchors from edge handles", () => {
    const anchor: Anchor = {
      type: "absolute",
      pos: { xEmu: pxToEmu(10), yEmu: pxToEmu(20) },
      size: { cx: pxToEmu(100), cy: pxToEmu(50) },
    };

    const resizedE = resizeAnchor(anchor, "e", 10, 0, geom);
    expect(resizedE.type).toBe("absolute");
    if (resizedE.type !== "absolute") throw new Error("unexpected anchor type");
    expect(resizedE.pos.xEmu).toBeCloseTo(anchor.pos.xEmu);
    expect(resizedE.size.cx).toBeCloseTo(pxToEmu(110));

    const resizedW = resizeAnchor(anchor, "w", 10, 0, geom);
    expect(resizedW.type).toBe("absolute");
    if (resizedW.type !== "absolute") throw new Error("unexpected anchor type");
    expect(resizedW.pos.xEmu).toBeCloseTo(pxToEmu(20));
    expect(resizedW.size.cx).toBeCloseTo(pxToEmu(90));

    const resizedS = resizeAnchor(anchor, "s", 0, 5, geom);
    expect(resizedS.type).toBe("absolute");
    if (resizedS.type !== "absolute") throw new Error("unexpected anchor type");
    expect(resizedS.pos.yEmu).toBeCloseTo(anchor.pos.yEmu);
    expect(resizedS.size.cy).toBeCloseTo(pxToEmu(55));

    const resizedN = resizeAnchor(anchor, "n", 0, 5, geom);
    expect(resizedN.type).toBe("absolute");
    if (resizedN.type !== "absolute") throw new Error("unexpected anchor type");
    expect(resizedN.pos.yEmu).toBeCloseTo(pxToEmu(25));
    expect(resizedN.size.cy).toBeCloseTo(pxToEmu(45));
  });

  it("resizes oneCell anchors from edge handles without crossing cells", () => {
    const anchor: Anchor = {
      type: "oneCell",
      from: { cell: { row: 0, col: 0 }, offset: { xEmu: pxToEmu(5), yEmu: pxToEmu(7) } },
      size: { cx: pxToEmu(40), cy: pxToEmu(30) },
    };

    const resizedE = resizeAnchor(anchor, "e", 10, 0, geom);
    expect(resizedE.type).toBe("oneCell");
    if (resizedE.type !== "oneCell") throw new Error("unexpected anchor type");
    expect(resizedE.from.cell).toEqual(anchor.from.cell);
    expect(resizedE.from.offset.xEmu).toBeCloseTo(anchor.from.offset.xEmu);
    expect(resizedE.size.cx).toBeCloseTo(pxToEmu(50));

    const resizedW = resizeAnchor(anchor, "w", 10, 0, geom);
    expect(resizedW.type).toBe("oneCell");
    if (resizedW.type !== "oneCell") throw new Error("unexpected anchor type");
    expect(resizedW.from.cell).toEqual(anchor.from.cell);
    expect(resizedW.from.offset.xEmu).toBeCloseTo(pxToEmu(15));
    expect(resizedW.size.cx).toBeCloseTo(pxToEmu(30));

    const resizedS = resizeAnchor(anchor, "s", 0, 5, geom);
    expect(resizedS.type).toBe("oneCell");
    if (resizedS.type !== "oneCell") throw new Error("unexpected anchor type");
    expect(resizedS.from.cell).toEqual(anchor.from.cell);
    expect(resizedS.from.offset.yEmu).toBeCloseTo(anchor.from.offset.yEmu);
    expect(resizedS.size.cy).toBeCloseTo(pxToEmu(35));

    const resizedN = resizeAnchor(anchor, "n", 0, 5, geom);
    expect(resizedN.type).toBe("oneCell");
    if (resizedN.type !== "oneCell") throw new Error("unexpected anchor type");
    expect(resizedN.from.cell).toEqual(anchor.from.cell);
    expect(resizedN.from.offset.yEmu).toBeCloseTo(pxToEmu(12));
    expect(resizedN.size.cy).toBeCloseTo(pxToEmu(25));
  });

  it("resizes twoCell anchors from edge handles", () => {
    const anchor: Anchor = {
      type: "twoCell",
      from: { cell: { row: 0, col: 0 }, offset: { xEmu: pxToEmu(0), yEmu: pxToEmu(0) } },
      to: { cell: { row: 0, col: 0 }, offset: { xEmu: pxToEmu(20), yEmu: pxToEmu(10) } },
    };

    const resizedE = resizeAnchor(anchor, "e", 10, 0, geom);
    expect(resizedE.type).toBe("twoCell");
    if (resizedE.type !== "twoCell") throw new Error("unexpected anchor type");
    expect(resizedE.from).toEqual(anchor.from);
    expect(resizedE.to.offset.xEmu).toBeCloseTo(pxToEmu(30));

    const resizedW = resizeAnchor(anchor, "w", 10, 0, geom);
    expect(resizedW.type).toBe("twoCell");
    if (resizedW.type !== "twoCell") throw new Error("unexpected anchor type");
    expect(resizedW.to).toEqual(anchor.to);
    expect(resizedW.from.offset.xEmu).toBeCloseTo(pxToEmu(10));

    const resizedS = resizeAnchor(anchor, "s", 0, 5, geom);
    expect(resizedS.type).toBe("twoCell");
    if (resizedS.type !== "twoCell") throw new Error("unexpected anchor type");
    expect(resizedS.from).toEqual(anchor.from);
    expect(resizedS.to.offset.yEmu).toBeCloseTo(pxToEmu(15));

    const resizedN = resizeAnchor(anchor, "n", 0, 5, geom);
    expect(resizedN.type).toBe("twoCell");
    if (resizedN.type !== "twoCell") throw new Error("unexpected anchor type");
    expect(resizedN.to).toEqual(anchor.to);
    expect(resizedN.from.offset.yEmu).toBeCloseTo(pxToEmu(5));
  });
});
