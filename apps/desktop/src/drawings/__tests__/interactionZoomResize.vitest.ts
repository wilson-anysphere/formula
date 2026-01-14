import { describe, expect, it } from "vitest";

import { resizeAnchor, shiftAnchor } from "../interaction";
import { pxToEmu, type GridGeometry } from "../overlay";
import type { Anchor } from "../types";

function createBaseGeom(opts: { origin?: { x: number; y: number }; colWidth: number; rowHeight: number }): GridGeometry {
  const origin = opts.origin ?? { x: 0, y: 0 };
  return {
    cellOriginPx: ({ row, col }) => ({ x: origin.x + col * opts.colWidth, y: origin.y + row * opts.rowHeight }),
    cellSizePx: () => ({ width: opts.colWidth, height: opts.rowHeight }),
  };
}

function createZoomedGeom(base: GridGeometry, zoom: number): GridGeometry {
  return {
    cellOriginPx: (cell) => {
      const p = base.cellOriginPx(cell);
      return { x: p.x * zoom, y: p.y * zoom };
    },
    cellSizePx: (cell) => {
      const s = base.cellSizePx(cell);
      return { width: s.width * zoom, height: s.height * zoom };
    },
  };
}

describe("resizeAnchor / shiftAnchor zoom handling", () => {
  it("resizes absolute anchors correctly under zoom (and ignores A1 origin offsets)", () => {
    const zoom = 2;
    const geomBase = createBaseGeom({ origin: { x: 40, y: 22 }, colWidth: 100, rowHeight: 20 });
    const geom = createZoomedGeom(geomBase, zoom);

    const anchor: Anchor = {
      type: "absolute",
      pos: { xEmu: pxToEmu(10), yEmu: pxToEmu(20) },
      size: { cx: pxToEmu(100), cy: pxToEmu(50) },
    };

    const resizedE = resizeAnchor(anchor, "e", 20, 0, geom, undefined, zoom);
    expect(resizedE.type).toBe("absolute");
    if (resizedE.type !== "absolute") throw new Error("unexpected anchor type");
    // dx=20 screen px at zoom=2 => +10 base px.
    expect(resizedE.pos.xEmu).toBeCloseTo(anchor.pos.xEmu);
    expect(resizedE.size.cx).toBeCloseTo(pxToEmu(110));

    const resizedW = resizeAnchor(anchor, "w", 20, 0, geom, undefined, zoom);
    expect(resizedW.type).toBe("absolute");
    if (resizedW.type !== "absolute") throw new Error("unexpected anchor type");
    expect(resizedW.pos.xEmu).toBeCloseTo(pxToEmu(20));
    expect(resizedW.size.cx).toBeCloseTo(pxToEmu(90));
  });

  it("resizes oneCell anchors correctly under zoom", () => {
    const zoom = 2;
    const geomBase = createBaseGeom({ colWidth: 50, rowHeight: 30 });
    const geom = createZoomedGeom(geomBase, zoom);

    const anchor: Anchor = {
      type: "oneCell",
      from: { cell: { row: 0, col: 0 }, offset: { xEmu: pxToEmu(5), yEmu: pxToEmu(7) } },
      size: { cx: pxToEmu(40), cy: pxToEmu(30) },
    };

    const resizedE = resizeAnchor(anchor, "e", 20, 0, geom, undefined, zoom);
    expect(resizedE.type).toBe("oneCell");
    if (resizedE.type !== "oneCell") throw new Error("unexpected anchor type");
    // dx=20 screen px at zoom=2 => +10 base px.
    expect(resizedE.from).toEqual(anchor.from);
    expect(resizedE.size.cx).toBeCloseTo(pxToEmu(50));

    const resizedW = resizeAnchor(anchor, "w", 20, 0, geom, undefined, zoom);
    expect(resizedW.type).toBe("oneCell");
    if (resizedW.type !== "oneCell") throw new Error("unexpected anchor type");
    expect(resizedW.from.cell).toEqual(anchor.from.cell);
    expect(resizedW.from.offset.xEmu).toBeCloseTo(pxToEmu(15));
    expect(resizedW.size.cx).toBeCloseTo(pxToEmu(30));
  });

  it("moves oneCell anchors across cell boundaries under zoom", () => {
    const zoom = 2;
    const geomBase = createBaseGeom({ colWidth: 50, rowHeight: 30 });
    const geom = createZoomedGeom(geomBase, zoom);

    const anchor: Anchor = {
      type: "oneCell",
      from: { cell: { row: 0, col: 0 }, offset: { xEmu: pxToEmu(0), yEmu: pxToEmu(0) } },
      size: { cx: pxToEmu(10), cy: pxToEmu(10) },
    };

    const moved = shiftAnchor(anchor, 120, 0, geom, zoom);
    expect(moved.type).toBe("oneCell");
    if (moved.type !== "oneCell") throw new Error("unexpected anchor type");
    // dx=120 screen px at zoom=2 => 60 base px: 1 full cell (50) + 10 remainder.
    expect(moved.from.cell.col).toBe(1);
    expect(moved.from.offset.xEmu).toBeCloseTo(pxToEmu(10));
  });
});

