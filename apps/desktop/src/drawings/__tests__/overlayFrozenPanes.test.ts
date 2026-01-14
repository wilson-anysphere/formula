import { describe, expect, it } from "vitest";

import { DrawingOverlay, pxToEmu, type GridGeometry, type Viewport } from "../overlay";
import type { DrawingObject, ImageStore } from "../types";

function createStubCanvasContext(): {
  ctx: CanvasRenderingContext2D;
  calls: Array<{ method: string; args: unknown[] }>;
} {
  const calls: Array<{ method: string; args: unknown[] }> = [];
  const ctx: any = {
    clearRect: (...args: unknown[]) => calls.push({ method: "clearRect", args }),
    drawImage: (...args: unknown[]) => calls.push({ method: "drawImage", args }),
    save: () => calls.push({ method: "save", args: [] }),
    restore: () => calls.push({ method: "restore", args: [] }),
    beginPath: () => calls.push({ method: "beginPath", args: [] }),
    rect: (...args: unknown[]) => calls.push({ method: "rect", args }),
    clip: () => calls.push({ method: "clip", args: [] }),
    setLineDash: (...args: unknown[]) => calls.push({ method: "setLineDash", args }),
    strokeRect: (...args: unknown[]) => calls.push({ method: "strokeRect", args }),
    fillText: (...args: unknown[]) => calls.push({ method: "fillText", args }),
  };
  return { ctx: ctx as CanvasRenderingContext2D, calls };
}

function createStubCanvas(ctx: CanvasRenderingContext2D): HTMLCanvasElement {
  const canvas: any = {
    width: 0,
    height: 0,
    style: {},
    getContext: (type: string) => (type === "2d" ? ctx : null),
  };
  return canvas as HTMLCanvasElement;
}

function createOneCellShapeObject(opts: { id: number; row: number; col: number; widthPx: number; heightPx: number }): DrawingObject {
  return {
    id: opts.id,
    kind: { type: "shape" },
    anchor: {
      type: "oneCell",
      from: { cell: { row: opts.row, col: opts.col }, offset: { xEmu: 0, yEmu: 0 } },
      size: { cx: pxToEmu(opts.widthPx), cy: pxToEmu(opts.heightPx) },
    },
    zOrder: 0,
  };
}

const images: ImageStore = {
  get: () => undefined,
  set: () => {},
  delete: () => {},
  clear: () => {},
};

const CELL = 10;
const geom: GridGeometry = {
  cellOriginPx: (cell) => ({ x: cell.col * CELL, y: cell.row * CELL }),
  cellSizePx: () => ({ width: CELL, height: CELL }),
};

describe("DrawingOverlay frozen panes", () => {
  it("keeps objects anchored in the frozen top-left pane pinned when scrolling", async () => {
    const { ctx, calls } = createStubCanvasContext();
    const canvas = createStubCanvas(ctx);

    const overlay = new DrawingOverlay(canvas, images, geom);
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

    await overlay.render([createOneCellShapeObject({ id: 1, row: 0, col: 0, widthPx: 5, heightPx: 5 })], viewport);

    const stroke = calls.find((c) => c.method === "strokeRect");
    expect(stroke?.args).toEqual([0, 0, 5, 5]);
  });

  it("keeps frozen objects pinned across renders when the scroll position changes", async () => {
    const { ctx, calls } = createStubCanvasContext();
    const canvas = createStubCanvas(ctx);

    const overlay = new DrawingOverlay(canvas, images, geom);
    const base: Omit<Viewport, "scrollX" | "scrollY"> = {
      width: 200,
      height: 200,
      dpr: 1,
      frozenRows: 1,
      frozenCols: 1,
      frozenWidthPx: CELL,
      frozenHeightPx: CELL,
    };

    const objects = [createOneCellShapeObject({ id: 1, row: 0, col: 0, widthPx: 5, heightPx: 5 })];

    await overlay.render(objects, { ...base, scrollX: 0, scrollY: 0 });
    await overlay.render(objects, { ...base, scrollX: 50, scrollY: 100 });

    const strokes = calls.filter((c) => c.method === "strokeRect");
    expect(strokes).toHaveLength(2);
    expect(strokes[0]?.args).toEqual([0, 0, 5, 5]);
    expect(strokes[1]?.args).toEqual([0, 0, 5, 5]);
  });

  it("scrolls objects in the top-right pane horizontally but not vertically", async () => {
    const { ctx, calls } = createStubCanvasContext();
    const canvas = createStubCanvas(ctx);

    const overlay = new DrawingOverlay(canvas, images, geom);
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

    await overlay.render([createOneCellShapeObject({ id: 1, row: 0, col: 2, widthPx: 5, heightPx: 5 })], viewport);

    const stroke = calls.find((c) => c.method === "strokeRect");
    // Base x = 2 * CELL, scrolls by scrollX; y stays pinned because it's in a frozen row.
    expect(stroke?.args).toEqual([2 * CELL - viewport.scrollX, 0, 5, 5]);
  });

  it("clips objects to the quadrant that contains their anchor cell", async () => {
    const { ctx, calls } = createStubCanvasContext();
    const canvas = createStubCanvas(ctx);

    const overlay = new DrawingOverlay(canvas, images, geom);
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

    // Large shape anchored in A1; should be clipped to the top-left frozen quadrant (CELL x CELL).
    await overlay.render([createOneCellShapeObject({ id: 1, row: 0, col: 0, widthPx: 30, heightPx: 30 })], viewport);

    const clipRectCall = calls.find((c) => c.method === "rect");
    expect(clipRectCall?.args).toEqual([0, 0, CELL, CELL]);

    const clipIndex = calls.findIndex((c) => c.method === "clip");
    const strokeIndex = calls.findIndex((c) => c.method === "strokeRect");
    expect(clipIndex).toBeGreaterThanOrEqual(0);
    expect(strokeIndex).toBeGreaterThan(clipIndex);
  });
});
