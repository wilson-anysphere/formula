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
    save: () => calls.push({ method: "save", args: [] }),
    restore: () => calls.push({ method: "restore", args: [] }),
    beginPath: () => calls.push({ method: "beginPath", args: [] }),
    rect: (...args: unknown[]) => calls.push({ method: "rect", args }),
    clip: (...args: unknown[]) => calls.push({ method: "clip", args }),
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

const images: ImageStore = {
  get: () => undefined,
  set: () => {},
};

const geom: GridGeometry = {
  cellOriginPx: ({ row, col }) => ({ x: col * 100, y: row * 20 }),
  cellSizePx: () => ({ width: 100, height: 20 }),
};

function shapeObject(opts: { id: number; row: number; col: number; zOrder: number }): DrawingObject {
  return {
    id: opts.id,
    kind: { type: "shape" },
    anchor: {
      type: "oneCell",
      from: { cell: { row: opts.row, col: opts.col }, offset: { xEmu: 0, yEmu: 0 } },
      size: { cx: pxToEmu(10), cy: pxToEmu(10) },
    },
    zOrder: opts.zOrder,
  };
}

describe("DrawingOverlay frozen pane clipping", () => {
  it("clips each object to its frozen-pane quadrant (including header offsets)", async () => {
    const { ctx, calls } = createStubCanvasContext();
    const canvas = createStubCanvas(ctx);
    const overlay = new DrawingOverlay(canvas, images, geom);

    const objects: DrawingObject[] = [
      shapeObject({ id: 1, row: 0, col: 0, zOrder: 0 }), // top-left (frozen row + col)
      shapeObject({ id: 2, row: 0, col: 2, zOrder: 1 }), // top-right (frozen row only)
      shapeObject({ id: 3, row: 2, col: 0, zOrder: 2 }), // bottom-left (frozen col only)
      shapeObject({ id: 4, row: 2, col: 2, zOrder: 3 }), // bottom-right (scrollable)
    ];

    const headerOffsetX = 50;
    const headerOffsetY = 30;
    const frozenContentWidth = 100;
    const frozenContentHeight = 20;

    const viewport: Viewport = {
      scrollX: 10,
      scrollY: 5,
      width: 400,
      height: 200,
      dpr: 1,
      frozenRows: 1,
      frozenCols: 1,
      headerOffsetX,
      headerOffsetY,
      frozenWidthPx: headerOffsetX + frozenContentWidth,
      frozenHeightPx: headerOffsetY + frozenContentHeight,
    };

    await overlay.render(objects, viewport);

    const rectCalls = calls.filter((c) => c.method === "rect").map((c) => c.args as number[]);
    expect(rectCalls).toEqual([
      // TL
      [headerOffsetX, headerOffsetY, frozenContentWidth, frozenContentHeight],
      // TR
      [headerOffsetX + frozenContentWidth, headerOffsetY, viewport.width - headerOffsetX - frozenContentWidth, frozenContentHeight],
      // BL
      [headerOffsetX, headerOffsetY + frozenContentHeight, frozenContentWidth, viewport.height - headerOffsetY - frozenContentHeight],
      // BR
      [
        headerOffsetX + frozenContentWidth,
        headerOffsetY + frozenContentHeight,
        viewport.width - headerOffsetX - frozenContentWidth,
        viewport.height - headerOffsetY - frozenContentHeight,
      ],
    ]);

    const strokeCalls = calls.filter((c) => c.method === "strokeRect").map((c) => c.args as number[]);
    expect(strokeCalls).toEqual([
      [headerOffsetX, headerOffsetY, 10, 10], // TL (no scroll)
      [200 - viewport.scrollX + headerOffsetX, headerOffsetY, 10, 10], // TR (scrollX only)
      [headerOffsetX, 40 - viewport.scrollY + headerOffsetY, 10, 10], // BL (scrollY only)
      [200 - viewport.scrollX + headerOffsetX, 40 - viewport.scrollY + headerOffsetY, 10, 10], // BR (scrollX+scrollY)
    ]);
  });
});

