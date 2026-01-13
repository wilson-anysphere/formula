import { describe, expect, it } from "vitest";

import { DrawingOverlay, pxToEmu, type GridGeometry, type Viewport } from "../overlay";
import { buildHitTestIndex, hitTestDrawings } from "../hitTest";
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
    fillRect: (...args: unknown[]) => calls.push({ method: "fillRect", args }),
    fillText: (...args: unknown[]) => calls.push({ method: "fillText", args }),
    fill: (...args: unknown[]) => calls.push({ method: "fill", args }),
    stroke: (...args: unknown[]) => calls.push({ method: "stroke", args }),
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

describe("DrawingOverlay viewport transforms", () => {
  it("pins frozen-pane anchored objects while scrollable objects subtract scroll", async () => {
    const { ctx, calls } = createStubCanvasContext();
    const canvas = createStubCanvas(ctx);
    const overlay = new DrawingOverlay(canvas, images, geom);

    const objects: DrawingObject[] = [
      {
        id: 1,
        kind: { type: "shape" },
        anchor: {
          type: "oneCell",
          from: { cell: { row: 0, col: 0 }, offset: { xEmu: 0, yEmu: 0 } },
          size: { cx: pxToEmu(10), cy: pxToEmu(10) },
        },
        zOrder: 0,
      },
      {
        id: 2,
        kind: { type: "shape" },
        anchor: {
          type: "oneCell",
          from: { cell: { row: 3, col: 2 }, offset: { xEmu: 0, yEmu: 0 } },
          size: { cx: pxToEmu(10), cy: pxToEmu(10) },
        },
        zOrder: 1,
      },
      {
        id: 3,
        kind: { type: "shape" },
        anchor: {
          type: "absolute",
          pos: { xEmu: pxToEmu(300), yEmu: pxToEmu(80) },
          size: { cx: pxToEmu(10), cy: pxToEmu(10) },
        },
        zOrder: 2,
      },
    ];

    const viewport: Viewport = {
      scrollX: 50,
      scrollY: 30,
      width: 500,
      height: 500,
      dpr: 1,
      frozenRows: 1,
      frozenCols: 1,
    };

    await overlay.render(objects, viewport);

    const strokeRects = calls.filter((call) => call.method === "strokeRect").map((call) => call.args);
    expect(strokeRects).toEqual([
      [0, 0, 10, 10], // frozen
      [150, 30, 10, 10], // scrolled (200-50, 60-30)
      [250, 50, 10, 10], // absolute always scrolls (300-50, 80-30)
    ]);

    // Hit testing should use the same frozen-aware transform.
    const index = buildHitTestIndex(objects, geom);
    const hit = hitTestDrawings(index, viewport, 5, 5, geom);
    expect(hit?.object.id).toBe(1);

    const absHit = hitTestDrawings(index, viewport, 251, 51, geom);
    expect(absHit?.object.id).toBe(3);
  });

  it("applies viewport zoom to EMU-derived pixel geometry", async () => {
    const { ctx, calls } = createStubCanvasContext();
    const canvas = createStubCanvas(ctx);
    const overlay = new DrawingOverlay(canvas, images, geom);

    const objects: DrawingObject[] = [
      {
        id: 1,
        kind: { type: "shape" },
        anchor: {
          type: "absolute",
          pos: { xEmu: pxToEmu(5), yEmu: pxToEmu(7) },
          size: { cx: pxToEmu(20), cy: pxToEmu(10) },
        },
        zOrder: 0,
      },
    ];

    const viewport: Viewport = { scrollX: 0, scrollY: 0, width: 500, height: 500, dpr: 1, zoom: 2 };
    await overlay.render(objects, viewport);

    const strokeRects = calls.filter((call) => call.method === "strokeRect").map((call) => call.args);
    expect(strokeRects).toEqual([[10, 14, 40, 20]]);

    const index = buildHitTestIndex(objects, geom, { zoom: viewport.zoom });
    const hit = hitTestDrawings(index, viewport, 11, 15, geom);
    expect(hit?.object.id).toBe(1);
    expect(hit?.bounds).toEqual({ x: 10, y: 14, width: 40, height: 20 });
  });

  it("renders selection handles using the same frozen/scroll transform as objects", async () => {
    const { ctx, calls } = createStubCanvasContext();
    const canvas = createStubCanvas(ctx);
    const overlay = new DrawingOverlay(canvas, images, geom);

    const objects: DrawingObject[] = [
      {
        id: 1,
        kind: { type: "shape" },
        anchor: {
          type: "oneCell",
          from: { cell: { row: 0, col: 0 }, offset: { xEmu: 0, yEmu: 0 } },
          size: { cx: pxToEmu(10), cy: pxToEmu(10) },
        },
        zOrder: 0,
      },
    ];

    overlay.setSelectedId(1);

    const viewport: Viewport = {
      scrollX: 50,
      scrollY: 30,
      width: 500,
      height: 500,
      dpr: 1,
      frozenRows: 1,
      frozenCols: 1,
    };

    await overlay.render(objects, viewport);

    const handleRects = calls
      .filter((call) => call.method === "rect" && call.args[2] === 8 && call.args[3] === 8)
      .map((call) => call.args);
    expect(handleRects[0]).toEqual([-4, -4, 8, 8]);
  });
});
