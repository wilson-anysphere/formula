import { describe, expect, it } from "vitest";

import { DrawingOverlay, pxToEmu, type GridGeometry, type Viewport } from "../overlay";
import { RESIZE_HANDLE_SIZE_PX } from "../selectionHandles";
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
  delete: () => {},
  clear: () => {},
};

const geom: GridGeometry = {
  cellOriginPx: () => ({ x: 0, y: 0 }),
  cellSizePx: () => ({ width: 0, height: 0 }),
};

describe("DrawingOverlay selection handles", () => {
  it("renders 8 handle rects when an object is selected", async () => {
    const { ctx, calls } = createStubCanvasContext();
    const canvas = createStubCanvas(ctx);

    const overlay = new DrawingOverlay(canvas, images, geom);
    overlay.setSelectedId(1);

    const objects: DrawingObject[] = [
      {
        id: 1,
        kind: { type: "unknown" },
        anchor: {
          type: "absolute",
          pos: { xEmu: pxToEmu(50), yEmu: pxToEmu(60) },
          size: { cx: pxToEmu(100), cy: pxToEmu(80) },
        },
        zOrder: 0,
      },
    ];

    const viewport: Viewport = { scrollX: 0, scrollY: 0, width: 300, height: 300, dpr: 1 };
    await overlay.render(objects, viewport);

    const handleRects = calls.filter(
      (call) =>
        call.method === "rect" &&
        call.args.length === 4 &&
        call.args[2] === RESIZE_HANDLE_SIZE_PX &&
        call.args[3] === RESIZE_HANDLE_SIZE_PX,
    );

    expect(handleRects).toHaveLength(8);
  });
});
