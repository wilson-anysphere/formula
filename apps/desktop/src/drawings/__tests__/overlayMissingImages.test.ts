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

function createImageObject(imageId: string): DrawingObject {
  return {
    id: 1,
    kind: { type: "image", imageId },
    anchor: {
      type: "absolute",
      pos: { xEmu: pxToEmu(5), yEmu: pxToEmu(7) },
      size: { cx: pxToEmu(20), cy: pxToEmu(10) },
    },
    zOrder: 0,
  };
}

const geom: GridGeometry = {
  cellOriginPx: () => ({ x: 0, y: 0 }),
  cellSizePx: () => ({ width: 0, height: 0 }),
};

const viewport: Viewport = { scrollX: 0, scrollY: 0, width: 100, height: 100, dpr: 1 };

describe("DrawingOverlay missing images", () => {
  it("renders a placeholder and selection handles when image bytes are missing", async () => {
    const { ctx, calls } = createStubCanvasContext();
    const canvas = createStubCanvas(ctx);

    const images: ImageStore = {
      get: () => undefined,
      set: () => {},
    };

    const overlay = new DrawingOverlay(canvas, images, geom);
    overlay.setSelectedId(1);

    await overlay.render([createImageObject("img_1")], viewport);

    expect(calls.some((call) => call.method === "drawImage")).toBe(false);

    const fillTextCall = calls.find((call) => call.method === "fillText");
    expect(fillTextCall).toBeTruthy();
    expect(String(fillTextCall!.args[0])).toMatch(/image/i);
    expect(fillTextCall!.args.slice(1)).toEqual([9, 21]);

    expect(
      calls.some(
        (call) =>
          call.method === "setLineDash" &&
          Array.isArray(call.args[0]) &&
          (call.args[0] as number[]).length === 2 &&
          (call.args[0] as number[])[0] === 4 &&
          (call.args[0] as number[])[1] === 2,
      ),
    ).toBe(true);

    // Placeholder border + selection border.
    expect(calls.filter((call) => call.method === "strokeRect")).toHaveLength(2);

    // Selection handles.
    const handleRects = calls.filter(
      (call) =>
        call.method === "rect" &&
        call.args.length >= 4 &&
        call.args[2] === 8 &&
        call.args[3] === 8,
    );
    expect(handleRects).toHaveLength(8);
    // 8 resize handles + 1 rotation handle.
    expect(calls.filter((call) => call.method === "fill")).toHaveLength(9);
    expect(calls.filter((call) => call.method === "stroke")).toHaveLength(9);
  });
});
