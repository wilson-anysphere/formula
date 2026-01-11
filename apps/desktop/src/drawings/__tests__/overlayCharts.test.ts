import { describe, expect, it } from "vitest";

import { DrawingOverlay, pxToEmu, type ChartRenderer, type GridGeometry, type Viewport } from "../overlay";
import type { DrawingObject, ImageStore } from "../types";

function createStubCanvasContext(): { ctx: CanvasRenderingContext2D; calls: Array<{ method: string; args: unknown[] }> } {
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

function createChartObject(chartId: string): DrawingObject {
  return {
    id: 1,
    kind: { type: "chart", chartId },
    anchor: {
      type: "absolute",
      pos: { xEmu: pxToEmu(5), yEmu: pxToEmu(7) },
      size: { cx: pxToEmu(20), cy: pxToEmu(10) },
    },
    zOrder: 0,
  };
}

const images: ImageStore = {
  get: () => undefined,
  set: () => {},
};

const geom: GridGeometry = {
  cellOriginPx: () => ({ x: 0, y: 0 }),
  cellSizePx: () => ({ width: 0, height: 0 }),
};

const viewport: Viewport = { scrollX: 0, scrollY: 0, width: 100, height: 100, dpr: 1 };

describe("DrawingOverlay charts", () => {
  it("delegates to chartRenderer when chartId is present", async () => {
    const { ctx, calls } = createStubCanvasContext();
    const canvas = createStubCanvas(ctx);

    let received: { chartId: string; rect: { x: number; y: number; width: number; height: number } } | null = null;
    const chartRenderer: ChartRenderer = {
      renderToCanvas: (renderCtx, chartId, rect) => {
        received = { chartId, rect };
        (renderCtx as any).fillStyle = "red";
        (renderCtx as any).fillRect(rect.x, rect.y, rect.width, rect.height);
      },
    };

    const overlay = new DrawingOverlay(canvas, images, geom, chartRenderer);
    await overlay.render([createChartObject("chart_1")], viewport);

    expect(received).toEqual({
      chartId: "chart_1",
      rect: { x: 5, y: 7, width: 20, height: 10 },
    });
    expect(calls.some((call) => call.method === "fillRect")).toBe(true);
    expect(calls.some((call) => call.method === "strokeRect")).toBe(false);
  });

  it("falls back to placeholder when chartRenderer throws", async () => {
    const { ctx, calls } = createStubCanvasContext();
    const canvas = createStubCanvas(ctx);

    const chartRenderer: ChartRenderer = {
      renderToCanvas: () => {
        throw new Error("boom");
      },
    };

    const overlay = new DrawingOverlay(canvas, images, geom, chartRenderer);
    await overlay.render([createChartObject("chart_1")], viewport);

    expect(calls.some((call) => call.method === "strokeRect")).toBe(true);
  });
});

