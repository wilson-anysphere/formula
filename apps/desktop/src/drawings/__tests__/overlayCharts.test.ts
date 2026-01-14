import { describe, expect, it } from "vitest";

import { DrawingOverlay, pxToEmu, type ChartRenderer, type GridGeometry, type Viewport } from "../overlay";
import type { DrawingObject, ImageStore } from "../types";

type CanvasCall = { method: string; args: unknown[]; strokeStyle?: unknown; fillStyle?: unknown };

function createStubCanvasContext(): { ctx: CanvasRenderingContext2D; calls: CanvasCall[] } {
  const calls: CanvasCall[] = [];
  const ctx: any = {
    clearRect: (...args: unknown[]) => calls.push({ method: "clearRect", args }),
    drawImage: (...args: unknown[]) => calls.push({ method: "drawImage", args }),
    save: () => calls.push({ method: "save", args: [] }),
    restore: () => calls.push({ method: "restore", args: [] }),
    beginPath: () => calls.push({ method: "beginPath", args: [] }),
    rect: (...args: unknown[]) => calls.push({ method: "rect", args }),
    clip: () => calls.push({ method: "clip", args: [] }),
    setLineDash: (...args: unknown[]) => calls.push({ method: "setLineDash", args }),
    fill: () => calls.push({ method: "fill", args: [], fillStyle: ctx.fillStyle }),
    stroke: () => calls.push({ method: "stroke", args: [], strokeStyle: ctx.strokeStyle }),
    strokeRect: (...args: unknown[]) => calls.push({ method: "strokeRect", args, strokeStyle: ctx.strokeStyle }),
    fillRect: (...args: unknown[]) => calls.push({ method: "fillRect", args }),
    fillText: (...args: unknown[]) => calls.push({ method: "fillText", args, fillStyle: ctx.fillStyle }),
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

  it("resolves nested CSS vars for overlay tokens", async () => {
    const hadDocument = Object.prototype.hasOwnProperty.call(globalThis, "document");
    const hadGetComputedStyle = Object.prototype.hasOwnProperty.call(globalThis, "getComputedStyle");
    const originalDocument = (globalThis as any).document;
    const originalGetComputedStyle = (globalThis as any).getComputedStyle;

    try {
      (globalThis as any).document = { documentElement: {} };
      const vars: Record<string, string> = {
        "--chart-series-1": "var(--chart-series-base)",
        "--chart-series-base": "rgb(1, 2, 3)",
        // Prefer grid-scoped tokens for overlay label/selection colors when running inside the spreadsheet.
        "--formula-grid-cell-text": "var(--grid-cell-text-base)",
        "--grid-cell-text-base": "rgb(10, 11, 12)",
        "--text-primary": "rgb(100, 101, 102)",
        // Exercise `var(--missing, fallback)` handling: the referenced token is undefined,
        // so the resolver should return the fallback value (and not the caller fallback).
        "--formula-grid-selection-border": "var(--selection-border-missing, rgb(4, 5, 6))",
        "--selection-border": "rgb(200, 201, 202)",
        "--formula-grid-bg": "var(--grid-bg-base)",
        "--grid-bg-base": "rgb(7, 8, 9)",
        "--bg-primary": "rgb(150, 151, 152)",
      };
      (globalThis as any).getComputedStyle = () => ({
        getPropertyValue: (name: string) => vars[name] ?? "",
      });

      const { ctx, calls } = createStubCanvasContext();
      const canvas = createStubCanvas(ctx);

      const overlay = new DrawingOverlay(canvas, images, geom);
      overlay.setSelectedId(1);
      await overlay.render([createChartObject("chart_1")], viewport);

      const strokeCalls = calls.filter((call) => call.method === "strokeRect");
      // First strokeRect is the placeholder outline.
      expect(strokeCalls.at(0)?.strokeStyle).toBe("rgb(1, 2, 3)");
      // The final strokeRect is the selection border.
      expect(strokeCalls.at(-1)?.strokeStyle).toBe("rgb(4, 5, 6)");

      const labelCall = calls.find((call) => call.method === "fillText");
      expect(labelCall?.fillStyle).toBe("rgb(10, 11, 12)");

      const handleFill = calls.find((call) => call.method === "fill");
      expect(handleFill?.fillStyle).toBe("rgb(7, 8, 9)");
    } finally {
      if (hadDocument) {
        (globalThis as any).document = originalDocument;
      } else {
        // eslint-disable-next-line @typescript-eslint/no-dynamic-delete
        delete (globalThis as any).document;
      }
      if (hadGetComputedStyle) {
        (globalThis as any).getComputedStyle = originalGetComputedStyle;
      } else {
        // eslint-disable-next-line @typescript-eslint/no-dynamic-delete
        delete (globalThis as any).getComputedStyle;
      }
    }
  });
});
