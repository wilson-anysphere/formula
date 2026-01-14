import { afterEach, beforeEach, describe, expect, it } from "vitest";

import { ChartRendererAdapter } from "../../charts/chartRendererAdapter";
import { FormulaChartModelStore } from "../../charts/formulaChartModelStore";
import type { ChartModel } from "../../charts/renderChart";
import { DrawingOverlay, pxToEmu, type GridGeometry, type Viewport } from "../overlay";
import type { DrawingObject, ImageStore } from "../types";

function createStubCanvasContext(): { ctx: CanvasRenderingContext2D; calls: Array<{ method: string; args: unknown[] }> } {
  const calls: Array<{ method: string; args: unknown[] }> = [];
  const record =
    (method: string) =>
    (...args: unknown[]) =>
      calls.push({ method, args });

  const ctx: any = {
    // Used by overlay.
    clearRect: record("clearRect"),
    drawImage: record("drawImage"),
    save: record("save"),
    restore: record("restore"),
    beginPath: record("beginPath"),
    rect: record("rect"),
    clip: record("clip"),
    setLineDash: record("setLineDash"),
    strokeRect: record("strokeRect"),
    fillText: record("fillText"),

    // Optional: used by ChartRendererAdapter.getContextScale.
    getTransform: () => ({ a: 1 }),
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

function createNoopSurfaceContext(): CanvasRenderingContext2D {
  const noop = () => {};
  const gradient = { addColorStop: noop } as any;
  // `charts/scene/color.ts` uses a Canvas context as a last-chance CSS parser by setting
  // `fillStyle` and reading it back. Real canvas contexts *reject* unknown values (including
  // CSS variables and many non-hex tokens) by leaving the prior fillStyle unchanged.
  //
  // In node tests we don't have a real canvas implementation, so mimic the relevant
  // semantics to avoid infinite recursion in `parseColor()`.
  const target: any = {
    globalAlpha: 1,
    strokeStyle: "black",
    lineWidth: 1,
    measureText: (text: string) => ({ width: text.length * 8 }),
    createLinearGradient: () => gradient,
    createPattern: () => null,
    getImageData: () => ({ data: new Uint8ClampedArray(), width: 0, height: 0 }),
    putImageData: noop,
  };

  let fillStyle = "#000000";
  Object.defineProperty(target, "fillStyle", {
    configurable: true,
    get: () => fillStyle,
    set: (value: unknown) => {
      const next = String(value ?? "").trim();
      if (!next) return;
      // Accept values `parseColor()` can understand without needing the browser's
      // CSS parser. Everything else is treated as invalid (leave prior value).
      if (next === "transparent" || next.startsWith("#") || /^rgba?\(/i.test(next)) {
        fillStyle = next;
      }
    },
  });

  const ctx: any = new Proxy(
    target,
    {
      get(target, prop) {
        if (prop in target) return (target as any)[prop];
        return noop;
      },
      set(target, prop, value) {
        (target as any)[prop] = value;
        return true;
      },
    },
  );
  return ctx as CanvasRenderingContext2D;
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

describe("ChartRendererAdapter + DrawingOverlay", () => {
  const originalOffscreen = (globalThis as any).OffscreenCanvas;

  beforeEach(() => {
    // Provide a minimal OffscreenCanvas implementation so ChartRendererAdapter can
    // create its offscreen surface in a node test environment.
    (globalThis as any).OffscreenCanvas = class OffscreenCanvas {
      width: number;
      height: number;
      #ctx: CanvasRenderingContext2D;

      constructor(width: number, height: number) {
        this.width = width;
        this.height = height;
        this.#ctx = createNoopSurfaceContext();
      }

      getContext(type: string) {
        if (type !== "2d") return null;
        return this.#ctx;
      }
    };
  });

  afterEach(() => {
    if (originalOffscreen == null) {
      delete (globalThis as any).OffscreenCanvas;
    } else {
      (globalThis as any).OffscreenCanvas = originalOffscreen;
    }
  });

  it("renders chart objects via ChartRendererAdapter (no placeholder)", async () => {
    const { ctx, calls } = createStubCanvasContext();
    const canvas = createStubCanvas(ctx);

    const store = new FormulaChartModelStore();
    const chartId = "sheet1:1";
    const model: ChartModel = {
      chartType: { kind: "bar" },
      title: "Imported Chart",
      legend: { position: "right", overlay: false },
      axes: [
        { kind: "category", position: "bottom" },
        { kind: "value", position: "left", majorGridlines: true, formatCode: "0" },
      ],
      series: [
        {
          name: "Series 1",
          categories: { cache: ["A", "B"] },
          values: { cache: [1, 2] },
        },
      ],
    };
    store.setChartModel(chartId, model);

    const chartRenderer = new ChartRendererAdapter(store);
    const overlay = new DrawingOverlay(canvas, images, geom, chartRenderer);
    await overlay.render([createChartObject(chartId)], viewport);

    expect(calls.some((call) => call.method === "drawImage")).toBe(true);
    expect(calls.some((call) => call.method === "strokeRect")).toBe(false);
  });

  it("prunes cached chart surfaces when chart objects are removed", async () => {
    const { ctx } = createStubCanvasContext();
    const canvas = createStubCanvas(ctx);

    const store = new FormulaChartModelStore();
    const chartId = "sheet1:prune";
    const model: ChartModel = {
      chartType: { kind: "bar" },
      title: "Chart",
      legend: { position: "right", overlay: false },
      axes: [
        { kind: "category", position: "bottom" },
        { kind: "value", position: "left", majorGridlines: true, formatCode: "0" },
      ],
      series: [
        {
          name: "Series 1",
          categories: { cache: ["A", "B"] },
          values: { cache: [1, 2] },
        },
      ],
    };
    store.setChartModel(chartId, model);

    const chartRenderer = new ChartRendererAdapter(store);
    const overlay = new DrawingOverlay(canvas, images, geom, chartRenderer);

    await overlay.render([createChartObject(chartId)], viewport);
    expect(((chartRenderer as any).surfaces as Map<string, unknown>).size).toBe(1);

    await overlay.render([], viewport);
    expect(((chartRenderer as any).surfaces as Map<string, unknown>).size).toBe(0);
  });
});
