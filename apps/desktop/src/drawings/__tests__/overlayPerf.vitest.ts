import { describe, expect, it, vi, afterEach } from "vitest";

import { DrawingOverlay, type GridGeometry, type Viewport } from "../overlay";
import type { ImageStore } from "../types";

function createStubCanvasContext(): CanvasRenderingContext2D {
  const ctx: any = {
    clearRect: vi.fn(),
    drawImage: vi.fn(),
    save: vi.fn(),
    restore: vi.fn(),
    beginPath: vi.fn(),
    rect: vi.fn(),
    clip: vi.fn(),
    setLineDash: vi.fn(),
    strokeRect: vi.fn(),
    fillRect: vi.fn(),
    fillText: vi.fn(),
  };

  return ctx as CanvasRenderingContext2D;
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
  cellOriginPx: () => ({ x: 0, y: 0 }),
  cellSizePx: () => ({ width: 0, height: 0 }),
};

const viewport: Viewport = { scrollX: 0, scrollY: 0, width: 100, height: 100, dpr: 1 };

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

describe("DrawingOverlay perf guards", () => {
  it("does not call getComputedStyle on every render()", async () => {
    const getPropertyValue = vi.fn((name: string) => {
      switch (name) {
        case "--chart-series-1":
          return "rgb(1, 2, 3)";
        case "--chart-series-2":
          return "rgb(4, 5, 6)";
        case "--chart-series-3":
          return "rgb(7, 8, 9)";
        case "--text-primary":
          return "rgb(10, 11, 12)";
        case "--selection-border":
          return "rgb(13, 14, 15)";
        case "--bg-primary":
          return "rgb(16, 17, 18)";
        default:
          return "";
      }
    });

    const getComputedStyleSpy = vi.fn(() => ({ getPropertyValue }) as any);

    vi.stubGlobal("document", { documentElement: {} } as any);
    vi.stubGlobal("getComputedStyle", getComputedStyleSpy as any);

    const ctx = createStubCanvasContext();
    const canvas = createStubCanvas(ctx);
    const overlay = new DrawingOverlay(canvas, images, geom);

    await overlay.render([], viewport);
    await overlay.render([], viewport);

    expect(getComputedStyleSpy).toHaveBeenCalledTimes(1);
  });

  it("does not sort objects on render() when already zOrder-sorted", async () => {
    const getPropertyValue = vi.fn(() => "");
    const getComputedStyleSpy = vi.fn(() => ({ getPropertyValue }) as any);

    vi.stubGlobal("document", { documentElement: {} } as any);
    vi.stubGlobal("getComputedStyle", getComputedStyleSpy as any);

    const ctx = createStubCanvasContext();
    const canvas = createStubCanvas(ctx);
    const overlay = new DrawingOverlay(canvas, images, geom);

    const objects = [
      {
        id: 1,
        kind: { type: "shape" as const },
        anchor: {
          type: "absolute" as const,
          pos: { xEmu: 0, yEmu: 0 },
          size: { cx: 0, cy: 0 },
        },
        zOrder: 0,
      },
      {
        id: 2,
        kind: { type: "shape" as const },
        anchor: {
          type: "absolute" as const,
          pos: { xEmu: 0, yEmu: 0 },
          size: { cx: 0, cy: 0 },
        },
        zOrder: 1,
      },
    ];

    const sortSpy = vi.spyOn(Array.prototype, "sort");
    await overlay.render(objects as any, viewport, { drawObjects: false });
    await overlay.render(objects.slice() as any, viewport, { drawObjects: false });
    expect(sortSpy).toHaveBeenCalledTimes(0);
  });
});
