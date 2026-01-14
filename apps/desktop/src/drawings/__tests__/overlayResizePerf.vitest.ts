import { describe, expect, it, vi } from "vitest";

import { DrawingOverlay, type GridGeometry, type Viewport } from "../overlay";
import type { ImageStore } from "../types";

const images: ImageStore = {
  get: () => undefined,
  set: () => {},
};

const geom: GridGeometry = {
  cellOriginPx: () => ({ x: 0, y: 0 }),
  cellSizePx: () => ({ width: 0, height: 0 }),
};

describe("DrawingOverlay.resize()", () => {
  it("avoids resetting canvas backing buffer when the viewport is unchanged", () => {
    const setTransform = vi.fn();
    const ctx: any = { setTransform };

    let width = 0;
    let height = 0;
    let widthSets = 0;
    let heightSets = 0;

    const style: any = {};
    let styleWidth = "";
    let styleHeight = "";
    let styleWidthSets = 0;
    let styleHeightSets = 0;

    Object.defineProperties(style, {
      width: {
        get() {
          return styleWidth;
        },
        set(value: string) {
          styleWidth = value;
          styleWidthSets += 1;
        },
      },
      height: {
        get() {
          return styleHeight;
        },
        set(value: string) {
          styleHeight = value;
          styleHeightSets += 1;
        },
      },
    });

    const canvas: any = {
      style,
      getContext: (type: string) => (type === "2d" ? ctx : null),
    };

    Object.defineProperties(canvas, {
      width: {
        get() {
          return width;
        },
        set(value: number) {
          width = value;
          widthSets += 1;
        },
      },
      height: {
        get() {
          return height;
        },
        set(value: number) {
          height = value;
          heightSets += 1;
        },
      },
    });

    const overlay = new DrawingOverlay(canvas as HTMLCanvasElement, images, geom);

    const viewport: Viewport = { scrollX: 0, scrollY: 0, width: 123, height: 456, dpr: 2 };

    overlay.resize(viewport);
    overlay.resize({ ...viewport });

    expect(width).toBe(Math.floor(viewport.width * viewport.dpr));
    expect(height).toBe(Math.floor(viewport.height * viewport.dpr));
    expect(widthSets).toBe(1);
    expect(heightSets).toBe(1);
    expect(styleWidthSets).toBe(1);
    expect(styleHeightSets).toBe(1);
    expect(setTransform).toHaveBeenCalledTimes(1);
  });

  it("still updates canvas sizing and transform when the DPR changes", () => {
    const setTransform = vi.fn();
    const ctx: any = { setTransform };

    const canvas: any = {
      width: 0,
      height: 0,
      style: {},
      getContext: (type: string) => (type === "2d" ? ctx : null),
    };

    const overlay = new DrawingOverlay(canvas as HTMLCanvasElement, images, geom);

    overlay.resize({ scrollX: 0, scrollY: 0, width: 100, height: 50, dpr: 1 });
    overlay.resize({ scrollX: 0, scrollY: 0, width: 100, height: 50, dpr: 2 });

    expect(canvas.width).toBe(200);
    expect(canvas.height).toBe(100);
    expect(setTransform).toHaveBeenCalledTimes(2);
  });
});

