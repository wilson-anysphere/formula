import { afterEach, describe, expect, it, vi } from "vitest";

import { buildHitTestIndex, hitTestDrawings } from "../hitTest";
import { pxToEmu, type GridGeometry, type Viewport } from "../overlay";
import type { DrawingObject } from "../types";

const geom: GridGeometry = {
  cellOriginPx: () => ({ x: 0, y: 0 }),
  cellSizePx: () => ({ width: 0, height: 0 }),
};

const viewport: Viewport = { scrollX: 0, scrollY: 0, width: 500, height: 500, dpr: 1 };

describe("drawings hit test transform perf", () => {
  afterEach(() => {
    vi.restoreAllMocks();
  });

  it("avoids trig calls during repeated hit tests for rotated objects", () => {
    const obj: DrawingObject = {
      id: 1,
      kind: { type: "shape" },
      anchor: {
        type: "absolute",
        pos: { xEmu: pxToEmu(100), yEmu: pxToEmu(100) },
        size: { cx: pxToEmu(100), cy: pxToEmu(50) },
      },
      zOrder: 0,
      transform: { rotationDeg: 45, flipH: false, flipV: false },
    };

    const cosSpy = vi.spyOn(Math, "cos");
    const sinSpy = vi.spyOn(Math, "sin");

    const index = buildHitTestIndex([obj], geom, { bucketSizePx: 64 });
    expect(cosSpy.mock.calls.length).toBeGreaterThan(0);
    expect(sinSpy.mock.calls.length).toBeGreaterThan(0);

    cosSpy.mockClear();
    sinSpy.mockClear();

    // This point is inside the rotated rect for a 45deg rotation about its center.
    for (let i = 0; i < 200; i += 1) {
      expect(hitTestDrawings(index, viewport, 150, 125, geom)?.object.id).toBe(1);
    }

    expect(cosSpy).toHaveBeenCalledTimes(0);
    expect(sinSpy).toHaveBeenCalledTimes(0);
  });
});

