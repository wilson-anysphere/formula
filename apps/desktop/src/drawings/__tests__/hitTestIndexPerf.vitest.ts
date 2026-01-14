import { afterEach, describe, expect, it, vi } from "vitest";

import { buildHitTestIndex, hitTestDrawings } from "../hitTest";
import { pxToEmu, type GridGeometry, type Viewport } from "../overlay";
import type { DrawingObject } from "../types";

const geom: GridGeometry = {
  cellOriginPx: () => ({ x: 0, y: 0 }),
  cellSizePx: () => ({ width: 64, height: 20 }),
};

const viewport: Viewport = { scrollX: 0, scrollY: 0, width: 800, height: 600, dpr: 1 };

function absoluteObject(id: number, zOrder: number, rect: { x: number; y: number; width: number; height: number }): DrawingObject {
  return {
    id,
    kind: { type: "shape" },
    anchor: {
      type: "absolute",
      pos: { xEmu: pxToEmu(rect.x), yEmu: pxToEmu(rect.y) },
      size: { cx: pxToEmu(rect.width), cy: pxToEmu(rect.height) },
    },
    zOrder,
  };
}

describe("drawings hit test index perf", () => {
  afterEach(() => {
    vi.restoreAllMocks();
  });

  it("avoids per-hit sort when reusing a cached index", () => {
    const objects: DrawingObject[] = [];
    for (let i = 0; i < 5_000; i += 1) {
      // Spread shapes far apart so each spatial bucket has very few candidates.
      objects.push(absoluteObject(i, i, { x: i * 50, y: i * 50, width: 10, height: 10 }));
    }

    // `buildHitTestIndex` is allowed to sort once; `hitTestDrawings` should not sort.
    const sortSpy = vi.spyOn(Array.prototype, "sort");
    const index = buildHitTestIndex(objects, geom, { bucketSizePx: 128 });
    // Objects are already zOrder-sorted, so index build should avoid sorting.
    expect(sortSpy).toHaveBeenCalledTimes(0);
    sortSpy.mockClear();

    for (let i = 0; i < 200; i += 1) {
      // No object is near (25, 25) so this exercises the "miss" path repeatedly.
      expect(hitTestDrawings(index, viewport, 25, 25, geom)).toBeNull();
    }

    expect(sortSpy).toHaveBeenCalledTimes(0);
  });

  it("avoids index-build sorting when objects are already zOrder-non-increasing (ties allowed)", () => {
    const objects: DrawingObject[] = [
      absoluteObject(1, 10, { x: 0, y: 0, width: 10, height: 10 }),
      absoluteObject(2, 10, { x: 20, y: 0, width: 10, height: 10 }),
      absoluteObject(3, 5, { x: 40, y: 0, width: 10, height: 10 }),
      absoluteObject(4, 5, { x: 60, y: 0, width: 10, height: 10 }),
      absoluteObject(5, 0, { x: 80, y: 0, width: 10, height: 10 }),
    ];

    const sortSpy = vi.spyOn(Array.prototype, "sort");
    buildHitTestIndex(objects, geom, { bucketSizePx: 128 });
    expect(sortSpy).toHaveBeenCalledTimes(0);
  });

  it("hit tests ties in reverse render order when input is already descending", () => {
    // Render order (after stable sort ascending by zOrder) will preserve the input
    // order for ties; the later object is drawn last and should be considered "on top".
    const objects: DrawingObject[] = [
      absoluteObject(1, 10, { x: 0, y: 0, width: 10, height: 10 }),
      absoluteObject(2, 10, { x: 0, y: 0, width: 10, height: 10 }),
    ];
    const index = buildHitTestIndex(objects, geom, { bucketSizePx: 128 });
    const hit = hitTestDrawings(index, viewport, 5, 5, geom);
    expect(hit?.object.id).toBe(2);
  });

  it("returns the top-most object when multiple overlap (including global candidates)", () => {
    const objects: DrawingObject[] = [
      // Large object spans many buckets and should end up in the `global` list.
      absoluteObject(1, 10, { x: 0, y: 0, width: 20_000, height: 20_000 }),
      // Smaller object overlaps, but is behind the global object.
      absoluteObject(2, 0, { x: 5, y: 5, width: 10, height: 10 }),
    ];

    const index = buildHitTestIndex(objects, geom, { bucketSizePx: 128, maxBucketsPerObject: 64 });
    const hit = hitTestDrawings(index, viewport, 6, 6, geom);
    expect(hit?.object.id).toBe(1);
  });
});
