import { describe, expect, it, vi } from "vitest";

import { DrawingSpatialIndex } from "../spatialIndex";
import { pxToEmu, type GridGeometry } from "../overlay";
import type { DrawingObject, Rect } from "../types";

function absObject(opts: { id: number; zOrder: number; x: number; y: number; w: number; h: number }): DrawingObject {
  return {
    id: opts.id,
    kind: { type: "shape", label: `obj_${opts.id}` },
    zOrder: opts.zOrder,
    anchor: {
      type: "absolute",
      pos: { xEmu: pxToEmu(opts.x), yEmu: pxToEmu(opts.y) },
      size: { cx: pxToEmu(opts.w), cy: pxToEmu(opts.h) },
    },
  };
}

const geom: GridGeometry = {
  cellOriginPx: () => ({ x: 0, y: 0 }),
  cellSizePx: () => ({ width: 0, height: 0 }),
};

describe("DrawingSpatialIndex", () => {
  it("absolute anchors include the A1 origin from grid geometry", () => {
    const originGeom: GridGeometry = {
      cellOriginPx: () => ({ x: 100, y: 200 }),
      cellSizePx: () => ({ width: 0, height: 0 }),
    };

    const objects: DrawingObject[] = [absObject({ id: 1, zOrder: 0, x: 10, y: 20, w: 30, h: 40 })];

    const index = new DrawingSpatialIndex({ tileSizePx: 512 });
    index.rebuild(objects, originGeom, 1);

    expect(index.getRect(1)).toEqual({ x: 110, y: 220, width: 30, height: 40 });
  });

  it("query returns objects intersecting viewport (in z-order)", () => {
    const objects: DrawingObject[] = [
      absObject({ id: 1, zOrder: 0, x: 10, y: 10, w: 20, h: 20 }),
      absObject({ id: 2, zOrder: 10, x: 700, y: 10, w: 20, h: 20 }),
      // Straddles the tile boundary (512px)
      absObject({ id: 3, zOrder: 5, x: 510, y: 10, w: 20, h: 20 }),
    ];

    const index = new DrawingSpatialIndex({ tileSizePx: 512 });
    index.rebuild(objects, geom, 1);

    const viewport: Rect = { x: 0, y: 0, width: 512, height: 512 };
    const result = index.query(viewport);

    expect(result.map((o) => o.id)).toEqual([1, 3]);
  });

  it("query includes objects whose transformed AABB intersects the viewport", () => {
    const rotated: DrawingObject = {
      ...absObject({ id: 1, zOrder: 0, x: 520, y: 0, w: 100, h: 100 }),
      transform: { rotationDeg: 45, flipH: false, flipV: false },
    };

    const index = new DrawingSpatialIndex({ tileSizePx: 512 });
    index.rebuild([rotated], geom, 1);

    // Use a viewport that stays within the first 512px tile so we verify the
    // index buckets based on transformed AABBs (not just the raw anchor rect).
    const viewport: Rect = { x: 0, y: 0, width: 511, height: 511 };
    const result = index.query(viewport);

    expect(result.map((o) => o.id)).toEqual([1]);
  });

  it("hitTest returns the topmost object among candidates", () => {
    const objects: DrawingObject[] = [
      absObject({ id: 1, zOrder: 0, x: 0, y: 0, w: 100, h: 100 }),
      absObject({ id: 2, zOrder: 1, x: 0, y: 0, w: 100, h: 100 }),
    ];

    const index = new DrawingSpatialIndex({ tileSizePx: 512 });
    index.rebuild(objects, geom, 1);

    const hit = index.hitTestWithBounds({ x: 50, y: 50 }, { scrollX: 0, scrollY: 0 });
    expect(hit?.object.id).toBe(2);
    expect(hit?.bounds).toEqual({ x: 0, y: 0, width: 100, height: 100 });
  });

  it("perf guard: query does not sort and returns a small subset", () => {
    const objects: DrawingObject[] = [];
    for (let i = 0; i < 10_000; i += 1) {
      objects.push(absObject({ id: i, zOrder: i, x: i * 600, y: 0, w: 10, h: 10 }));
    }

    const index = new DrawingSpatialIndex({ tileSizePx: 512 });
    index.rebuild(objects, geom, 1);

    const sortSpy = vi.spyOn(Array.prototype, "sort");
    const viewport: Rect = { x: 0, y: 0, width: 512, height: 512 };
    for (let i = 0; i < 50; i += 1) {
      const result = index.query(viewport);
      expect(result.length).toBeLessThan(50);
    }
    expect(sortSpy).not.toHaveBeenCalled();
    sortSpy.mockRestore();
  });
});
