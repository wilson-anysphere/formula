import { afterEach, describe, expect, it, vi } from "vitest";

import { pxToEmu, type GridGeometry, type Viewport } from "../overlay";
import { DrawingInteractionController } from "../interaction";
import type { DrawingObject } from "../types";

function absoluteObject(id: number, zOrder: number, x: number, y: number): DrawingObject {
  return {
    id,
    kind: { type: "shape" },
    anchor: {
      type: "absolute",
      pos: { xEmu: pxToEmu(x), yEmu: pxToEmu(y) },
      size: { cx: pxToEmu(10), cy: pxToEmu(10) },
    },
    zOrder,
  };
}

describe("DrawingInteractionController hit-test perf", () => {
  afterEach(() => {
    vi.restoreAllMocks();
  });

  it("does not re-sort on repeated pointermove when objects are unchanged", () => {
    const listeners = new Map<string, (e: any) => void>();
    const viewport: Viewport = { scrollX: 0, scrollY: 0, width: 800, height: 600, dpr: 1 };
    const canvas: any = {
      style: { cursor: "" },
      getBoundingClientRect: () => ({ left: 0, top: 0 } as DOMRect),
      addEventListener: (type: string, cb: (e: any) => void) => listeners.set(type, cb),
      removeEventListener: (type: string) => listeners.delete(type),
      setPointerCapture: () => {},
      releasePointerCapture: () => {},
    };

    const geom: GridGeometry = {
      cellOriginPx: () => ({ x: 0, y: 0 }),
      cellSizePx: () => ({ width: 10, height: 10 }),
    };

    const objects: DrawingObject[] = [];
    for (let i = 0; i < 5_000; i += 1) {
      // Spread the objects out so most buckets contain few candidates.
      objects.push(absoluteObject(i, i, i * 50, i * 50));
    }

    const callbacks = {
      getViewport: () => viewport,
      getObjects: () => objects,
      setObjects: () => {},
    };

    new DrawingInteractionController(canvas as HTMLCanvasElement, geom, callbacks);

    const pointerMove = listeners.get("pointermove");
    expect(pointerMove).toBeTypeOf("function");

    const sortSpy = vi.spyOn(Array.prototype, "sort");

    // First pointer move builds the index (one sort).
    pointerMove!({ clientX: 25, clientY: 25, pointerId: 1 });
    expect(sortSpy).toHaveBeenCalledTimes(1);

    // Subsequent moves should reuse the cached index (no more sorts).
    sortSpy.mockClear();
    for (let i = 0; i < 200; i += 1) {
      pointerMove!({ clientX: 25 + (i % 3), clientY: 25 + (i % 3), pointerId: 1 });
    }
    expect(sortSpy).toHaveBeenCalledTimes(0);
  });
});
