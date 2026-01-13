import { describe, expect, it, vi } from "vitest";

import { pxToEmu } from "../overlay";
import { DrawingInteractionController } from "../interaction";
import {
  cursorForResizeHandle,
  getResizeHandleCenters,
  hitTestResizeHandle,
  type ResizeHandle,
} from "../selectionHandles";
import type { DrawingObject } from "../types";
import type { GridGeometry, Viewport } from "../overlay";

describe("drawings selection handles", () => {
  it("hitTestResizeHandle detects all 8 handles", () => {
    const bounds = { x: 100, y: 200, width: 80, height: 40 };
    const centers = getResizeHandleCenters(bounds);
    for (const c of centers) {
      expect(hitTestResizeHandle(bounds, c.x, c.y)).toBe(c.handle);
    }

    // Center should not register as a handle.
    expect(hitTestResizeHandle(bounds, 140, 220)).toBeNull();
  });

  it("cursorForResizeHandle matches expected CSS cursors", () => {
    const expected: Record<ResizeHandle, string> = {
      nw: "nwse-resize",
      n: "ns-resize",
      ne: "nesw-resize",
      e: "ew-resize",
      se: "nwse-resize",
      s: "ns-resize",
      sw: "nesw-resize",
      w: "ew-resize",
    };

    for (const [handle, cursor] of Object.entries(expected) as Array<[ResizeHandle, string]>) {
      expect(cursorForResizeHandle(handle)).toBe(cursor);
    }
  });

  it("DrawingInteractionController updates cursor on hover for selected object handles", () => {
    const listeners = new Map<string, (e: any) => void>();
    const canvas: any = {
      style: { cursor: "" },
      addEventListener: (type: string, cb: (e: any) => void) => listeners.set(type, cb),
      removeEventListener: (type: string) => listeners.delete(type),
      setPointerCapture: vi.fn(),
      releasePointerCapture: vi.fn(),
    };

    const geom: GridGeometry = {
      cellOriginPx: () => ({ x: 0, y: 0 }),
      cellSizePx: () => ({ width: 0, height: 0 }),
    };

    const viewport: Viewport = { scrollX: 0, scrollY: 0, width: 500, height: 500, dpr: 1 };

    let objects: DrawingObject[] = [
      {
        id: 1,
        kind: { type: "shape", label: "shape" },
        anchor: {
          type: "absolute",
          pos: { xEmu: pxToEmu(100), yEmu: pxToEmu(200) },
          size: { cx: pxToEmu(80), cy: pxToEmu(40) },
        },
        zOrder: 0,
      },
    ];

    const callbacks = {
      getViewport: () => viewport,
      getObjects: () => objects,
      setObjects: (next: DrawingObject[]) => {
        objects = next;
      },
      onSelectionChange: vi.fn(),
    };

    new DrawingInteractionController(canvas as HTMLCanvasElement, geom, callbacks);

    const pointerDown = listeners.get("pointerdown");
    const pointerUp = listeners.get("pointerup");
    const pointerMove = listeners.get("pointermove");
    expect(pointerDown).toBeTypeOf("function");
    expect(pointerUp).toBeTypeOf("function");
    expect(pointerMove).toBeTypeOf("function");

    // Click once to select (and finish the drag sequence).
    pointerDown!({ offsetX: 140, offsetY: 220, pointerId: 1 });
    pointerUp!({ offsetX: 140, offsetY: 220, pointerId: 1 });

    const bounds = { x: 100, y: 200, width: 80, height: 40 };
    for (const c of getResizeHandleCenters(bounds)) {
      pointerMove!({ offsetX: c.x, offsetY: c.y, pointerId: 1 });
      expect(canvas.style.cursor).toBe(cursorForResizeHandle(c.handle));
    }
  });

  it("allows starting a resize from handle hit areas outside the object bounds", () => {
    const listeners = new Map<string, (e: any) => void>();
    const canvas: any = {
      style: { cursor: "" },
      addEventListener: (type: string, cb: (e: any) => void) => listeners.set(type, cb),
      removeEventListener: (type: string) => listeners.delete(type),
      setPointerCapture: vi.fn(),
      releasePointerCapture: vi.fn(),
    };

    const geom: GridGeometry = {
      cellOriginPx: () => ({ x: 0, y: 0 }),
      cellSizePx: () => ({ width: 0, height: 0 }),
    };

    const viewport: Viewport = { scrollX: 0, scrollY: 0, width: 500, height: 500, dpr: 1 };

    const initial: DrawingObject = {
      id: 1,
      kind: { type: "shape", label: "shape" },
      anchor: {
        type: "absolute",
        pos: { xEmu: pxToEmu(100), yEmu: pxToEmu(200) },
        size: { cx: pxToEmu(80), cy: pxToEmu(40) },
      },
      zOrder: 0,
    };

    let objects: DrawingObject[] = [initial];

    const callbacks = {
      getViewport: () => viewport,
      getObjects: () => objects,
      setObjects: (next: DrawingObject[]) => {
        objects = next;
      },
      onSelectionChange: vi.fn(),
    };

    new DrawingInteractionController(canvas as HTMLCanvasElement, geom, callbacks);

    const pointerDown = listeners.get("pointerdown");
    const pointerUp = listeners.get("pointerup");
    const pointerMove = listeners.get("pointermove");
    expect(pointerDown).toBeTypeOf("function");
    expect(pointerUp).toBeTypeOf("function");
    expect(pointerMove).toBeTypeOf("function");

    // Select the object first.
    pointerDown!({ offsetX: 140, offsetY: 220, pointerId: 1 });
    pointerUp!({ offsetX: 140, offsetY: 220, pointerId: 1 });
    expect(callbacks.onSelectionChange).toHaveBeenCalledTimes(1);

    // Click slightly outside the object's bounds but within the top-left handle hit box.
    // Object bounds: x in [100, 180], y in [200, 240]. This point is outside (x<100,y<200).
    pointerDown!({ offsetX: 97, offsetY: 197, pointerId: 1 });
    expect(callbacks.onSelectionChange).toHaveBeenCalledTimes(1);
    expect(canvas.style.cursor).toBe("nwse-resize");

    // Drag the handle up/left by 10px; this should expand the rect and move the origin.
    pointerMove!({ offsetX: 87, offsetY: 187, pointerId: 1 });
    const resized = objects[0]!.anchor;
    expect(resized.type).toBe("absolute");
    if (resized.type !== "absolute") throw new Error("unexpected anchor type");
    expect(resized.pos.xEmu).toBeCloseTo(pxToEmu(90));
    expect(resized.pos.yEmu).toBeCloseTo(pxToEmu(190));
    expect(resized.size.cx).toBeCloseTo(pxToEmu(90));
    expect(resized.size.cy).toBeCloseTo(pxToEmu(50));

    pointerUp!({ offsetX: 87, offsetY: 187, pointerId: 1 });
  });
});
