import { describe, expect, it } from "vitest";

import { hitTestResizeHandle, RESIZE_HANDLE_SIZE_PX } from "../selectionHandles";
import { DrawingInteractionController } from "../interaction";
import { pxToEmu, type GridGeometry, type Viewport } from "../overlay";
import type { DrawingObject } from "../types";

function createStubElement(): {
  element: HTMLElement;
  dispatch(type: string, event: PointerEvent): void;
} {
  const listeners = new Map<string, Set<(e: any) => void>>();
  const element: any = {
    style: { cursor: "default" },
    addEventListener: (type: string, listener: (e: any) => void) => {
      let set = listeners.get(type);
      if (!set) {
        set = new Set();
        listeners.set(type, set);
      }
      set.add(listener);
    },
    removeEventListener: (type: string, listener: (e: any) => void) => {
      listeners.get(type)?.delete(listener);
    },
    getBoundingClientRect: () => ({ left: 0, top: 0, width: 0, height: 0 } as any),
    setPointerCapture: () => {},
    releasePointerCapture: () => {},
  };

  return {
    element: element as HTMLElement,
    dispatch: (type, event) => {
      for (const listener of listeners.get(type) ?? []) {
        listener(event as any);
      }
    },
  };
}

function createPointerEvent(clientX: number, clientY: number, pointerId: number): PointerEvent {
  return {
    clientX,
    clientY,
    pointerId,
    preventDefault: () => {},
    stopPropagation: () => {},
    stopImmediatePropagation: () => {},
  } as any;
}

function createAbsoluteObject(bounds: { x: number; y: number; width: number; height: number }): DrawingObject {
  return {
    id: 1,
    kind: { type: "shape", label: "box" },
    anchor: {
      type: "absolute",
      pos: { xEmu: pxToEmu(bounds.x), yEmu: pxToEmu(bounds.y) },
      size: { cx: pxToEmu(bounds.width), cy: pxToEmu(bounds.height) },
    },
    zOrder: 0,
  };
}

const geom: GridGeometry = {
  cellOriginPx: () => ({ x: 0, y: 0 }),
  cellSizePx: () => ({ width: 0, height: 0 }),
};

const viewport: Viewport = { scrollX: 0, scrollY: 0, width: 200, height: 200, dpr: 1 };

describe("drawing resize handle hit testing", () => {
  it("matches the rendered handle geometry", () => {
    const bounds = { x: 10, y: 10, width: 20, height: 20 };
    const half = RESIZE_HANDLE_SIZE_PX / 2;

    // Inside the rendered NW handle square.
    expect(hitTestResizeHandle(bounds, bounds.x + half - 1, bounds.y + half - 1)).toBe("nw");

    // Just outside the rendered handle square (this used to hit when the hit box was larger).
    expect(hitTestResizeHandle(bounds, bounds.x + half + 1, bounds.y + half + 1)).toBeNull();
  });

  it("does not enter resize mode when clicking near (but outside) the visible handle square", () => {
    const { element, dispatch } = createStubElement();

    const bounds = { x: 10, y: 10, width: 20, height: 20 };
    let objects: DrawingObject[] = [createAbsoluteObject(bounds)];
    const setCalls: DrawingObject[][] = [];

    const controller = new DrawingInteractionController(element, geom, {
      getViewport: () => viewport,
      getObjects: () => objects,
      setObjects: (next) => {
        objects = next;
        setCalls.push(next);
      },
    });

    // First click selects the object (and would start dragging, but no move happens).
    dispatch("pointerdown", createPointerEvent(20, 20, 1));
    dispatch("pointerup", createPointerEvent(20, 20, 1));

    // Click just outside the rendered NW handle square. With the previous 10px hit box,
    // this would incorrectly start resizing.
    const half = RESIZE_HANDLE_SIZE_PX / 2;
    const startX = bounds.x + half + 1;
    const startY = bounds.y + half + 1;
    dispatch("pointerdown", createPointerEvent(startX, startY, 1));
    dispatch("pointermove", createPointerEvent(startX + 10, startY + 10, 1));

    expect(setCalls.length).toBe(1);
    const anchor = objects[0]!.anchor;
    expect(anchor.type).toBe("absolute");
    if (anchor.type !== "absolute") throw new Error("expected absolute anchor");

    // Dragging moves the object but keeps its size unchanged.
    expect(anchor.pos.xEmu).toBe(pxToEmu(20));
    expect(anchor.pos.yEmu).toBe(pxToEmu(20));
    expect(anchor.size.cx).toBe(pxToEmu(20));
    expect(anchor.size.cy).toBe(pxToEmu(20));

    controller.dispose();
  });

  it("enters resize mode when clicking within the rendered handle square (even outside the object bounds)", () => {
    const { element, dispatch } = createStubElement();

    const bounds = { x: 10, y: 10, width: 20, height: 20 };
    let objects: DrawingObject[] = [createAbsoluteObject(bounds)];
    const setCalls: DrawingObject[][] = [];

    const controller = new DrawingInteractionController(element, geom, {
      getViewport: () => viewport,
      getObjects: () => objects,
      setObjects: (next) => {
        objects = next;
        setCalls.push(next);
      },
    });

    // Select first (so we can grab the portion of the handle that extends outside the object's bounds).
    dispatch("pointerdown", createPointerEvent(20, 20, 1));
    dispatch("pointerup", createPointerEvent(20, 20, 1));

    const half = RESIZE_HANDLE_SIZE_PX / 2;
    const startX = bounds.x - half + 1;
    const startY = bounds.y - half + 1;
    dispatch("pointerdown", createPointerEvent(startX, startY, 1));
    dispatch("pointermove", createPointerEvent(startX + 10, startY + 10, 1));

    expect(setCalls.length).toBe(1);
    const anchor = objects[0]!.anchor;
    expect(anchor.type).toBe("absolute");
    if (anchor.type !== "absolute") throw new Error("expected absolute anchor");

    // NW resize moves the origin and shrinks the size.
    expect(anchor.pos.xEmu).toBe(pxToEmu(20));
    expect(anchor.pos.yEmu).toBe(pxToEmu(20));
    expect(anchor.size.cx).toBe(pxToEmu(10));
    expect(anchor.size.cy).toBe(pxToEmu(10));

    controller.dispose();
  });
});
