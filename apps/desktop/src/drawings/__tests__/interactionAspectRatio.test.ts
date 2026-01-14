import { describe, expect, it } from "vitest";

import { pxToEmu, type GridGeometry, type Viewport } from "../overlay";
import { DrawingInteractionController } from "../interaction";
import type { DrawingObject } from "../types";

type Listener = (e: any) => void;

class StubEventTarget {
  private readonly listeners = new Map<string, Array<{ listener: Listener; capture: boolean }>>();
  readonly style: { cursor?: string } = {};

  constructor(private readonly rect: { left: number; top: number; width?: number; height?: number }) {}

  addEventListener(type: string, listener: Listener, options?: boolean | AddEventListenerOptions): void {
    const capture = typeof options === "boolean" ? options : Boolean(options?.capture);
    const entries = this.listeners.get(type) ?? [];
    entries.push({ listener, capture });
    this.listeners.set(type, entries);
  }

  removeEventListener(type: string, listener: Listener, options?: boolean | EventListenerOptions): void {
    const capture = typeof options === "boolean" ? options : Boolean(options?.capture);
    const entries = this.listeners.get(type) ?? [];
    this.listeners.set(
      type,
      entries.filter((entry) => entry.listener !== listener || entry.capture !== capture),
    );
  }

  getBoundingClientRect(): DOMRect {
    const width = this.rect.width ?? 300;
    const height = this.rect.height ?? 200;
    const left = this.rect.left;
    const top = this.rect.top;
    return {
      left,
      top,
      right: left + width,
      bottom: top + height,
      width,
      height,
      x: left,
      y: top,
      toJSON: () => ({}),
    } as unknown as DOMRect;
  }

  setPointerCapture(): void {
    // no-op for tests
  }

  releasePointerCapture(): void {
    // no-op for tests
  }

  dispatchPointerEvent(type: string, e: any): void {
    const entries = this.listeners.get(type) ?? [];
    const capture = entries.filter((entry) => entry.capture);
    const bubble = entries.filter((entry) => !entry.capture);

    for (const entry of capture) {
      if (e._immediateStopped) return;
      entry.listener(e);
    }

    if (e._propagationStopped || e._immediateStopped) return;

    for (const entry of bubble) {
      if (e._immediateStopped) return;
      entry.listener(e);
    }
  }
}

function createPointerEvent(init: { clientX: number; clientY: number; pointerId: number; shiftKey?: boolean }): any {
  return {
    ...init,
    defaultPrevented: false,
    _propagationStopped: false,
    _immediateStopped: false,
    preventDefault() {
      this.defaultPrevented = true;
    },
    stopPropagation() {
      this._propagationStopped = true;
    },
    stopImmediatePropagation() {
      this._immediateStopped = true;
      this._propagationStopped = true;
    },
  };
}

const geom: GridGeometry = {
  cellOriginPx: () => ({ x: 0, y: 0 }),
  cellSizePx: () => ({ width: 100, height: 20 }),
};

const viewport: Viewport = { scrollX: 0, scrollY: 0, width: 1_000, height: 1_000, dpr: 1, zoom: 1 };

function createImageObject(opts?: { transform?: DrawingObject["transform"] }): DrawingObject {
  return {
    id: 1,
    kind: { type: "image", imageId: "img_1" },
    anchor: {
      type: "absolute",
      pos: { xEmu: 0, yEmu: 0 },
      size: { cx: pxToEmu(200), cy: pxToEmu(100) },
    },
    zOrder: 0,
    transform: opts?.transform,
  };
}

function createShapeObject(): DrawingObject {
  return {
    id: 1,
    kind: { type: "shape" },
    anchor: {
      type: "absolute",
      pos: { xEmu: 0, yEmu: 0 },
      size: { cx: pxToEmu(200), cy: pxToEmu(100) },
    },
    zOrder: 0,
  };
}

describe("DrawingInteractionController image resize aspect ratio", () => {
  it("keeps the original aspect ratio when Shift is held during corner resize", () => {
    const el = new StubEventTarget({ left: 0, top: 0 });
    let objects: DrawingObject[] = [createImageObject()];

    new DrawingInteractionController(el as unknown as HTMLElement, geom, {
      getViewport: () => viewport,
      getObjects: () => objects,
      setObjects: (next) => {
        objects = next;
      },
    });

    // Start resizing from the south-east corner.
    el.dispatchPointerEvent("pointerdown", createPointerEvent({ clientX: 200, clientY: 100, pointerId: 1 }));

    // Drag horizontally while holding Shift; height should be adjusted to keep 2:1.
    el.dispatchPointerEvent("pointermove", createPointerEvent({ clientX: 250, clientY: 100, pointerId: 1, shiftKey: true }));

    expect(objects[0]?.anchor).toMatchObject({
      type: "absolute",
      size: { cx: pxToEmu(250), cy: pxToEmu(125) },
    });
  });

  it("does not lock aspect ratio when resizing from an edge handle (even for images)", () => {
    const el = new StubEventTarget({ left: 0, top: 0 });
    let objects: DrawingObject[] = [createImageObject()];

    new DrawingInteractionController(el as unknown as HTMLElement, geom, {
      getViewport: () => viewport,
      getObjects: () => objects,
      setObjects: (next) => {
        objects = next;
      },
    });

    // Start resizing from the east (right) edge handle.
    el.dispatchPointerEvent("pointerdown", createPointerEvent({ clientX: 200, clientY: 50, pointerId: 1 }));

    // Drag right while holding Shift; height should remain unchanged for edge handles.
    el.dispatchPointerEvent("pointermove", createPointerEvent({ clientX: 250, clientY: 50, pointerId: 1, shiftKey: true }));

    expect(objects[0]?.anchor).toMatchObject({
      type: "absolute",
      size: { cx: pxToEmu(250), cy: pxToEmu(100) },
    });
  });

  it("does not lock aspect ratio for non-image objects", () => {
    const el = new StubEventTarget({ left: 0, top: 0 });
    let objects: DrawingObject[] = [createShapeObject()];

    new DrawingInteractionController(el as unknown as HTMLElement, geom, {
      getViewport: () => viewport,
      getObjects: () => objects,
      setObjects: (next) => {
        objects = next;
      },
    });

    // Start resizing from the south-east corner.
    el.dispatchPointerEvent("pointerdown", createPointerEvent({ clientX: 200, clientY: 100, pointerId: 1 }));

    // Holding Shift should not change behavior for non-images.
    el.dispatchPointerEvent("pointermove", createPointerEvent({ clientX: 250, clientY: 100, pointerId: 1, shiftKey: true }));

    expect(objects[0]?.anchor).toMatchObject({
      type: "absolute",
      size: { cx: pxToEmu(250), cy: pxToEmu(100) },
    });
  });

  it("keeps the original aspect ratio for rotated images (lock is applied in local coords)", () => {
    const el = new StubEventTarget({ left: 0, top: 0 });
    let objects: DrawingObject[] = [
      createImageObject({ transform: { rotationDeg: 90, flipH: false, flipV: false } }),
    ];

    new DrawingInteractionController(el as unknown as HTMLElement, geom, {
      getViewport: () => viewport,
      getObjects: () => objects,
      setObjects: (next) => {
        objects = next;
      },
    });

    // For a 200x100 rect rotated 90deg, the "se" local handle center is at (50, 150).
    el.dispatchPointerEvent("pointerdown", createPointerEvent({ clientX: 50, clientY: 150, pointerId: 1 }));

    // Move down by 50px while holding Shift. This corresponds to +50px local width delta.
    el.dispatchPointerEvent("pointermove", createPointerEvent({ clientX: 50, clientY: 200, pointerId: 1, shiftKey: true }));

    expect(objects[0]?.anchor).toMatchObject({
      type: "absolute",
      size: { cx: pxToEmu(250), cy: pxToEmu(125) },
    });
  });

  it("keeps the original aspect ratio for flipped images (lock is applied in local coords)", () => {
    const el = new StubEventTarget({ left: 0, top: 0 });
    let objects: DrawingObject[] = [
      createImageObject({ transform: { rotationDeg: 0, flipH: true, flipV: false } }),
    ];

    new DrawingInteractionController(el as unknown as HTMLElement, geom, {
      getViewport: () => viewport,
      getObjects: () => objects,
      setObjects: (next) => {
        objects = next;
      },
    });

    // For a 200x100 rect flipped horizontally, the local "se" handle center maps to the bottom-left corner (0, 100).
    el.dispatchPointerEvent("pointerdown", createPointerEvent({ clientX: 0, clientY: 100, pointerId: 1 }));

    // Drag left by 50px while holding Shift. This corresponds to +50px local width delta.
    el.dispatchPointerEvent("pointermove", createPointerEvent({ clientX: -50, clientY: 100, pointerId: 1, shiftKey: true }));

    expect(objects[0]?.anchor).toMatchObject({
      type: "absolute",
      size: { cx: pxToEmu(250), cy: pxToEmu(125) },
    });
  });

  it("allows width/height to change independently when Shift is not held", () => {
    const el = new StubEventTarget({ left: 0, top: 0 });
    let objects: DrawingObject[] = [createImageObject()];

    new DrawingInteractionController(el as unknown as HTMLElement, geom, {
      getViewport: () => viewport,
      getObjects: () => objects,
      setObjects: (next) => {
        objects = next;
      },
    });

    // Start resizing from the south-east corner.
    el.dispatchPointerEvent("pointerdown", createPointerEvent({ clientX: 200, clientY: 100, pointerId: 1 }));

    // Drag horizontally without Shift; height should remain unchanged.
    el.dispatchPointerEvent("pointermove", createPointerEvent({ clientX: 250, clientY: 100, pointerId: 1, shiftKey: false }));

    expect(objects[0]?.anchor).toMatchObject({
      type: "absolute",
      size: { cx: pxToEmu(250), cy: pxToEmu(100) },
    });
  });
});
