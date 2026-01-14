import { describe, expect, it } from "vitest";

import { DrawingInteractionController } from "../interaction";
import { pxToEmu, type GridGeometry, type Viewport } from "../overlay";
import type { DrawingObject } from "../types";

type Listener = (e: any) => void;

class StubEventTarget {
  private readonly listeners = new Map<string, Array<{ listener: Listener; capture: boolean }>>();
  private readonly pointerCapture = new Set<number>();
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

    // Return a plain object; the controller only uses `left`/`top`.
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

  setPointerCapture(pointerId: number): void {
    this.pointerCapture.add(pointerId);
  }

  releasePointerCapture(pointerId: number): void {
    this.pointerCapture.delete(pointerId);
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

function createPointerEvent(init: { clientX: number; clientY: number; pointerId: number; offsetX?: number; offsetY?: number }): any {
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
  cellSizePx: () => ({ width: 0, height: 0 }),
};

function createAbsoluteObject(): DrawingObject {
  return {
    id: 1,
    kind: { type: "shape" },
    anchor: {
      type: "absolute",
      pos: { xEmu: pxToEmu(50), yEmu: pxToEmu(50) },
      size: { cx: pxToEmu(40), cy: pxToEmu(40) },
    },
    zOrder: 0,
  };
}

describe("DrawingInteractionController zoom + capture", () => {
  it("applies zoom when converting pointer deltas to EMU", () => {
    const el = new StubEventTarget({ left: 100, top: 50 });

    const viewport: Viewport = { scrollX: 0, scrollY: 0, width: 300, height: 200, dpr: 1, zoom: 2 };

    let objects: DrawingObject[] = [createAbsoluteObject()];
    let lastSet: DrawingObject[] | null = null;

    const controller = new DrawingInteractionController(el as unknown as HTMLElement, geom, {
      getViewport: () => viewport,
      getObjects: () => objects,
      setObjects: (next) => {
        lastSet = next;
        objects = next;
      },
    });

    // Local coordinates are (client - rect). Set bogus `offsetX/Y` to ensure we don't rely on them.
    // Object starts at (50,50) with size 40x40 in document-space; zoom=2 means the on-screen rect
    // begins at (100,100) with size 80x80. (Rect is relative to the element's top-left.)
    el.dispatchPointerEvent("pointerdown", createPointerEvent({ clientX: 210, clientY: 160, pointerId: 1, offsetX: 0, offsetY: 0 }));
    el.dispatchPointerEvent("pointermove", createPointerEvent({ clientX: 230, clientY: 160, pointerId: 1, offsetX: 0, offsetY: 0 }));

    expect(lastSet).not.toBeNull();
    const moved = (lastSet ?? [])[0]!;
    expect(moved.anchor.type).toBe("absolute");
    if (moved.anchor.type === "absolute") {
      // Screen dx = 20px; zoom=2 => sheet dx = 10px.
      expect(moved.anchor.pos.xEmu).toBe(pxToEmu(60));
      expect(moved.anchor.pos.yEmu).toBe(pxToEmu(50));
    }

    controller.dispose();
  });

  it("stops propagation on hit, but not on miss (capture listener)", () => {
    const el = new StubEventTarget({ left: 0, top: 0 });

    const viewport: Viewport = { scrollX: 0, scrollY: 0, width: 300, height: 200, dpr: 1, zoom: 1 };
    let objects: DrawingObject[] = [createAbsoluteObject()];

    const counters = { down: 0, move: 0, up: 0 };
    el.addEventListener("pointerdown", () => counters.down++, false);
    el.addEventListener("pointermove", () => counters.move++, false);
    el.addEventListener("pointerup", () => counters.up++, false);

    const controller = new DrawingInteractionController(
      el as unknown as HTMLElement,
      geom,
      {
        getViewport: () => viewport,
        getObjects: () => objects,
        setObjects: (next) => {
          objects = next;
        },
      },
      { capture: true },
    );

    // Hit: should not reach bubble listeners.
    el.dispatchPointerEvent("pointerdown", createPointerEvent({ clientX: 60, clientY: 60, pointerId: 1 }));
    el.dispatchPointerEvent("pointermove", createPointerEvent({ clientX: 70, clientY: 60, pointerId: 1 }));
    el.dispatchPointerEvent("pointerup", createPointerEvent({ clientX: 70, clientY: 60, pointerId: 1 }));
    expect(counters).toEqual({ down: 0, move: 0, up: 0 });

    // Miss: should be allowed through (normal grid selection).
    el.dispatchPointerEvent("pointerdown", createPointerEvent({ clientX: 10, clientY: 10, pointerId: 2 }));
    expect(counters.down).toBe(1);

    controller.dispose();
  });
});
