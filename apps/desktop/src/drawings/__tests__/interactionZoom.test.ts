import { describe, expect, it } from "vitest";

import { DrawingInteractionController } from "../interaction";
import type { GridGeometry, Viewport } from "../overlay";
import { pxToEmu } from "../overlay";
import type { DrawingObject } from "../types";

function createStubCanvas(): { canvas: HTMLCanvasElement; listeners: Map<string, (e: any) => void> } {
  const listeners = new Map<string, (e: any) => void>();
  const rect = { left: 0, top: 0, right: 0, bottom: 0, width: 0, height: 0 } as DOMRect;
  const canvas: any = {
    style: { cursor: "" },
    addEventListener: (type: string, handler: (e: any) => void) => {
      listeners.set(type, handler);
    },
    removeEventListener: (type: string) => {
      listeners.delete(type);
    },
    getBoundingClientRect: () => rect,
    setPointerCapture: () => {},
    releasePointerCapture: () => {},
  };
  return { canvas: canvas as HTMLCanvasElement, listeners };
}

const geom: GridGeometry = {
  cellOriginPx: () => ({ x: 0, y: 0 }),
  cellSizePx: () => ({ width: 0, height: 0 }),
};

describe("DrawingInteractionController zoom", () => {
  it("converts drag deltas using the current zoom factor", () => {
    const { canvas, listeners } = createStubCanvas();

    const viewport: Viewport = { scrollX: 0, scrollY: 0, width: 500, height: 500, dpr: 1, zoom: 2 };

    const startObject: DrawingObject = {
      id: 1,
      kind: { type: "shape" },
      anchor: {
        type: "absolute",
        pos: { xEmu: 0, yEmu: 0 },
        size: { cx: pxToEmu(100), cy: pxToEmu(100) },
      },
      zOrder: 0,
    };

    let latest: DrawingObject[] | null = null;
    const controller = new DrawingInteractionController(canvas, geom, {
      getViewport: () => viewport,
      getObjects: () => [startObject],
      setObjects: (next) => {
        latest = next;
      },
    });

    const stopFns = { preventDefault: () => {}, stopPropagation: () => {}, stopImmediatePropagation: () => {} };

    // Click somewhere away from resize handles.
    listeners.get("pointerdown")?.({ clientX: 20, clientY: 20, pointerId: 1, ...stopFns } as any);
    // Drag 10 screen pixels to the right at zoom=2 => 5 base pixels worth of EMU.
    listeners.get("pointermove")?.({ clientX: 30, clientY: 20, pointerId: 1, ...stopFns } as any);

    expect(latest).not.toBeNull();
    const moved = latest![0]!;
    expect(moved.anchor.type).toBe("absolute");
    if (moved.anchor.type === "absolute") {
      expect(moved.anchor.pos.xEmu).toBeCloseTo(pxToEmu(5), 6);
      expect(moved.anchor.pos.yEmu).toBeCloseTo(0, 6);
    }

    controller.dispose();
  });
});
