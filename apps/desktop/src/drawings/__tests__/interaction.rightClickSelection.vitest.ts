/**
 * @vitest-environment jsdom
 */

import { describe, expect, it, vi } from "vitest";

import { DrawingInteractionController } from "../interaction";
import { pxToEmu, type GridGeometry, type Viewport } from "../overlay";
import type { DrawingObject } from "../types";

function createPointerLikeMouseEvent(
  type: string,
  options: { offsetX: number; offsetY: number; button: number; pointerId?: number; pointerType?: string },
): MouseEvent {
  const event = new MouseEvent(type, { bubbles: true, cancelable: true, button: options.button });
  Object.defineProperty(event, "pointerId", { configurable: true, value: options.pointerId ?? 1 });
  Object.defineProperty(event, "pointerType", { configurable: true, value: options.pointerType ?? "mouse" });
  Object.defineProperty(event, "offsetX", { configurable: true, value: options.offsetX });
  Object.defineProperty(event, "offsetY", { configurable: true, value: options.offsetY });
  return event;
}

describe("DrawingInteractionController mouse right-click", () => {
  it("selects the drawing but does not drag/resize or stop propagation", () => {
    const canvas = document.createElement("canvas");
    const container = document.createElement("div");
    container.appendChild(canvas);
    document.body.appendChild(container);

    const viewport: Viewport = { scrollX: 0, scrollY: 0, width: 500, height: 500, dpr: 1 };
    const geom: GridGeometry = {
      cellOriginPx: () => ({ x: 0, y: 0 }),
      cellSizePx: () => ({ width: 0, height: 0 }),
    };

    const initialObjects: DrawingObject[] = [
      {
        id: 1,
        kind: { type: "image", imageId: "img_1" },
        anchor: {
          type: "absolute",
          pos: { xEmu: pxToEmu(0), yEmu: pxToEmu(0) },
          size: { cx: pxToEmu(100), cy: pxToEmu(100) },
        },
        zOrder: 0,
      },
    ];

    let objects = initialObjects;
    let selectedId: number | null = null;
    const setObjects = vi.fn((next: DrawingObject[]) => {
      objects = next;
    });

    const controller = new DrawingInteractionController(canvas, geom, {
      getViewport: () => viewport,
      getObjects: () => objects,
      setObjects,
      onSelectionChange: (id) => {
        selectedId = id;
      },
    });

    const bubbled = vi.fn();
    container.addEventListener("pointerdown", bubbled);

    const down = createPointerLikeMouseEvent("pointerdown", { offsetX: 10, offsetY: 10, button: 2 });
    canvas.dispatchEvent(down);

    expect(selectedId).toBe(1);
    expect(down.defaultPrevented).toBe(false);
    expect(bubbled).toHaveBeenCalledTimes(1);

    // Ensure we did not enter a drag/resize state.
    canvas.dispatchEvent(createPointerLikeMouseEvent("pointermove", { offsetX: 50, offsetY: 50, button: 2 }));
    expect(setObjects).not.toHaveBeenCalled();
    expect(objects).toEqual(initialObjects);

    controller.dispose();
    container.remove();
  });
});

