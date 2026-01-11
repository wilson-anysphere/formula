// @vitest-environment jsdom
import React, { act } from "react";
import { createRoot } from "react-dom/client";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { CanvasGrid, type GridApi } from "../CanvasGrid";

// React 18 relies on this flag to suppress act() warnings in test runners.
// eslint-disable-next-line @typescript-eslint/no-explicit-any
(globalThis as any).IS_REACT_ACT_ENVIRONMENT = true;

function createTouchPointerEvent(type: string, options: { clientX: number; clientY: number; pointerId: number }): Event {
  const event = new MouseEvent(type, { bubbles: true, cancelable: true, clientX: options.clientX, clientY: options.clientY });
  Object.defineProperty(event, "pointerId", { value: options.pointerId });
  Object.defineProperty(event, "pointerType", { value: "touch" });
  return event;
}

describe("CanvasGrid touch interactions", () => {
  beforeEach(() => {
    vi.stubGlobal(
      "ResizeObserver",
      class ResizeObserver {
        observe(): void {}
        unobserve(): void {}
        disconnect(): void {}
      }
    );

    // Avoid running full render frames; these tests only validate scroll/zoom state.
    vi.stubGlobal("requestAnimationFrame", vi.fn((_cb: FrameRequestCallback) => 0));

    vi.spyOn(HTMLElement.prototype, "getBoundingClientRect").mockReturnValue({
      left: 0,
      top: 0,
      right: 200,
      bottom: 200,
      width: 200,
      height: 200,
      x: 0,
      y: 0,
      toJSON: () => ({})
    } as unknown as DOMRect);

    const ctxStub: Partial<CanvasRenderingContext2D> = {
      setTransform: vi.fn(),
      measureText: (text: string) =>
        ({
          width: text.length * 6,
          actualBoundingBoxAscent: 8,
          actualBoundingBoxDescent: 2
        }) as TextMetrics
    };

    vi.spyOn(HTMLCanvasElement.prototype, "getContext").mockImplementation(
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      () => ctxStub as any
    );
  });

  afterEach(() => {
    vi.restoreAllMocks();
    vi.unstubAllGlobals();
  });

  it("pans the grid on single-finger drag and does not start selection", async () => {
    const apiRef = React.createRef<GridApi>();

    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(
        <CanvasGrid provider={{ getCell: () => null }} rowCount={100} colCount={10} defaultRowHeight={10} defaultColWidth={10} apiRef={apiRef} />
      );
    });

    const selectionCanvas = host.querySelectorAll("canvas")[2] as HTMLCanvasElement;
    expect(selectionCanvas).toBeTruthy();

    expect(apiRef.current?.getScroll().y).toBe(0);

    // Drag finger upward to scroll down.
    await act(async () => {
      selectionCanvas.dispatchEvent(createTouchPointerEvent("pointerdown", { clientX: 50, clientY: 100, pointerId: 1 }));
      selectionCanvas.dispatchEvent(createTouchPointerEvent("pointermove", { clientX: 50, clientY: 50, pointerId: 1 }));
      selectionCanvas.dispatchEvent(createTouchPointerEvent("pointerup", { clientX: 50, clientY: 50, pointerId: 1 }));
    });

    expect(apiRef.current?.getScroll().y).toBeGreaterThan(0);
    expect(apiRef.current?.getSelection()).toBeNull();

    await act(async () => {
      root.unmount();
    });
    host.remove();
  });

  it("pinch zooms the grid with two touch pointers", async () => {
    const apiRef = React.createRef<GridApi>();

    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(
        <CanvasGrid provider={{ getCell: () => null }} rowCount={100} colCount={10} defaultRowHeight={10} defaultColWidth={10} apiRef={apiRef} />
      );
    });

    const selectionCanvas = host.querySelectorAll("canvas")[2] as HTMLCanvasElement;
    expect(selectionCanvas).toBeTruthy();

    expect(apiRef.current?.getZoom()).toBe(1);

    await act(async () => {
      selectionCanvas.dispatchEvent(createTouchPointerEvent("pointerdown", { clientX: 100, clientY: 100, pointerId: 1 }));
      selectionCanvas.dispatchEvent(createTouchPointerEvent("pointerdown", { clientX: 150, clientY: 100, pointerId: 2 }));
      // Increase pinch distance from 50px to 100px -> zoom 2x.
      selectionCanvas.dispatchEvent(createTouchPointerEvent("pointermove", { clientX: 200, clientY: 100, pointerId: 2 }));
      selectionCanvas.dispatchEvent(createTouchPointerEvent("pointerup", { clientX: 200, clientY: 100, pointerId: 2 }));
      selectionCanvas.dispatchEvent(createTouchPointerEvent("pointerup", { clientX: 100, clientY: 100, pointerId: 1 }));
    });

    expect(apiRef.current?.getZoom()).toBeCloseTo(2, 5);
    expect(apiRef.current?.getSelection()).toBeNull();

    await act(async () => {
      root.unmount();
    });
    host.remove();
  });
});

