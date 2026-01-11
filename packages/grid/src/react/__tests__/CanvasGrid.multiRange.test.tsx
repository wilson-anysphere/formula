// @vitest-environment jsdom
import React, { act } from "react";
import { createRoot } from "react-dom/client";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import type { CellProvider } from "../../model/CellProvider";
import { CanvasGrid, type GridApi } from "../CanvasGrid";

// React 18 relies on this flag to suppress act() warnings in test runners.
// eslint-disable-next-line @typescript-eslint/no-explicit-any
(globalThis as any).IS_REACT_ACT_ENVIRONMENT = true;

function createPointerEvent(type: string, options: { clientX: number; clientY: number; pointerId: number; ctrlKey?: boolean; shiftKey?: boolean }): Event {
  const PointerEventCtor = (window as unknown as { PointerEvent?: typeof PointerEvent }).PointerEvent;
  if (PointerEventCtor) {
    return new PointerEventCtor(type, {
      bubbles: true,
      cancelable: true,
      clientX: options.clientX,
      clientY: options.clientY,
      buttons: 1,
      pointerId: options.pointerId,
      ctrlKey: options.ctrlKey,
      shiftKey: options.shiftKey
    } as PointerEventInit);
  }

  const event = new MouseEvent(type, {
    bubbles: true,
    cancelable: true,
    clientX: options.clientX,
    clientY: options.clientY,
    buttons: 1,
    ctrlKey: options.ctrlKey,
    shiftKey: options.shiftKey
  });
  Object.defineProperty(event, "pointerId", { value: options.pointerId });
  return event;
}

describe("CanvasGrid multi-range selection", () => {
  const provider: CellProvider = {
    getCell: () => null
  };

  beforeEach(() => {
    vi.stubGlobal(
      "ResizeObserver",
      class ResizeObserver {
        observe(): void {}
        unobserve(): void {}
        disconnect(): void {}
      }
    );

    // Avoid running full render frames; these tests only validate selection state.
    vi.stubGlobal("requestAnimationFrame", vi.fn((_cb: FrameRequestCallback) => 0));

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

  it("adds a new range on Ctrl+click and preserves existing ranges", async () => {
    vi.spyOn(HTMLElement.prototype, "getBoundingClientRect").mockReturnValue({
      x: 0,
      y: 0,
      top: 0,
      left: 0,
      bottom: 200,
      right: 200,
      width: 200,
      height: 200,
      toJSON: () => ({})
    } as unknown as DOMRect);

    const apiRef = React.createRef<GridApi>();

    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(
        <CanvasGrid provider={provider} rowCount={50} colCount={50} defaultRowHeight={10} defaultColWidth={10} apiRef={apiRef} />
      );
    });

    const selectionCanvas = host.querySelectorAll("canvas")[2] as HTMLCanvasElement;
    expect(selectionCanvas).toBeTruthy();

    // Click cell (4, 4).
    await act(async () => {
      selectionCanvas.dispatchEvent(createPointerEvent("pointerdown", { clientX: 45, clientY: 45, pointerId: 1 }));
      selectionCanvas.dispatchEvent(createPointerEvent("pointerup", { clientX: 45, clientY: 45, pointerId: 1 }));
    });

    expect(apiRef.current?.getSelectionRanges()).toEqual([{ startRow: 4, endRow: 5, startCol: 4, endCol: 5 }]);
    expect(apiRef.current?.getActiveSelectionRangeIndex()).toBe(0);

    // Ctrl+click cell (6, 7) to add a second 1x1 selection.
    await act(async () => {
      selectionCanvas.dispatchEvent(
        createPointerEvent("pointerdown", { clientX: 75, clientY: 65, pointerId: 2, ctrlKey: true })
      );
      selectionCanvas.dispatchEvent(createPointerEvent("pointerup", { clientX: 75, clientY: 65, pointerId: 2, ctrlKey: true }));
    });

    expect(apiRef.current?.getSelectionRanges()).toEqual([
      { startRow: 4, endRow: 5, startCol: 4, endCol: 5 },
      { startRow: 6, endRow: 7, startCol: 7, endCol: 8 }
    ]);
    expect(apiRef.current?.getActiveSelectionRangeIndex()).toBe(1);
    expect(apiRef.current?.getSelectionRange()).toEqual({ startRow: 6, endRow: 7, startCol: 7, endCol: 8 });

    // Shift+click extends the active range without clearing the previous range.
    await act(async () => {
      selectionCanvas.dispatchEvent(
        createPointerEvent("pointerdown", { clientX: 95, clientY: 85, pointerId: 3, shiftKey: true })
      );
      selectionCanvas.dispatchEvent(createPointerEvent("pointerup", { clientX: 95, clientY: 85, pointerId: 3, shiftKey: true }));
    });

    expect(apiRef.current?.getSelectionRanges()).toEqual([
      { startRow: 4, endRow: 5, startCol: 4, endCol: 5 },
      { startRow: 6, endRow: 9, startCol: 7, endCol: 10 }
    ]);

    await act(async () => {
      root.unmount();
    });
    host.remove();
  });

  it("scrollToCell positions the requested cell within the viewport", async () => {
    const viewportWidth = 50;
    const viewportHeight = 50;

    vi.spyOn(HTMLElement.prototype, "getBoundingClientRect").mockReturnValue({
      x: 0,
      y: 0,
      top: 0,
      left: 0,
      bottom: viewportHeight,
      right: viewportWidth,
      width: viewportWidth,
      height: viewportHeight,
      toJSON: () => ({})
    } as unknown as DOMRect);

    const apiRef = React.createRef<GridApi>();

    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(
        <CanvasGrid provider={provider} rowCount={200} colCount={200} defaultRowHeight={10} defaultColWidth={10} apiRef={apiRef} />
      );
    });

    apiRef.current?.scrollToCell(50, 50, { align: "auto" });

    const rect = apiRef.current?.getCellRect(50, 50);
    expect(rect).not.toBeNull();
    expect(rect!.x).toBeGreaterThanOrEqual(0);
    expect(rect!.y).toBeGreaterThanOrEqual(0);
    expect(rect!.x + rect!.width).toBeLessThanOrEqual(viewportWidth);
    expect(rect!.y + rect!.height).toBeLessThanOrEqual(viewportHeight);

    await act(async () => {
      root.unmount();
    });
    host.remove();
  });
});

