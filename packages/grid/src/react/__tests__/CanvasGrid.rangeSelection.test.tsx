// @vitest-environment jsdom
import React from "react";
import { createRoot } from "react-dom/client";
import { act } from "react-dom/test-utils";
import { describe, expect, it, vi, beforeEach, afterEach } from "vitest";

import type { CellProvider } from "../../model/CellProvider";
import { CanvasGrid, type GridApi } from "../CanvasGrid";

function createMockContext(): CanvasRenderingContext2D {
  const noop = () => {};

  // A deliberately minimal mock; this test only verifies selection plumbing.
  return {
    canvas: document.createElement("canvas"),
    save: noop,
    restore: noop,
    beginPath: noop,
    rect: noop,
    clip: noop,
    fillRect: noop,
    clearRect: noop,
    strokeRect: noop,
    drawImage: noop,
    fillText: noop,
    moveTo: noop,
    lineTo: noop,
    stroke: noop,
    fill: noop,
    closePath: noop,
    translate: noop,
    rotate: noop,
    setTransform: noop,
    measureText: (text: string) =>
      ({
        width: text.length * 6,
        actualBoundingBoxAscent: 8,
        actualBoundingBoxDescent: 2
      }) as TextMetrics,
    imageSmoothingEnabled: false
  } as unknown as CanvasRenderingContext2D;
}

function createPointerEvent(type: string, options: { clientX: number; clientY: number; pointerId: number }): Event {
  const PointerEventCtor = (window as unknown as { PointerEvent?: typeof PointerEvent }).PointerEvent;
  if (PointerEventCtor) {
    return new PointerEventCtor(type, {
      bubbles: true,
      cancelable: true,
      clientX: options.clientX,
      clientY: options.clientY,
      buttons: 1,
      pointerId: options.pointerId
    } as PointerEventInit);
  }

  const event = new MouseEvent(type, {
    bubbles: true,
    cancelable: true,
    clientX: options.clientX,
    clientY: options.clientY,
    buttons: 1
  });
  Object.defineProperty(event, "pointerId", { value: options.pointerId });
  return event;
}

describe("CanvasGrid range selection", () => {
  const provider: CellProvider = {
    getCell: () => null
  };

  const rect = {
    x: 0,
    y: 0,
    top: 0,
    left: 0,
    bottom: 200,
    right: 200,
    width: 200,
    height: 200,
    toJSON: () => ({})
  };

  beforeEach(() => {
    vi.spyOn(HTMLElement.prototype, "getBoundingClientRect").mockReturnValue(rect as unknown as DOMRect);
    vi.spyOn(HTMLCanvasElement.prototype, "getContext").mockImplementation(() => createMockContext());

    class ResizeObserverMock {
      observe() {}
      unobserve() {}
      disconnect() {}
    }

    (globalThis as unknown as { ResizeObserver?: typeof ResizeObserver }).ResizeObserver = ResizeObserverMock as unknown as typeof ResizeObserver;
  });

  afterEach(() => {
    vi.restoreAllMocks();
  });

  it("supports drag range selection and exposes it via callbacks/api", async () => {
    const apiRef = React.createRef<GridApi>();
    const onSelectionChange = vi.fn();
    const onSelectionRangeChange = vi.fn();

    const host = document.createElement("div");
    document.body.appendChild(host);

    const root = createRoot(host);
    await act(async () => {
      root.render(
        <CanvasGrid
          provider={provider}
          rowCount={10}
          colCount={10}
          defaultRowHeight={10}
          defaultColWidth={10}
          apiRef={apiRef}
          onSelectionChange={onSelectionChange}
          onSelectionRangeChange={onSelectionRangeChange}
        />
      );
    });

    const canvases = host.querySelectorAll("canvas");
    expect(canvases.length).toBe(3);
    const selectionCanvas = canvases[2] as HTMLCanvasElement;

    await act(async () => {
      selectionCanvas.dispatchEvent(createPointerEvent("pointerdown", { clientX: 5, clientY: 5, pointerId: 1 }));
      selectionCanvas.dispatchEvent(createPointerEvent("pointermove", { clientX: 25, clientY: 15, pointerId: 1 }));
      selectionCanvas.dispatchEvent(createPointerEvent("pointerup", { clientX: 25, clientY: 15, pointerId: 1 }));
    });

    expect(onSelectionChange).toHaveBeenCalledWith({ row: 0, col: 0 });

    const expectedRange = { startRow: 0, endRow: 2, startCol: 0, endCol: 3 };
    expect(onSelectionRangeChange).toHaveBeenCalled();
    const lastRangeCall = onSelectionRangeChange.mock.calls[onSelectionRangeChange.mock.calls.length - 1]?.[0];
    expect(lastRangeCall).toEqual(expectedRange);

    expect(apiRef.current).not.toBeNull();
    expect(apiRef.current?.getSelectionRange()).toEqual(expectedRange);

    await act(async () => {
      root.unmount();
    });
    host.remove();
  });
});
