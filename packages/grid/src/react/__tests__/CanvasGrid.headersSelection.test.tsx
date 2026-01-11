// @vitest-environment jsdom
import React, { act } from "react";
import { createRoot } from "react-dom/client";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import type { CellProvider } from "../../model/CellProvider";
import { CanvasGrid, type GridApi } from "../CanvasGrid";

// React 18 relies on this flag to suppress act() warnings in test runners.
// eslint-disable-next-line @typescript-eslint/no-explicit-any
(globalThis as any).IS_REACT_ACT_ENVIRONMENT = true;

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

describe("CanvasGrid header selection", () => {
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

  it("selects entire columns/rows/all when clicking headers", async () => {
    const apiRef = React.createRef<GridApi>();

    const host = document.createElement("div");
    document.body.appendChild(host);

    const root = createRoot(host);
    await act(async () => {
      root.render(
        <CanvasGrid
          provider={provider}
          rowCount={10}
          colCount={10}
          headerRows={1}
          headerCols={1}
          defaultRowHeight={10}
          defaultColWidth={10}
          apiRef={apiRef}
        />
      );
    });

    const canvases = host.querySelectorAll("canvas");
    expect(canvases.length).toBe(3);
    const selectionCanvas = canvases[2] as HTMLCanvasElement;

    // Click column header (row 0, col 2).
    await act(async () => {
      selectionCanvas.dispatchEvent(createPointerEvent("pointerdown", { clientX: 25, clientY: 5, pointerId: 1 }));
      selectionCanvas.dispatchEvent(createPointerEvent("pointerup", { clientX: 25, clientY: 5, pointerId: 1 }));
    });
    expect(apiRef.current?.getSelectionRange()).toEqual({ startRow: 1, endRow: 10, startCol: 2, endCol: 3 });

    // Click row header (row 3, col 0).
    await act(async () => {
      selectionCanvas.dispatchEvent(createPointerEvent("pointerdown", { clientX: 5, clientY: 35, pointerId: 2 }));
      selectionCanvas.dispatchEvent(createPointerEvent("pointerup", { clientX: 5, clientY: 35, pointerId: 2 }));
    });
    expect(apiRef.current?.getSelectionRange()).toEqual({ startRow: 3, endRow: 4, startCol: 1, endCol: 10 });

    // Click corner (row 0, col 0).
    await act(async () => {
      selectionCanvas.dispatchEvent(createPointerEvent("pointerdown", { clientX: 5, clientY: 5, pointerId: 3 }));
      selectionCanvas.dispatchEvent(createPointerEvent("pointerup", { clientX: 5, clientY: 5, pointerId: 3 }));
    });
    expect(apiRef.current?.getSelectionRange()).toEqual({ startRow: 1, endRow: 10, startCol: 1, endCol: 10 });

    await act(async () => {
      root.unmount();
    });
    host.remove();
  });
});

