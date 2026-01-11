// @vitest-environment jsdom
import React, { act } from "react";
import { createRoot } from "react-dom/client";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { CanvasGrid, type GridApi } from "../CanvasGrid";

// React 18 relies on this flag to suppress act() warnings in test runners.
// eslint-disable-next-line @typescript-eslint/no-explicit-any
(globalThis as any).IS_REACT_ACT_ENVIRONMENT = true;

function createPointerEvent(
  type: string,
  options: {
    clientX: number;
    clientY: number;
    pointerId: number;
    ctrlKey?: boolean;
    shiftKey?: boolean;
    pointerType?: string;
  }
): Event {
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
  Object.defineProperty(event, "pointerType", { value: options.pointerType ?? "mouse" });
  return event;
}

describe("CanvasGrid header selection", () => {
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

  it("selects rows/cols/all when clicking header cells", async () => {
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
        <CanvasGrid
          provider={{ getCell: () => null }}
          rowCount={20}
          colCount={20}
          headerRows={1}
          headerCols={1}
          frozenRows={1}
          frozenCols={1}
          defaultRowHeight={10}
          defaultColWidth={10}
          apiRef={apiRef}
        />
      );
    });

    const selectionCanvas = host.querySelectorAll("canvas")[2] as HTMLCanvasElement;
    expect(selectionCanvas).toBeTruthy();

    // Click the top-left corner header: select all data cells (excluding headers).
    await act(async () => {
      selectionCanvas.dispatchEvent(createPointerEvent("pointerdown", { clientX: 5, clientY: 5, pointerId: 1 }));
      selectionCanvas.dispatchEvent(createPointerEvent("pointerup", { clientX: 5, clientY: 5, pointerId: 1 }));
    });

    expect(apiRef.current?.getSelectionRange()).toEqual({ startRow: 1, endRow: 20, startCol: 1, endCol: 20 });

    // Click a column header (row 0, col 3): select entire column (excluding header row).
    await act(async () => {
      selectionCanvas.dispatchEvent(createPointerEvent("pointerdown", { clientX: 35, clientY: 5, pointerId: 2 }));
      selectionCanvas.dispatchEvent(createPointerEvent("pointerup", { clientX: 35, clientY: 5, pointerId: 2 }));
    });

    expect(apiRef.current?.getSelectionRange()).toEqual({ startRow: 1, endRow: 20, startCol: 3, endCol: 4 });

    // Ctrl+click another column header to add a second column selection.
    await act(async () => {
      selectionCanvas.dispatchEvent(
        createPointerEvent("pointerdown", { clientX: 55, clientY: 5, pointerId: 3, ctrlKey: true })
      );
      selectionCanvas.dispatchEvent(
        createPointerEvent("pointerup", { clientX: 55, clientY: 5, pointerId: 3, ctrlKey: true })
      );
    });

    expect(apiRef.current?.getSelectionRanges()).toEqual([
      { startRow: 1, endRow: 20, startCol: 3, endCol: 4 },
      { startRow: 1, endRow: 20, startCol: 5, endCol: 6 }
    ]);
    expect(apiRef.current?.getActiveSelectionRangeIndex()).toBe(1);

    // Click a row header (row 5, col 0): select entire row (excluding header col).
    await act(async () => {
      selectionCanvas.dispatchEvent(createPointerEvent("pointerdown", { clientX: 5, clientY: 55, pointerId: 4 }));
      selectionCanvas.dispatchEvent(createPointerEvent("pointerup", { clientX: 5, clientY: 55, pointerId: 4 }));
    });

    expect(apiRef.current?.getSelectionRange()).toEqual({ startRow: 5, endRow: 6, startCol: 1, endCol: 20 });

    // Shift+clicking headers extends across full rows/cols.
    await act(async () => {
      apiRef.current?.setSelection(10, 4);
    });

    await act(async () => {
      selectionCanvas.dispatchEvent(
        createPointerEvent("pointerdown", { clientX: 65, clientY: 5, pointerId: 5, shiftKey: true })
      );
      selectionCanvas.dispatchEvent(
        createPointerEvent("pointerup", { clientX: 65, clientY: 5, pointerId: 5, shiftKey: true })
      );
    });

    // Selected columns 4..6 (inclusive) across all rows, excluding the header row.
    expect(apiRef.current?.getSelectionRange()).toEqual({ startRow: 1, endRow: 20, startCol: 4, endCol: 7 });

    await act(async () => {
      selectionCanvas.dispatchEvent(
        createPointerEvent("pointerdown", { clientX: 5, clientY: 125, pointerId: 6, shiftKey: true })
      );
      selectionCanvas.dispatchEvent(
        createPointerEvent("pointerup", { clientX: 5, clientY: 125, pointerId: 6, shiftKey: true })
      );
    });

    // Selected rows 10..12 (inclusive) across all cols, excluding the header col.
    expect(apiRef.current?.getSelectionRange()).toEqual({ startRow: 10, endRow: 13, startCol: 1, endCol: 20 });

    await act(async () => {
      root.unmount();
    });
    host.remove();
  });

  it("selects rows/cols/all when tapping header cells on touch", async () => {
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
        <CanvasGrid
          provider={{ getCell: () => null }}
          rowCount={20}
          colCount={20}
          headerRows={1}
          headerCols={1}
          frozenRows={1}
          frozenCols={1}
          defaultRowHeight={10}
          defaultColWidth={10}
          apiRef={apiRef}
        />
      );
    });

    const selectionCanvas = host.querySelectorAll("canvas")[2] as HTMLCanvasElement;
    expect(selectionCanvas).toBeTruthy();

    // Tap the top-left corner header: select all data cells (excluding headers).
    await act(async () => {
      selectionCanvas.dispatchEvent(createPointerEvent("pointerdown", { clientX: 5, clientY: 5, pointerId: 1, pointerType: "touch" }));
      selectionCanvas.dispatchEvent(createPointerEvent("pointerup", { clientX: 5, clientY: 5, pointerId: 1, pointerType: "touch" }));
    });
    expect(apiRef.current?.getSelectionRange()).toEqual({ startRow: 1, endRow: 20, startCol: 1, endCol: 20 });

    // Tap a column header (row 0, col 3): select entire column (excluding header row).
    await act(async () => {
      selectionCanvas.dispatchEvent(createPointerEvent("pointerdown", { clientX: 35, clientY: 5, pointerId: 2, pointerType: "touch" }));
      selectionCanvas.dispatchEvent(createPointerEvent("pointerup", { clientX: 35, clientY: 5, pointerId: 2, pointerType: "touch" }));
    });
    expect(apiRef.current?.getSelectionRange()).toEqual({ startRow: 1, endRow: 20, startCol: 3, endCol: 4 });

    // Tap a row header (row 5, col 0): select entire row (excluding header col).
    await act(async () => {
      selectionCanvas.dispatchEvent(createPointerEvent("pointerdown", { clientX: 5, clientY: 55, pointerId: 3, pointerType: "touch" }));
      selectionCanvas.dispatchEvent(createPointerEvent("pointerup", { clientX: 5, clientY: 55, pointerId: 3, pointerType: "touch" }));
    });
    expect(apiRef.current?.getSelectionRange()).toEqual({ startRow: 5, endRow: 6, startCol: 1, endCol: 20 });

    await act(async () => {
      root.unmount();
    });
    host.remove();
  });
});
