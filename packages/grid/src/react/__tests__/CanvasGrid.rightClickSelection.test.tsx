// @vitest-environment jsdom
import React, { act } from "react";
import { createRoot } from "react-dom/client";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { CanvasGrid, MockCellProvider, type GridApi } from "../../index";

// React 18 relies on this flag to suppress act() warnings in test runners.
// eslint-disable-next-line @typescript-eslint/no-explicit-any
(globalThis as any).IS_REACT_ACT_ENVIRONMENT = true;

function createMock2dContext(canvas: HTMLCanvasElement): CanvasRenderingContext2D {
  const noop = () => {};
  const gradient = { addColorStop: noop } as any;

  const context = new Proxy(
    {
      canvas,
      measureText: (text: string) =>
        ({
          width: text.length * 8,
          actualBoundingBoxAscent: 8,
          actualBoundingBoxDescent: 2
        }) as TextMetrics,
      createLinearGradient: () => gradient,
      createPattern: () => null,
      getImageData: () => ({ data: new Uint8ClampedArray(), width: 0, height: 0 }),
      putImageData: noop
    },
    {
      get(target, prop) {
        if (prop in target) return (target as any)[prop];
        return noop;
      },
      set(target, prop, value) {
        (target as any)[prop] = value;
        return true;
      }
    }
  );
  return context as any;
}

function createPointerLikeEvent(
  type: string,
  options: {
    clientX: number;
    clientY: number;
    button: number;
    pointerId: number;
    ctrlKey?: boolean;
    metaKey?: boolean;
    shiftKey?: boolean;
  }
): MouseEvent {
  const event = new MouseEvent(type, {
    bubbles: true,
    cancelable: true,
    clientX: options.clientX,
    clientY: options.clientY,
    button: options.button,
    ctrlKey: options.ctrlKey,
    metaKey: options.metaKey,
    shiftKey: options.shiftKey
  });
  Object.defineProperty(event, "pointerId", { configurable: true, value: options.pointerId });
  return event;
}

describe("CanvasGrid right-click selection semantics", () => {
  const originalGetContext = HTMLCanvasElement.prototype.getContext;
  const originalPlatform = navigator.platform;

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
    vi.stubGlobal("cancelAnimationFrame", vi.fn());

    vi.spyOn(HTMLElement.prototype, "getBoundingClientRect").mockReturnValue({
      left: 0,
      top: 0,
      right: 400,
      bottom: 200,
      width: 400,
      height: 200,
      x: 0,
      y: 0,
      toJSON: () => ({})
    } as unknown as DOMRect);

    HTMLCanvasElement.prototype.getContext = vi.fn(function (this: HTMLCanvasElement) {
      return createMock2dContext(this);
    }) as unknown as typeof HTMLCanvasElement.prototype.getContext;
  });

  afterEach(() => {
    HTMLCanvasElement.prototype.getContext = originalGetContext;
    vi.restoreAllMocks();
    vi.unstubAllGlobals();

    try {
      Object.defineProperty(navigator, "platform", { configurable: true, value: originalPlatform });
    } catch {
      // ignore
    }
  });

  it("does not change selection/range when right-clicking inside an existing selection", async () => {
    const provider = new MockCellProvider({ rowCount: 100, colCount: 100 });
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
          rowCount={100}
          colCount={100}
          frozenRows={1}
          frozenCols={1}
          headerRows={1}
          headerCols={1}
          defaultRowHeight={24}
          defaultColWidth={100}
          enableResize={false}
          apiRef={apiRef}
          onSelectionChange={onSelectionChange}
          onSelectionRangeChange={onSelectionRangeChange}
        />
      );
    });

    const gridContainer = host.querySelector('[data-testid="canvas-grid"]') as HTMLDivElement;
    const selectionCanvas = host.querySelectorAll("canvas")[2] as HTMLCanvasElement;

    await act(async () => {
      apiRef.current?.setColWidth(0, 48);
      apiRef.current?.setRowHeight(0, 24);
      apiRef.current?.setSelectionRanges([{ startRow: 1, endRow: 3, startCol: 1, endCol: 3 }]);
    });

    onSelectionChange.mockClear();
    onSelectionRangeChange.mockClear();

    const outsideFocusTarget = document.createElement("button");
    outsideFocusTarget.textContent = "outside";
    document.body.appendChild(outsideFocusTarget);
    outsideFocusTarget.focus();
    expect(document.activeElement).toBe(outsideFocusTarget);

    // Right click B2 (within selection). This should not move the active cell.
    await act(async () => {
      selectionCanvas.dispatchEvent(
        createPointerLikeEvent("pointerdown", {
          clientX: 48 + 100 + 10,
          clientY: 24 + 24 + 10,
          button: 2,
          pointerId: 1
        })
      );
    });

    expect(apiRef.current?.getSelection()).toEqual({ row: 1, col: 1 });
    expect(apiRef.current?.getSelectionRanges()).toEqual([{ startRow: 1, endRow: 3, startCol: 1, endCol: 3 }]);
    expect(onSelectionChange).not.toHaveBeenCalled();
    expect(onSelectionRangeChange).not.toHaveBeenCalled();
    expect(document.activeElement).toBe(gridContainer);

    await act(async () => {
      root.unmount();
    });
    host.remove();
    outsideFocusTarget.remove();
  });

  it("moves the active cell when right-clicking outside the selection", async () => {
    const provider = new MockCellProvider({ rowCount: 100, colCount: 100 });
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
          rowCount={100}
          colCount={100}
          frozenRows={1}
          frozenCols={1}
          headerRows={1}
          headerCols={1}
          defaultRowHeight={24}
          defaultColWidth={100}
          enableResize={false}
          apiRef={apiRef}
          onSelectionChange={onSelectionChange}
          onSelectionRangeChange={onSelectionRangeChange}
        />
      );
    });

    const gridContainer = host.querySelector('[data-testid="canvas-grid"]') as HTMLDivElement;
    const selectionCanvas = host.querySelectorAll("canvas")[2] as HTMLCanvasElement;

    await act(async () => {
      apiRef.current?.setColWidth(0, 48);
      apiRef.current?.setRowHeight(0, 24);
      apiRef.current?.setSelectionRanges([{ startRow: 1, endRow: 3, startCol: 1, endCol: 3 }]);
    });

    onSelectionChange.mockClear();
    onSelectionRangeChange.mockClear();

    // Right click D4 (outside selection). This should collapse selection to D4.
    await act(async () => {
      selectionCanvas.dispatchEvent(
        createPointerLikeEvent("pointerdown", {
          clientX: 48 + 3 * 100 + 10,
          clientY: 24 + 3 * 24 + 10,
          button: 2,
          pointerId: 1
        })
      );
    });

    expect(apiRef.current?.getSelection()).toEqual({ row: 4, col: 4 });
    expect(apiRef.current?.getSelectionRanges()).toEqual([{ startRow: 4, endRow: 5, startCol: 4, endCol: 5 }]);
    expect(onSelectionChange).toHaveBeenCalledTimes(1);
    expect(onSelectionChange).toHaveBeenLastCalledWith({ row: 4, col: 4 });
    expect(onSelectionRangeChange).toHaveBeenCalledTimes(1);
    expect(onSelectionRangeChange).toHaveBeenLastCalledWith({ startRow: 4, endRow: 5, startCol: 4, endCol: 5 });
    expect(document.activeElement).toBe(gridContainer);

    await act(async () => {
      root.unmount();
    });
    host.remove();
  });

  it("treats Ctrl+click as a context-click on macOS (does not add to the selection)", async () => {
    try {
      Object.defineProperty(navigator, "platform", { configurable: true, value: "MacIntel" });
    } catch {
      // If the runtime doesn't allow stubbing `navigator.platform`, skip the test.
      return;
    }

    const provider = new MockCellProvider({ rowCount: 100, colCount: 100 });
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
          rowCount={100}
          colCount={100}
          frozenRows={1}
          frozenCols={1}
          headerRows={1}
          headerCols={1}
          defaultRowHeight={24}
          defaultColWidth={100}
          enableResize={false}
          apiRef={apiRef}
          onSelectionChange={onSelectionChange}
          onSelectionRangeChange={onSelectionRangeChange}
        />
      );
    });

    const selectionCanvas = host.querySelectorAll("canvas")[2] as HTMLCanvasElement;

    await act(async () => {
      apiRef.current?.setColWidth(0, 48);
      apiRef.current?.setRowHeight(0, 24);
      apiRef.current?.setSelectionRanges([{ startRow: 1, endRow: 3, startCol: 1, endCol: 3 }]);
    });

    onSelectionChange.mockClear();
    onSelectionRangeChange.mockClear();

    // Ctrl+click B2 (within selection) should keep the selection as-is.
    await act(async () => {
      selectionCanvas.dispatchEvent(
        createPointerLikeEvent("pointerdown", {
          clientX: 48 + 100 + 10,
          clientY: 24 + 24 + 10,
          button: 0,
          pointerId: 1,
          ctrlKey: true
        })
      );
    });

    expect(apiRef.current?.getSelection()).toEqual({ row: 1, col: 1 });
    expect(apiRef.current?.getSelectionRanges()).toEqual([{ startRow: 1, endRow: 3, startCol: 1, endCol: 3 }]);
    expect(onSelectionChange).not.toHaveBeenCalled();
    expect(onSelectionRangeChange).not.toHaveBeenCalled();

    // Ctrl+click D4 (outside) should collapse selection to D4 (not additive).
    await act(async () => {
      selectionCanvas.dispatchEvent(
        createPointerLikeEvent("pointerdown", {
          clientX: 48 + 3 * 100 + 10,
          clientY: 24 + 3 * 24 + 10,
          button: 0,
          pointerId: 2,
          ctrlKey: true
        })
      );
    });

    expect(apiRef.current?.getSelection()).toEqual({ row: 4, col: 4 });
    expect(apiRef.current?.getSelectionRanges()).toEqual([{ startRow: 4, endRow: 5, startCol: 4, endCol: 5 }]);
    expect(onSelectionChange).toHaveBeenCalledTimes(1);
    expect(onSelectionChange).toHaveBeenLastCalledWith({ row: 4, col: 4 });
    expect(onSelectionRangeChange).toHaveBeenCalledTimes(1);
    expect(onSelectionRangeChange).toHaveBeenLastCalledWith({ startRow: 4, endRow: 5, startCol: 4, endCol: 5 });

    await act(async () => {
      root.unmount();
    });
    host.remove();
  });
});

