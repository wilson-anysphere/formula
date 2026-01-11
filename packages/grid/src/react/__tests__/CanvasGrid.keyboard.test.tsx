// @vitest-environment jsdom
import React, { act } from "react";
import { createRoot } from "react-dom/client";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { CanvasGrid, type GridApi } from "../CanvasGrid";

// React 18 relies on this flag to suppress act() warnings in test runners.
// eslint-disable-next-line @typescript-eslint/no-explicit-any
(globalThis as any).IS_REACT_ACT_ENVIRONMENT = true;

function createMock2dContext(canvas: HTMLCanvasElement): CanvasRenderingContext2D {
  const noop = () => {};
  return {
    canvas,
    fillStyle: "#000",
    strokeStyle: "#000",
    lineWidth: 1,
    font: "",
    textAlign: "left",
    textBaseline: "alphabetic",
    globalAlpha: 1,
    imageSmoothingEnabled: false,
    setTransform: noop,
    clearRect: noop,
    fillRect: noop,
    strokeRect: noop,
    beginPath: noop,
    rect: noop,
    clip: noop,
    fill: noop,
    stroke: noop,
    moveTo: noop,
    lineTo: noop,
    closePath: noop,
    save: noop,
    restore: noop,
    drawImage: noop,
    translate: noop,
    rotate: noop,
    fillText: noop,
    measureText: (text: string) =>
      ({
        width: text.length * 6,
        actualBoundingBoxAscent: 8,
        actualBoundingBoxDescent: 2
      }) as TextMetrics
  } as unknown as CanvasRenderingContext2D;
}

describe("CanvasGrid keyboard navigation", () => {
  const originalGetContext = HTMLCanvasElement.prototype.getContext;

  beforeEach(() => {
    vi.stubGlobal(
      "ResizeObserver",
      class ResizeObserver {
        observe(): void {}
        unobserve(): void {}
        disconnect(): void {}
      }
    );

    vi.stubGlobal("requestAnimationFrame", (cb: FrameRequestCallback) => {
      cb(0);
      return 0;
    });

    HTMLCanvasElement.prototype.getContext = vi.fn(function (this: HTMLCanvasElement) {
      return createMock2dContext(this);
    }) as unknown as typeof HTMLCanvasElement.prototype.getContext;

    vi.spyOn(HTMLElement.prototype, "getBoundingClientRect").mockReturnValue({
      left: 0,
      top: 0,
      right: 400,
      bottom: 200,
      width: 400,
      height: 200,
      x: 0,
      y: 0,
      toJSON: () => {}
    } as DOMRect);
  });

  afterEach(() => {
    HTMLCanvasElement.prototype.getContext = originalGetContext;
    vi.restoreAllMocks();
    vi.unstubAllGlobals();
  });

  it("moves selection with arrow keys when the grid container is focused", async () => {
    const apiRef = React.createRef<GridApi>();
    const onSelectionChange = vi.fn();
    const onSelectionRangeChange = vi.fn();

    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(
        <CanvasGrid
          provider={{ getCell: (row, col) => ({ row, col, value: `${row},${col}` }) }}
          rowCount={10}
          colCount={10}
          apiRef={apiRef}
          onSelectionChange={onSelectionChange}
          onSelectionRangeChange={onSelectionRangeChange}
        />
      );
    });

    // Seed selection at A1 (row 0, col 0) and clear initial calls.
    await act(async () => {
      apiRef.current?.setSelection(0, 0);
    });
    onSelectionChange.mockClear();
    onSelectionRangeChange.mockClear();

    const container = host.querySelector('[data-testid="canvas-grid"]') as HTMLDivElement;
    expect(container).toBeTruthy();
    container.focus();

    await act(async () => {
      container.dispatchEvent(new KeyboardEvent("keydown", { key: "ArrowRight", bubbles: true, cancelable: true }));
    });

    expect(onSelectionChange).toHaveBeenCalledWith({ row: 0, col: 1 });
    expect(onSelectionRangeChange).toHaveBeenCalledWith({ startRow: 0, endRow: 1, startCol: 1, endCol: 2 });

    const status = host.querySelector('[data-testid="canvas-grid-a11y-status"]') as HTMLDivElement;
    expect(status.textContent).toContain("Active cell B1");
    expect(status.textContent).toContain("value 0,1");

    await act(async () => {
      root.unmount();
    });
    host.remove();
  });

  it("supports Excel-style selection shortcuts (Ctrl+A, Ctrl+Space, Shift+Space)", async () => {
    const apiRef = React.createRef<GridApi>();
    const onSelectionChange = vi.fn();
    const onSelectionRangeChange = vi.fn();

    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(
        <CanvasGrid
          provider={{ getCell: (row, col) => ({ row, col, value: `${row},${col}` }) }}
          rowCount={10}
          colCount={10}
          apiRef={apiRef}
          onSelectionChange={onSelectionChange}
          onSelectionRangeChange={onSelectionRangeChange}
        />
      );
    });

    await act(async () => {
      apiRef.current?.setSelection(2, 3);
    });
    onSelectionChange.mockClear();
    onSelectionRangeChange.mockClear();

    const container = host.querySelector('[data-testid="canvas-grid"]') as HTMLDivElement;
    container.focus();

    await act(async () => {
      container.dispatchEvent(
        new KeyboardEvent("keydown", { key: " ", code: "Space", ctrlKey: true, bubbles: true, cancelable: true })
      );
    });

    expect(onSelectionChange).not.toHaveBeenCalled();
    expect(onSelectionRangeChange).toHaveBeenCalledWith({ startRow: 0, endRow: 10, startCol: 3, endCol: 4 });
    expect(apiRef.current?.getSelectionRange()).toEqual({ startRow: 0, endRow: 10, startCol: 3, endCol: 4 });

    onSelectionChange.mockClear();
    onSelectionRangeChange.mockClear();

    await act(async () => {
      container.dispatchEvent(
        new KeyboardEvent("keydown", { key: " ", code: "Space", shiftKey: true, bubbles: true, cancelable: true })
      );
    });

    expect(onSelectionChange).not.toHaveBeenCalled();
    expect(onSelectionRangeChange).toHaveBeenCalledWith({ startRow: 2, endRow: 3, startCol: 0, endCol: 10 });
    expect(apiRef.current?.getSelectionRange()).toEqual({ startRow: 2, endRow: 3, startCol: 0, endCol: 10 });

    onSelectionChange.mockClear();
    onSelectionRangeChange.mockClear();

    await act(async () => {
      container.dispatchEvent(new KeyboardEvent("keydown", { key: "a", ctrlKey: true, bubbles: true, cancelable: true }));
    });

    expect(onSelectionChange).not.toHaveBeenCalled();
    expect(onSelectionRangeChange).toHaveBeenCalledWith({ startRow: 0, endRow: 10, startCol: 0, endCol: 10 });
    expect(apiRef.current?.getSelectionRange()).toEqual({ startRow: 0, endRow: 10, startCol: 0, endCol: 10 });

    await act(async () => {
      root.unmount();
    });
    host.remove();
  });

  it("moves horizontally with Alt+PageDown/PageUp", async () => {
    const apiRef = React.createRef<GridApi>();
    const onSelectionChange = vi.fn();
    const onSelectionRangeChange = vi.fn();

    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(
        <CanvasGrid
          provider={{ getCell: (row, col) => ({ row, col, value: `${row},${col}` }) }}
          rowCount={10}
          colCount={10}
          apiRef={apiRef}
          onSelectionChange={onSelectionChange}
          onSelectionRangeChange={onSelectionRangeChange}
        />
      );
    });

    await act(async () => {
      apiRef.current?.setSelection(0, 0);
    });
    onSelectionChange.mockClear();
    onSelectionRangeChange.mockClear();

    const viewport = apiRef.current?.getViewportState();
    expect(viewport).not.toBeNull();
    const pageCols = Math.max(1, (viewport?.main.cols.end ?? 0) - (viewport?.main.cols.start ?? 0));
    const expectedCol = Math.min(9, pageCols);

    const container = host.querySelector('[data-testid="canvas-grid"]') as HTMLDivElement;
    container.focus();

    await act(async () => {
      container.dispatchEvent(
        new KeyboardEvent("keydown", { key: "PageDown", altKey: true, bubbles: true, cancelable: true })
      );
    });

    expect(onSelectionChange).toHaveBeenCalledWith({ row: 0, col: expectedCol });
    expect(onSelectionRangeChange).toHaveBeenCalledWith({
      startRow: 0,
      endRow: 1,
      startCol: expectedCol,
      endCol: expectedCol + 1
    });

    onSelectionChange.mockClear();
    onSelectionRangeChange.mockClear();

    await act(async () => {
      container.dispatchEvent(
        new KeyboardEvent("keydown", { key: "PageUp", altKey: true, bubbles: true, cancelable: true })
      );
    });

    expect(onSelectionChange).toHaveBeenCalledWith({ row: 0, col: 0 });
    expect(onSelectionRangeChange).toHaveBeenCalledWith({ startRow: 0, endRow: 1, startCol: 0, endCol: 1 });

    await act(async () => {
      root.unmount();
    });
    host.remove();
  });

  it("moves within an existing selection range with Tab/Enter (including shift reverse)", async () => {
    const apiRef = React.createRef<GridApi>();
    const onSelectionChange = vi.fn();
    const onSelectionRangeChange = vi.fn();

    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(
        <CanvasGrid
          provider={{ getCell: (row, col) => ({ row, col, value: `${row},${col}` }) }}
          rowCount={10}
          colCount={10}
          apiRef={apiRef}
          onSelectionChange={onSelectionChange}
          onSelectionRangeChange={onSelectionRangeChange}
        />
      );
    });

    await act(async () => {
      apiRef.current?.setSelectionRange({ startRow: 0, endRow: 2, startCol: 0, endCol: 2 });
    });
    onSelectionChange.mockClear();
    onSelectionRangeChange.mockClear();

    const container = host.querySelector('[data-testid="canvas-grid"]') as HTMLDivElement;
    container.focus();

    await act(async () => {
      container.dispatchEvent(new KeyboardEvent("keydown", { key: "Tab", bubbles: true, cancelable: true }));
    });

    expect(apiRef.current?.getSelection()).toEqual({ row: 0, col: 1 });
    expect(apiRef.current?.getSelectionRange()).toEqual({ startRow: 0, endRow: 2, startCol: 0, endCol: 2 });
    expect(onSelectionChange).toHaveBeenCalledWith({ row: 0, col: 1 });
    expect(onSelectionRangeChange).not.toHaveBeenCalled();

    onSelectionChange.mockClear();

    await act(async () => {
      container.dispatchEvent(new KeyboardEvent("keydown", { key: "Tab", bubbles: true, cancelable: true }));
    });

    // Wrap to the next row start.
    expect(apiRef.current?.getSelection()).toEqual({ row: 1, col: 0 });
    expect(apiRef.current?.getSelectionRange()).toEqual({ startRow: 0, endRow: 2, startCol: 0, endCol: 2 });
    expect(onSelectionChange).toHaveBeenCalledWith({ row: 1, col: 0 });

    onSelectionChange.mockClear();

    await act(async () => {
      container.dispatchEvent(new KeyboardEvent("keydown", { key: "Tab", shiftKey: true, bubbles: true, cancelable: true }));
    });

    expect(apiRef.current?.getSelection()).toEqual({ row: 0, col: 1 });
    expect(apiRef.current?.getSelectionRange()).toEqual({ startRow: 0, endRow: 2, startCol: 0, endCol: 2 });
    expect(onSelectionChange).toHaveBeenCalledWith({ row: 0, col: 1 });

    onSelectionChange.mockClear();

    await act(async () => {
      container.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", bubbles: true, cancelable: true }));
    });

    expect(apiRef.current?.getSelection()).toEqual({ row: 1, col: 1 });
    expect(apiRef.current?.getSelectionRange()).toEqual({ startRow: 0, endRow: 2, startCol: 0, endCol: 2 });
    expect(onSelectionChange).toHaveBeenCalledWith({ row: 1, col: 1 });

    onSelectionChange.mockClear();

    await act(async () => {
      // Shift+Enter moves backward within the range (no selection extension).
      container.dispatchEvent(
        new KeyboardEvent("keydown", { key: "Enter", shiftKey: true, bubbles: true, cancelable: true })
      );
    });

    expect(apiRef.current?.getSelection()).toEqual({ row: 0, col: 1 });
    expect(apiRef.current?.getSelectionRange()).toEqual({ startRow: 0, endRow: 2, startCol: 0, endCol: 2 });
    expect(onSelectionChange).toHaveBeenCalledWith({ row: 0, col: 1 });
    expect(onSelectionRangeChange).not.toHaveBeenCalled();

    await act(async () => {
      root.unmount();
    });
    host.remove();
  });

  it("skips merged cell interiors when navigating within a selection range", async () => {
    const merged = { startRow: 0, endRow: 2, startCol: 0, endCol: 2 };
    const provider = {
      getCell: (row: number, col: number) => ({ row, col, value: `${row},${col}` }),
      getMergedRangeAt: (row: number, col: number) =>
        row >= merged.startRow && row < merged.endRow && col >= merged.startCol && col < merged.endCol ? merged : null,
      getMergedRangesInRange: (range: { startRow: number; endRow: number; startCol: number; endCol: number }) =>
        range.startRow < merged.endRow && range.endRow > merged.startRow && range.startCol < merged.endCol && range.endCol > merged.startCol
          ? [merged]
          : []
    };

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
          apiRef={apiRef}
          onSelectionChange={onSelectionChange}
          onSelectionRangeChange={onSelectionRangeChange}
        />
      );
    });

    await act(async () => {
      apiRef.current?.setSelectionRange({ startRow: 0, endRow: 2, startCol: 0, endCol: 3 });
    });
    onSelectionChange.mockClear();
    onSelectionRangeChange.mockClear();

    const container = host.querySelector('[data-testid="canvas-grid"]') as HTMLDivElement;
    container.focus();

    await act(async () => {
      container.dispatchEvent(new KeyboardEvent("keydown", { key: "Tab", bubbles: true, cancelable: true }));
    });

    // Skip over the merged range (A1:B2) and land on the first non-merged cell (C1).
    expect(apiRef.current?.getSelection()).toEqual({ row: 0, col: 2 });
    expect(apiRef.current?.getSelectionRange()).toEqual({ startRow: 0, endRow: 2, startCol: 0, endCol: 3 });

    onSelectionChange.mockClear();
    onSelectionRangeChange.mockClear();

    await act(async () => {
      container.dispatchEvent(new KeyboardEvent("keydown", { key: "Tab", bubbles: true, cancelable: true }));
    });

    // Wrap to the next row start, which is inside the merge; skip to C2 instead.
    expect(apiRef.current?.getSelection()).toEqual({ row: 1, col: 2 });
    expect(apiRef.current?.getSelectionRange()).toEqual({ startRow: 0, endRow: 2, startCol: 0, endCol: 3 });

    onSelectionChange.mockClear();
    onSelectionRangeChange.mockClear();

    await act(async () => {
      container.dispatchEvent(new KeyboardEvent("keydown", { key: "Tab", bubbles: true, cancelable: true }));
    });

    // Wrap back to the beginning.
    expect(apiRef.current?.getSelection()).toEqual({ row: 0, col: 0 });
    expect(apiRef.current?.getSelectionRange()).toEqual({ startRow: 0, endRow: 2, startCol: 0, endCol: 3 });

    onSelectionChange.mockClear();
    onSelectionRangeChange.mockClear();

    await act(async () => {
      container.dispatchEvent(new KeyboardEvent("keydown", { key: "Tab", shiftKey: true, bubbles: true, cancelable: true }));
    });

    // Backwards wraps to the last non-merged cell.
    expect(apiRef.current?.getSelection()).toEqual({ row: 1, col: 2 });
    expect(apiRef.current?.getSelectionRange()).toEqual({ startRow: 0, endRow: 2, startCol: 0, endCol: 3 });

    await act(async () => {
      root.unmount();
    });
    host.remove();
  });

  it("treats frozen header rows/cols as the navigation origin", async () => {
    const apiRef = React.createRef<GridApi>();
    const onSelectionChange = vi.fn();
    const onSelectionRangeChange = vi.fn();

    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(
        <CanvasGrid
          provider={{ getCell: (row, col) => ({ row, col, value: `${row},${col}` }) }}
          rowCount={20}
          colCount={20}
          frozenRows={1}
          frozenCols={1}
          apiRef={apiRef}
          onSelectionChange={onSelectionChange}
          onSelectionRangeChange={onSelectionRangeChange}
        />
      );
    });

    const container = host.querySelector('[data-testid="canvas-grid"]') as HTMLDivElement;
    container.focus();

    await act(async () => {
      apiRef.current?.setSelection(5, 5);
    });
    onSelectionChange.mockClear();
    onSelectionRangeChange.mockClear();

    await act(async () => {
      container.dispatchEvent(new KeyboardEvent("keydown", { key: "Home", bubbles: true, cancelable: true }));
    });

    expect(apiRef.current?.getSelection()).toEqual({ row: 5, col: 1 });
    expect(apiRef.current?.getSelectionRange()).toEqual({ startRow: 5, endRow: 6, startCol: 1, endCol: 2 });

    onSelectionChange.mockClear();
    onSelectionRangeChange.mockClear();

    await act(async () => {
      apiRef.current?.setSelection(5, 5);
    });
    onSelectionChange.mockClear();
    onSelectionRangeChange.mockClear();

    await act(async () => {
      container.dispatchEvent(
        new KeyboardEvent("keydown", { key: "Home", ctrlKey: true, bubbles: true, cancelable: true })
      );
    });

    expect(apiRef.current?.getSelection()).toEqual({ row: 1, col: 1 });
    expect(apiRef.current?.getSelectionRange()).toEqual({ startRow: 1, endRow: 2, startCol: 1, endCol: 2 });

    await act(async () => {
      root.unmount();
    });
    host.remove();
  });

  it("skips merged cell interiors when navigating", async () => {
    const merged = { startRow: 0, endRow: 2, startCol: 0, endCol: 2 };
    const provider = {
      getCell: (row: number, col: number) => ({ row, col, value: `${row},${col}` }),
      getMergedRangeAt: (row: number, col: number) =>
        row >= merged.startRow && row < merged.endRow && col >= merged.startCol && col < merged.endCol ? merged : null,
      getMergedRangesInRange: (range: { startRow: number; endRow: number; startCol: number; endCol: number }) =>
        range.startRow < merged.endRow && range.endRow > merged.startRow && range.startCol < merged.endCol && range.endCol > merged.startCol
          ? [merged]
          : []
    };

    const apiRef = React.createRef<GridApi>();

    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(<CanvasGrid provider={provider} rowCount={10} colCount={10} apiRef={apiRef} />);
    });

    const container = host.querySelector('[data-testid="canvas-grid"]') as HTMLDivElement;
    container.focus();

    await act(async () => {
      apiRef.current?.setSelection(0, 0);
    });

    await act(async () => {
      container.dispatchEvent(new KeyboardEvent("keydown", { key: "ArrowRight", bubbles: true, cancelable: true }));
    });
    expect(apiRef.current?.getSelection()).toEqual({ row: 0, col: 2 });

    await act(async () => {
      apiRef.current?.setSelection(0, 0);
    });

    await act(async () => {
      container.dispatchEvent(new KeyboardEvent("keydown", { key: "ArrowDown", bubbles: true, cancelable: true }));
    });
    expect(apiRef.current?.getSelection()).toEqual({ row: 2, col: 0 });

    await act(async () => {
      apiRef.current?.setSelection(0, 0);
    });

    await act(async () => {
      container.dispatchEvent(new KeyboardEvent("keydown", { key: "Tab", bubbles: true, cancelable: true }));
    });
    expect(apiRef.current?.getSelection()).toEqual({ row: 0, col: 2 });

    await act(async () => {
      apiRef.current?.setSelection(0, 0);
    });

    await act(async () => {
      container.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", bubbles: true, cancelable: true }));
    });
    expect(apiRef.current?.getSelection()).toEqual({ row: 2, col: 0 });

    await act(async () => {
      root.unmount();
    });
    host.remove();
  });
});
