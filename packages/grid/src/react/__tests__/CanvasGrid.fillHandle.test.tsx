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
    setLineDash: noop,
    measureText: (text: string) =>
      ({
        width: text.length * 6,
        actualBoundingBoxAscent: 8,
        actualBoundingBoxDescent: 2
      }) as TextMetrics
  } as unknown as CanvasRenderingContext2D;
}

function createPointerEvent(
  type: string,
  options: { clientX: number; clientY: number; pointerId: number; ctrlKey?: boolean; metaKey?: boolean; altKey?: boolean }
): Event {
  const event = new MouseEvent(type, {
    bubbles: true,
    cancelable: true,
    clientX: options.clientX,
    clientY: options.clientY,
    ctrlKey: options.ctrlKey,
    metaKey: options.metaKey,
    altKey: options.altKey
  });
  Object.defineProperty(event, "pointerId", { value: options.pointerId });
  return event;
}

describe("CanvasGrid fill handle", () => {
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

  it("calls onFillHandleCommit with the extended range", async () => {
    const apiRef = React.createRef<GridApi>();
    const onFillHandleCommit = vi.fn();
    const onSelectionRangeChange = vi.fn();

    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(
        <CanvasGrid
          provider={{ getCell: (row, col) => ({ row, col, value: `${row},${col}` }) }}
          rowCount={20}
          colCount={10}
          apiRef={apiRef}
          onSelectionRangeChange={onSelectionRangeChange}
          onFillHandleCommit={onFillHandleCommit}
        />
      );
    });

    const sourceRange = { startRow: 0, endRow: 2, startCol: 0, endCol: 1 };
    await act(async () => {
      apiRef.current?.setSelectionRange(sourceRange);
    });

    onSelectionRangeChange.mockClear();
    onFillHandleCommit.mockClear();

    const handle = apiRef.current?.getFillHandleRect();
    expect(handle).not.toBeNull();

    const targetCell = apiRef.current?.getCellRect(3, 0);
    expect(targetCell).not.toBeNull();

    const selectionCanvas = host.querySelectorAll("canvas")[2] as HTMLCanvasElement;
    expect(selectionCanvas).toBeTruthy();

    const start = { clientX: handle!.x + handle!.width / 2, clientY: handle!.y + handle!.height / 2 };
    const end = { clientX: targetCell!.x + targetCell!.width / 2, clientY: targetCell!.y + targetCell!.height / 2 };

    await act(async () => {
      selectionCanvas.dispatchEvent(createPointerEvent("pointerdown", { ...start, pointerId: 1 }));
      selectionCanvas.dispatchEvent(createPointerEvent("pointermove", { ...end, pointerId: 1 }));
      selectionCanvas.dispatchEvent(createPointerEvent("pointerup", { ...end, pointerId: 1 }));
    });

    const expectedTarget = { startRow: 0, endRow: 4, startCol: 0, endCol: 1 };

    expect(onFillHandleCommit).toHaveBeenCalledTimes(1);
    expect(onFillHandleCommit).toHaveBeenCalledWith({ source: sourceRange, target: expectedTarget });
    expect(apiRef.current?.getSelectionRange()).toEqual(expectedTarget);
    expect(onSelectionRangeChange).toHaveBeenCalledWith(expectedTarget);

    await act(async () => {
      root.unmount();
    });
    host.remove();
  });

  it("calls onFillCommit with the target-only range and inferred mode", async () => {
    const apiRef = React.createRef<GridApi>();
    const onFillCommit = vi.fn();
    const onFillHandleCommit = vi.fn();

    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(
        <CanvasGrid
          provider={{ getCell: (row, col) => ({ row, col, value: `${row},${col}` }) }}
          rowCount={20}
          colCount={10}
          apiRef={apiRef}
          onFillCommit={onFillCommit}
          onFillHandleCommit={onFillHandleCommit}
        />
      );
    });

    const sourceRange = { startRow: 0, endRow: 2, startCol: 0, endCol: 1 };
    await act(async () => {
      apiRef.current?.setSelectionRange(sourceRange);
    });

    onFillCommit.mockClear();
    onFillHandleCommit.mockClear();

    const handle = apiRef.current?.getFillHandleRect();
    expect(handle).not.toBeNull();

    const targetCell = apiRef.current?.getCellRect(3, 0);
    expect(targetCell).not.toBeNull();

    const selectionCanvas = host.querySelectorAll("canvas")[2] as HTMLCanvasElement;
    expect(selectionCanvas).toBeTruthy();

    const start = { clientX: handle!.x + handle!.width / 2, clientY: handle!.y + handle!.height / 2 };
    const end = { clientX: targetCell!.x + targetCell!.width / 2, clientY: targetCell!.y + targetCell!.height / 2 };

    await act(async () => {
      selectionCanvas.dispatchEvent(createPointerEvent("pointerdown", { ...start, pointerId: 1, ctrlKey: true }));
      selectionCanvas.dispatchEvent(createPointerEvent("pointermove", { ...end, pointerId: 1 }));
      selectionCanvas.dispatchEvent(createPointerEvent("pointerup", { ...end, pointerId: 1 }));
    });

    expect(onFillHandleCommit).not.toHaveBeenCalled();
    expect(onFillCommit).toHaveBeenCalledTimes(1);
    expect(onFillCommit).toHaveBeenCalledWith({
      sourceRange,
      targetRange: { startRow: 2, endRow: 4, startCol: 0, endCol: 1 },
      mode: "copy"
    });

    expect(apiRef.current?.getSelectionRange()).toEqual({ startRow: 0, endRow: 4, startCol: 0, endCol: 1 });

    await act(async () => {
      root.unmount();
    });
    host.remove();
  });

  it("emits onFillPreviewChange with the target-only range and clears it", async () => {
    const apiRef = React.createRef<GridApi>();
    const onFillPreviewChange = vi.fn();

    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(
        <CanvasGrid
          provider={{ getCell: (row, col) => ({ row, col, value: `${row},${col}` }) }}
          rowCount={20}
          colCount={10}
          apiRef={apiRef}
          onFillPreviewChange={onFillPreviewChange}
          onFillHandleCommit={vi.fn()}
        />
      );
    });

    const sourceRange = { startRow: 0, endRow: 2, startCol: 0, endCol: 1 };
    await act(async () => {
      apiRef.current?.setSelectionRange(sourceRange);
    });

    onFillPreviewChange.mockClear();

    const handle = apiRef.current?.getFillHandleRect();
    expect(handle).not.toBeNull();

    const targetCell = apiRef.current?.getCellRect(3, 0);
    expect(targetCell).not.toBeNull();

    const selectionCanvas = host.querySelectorAll("canvas")[2] as HTMLCanvasElement;
    expect(selectionCanvas).toBeTruthy();

    const start = { clientX: handle!.x + handle!.width / 2, clientY: handle!.y + handle!.height / 2 };
    const end = { clientX: targetCell!.x + targetCell!.width / 2, clientY: targetCell!.y + targetCell!.height / 2 };

    await act(async () => {
      selectionCanvas.dispatchEvent(createPointerEvent("pointerdown", { ...start, pointerId: 1 }));
      selectionCanvas.dispatchEvent(createPointerEvent("pointermove", { ...end, pointerId: 1 }));
    });

    expect(onFillPreviewChange).toHaveBeenCalled();
    expect(onFillPreviewChange).toHaveBeenLastCalledWith({ startRow: 2, endRow: 4, startCol: 0, endCol: 1 });

    await act(async () => {
      selectionCanvas.dispatchEvent(createPointerEvent("pointerup", { ...end, pointerId: 1 }));
    });

    expect(onFillPreviewChange).toHaveBeenLastCalledWith(null);

    await act(async () => {
      root.unmount();
    });
    host.remove();
  });

  it("does not commit a fill operation on pointercancel", async () => {
    const apiRef = React.createRef<GridApi>();
    const onFillCommit = vi.fn();

    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(
        <CanvasGrid
          provider={{ getCell: (row, col) => ({ row, col, value: `${row},${col}` }) }}
          rowCount={20}
          colCount={10}
          apiRef={apiRef}
          onFillCommit={onFillCommit}
        />
      );
    });

    const sourceRange = { startRow: 0, endRow: 2, startCol: 0, endCol: 1 };
    await act(async () => {
      apiRef.current?.setSelectionRange(sourceRange);
    });

    onFillCommit.mockClear();

    const handle = apiRef.current?.getFillHandleRect();
    expect(handle).not.toBeNull();

    const targetCell = apiRef.current?.getCellRect(3, 0);
    expect(targetCell).not.toBeNull();

    const selectionCanvas = host.querySelectorAll("canvas")[2] as HTMLCanvasElement;
    expect(selectionCanvas).toBeTruthy();

    const start = { clientX: handle!.x + handle!.width / 2, clientY: handle!.y + handle!.height / 2 };
    const end = { clientX: targetCell!.x + targetCell!.width / 2, clientY: targetCell!.y + targetCell!.height / 2 };

    await act(async () => {
      selectionCanvas.dispatchEvent(createPointerEvent("pointerdown", { ...start, pointerId: 1 }));
      selectionCanvas.dispatchEvent(createPointerEvent("pointermove", { ...end, pointerId: 1 }));
      selectionCanvas.dispatchEvent(createPointerEvent("pointercancel", { ...end, pointerId: 1 }));
    });

    expect(onFillCommit).not.toHaveBeenCalled();
    expect(apiRef.current?.getSelectionRange()).toEqual(sourceRange);

    await act(async () => {
      root.unmount();
    });
    host.remove();
  });

  it("cancels fill handle drag on Escape without committing", async () => {
    const apiRef = React.createRef<GridApi>();
    const onFillCommit = vi.fn();
    const onFillPreviewChange = vi.fn();

    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(
        <CanvasGrid
          provider={{ getCell: (row, col) => ({ row, col, value: `${row},${col}` }) }}
          rowCount={20}
          colCount={10}
          apiRef={apiRef}
          onFillCommit={onFillCommit}
          onFillPreviewChange={onFillPreviewChange}
        />
      );
    });

    const sourceRange = { startRow: 0, endRow: 2, startCol: 0, endCol: 1 };
    await act(async () => {
      apiRef.current?.setSelectionRange(sourceRange);
    });

    onFillCommit.mockClear();
    onFillPreviewChange.mockClear();

    const handle = apiRef.current?.getFillHandleRect();
    expect(handle).not.toBeNull();

    const targetCell = apiRef.current?.getCellRect(3, 0);
    expect(targetCell).not.toBeNull();

    const selectionCanvas = host.querySelectorAll("canvas")[2] as HTMLCanvasElement;
    expect(selectionCanvas).toBeTruthy();

    const container = host.querySelector('[data-testid="canvas-grid"]') as HTMLDivElement;
    expect(container).toBeTruthy();

    const start = { clientX: handle!.x + handle!.width / 2, clientY: handle!.y + handle!.height / 2 };
    const end = { clientX: targetCell!.x + targetCell!.width / 2, clientY: targetCell!.y + targetCell!.height / 2 };

    await act(async () => {
      selectionCanvas.dispatchEvent(createPointerEvent("pointerdown", { ...start, pointerId: 1 }));
      selectionCanvas.dispatchEvent(createPointerEvent("pointermove", { ...end, pointerId: 1 }));
    });

    await act(async () => {
      container.dispatchEvent(new KeyboardEvent("keydown", { bubbles: true, cancelable: true, key: "Escape" }));
    });

    expect(onFillCommit).not.toHaveBeenCalled();
    expect(onFillPreviewChange).toHaveBeenLastCalledWith(null);
    expect(apiRef.current?.getSelectionRange()).toEqual(sourceRange);

    await act(async () => {
      selectionCanvas.dispatchEvent(createPointerEvent("pointerup", { ...end, pointerId: 1 }));
    });

    expect(onFillCommit).not.toHaveBeenCalled();
    expect(apiRef.current?.getSelectionRange()).toEqual(sourceRange);

    await act(async () => {
      root.unmount();
    });
    host.remove();
  });

  it("infers horizontal fill when dragging the handle sideways", async () => {
    const apiRef = React.createRef<GridApi>();
    const onFillHandleCommit = vi.fn();

    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(
        <CanvasGrid
          provider={{ getCell: (row, col) => ({ row, col, value: `${row},${col}` }) }}
          rowCount={20}
          colCount={20}
          apiRef={apiRef}
          onFillHandleCommit={onFillHandleCommit}
        />
      );
    });

    const sourceRange = { startRow: 0, endRow: 1, startCol: 0, endCol: 2 };
    await act(async () => {
      apiRef.current?.setSelectionRange(sourceRange);
    });

    onFillHandleCommit.mockClear();

    const handle = apiRef.current?.getFillHandleRect();
    expect(handle).not.toBeNull();

    const targetCell = apiRef.current?.getCellRect(0, 3);
    expect(targetCell).not.toBeNull();

    const selectionCanvas = host.querySelectorAll("canvas")[2] as HTMLCanvasElement;
    expect(selectionCanvas).toBeTruthy();

    const start = { clientX: handle!.x + handle!.width / 2, clientY: handle!.y + handle!.height / 2 };
    const end = { clientX: targetCell!.x + targetCell!.width / 2, clientY: start.clientY };

    await act(async () => {
      selectionCanvas.dispatchEvent(createPointerEvent("pointerdown", { ...start, pointerId: 1 }));
      selectionCanvas.dispatchEvent(createPointerEvent("pointermove", { ...end, pointerId: 1 }));
      selectionCanvas.dispatchEvent(createPointerEvent("pointerup", { ...end, pointerId: 1 }));
    });

    const expectedTarget = { startRow: 0, endRow: 1, startCol: 0, endCol: 4 };

    expect(onFillHandleCommit).toHaveBeenCalledTimes(1);
    expect(onFillHandleCommit).toHaveBeenCalledWith({ source: sourceRange, target: expectedTarget });
    expect(apiRef.current?.getSelectionRange()).toEqual(expectedTarget);

    await act(async () => {
      root.unmount();
    });
    host.remove();
  });

  it("does not extend the selection into header rows/cols", async () => {
    const apiRef = React.createRef<GridApi>();
    const onFillHandleCommit = vi.fn();
    const onSelectionRangeChange = vi.fn();

    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(
        <CanvasGrid
          provider={{ getCell: (row, col) => ({ row, col, value: `${row},${col}` }) }}
          rowCount={20}
          colCount={10}
          headerRows={1}
          headerCols={1}
          frozenRows={1}
          frozenCols={1}
          apiRef={apiRef}
          onSelectionRangeChange={onSelectionRangeChange}
          onFillHandleCommit={onFillHandleCommit}
        />
      );
    });

    const sourceRange = { startRow: 1, endRow: 3, startCol: 1, endCol: 2 };
    await act(async () => {
      apiRef.current?.setSelectionRange(sourceRange);
    });

    onSelectionRangeChange.mockClear();
    onFillHandleCommit.mockClear();

    const handle = apiRef.current?.getFillHandleRect();
    expect(handle).not.toBeNull();

    const headerCell = apiRef.current?.getCellRect(0, 1);
    expect(headerCell).not.toBeNull();

    const selectionCanvas = host.querySelectorAll("canvas")[2] as HTMLCanvasElement;
    expect(selectionCanvas).toBeTruthy();

    const start = { clientX: handle!.x + handle!.width / 2, clientY: handle!.y + handle!.height / 2 };
    const end = { clientX: headerCell!.x + headerCell!.width / 2, clientY: headerCell!.y + headerCell!.height / 2 };

    await act(async () => {
      selectionCanvas.dispatchEvent(createPointerEvent("pointerdown", { ...start, pointerId: 1 }));
      selectionCanvas.dispatchEvent(createPointerEvent("pointermove", { ...end, pointerId: 1 }));
      selectionCanvas.dispatchEvent(createPointerEvent("pointerup", { ...end, pointerId: 1 }));
    });

    expect(onFillHandleCommit).not.toHaveBeenCalled();
    expect(apiRef.current?.getSelectionRange()).toEqual(sourceRange);
    expect(onSelectionRangeChange).not.toHaveBeenCalled();

    await act(async () => {
      root.unmount();
    });
    host.remove();
  });

  it("hides the fill handle when interactionMode is rangeSelection", async () => {
    const apiRef = React.createRef<GridApi>();

    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(
        <CanvasGrid
          provider={{ getCell: (row, col) => ({ row, col, value: `${row},${col}` }) }}
          rowCount={20}
          colCount={10}
          interactionMode="rangeSelection"
          apiRef={apiRef}
        />
      );
    });

    await act(async () => {
      apiRef.current?.setSelectionRange({ startRow: 0, endRow: 1, startCol: 0, endCol: 1 });
    });

    expect(apiRef.current?.getFillHandleRect()).toBeNull();

    await act(async () => {
      root.unmount();
    });
    host.remove();
  });

  it("only updates the active selection range when using multi-range selection", async () => {
    const apiRef = React.createRef<GridApi>();
    const onFillHandleCommit = vi.fn();

    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(
        <CanvasGrid
          provider={{ getCell: (row, col) => ({ row, col, value: `${row},${col}` }) }}
          rowCount={20}
          colCount={10}
          apiRef={apiRef}
          onFillHandleCommit={onFillHandleCommit}
        />
      );
    });

    const rangeA = { startRow: 0, endRow: 1, startCol: 0, endCol: 1 };
    const rangeB = { startRow: 2, endRow: 3, startCol: 2, endCol: 3 };

    await act(async () => {
      apiRef.current?.setSelectionRanges([rangeA, rangeB], { activeIndex: 1 });
    });

    const handle = apiRef.current?.getFillHandleRect();
    expect(handle).not.toBeNull();

    const targetCell = apiRef.current?.getCellRect(4, 2);
    expect(targetCell).not.toBeNull();

    const selectionCanvas = host.querySelectorAll("canvas")[2] as HTMLCanvasElement;
    expect(selectionCanvas).toBeTruthy();

    const start = { clientX: handle!.x + handle!.width / 2, clientY: handle!.y + handle!.height / 2 };
    const end = { clientX: targetCell!.x + targetCell!.width / 2, clientY: targetCell!.y + targetCell!.height / 2 };

    await act(async () => {
      selectionCanvas.dispatchEvent(createPointerEvent("pointerdown", { ...start, pointerId: 1 }));
      selectionCanvas.dispatchEvent(createPointerEvent("pointermove", { ...end, pointerId: 1 }));
      selectionCanvas.dispatchEvent(createPointerEvent("pointerup", { ...end, pointerId: 1 }));
    });

    const expectedTarget = { startRow: 2, endRow: 5, startCol: 2, endCol: 3 };
    expect(onFillHandleCommit).toHaveBeenCalledWith({ source: rangeB, target: expectedTarget });

    expect(apiRef.current?.getSelectionRanges()).toEqual([rangeA, expectedTarget]);
    expect(apiRef.current?.getActiveSelectionRangeIndex()).toBe(1);

    await act(async () => {
      root.unmount();
    });
    host.remove();
  });
});
