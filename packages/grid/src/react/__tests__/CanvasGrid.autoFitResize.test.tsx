// @vitest-environment jsdom
import React, { act } from "react";
import { createRoot } from "react-dom/client";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { CanvasGrid, type GridApi, type GridAxisSizeChange } from "../CanvasGrid";
import type { CellProvider } from "../../model/CellProvider";

// React 18 relies on this flag to suppress act() warnings in test runners.
// eslint-disable-next-line @typescript-eslint/no-explicit-any
(globalThis as any).IS_REACT_ACT_ENVIRONMENT = true;

function createPointerEvent(
  type: string,
  options: { clientX: number; clientY: number; pointerId: number; pointerType?: string }
): Event {
  const PointerEventCtor = (window as unknown as { PointerEvent?: typeof PointerEvent }).PointerEvent;
  if (PointerEventCtor) {
    return new PointerEventCtor(type, {
      bubbles: true,
      cancelable: true,
      clientX: options.clientX,
      clientY: options.clientY,
      buttons: 1,
      pointerId: options.pointerId,
      pointerType: options.pointerType
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
  if (options.pointerType) Object.defineProperty(event, "pointerType", { value: options.pointerType });
  return event;
}

describe("CanvasGrid resize auto-fit interactions", () => {
  beforeEach(() => {
    vi.stubGlobal(
      "ResizeObserver",
      class ResizeObserver {
        observe(): void {}
        unobserve(): void {}
        disconnect(): void {}
      }
    );

    // Avoid running full render frames; these tests only validate interaction wiring.
    vi.stubGlobal("requestAnimationFrame", vi.fn(() => 0));

    vi.spyOn(HTMLElement.prototype, "getBoundingClientRect").mockReturnValue({
      left: 0,
      top: 0,
      right: 300,
      bottom: 200,
      width: 300,
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

  it("double clicking a header resize boundary auto-fits the column and emits a single onAxisSizeChange event", async () => {
    const provider: CellProvider = {
      getCell: (row, col) => {
        // Make column A (col=1) require a different width than the default.
        if (col === 1 && row === 1) return { row, col, value: "This is a long cell value" };
        // Basic headers so address/selection math behaves as expected.
        if (row === 0 && col === 1) return { row, col, value: "A" };
        if (row === 1 && col === 0) return { row, col, value: 1 };
        return null;
      }
    };

    const apiRef = React.createRef<GridApi>();
    const onAxisSizeChange = vi.fn<(change: GridAxisSizeChange) => void>();

    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(
        <CanvasGrid
          provider={provider}
          rowCount={20}
          colCount={10}
          defaultRowHeight={10}
          defaultColWidth={50}
          frozenRows={1}
          frozenCols={1}
          enableResize
          apiRef={apiRef}
          onAxisSizeChange={onAxisSizeChange}
        />
      );
    });

    const api = apiRef.current;
    expect(api).toBeTruthy();

    const before = api!.getColWidth(1);

    const colARect = api!.getCellRect(0, 1);
    expect(colARect).not.toBeNull();

    const boundaryX = colARect!.x + colARect!.width;
    const boundaryY = colARect!.y + colARect!.height / 2;

    const selectionCanvas = host.querySelector('[data-testid="canvas-grid-selection"]') as HTMLCanvasElement;
    expect(selectionCanvas).toBeTruthy();

    await act(async () => {
      selectionCanvas.dispatchEvent(createPointerEvent("pointerdown", { clientX: boundaryX, clientY: boundaryY, pointerId: 1 }));
      selectionCanvas.dispatchEvent(createPointerEvent("pointerup", { clientX: boundaryX, clientY: boundaryY, pointerId: 1 }));
      selectionCanvas.dispatchEvent(createPointerEvent("pointerdown", { clientX: boundaryX, clientY: boundaryY, pointerId: 1 }));
      selectionCanvas.dispatchEvent(createPointerEvent("pointerup", { clientX: boundaryX, clientY: boundaryY, pointerId: 1 }));
    });

    const after = api!.getColWidth(1);
    expect(after).not.toBe(before);

    expect(onAxisSizeChange).toHaveBeenCalledTimes(1);
    expect(onAxisSizeChange.mock.calls[0]?.[0]).toMatchObject({
      kind: "col",
      index: 1,
      source: "autoFit"
    });

    await act(async () => {
      root.unmount();
    });
    host.remove();
  });

  it("dragging a header resize boundary updates the column and emits a single onAxisSizeChange commit event", async () => {
    const provider: CellProvider = {
      getCell: (row, col) => {
        if (row === 0 && col === 1) return { row, col, value: "A" };
        if (row === 1 && col === 0) return { row, col, value: 1 };
        return null;
      }
    };

    const apiRef = React.createRef<GridApi>();
    const onAxisSizeChange = vi.fn<(change: GridAxisSizeChange) => void>();

    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(
        <CanvasGrid
          provider={provider}
          rowCount={20}
          colCount={10}
          defaultRowHeight={10}
          defaultColWidth={50}
          frozenRows={1}
          frozenCols={1}
          enableResize
          apiRef={apiRef}
          onAxisSizeChange={onAxisSizeChange}
        />
      );
    });

    const api = apiRef.current;
    expect(api).toBeTruthy();

    const before = api!.getColWidth(1);
    expect(before).toBe(50);

    const colARect = api!.getCellRect(0, 1);
    expect(colARect).not.toBeNull();

    const boundaryX = colARect!.x + colARect!.width;
    const boundaryY = colARect!.y + colARect!.height / 2;

    const selectionCanvas = host.querySelector('[data-testid="canvas-grid-selection"]') as HTMLCanvasElement;
    expect(selectionCanvas).toBeTruthy();

    await act(async () => {
      selectionCanvas.dispatchEvent(createPointerEvent("pointerdown", { clientX: boundaryX, clientY: boundaryY, pointerId: 3 }));
      selectionCanvas.dispatchEvent(createPointerEvent("pointermove", { clientX: boundaryX + 40, clientY: boundaryY, pointerId: 3 }));
    });

    expect(onAxisSizeChange).toHaveBeenCalledTimes(0);

    await act(async () => {
      selectionCanvas.dispatchEvent(createPointerEvent("pointerup", { clientX: boundaryX + 40, clientY: boundaryY, pointerId: 3 }));
    });

    const after = api!.getColWidth(1);
    expect(after).toBeGreaterThan(before);

    expect(onAxisSizeChange).toHaveBeenCalledTimes(1);
    expect(onAxisSizeChange.mock.calls[0]?.[0]).toMatchObject({
      kind: "col",
      index: 1,
      previousSize: before,
      size: after,
      defaultSize: 50,
      zoom: 1,
      source: "resize"
    });

    await act(async () => {
      root.unmount();
    });
    host.remove();
  });

  it("double tapping a header resize boundary auto-fits the column on touch", async () => {
    const provider: CellProvider = {
      getCell: (row, col) => {
        if (col === 1 && row === 1) return { row, col, value: "This is a long cell value" };
        if (row === 0 && col === 1) return { row, col, value: "A" };
        if (row === 1 && col === 0) return { row, col, value: 1 };
        return null;
      }
    };

    const apiRef = React.createRef<GridApi>();
    const onAxisSizeChange = vi.fn<(change: GridAxisSizeChange) => void>();

    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(
        <CanvasGrid
          provider={provider}
          rowCount={20}
          colCount={10}
          defaultRowHeight={10}
          defaultColWidth={50}
          frozenRows={1}
          frozenCols={1}
          enableResize
          apiRef={apiRef}
          onAxisSizeChange={onAxisSizeChange}
        />
      );
    });

    const api = apiRef.current;
    expect(api).toBeTruthy();

    const before = api!.getColWidth(1);

    const colARect = api!.getCellRect(0, 1);
    expect(colARect).not.toBeNull();

    const boundaryX = colARect!.x + colARect!.width;
    const boundaryY = colARect!.y + colARect!.height / 2;

    const selectionCanvas = host.querySelector('[data-testid="canvas-grid-selection"]') as HTMLCanvasElement;
    expect(selectionCanvas).toBeTruthy();

    await act(async () => {
      selectionCanvas.dispatchEvent(
        createPointerEvent("pointerdown", { clientX: boundaryX, clientY: boundaryY, pointerId: 1, pointerType: "touch" })
      );
      selectionCanvas.dispatchEvent(
        createPointerEvent("pointerup", { clientX: boundaryX, clientY: boundaryY, pointerId: 1, pointerType: "touch" })
      );
      selectionCanvas.dispatchEvent(
        createPointerEvent("pointerdown", { clientX: boundaryX, clientY: boundaryY, pointerId: 1, pointerType: "touch" })
      );
      selectionCanvas.dispatchEvent(
        createPointerEvent("pointerup", { clientX: boundaryX, clientY: boundaryY, pointerId: 1, pointerType: "touch" })
      );
    });

    const after = api!.getColWidth(1);
    expect(after).not.toBe(before);
    // Touch resize-handle taps should not select cells.
    expect(api!.getSelection()).toBeNull();

    expect(onAxisSizeChange).toHaveBeenCalledTimes(1);
    expect(onAxisSizeChange.mock.calls[0]?.[0]).toMatchObject({
      kind: "col",
      index: 1,
      source: "autoFit"
    });

    await act(async () => {
      root.unmount();
    });
    host.remove();
  });

  it("double clicking a row header resize boundary auto-fits the row", async () => {
    const provider: CellProvider = {
      getCell: (row, col) => {
        if (row === 1 && col === 1) return { row, col, value: "A", style: { fontSize: 30 } };
        if (row === 0 && col === 1) return { row, col, value: "A" };
        if (row === 1 && col === 0) return { row, col, value: 1 };
        return null;
      }
    };

    const apiRef = React.createRef<GridApi>();
    const onAxisSizeChange = vi.fn<(change: GridAxisSizeChange) => void>();

    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(
        <CanvasGrid
          provider={provider}
          rowCount={20}
          colCount={10}
          defaultRowHeight={10}
          defaultColWidth={50}
          frozenRows={1}
          frozenCols={1}
          enableResize
          apiRef={apiRef}
          onAxisSizeChange={onAxisSizeChange}
        />
      );
    });

    const api = apiRef.current;
    expect(api).toBeTruthy();

    const before = api!.getRowHeight(1);
    expect(before).toBe(10);

    const rowHeaderRect = api!.getCellRect(1, 0);
    expect(rowHeaderRect).not.toBeNull();

    const boundaryX = rowHeaderRect!.x + rowHeaderRect!.width / 2;
    const boundaryY = rowHeaderRect!.y + rowHeaderRect!.height;

    const selectionCanvas = host.querySelector('[data-testid="canvas-grid-selection"]') as HTMLCanvasElement;
    expect(selectionCanvas).toBeTruthy();

    await act(async () => {
      selectionCanvas.dispatchEvent(createPointerEvent("pointerdown", { clientX: boundaryX, clientY: boundaryY, pointerId: 2 }));
      selectionCanvas.dispatchEvent(createPointerEvent("pointerup", { clientX: boundaryX, clientY: boundaryY, pointerId: 2 }));
      selectionCanvas.dispatchEvent(createPointerEvent("pointerdown", { clientX: boundaryX, clientY: boundaryY, pointerId: 2 }));
      selectionCanvas.dispatchEvent(createPointerEvent("pointerup", { clientX: boundaryX, clientY: boundaryY, pointerId: 2 }));
    });

    const after = api!.getRowHeight(1);
    expect(after).toBeGreaterThan(before);

    expect(onAxisSizeChange).toHaveBeenCalledTimes(1);
    expect(onAxisSizeChange.mock.calls[0]?.[0]).toMatchObject({
      kind: "row",
      index: 1,
      source: "autoFit"
    });

    await act(async () => {
      root.unmount();
    });
    host.remove();
  });
});
