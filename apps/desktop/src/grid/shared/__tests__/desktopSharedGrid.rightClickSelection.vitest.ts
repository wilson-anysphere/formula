/**
 * @vitest-environment jsdom
 */

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { MockCellProvider } from "@formula/grid";

import { DesktopSharedGrid } from "../desktopSharedGrid";

function createMockCanvasContext(): CanvasRenderingContext2D {
  const noop = () => {};
  const gradient = { addColorStop: noop } as any;
  const context = new Proxy(
    {
      canvas: document.createElement("canvas"),
      measureText: (text: string) => ({ width: text.length * 8 }),
      createLinearGradient: () => gradient,
      createPattern: () => null,
      getImageData: () => ({ data: new Uint8ClampedArray(), width: 0, height: 0 }),
      putImageData: noop,
    },
    {
      get(target, prop) {
        if (prop in target) return (target as any)[prop];
        return noop;
      },
      set(target, prop, value) {
        (target as any)[prop] = value;
        return true;
      },
    },
  );
  return context as any;
}

function createPointerLikeMouseEvent(
  type: string,
  options: { clientX: number; clientY: number; button: number; pointerId?: number; pointerType?: string },
): MouseEvent {
  const event = new MouseEvent(type, {
    bubbles: true,
    cancelable: true,
    clientX: options.clientX,
    clientY: options.clientY,
    button: options.button,
  });
  Object.defineProperty(event, "pointerId", { configurable: true, value: options.pointerId ?? 1 });
  Object.defineProperty(event, "pointerType", { configurable: true, value: options.pointerType ?? "mouse" });
  return event;
}

describe("DesktopSharedGrid right-click selection semantics", () => {
  afterEach(() => {
    vi.restoreAllMocks();
    vi.unstubAllGlobals();
  });

  beforeEach(() => {
    document.body.innerHTML = "";

    vi.stubGlobal("requestAnimationFrame", () => 0);
    vi.stubGlobal("cancelAnimationFrame", () => {});

    Object.defineProperty(HTMLCanvasElement.prototype, "getContext", {
      configurable: true,
      value: () => createMockCanvasContext(),
    });
  });

  it("keeps selection intact when right-clicking inside the selection, but moves it when right-clicking outside", () => {
    const container = document.createElement("div");
    container.tabIndex = 0;
    document.body.appendChild(container);

    const gridCanvas = document.createElement("canvas");
    const contentCanvas = document.createElement("canvas");
    const selectionCanvas = document.createElement("canvas");
    container.appendChild(gridCanvas);
    container.appendChild(contentCanvas);
    container.appendChild(selectionCanvas);

    const vTrack = document.createElement("div");
    const vThumb = document.createElement("div");
    const hTrack = document.createElement("div");
    const hThumb = document.createElement("div");
    container.appendChild(vTrack);
    container.appendChild(vThumb);
    container.appendChild(hTrack);
    container.appendChild(hThumb);

    vi.spyOn(selectionCanvas, "getBoundingClientRect").mockReturnValue({
      left: 0,
      top: 0,
      right: 400,
      bottom: 200,
      width: 400,
      height: 200,
      x: 0,
      y: 0,
      toJSON: () => {},
    } as DOMRect);

    const outsideFocusTarget = document.createElement("button");
    outsideFocusTarget.textContent = "outside";
    document.body.appendChild(outsideFocusTarget);
    outsideFocusTarget.focus();
    expect(document.activeElement).toBe(outsideFocusTarget);

    const grid = new DesktopSharedGrid({
      container,
      provider: new MockCellProvider({ rowCount: 100, colCount: 100 }),
      rowCount: 100,
      colCount: 100,
      frozenRows: 1,
      frozenCols: 1,
      defaultRowHeight: 24,
      defaultColWidth: 100,
      canvases: { grid: gridCanvas, content: contentCanvas, selection: selectionCanvas },
      scrollbars: { vTrack, vThumb, hTrack, hThumb },
      enableResize: false,
      enableKeyboard: false,
      enableWheel: false,
    });
    grid.renderer.setColWidth(0, 48);
    grid.renderer.setRowHeight(0, 24);
    grid.resize(400, 200, 1);

    // Select A1:B2 with active cell at A1.
    grid.setSelectionRanges(
      [{ startRow: 1, endRow: 3, startCol: 1, endCol: 3 }],
      { activeIndex: 0, activeCell: { row: 1, col: 1 }, scrollIntoView: false },
    );

    // Right click B2 (within selection). This should not move the active cell.
    selectionCanvas.dispatchEvent(
      createPointerLikeMouseEvent("pointerdown", {
        clientX: 48 + 100 + 10, // col 2 (B), account for row header + col width
        clientY: 24 + 24 + 10, // row 2 (2), account for header row + row height
        button: 2,
      }),
    );
    expect(grid.renderer.getSelection()).toEqual({ row: 1, col: 1 });
    expect(grid.renderer.getSelectionRanges()).toEqual([{ startRow: 1, endRow: 3, startCol: 1, endCol: 3 }]);
    expect(document.activeElement).toBe(container);

    // Right click D4 (outside selection). This should collapse selection to D4.
    selectionCanvas.dispatchEvent(
      createPointerLikeMouseEvent("pointerdown", {
        clientX: 48 + 3 * 100 + 10, // col 4 (D)
        clientY: 24 + 3 * 24 + 10, // row 4
        button: 2,
      }),
    );
    expect(grid.renderer.getSelection()).toEqual({ row: 4, col: 4 });
    expect(grid.renderer.getSelectionRanges()).toEqual([{ startRow: 4, endRow: 5, startCol: 4, endCol: 5 }]);

    grid.destroy();
    container.remove();
    outsideFocusTarget.remove();
  });
});

