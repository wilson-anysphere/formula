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

describe("DesktopSharedGrid fill handle", () => {
  const originalGetContext = HTMLCanvasElement.prototype.getContext;

  beforeEach(() => {
    document.body.innerHTML = "";

    vi.stubGlobal("requestAnimationFrame", (cb: FrameRequestCallback) => {
      cb(0);
      return 0;
    });
    vi.stubGlobal("cancelAnimationFrame", () => {});

    Object.defineProperty(HTMLCanvasElement.prototype, "getContext", {
      configurable: true,
      writable: true,
      value: () => createMockCanvasContext()
    });
  });

  afterEach(() => {
    Object.defineProperty(HTMLCanvasElement.prototype, "getContext", {
      configurable: true,
      writable: true,
      value: originalGetContext
    });
    vi.restoreAllMocks();
    vi.unstubAllGlobals();
  });

  it("calls onFillCommit with the target-only range and inferred mode", () => {
    const onFillCommit = vi.fn();
    const onSelectionRangeChange = vi.fn();

    const container = document.createElement("div");
    container.tabIndex = 0;
    document.body.appendChild(container);

    const canvases = {
      grid: document.createElement("canvas"),
      content: document.createElement("canvas"),
      selection: document.createElement("canvas")
    };
    container.appendChild(canvases.grid);
    container.appendChild(canvases.content);
    container.appendChild(canvases.selection);

    const scrollbars = {
      vTrack: document.createElement("div"),
      vThumb: document.createElement("div"),
      hTrack: document.createElement("div"),
      hThumb: document.createElement("div")
    };
    container.appendChild(scrollbars.vTrack);
    container.appendChild(scrollbars.vThumb);
    container.appendChild(scrollbars.hTrack);
    container.appendChild(scrollbars.hThumb);

    vi.spyOn(canvases.selection, "getBoundingClientRect").mockReturnValue({
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

    const grid = new DesktopSharedGrid({
      container,
      provider: new MockCellProvider({ rowCount: 20, colCount: 10 }),
      rowCount: 20,
      colCount: 10,
      canvases,
      scrollbars,
      enableResize: false,
      enableKeyboard: false,
      enableWheel: false,
      callbacks: {
        onFillCommit,
        onSelectionRangeChange
      }
    });
    grid.resize(400, 200, 1);

    const sourceRange = { startRow: 0, endRow: 2, startCol: 0, endCol: 1 };
    grid.renderer.setSelectionRange(sourceRange);
    onFillCommit.mockClear();
    onSelectionRangeChange.mockClear();

    const handle = grid.renderer.getFillHandleRect();
    expect(handle).not.toBeNull();
    const targetCell = grid.renderer.getCellRect(3, 0);
    expect(targetCell).not.toBeNull();

    const start = { clientX: handle!.x + handle!.width / 2, clientY: handle!.y + handle!.height / 2 };
    const end = { clientX: targetCell!.x + targetCell!.width / 2, clientY: targetCell!.y + targetCell!.height / 2 };

    canvases.selection.dispatchEvent(createPointerEvent("pointerdown", { ...start, pointerId: 1, ctrlKey: true }));
    canvases.selection.dispatchEvent(createPointerEvent("pointermove", { ...end, pointerId: 1 }));
    canvases.selection.dispatchEvent(createPointerEvent("pointerup", { ...end, pointerId: 1 }));

    expect(onFillCommit).toHaveBeenCalledTimes(1);
    expect(onFillCommit).toHaveBeenCalledWith({
      sourceRange,
      targetRange: { startRow: 2, endRow: 4, startCol: 0, endCol: 1 },
      mode: "copy"
    });

    expect(grid.renderer.getSelectionRange()).toEqual({ startRow: 0, endRow: 4, startCol: 0, endCol: 1 });
    expect(onSelectionRangeChange).toHaveBeenLastCalledWith({ startRow: 0, endRow: 4, startCol: 0, endCol: 1 });
    // Desktop behavior: active cell tracks the end of the fill handle drag.
    expect(grid.renderer.getSelection()).toEqual({ row: 3, col: 0 });

    grid.destroy();
    container.remove();
  });

  it("does not extend the selection into header rows/cols", () => {
    const onFillCommit = vi.fn();

    const container = document.createElement("div");
    container.tabIndex = 0;
    document.body.appendChild(container);

    const canvases = {
      grid: document.createElement("canvas"),
      content: document.createElement("canvas"),
      selection: document.createElement("canvas")
    };
    container.appendChild(canvases.grid);
    container.appendChild(canvases.content);
    container.appendChild(canvases.selection);

    const scrollbars = {
      vTrack: document.createElement("div"),
      vThumb: document.createElement("div"),
      hTrack: document.createElement("div"),
      hThumb: document.createElement("div")
    };
    container.appendChild(scrollbars.vTrack);
    container.appendChild(scrollbars.vThumb);
    container.appendChild(scrollbars.hTrack);
    container.appendChild(scrollbars.hThumb);

    vi.spyOn(canvases.selection, "getBoundingClientRect").mockReturnValue({
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

    const grid = new DesktopSharedGrid({
      container,
      provider: new MockCellProvider({ rowCount: 20, colCount: 10 }),
      rowCount: 20,
      colCount: 10,
      canvases,
      scrollbars,
      frozenRows: 1,
      frozenCols: 1,
      enableResize: false,
      enableKeyboard: false,
      enableWheel: false,
      callbacks: {
        onFillCommit
      }
    });
    grid.resize(400, 200, 1);

    const sourceRange = { startRow: 1, endRow: 3, startCol: 1, endCol: 2 };
    grid.renderer.setSelectionRange(sourceRange);
    onFillCommit.mockClear();

    const handle = grid.renderer.getFillHandleRect();
    expect(handle).not.toBeNull();
    const headerCell = grid.renderer.getCellRect(0, 1);
    expect(headerCell).not.toBeNull();

    const start = { clientX: handle!.x + handle!.width / 2, clientY: handle!.y + handle!.height / 2 };
    const end = { clientX: headerCell!.x + headerCell!.width / 2, clientY: headerCell!.y + headerCell!.height / 2 };

    canvases.selection.dispatchEvent(createPointerEvent("pointerdown", { ...start, pointerId: 1 }));
    canvases.selection.dispatchEvent(createPointerEvent("pointermove", { ...end, pointerId: 1 }));
    canvases.selection.dispatchEvent(createPointerEvent("pointerup", { ...end, pointerId: 1 }));

    expect(onFillCommit).not.toHaveBeenCalled();
    expect(grid.renderer.getSelectionRange()).toEqual(sourceRange);

    grid.destroy();
    container.remove();
  });
});

