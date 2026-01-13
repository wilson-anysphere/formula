// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { MockCellProvider } from "@formula/grid";
import { DesktopSharedGrid } from "../desktopSharedGrid";

function createMockCanvasContext(canvas: HTMLCanvasElement): CanvasRenderingContext2D {
  const noop = () => {};
  const gradient = { addColorStop: noop } as any;
  const context = new Proxy(
    {
      canvas,
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

function createPointerEvent(type: string, options: { clientX: number; clientY: number; pointerId: number }): Event {
  const event = new MouseEvent(type, {
    bubbles: true,
    cancelable: true,
    clientX: options.clientX,
    clientY: options.clientY
  });
  Object.defineProperty(event, "pointerId", { value: options.pointerId });
  return event;
}

describe("DesktopSharedGrid resize hit testing", () => {
  const originalGetContext = HTMLCanvasElement.prototype.getContext;

  beforeEach(() => {
    document.body.innerHTML = "";

    vi.stubGlobal("requestAnimationFrame", vi.fn(() => 0));

    Object.defineProperty(HTMLCanvasElement.prototype, "getContext", {
      configurable: true,
      writable: true,
      value: function (this: HTMLCanvasElement) {
        return createMockCanvasContext(this);
      }
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

  it("only enables resize cursors within headerRows/headerCols (not extra frozen panes)", () => {
    const rowCount = 100;
    const colCount = 100;

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
    vTrack.appendChild(vThumb);
    hTrack.appendChild(hThumb);
    container.appendChild(vTrack);
    container.appendChild(hTrack);

    vi.spyOn(selectionCanvas, "getBoundingClientRect").mockReturnValue({
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

    const grid = new DesktopSharedGrid({
      container,
      provider: new MockCellProvider({ rowCount, colCount }),
      rowCount,
      colCount,
      canvases: { grid: gridCanvas, content: contentCanvas, selection: selectionCanvas },
      scrollbars: { vTrack, vThumb, hTrack, hThumb },
      frozenRows: 1,
      frozenCols: 1,
      defaultRowHeight: 20,
      defaultColWidth: 50,
      enableResize: true,
      enableKeyboard: false,
      enableWheel: false
    });

    grid.resize(300, 200, 1);
    grid.renderer.setFrozen(3, 3);

    // Control: inside the true column header row (y < headerHeight) near a column boundary.
    selectionCanvas.dispatchEvent(createPointerEvent("pointermove", { clientX: 100, clientY: 10, pointerId: 1 }));
    expect(selectionCanvas.style.cursor).toBe("col-resize");

    // Frozen but below the header row should not show a column resize cursor.
    selectionCanvas.dispatchEvent(createPointerEvent("pointermove", { clientX: 100, clientY: 30, pointerId: 1 }));
    expect(selectionCanvas.style.cursor).toBe("default");

    // Control: inside the true row header column (x < headerWidth) near a row boundary.
    selectionCanvas.dispatchEvent(createPointerEvent("pointermove", { clientX: 10, clientY: 40, pointerId: 1 }));
    expect(selectionCanvas.style.cursor).toBe("row-resize");

    // Frozen but to the right of the row header column should not show a row resize cursor.
    selectionCanvas.dispatchEvent(createPointerEvent("pointermove", { clientX: 60, clientY: 40, pointerId: 1 }));
    expect(selectionCanvas.style.cursor).toBe("default");

    grid.destroy();
    container.remove();
  });
});
