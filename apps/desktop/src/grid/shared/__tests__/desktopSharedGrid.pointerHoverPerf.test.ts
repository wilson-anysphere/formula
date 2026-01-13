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

function createPointerEvent(type: string, options: { clientX: number; clientY: number; offsetX: number; offsetY: number }): Event {
  const event = new MouseEvent(type, {
    bubbles: true,
    cancelable: true,
    clientX: options.clientX,
    clientY: options.clientY
  });
  Object.defineProperty(event, "offsetX", { value: options.offsetX });
  Object.defineProperty(event, "offsetY", { value: options.offsetY });
  return event;
}

describe("DesktopSharedGrid pointer hover perf", () => {
  const originalGetContext = HTMLCanvasElement.prototype.getContext;

  beforeEach(() => {
    document.body.innerHTML = "";

    vi.stubGlobal("requestAnimationFrame", (cb: FrameRequestCallback) => {
      cb(0);
      return 0;
    });

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

  it("does not call getBoundingClientRect on repeated hover pointermove events", () => {
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

    const rectSpy = vi.spyOn(selectionCanvas, "getBoundingClientRect").mockReturnValue({
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
      provider: new MockCellProvider({ rowCount: 100, colCount: 100 }),
      rowCount: 100,
      colCount: 100,
      canvases: { grid: gridCanvas, content: contentCanvas, selection: selectionCanvas },
      scrollbars: { vTrack, vThumb, hTrack, hThumb },
      enableResize: true,
      enableKeyboard: false,
      enableWheel: false
    });

    rectSpy.mockClear();

    for (let i = 0; i < 50; i++) {
      selectionCanvas.dispatchEvent(createPointerEvent("pointermove", { clientX: 120 + i, clientY: 80, offsetX: 120 + i, offsetY: 80 }));
    }

    expect(rectSpy).not.toHaveBeenCalled();

    grid.destroy();
    container.remove();
  });
});

