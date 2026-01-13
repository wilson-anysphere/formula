// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { MockCellProvider } from "@formula/grid";
import { DesktopSharedGrid } from "../desktopSharedGrid";

function createMock2dContext(canvas: HTMLCanvasElement): CanvasRenderingContext2D {
  const noop = () => {};

  // DesktopSharedGrid relies on CanvasGridRenderer, which touches a broad surface
  // area of the 2D canvas context. For scrollbar unit tests, a no-op context is
  // sufficient as long as the used methods/properties exist.
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
    setLineDash: noop,
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

function createPointerEvent(
  type: string,
  options: { pointerId: number; clientX?: number; clientY?: number }
): MouseEvent {
  const event = new MouseEvent(type, {
    bubbles: true,
    cancelable: true,
    clientX: options.clientX ?? 0,
    clientY: options.clientY ?? 0
  });
  Object.defineProperty(event, "pointerId", { value: options.pointerId });
  return event;
}

describe("DesktopSharedGrid scrollbar thumb pointercancel", () => {
  const originalGetContext = HTMLCanvasElement.prototype.getContext;

  beforeEach(() => {
    vi.stubGlobal("requestAnimationFrame", (cb: FrameRequestCallback) => {
      cb(0);
      return 0;
    });

    HTMLCanvasElement.prototype.getContext = vi.fn(function (this: HTMLCanvasElement) {
      return createMock2dContext(this);
    }) as unknown as typeof HTMLCanvasElement.prototype.getContext;
  });

  afterEach(() => {
    HTMLCanvasElement.prototype.getContext = originalGetContext;
    vi.unstubAllGlobals();
  });

  function createGrid(options: {
    rowCount: number;
    colCount: number;
    defaultRowHeight?: number;
    defaultColWidth?: number;
  }): {
    grid: DesktopSharedGrid;
    container: HTMLDivElement;
    scrollbars: { vTrack: HTMLDivElement; vThumb: HTMLDivElement; hTrack: HTMLDivElement; hThumb: HTMLDivElement };
  } {
    const { rowCount, colCount } = options;
    const provider = new MockCellProvider({ rowCount, colCount });

    const container = document.createElement("div");
    document.body.appendChild(container);

    const canvases = {
      grid: document.createElement("canvas"),
      content: document.createElement("canvas"),
      selection: document.createElement("canvas")
    };

    // Mirror production DOM structure: canvases and scrollbars live inside the container.
    container.appendChild(canvases.grid);
    container.appendChild(canvases.content);
    container.appendChild(canvases.selection);

    const scrollbars = {
      vTrack: document.createElement("div"),
      vThumb: document.createElement("div"),
      hTrack: document.createElement("div"),
      hThumb: document.createElement("div")
    };

    scrollbars.vTrack.appendChild(scrollbars.vThumb);
    scrollbars.hTrack.appendChild(scrollbars.hThumb);
    container.appendChild(scrollbars.vTrack);
    container.appendChild(scrollbars.hTrack);

    const grid = new DesktopSharedGrid({
      container,
      provider,
      rowCount,
      colCount,
      canvases,
      scrollbars,
      defaultRowHeight: options.defaultRowHeight,
      defaultColWidth: options.defaultColWidth,
      enableWheel: false,
      enableKeyboard: false,
      enableResize: false
    });

    return { grid, container, scrollbars };
  }

  it("cleans up vertical thumb drag listeners on pointercancel", () => {
    const { grid, container, scrollbars } = createGrid({
      rowCount: 100,
      colCount: 10,
      defaultRowHeight: 10,
      defaultColWidth: 10
    });

    grid.resize(200, 200, 1);
    expect(grid.renderer.getViewportState().maxScrollY).toBeGreaterThan(0);

    vi.spyOn(scrollbars.vTrack, "getBoundingClientRect").mockReturnValue({
      left: 0,
      top: 0,
      right: 10,
      bottom: 200,
      width: 10,
      height: 200,
      x: 0,
      y: 0,
      toJSON: () => {}
    } as DOMRect);

    vi.spyOn(scrollbars.vThumb, "getBoundingClientRect").mockReturnValue({
      left: 0,
      top: 0,
      right: 10,
      bottom: 20,
      width: 10,
      height: 20,
      x: 0,
      y: 0,
      toJSON: () => {}
    } as DOMRect);

    scrollbars.vThumb.dispatchEvent(createPointerEvent("pointerdown", { pointerId: 1, clientY: 0 }));
    window.dispatchEvent(createPointerEvent("pointermove", { pointerId: 1, clientY: 60 }));

    const afterMove = grid.getScroll().y;
    expect(afterMove).toBeGreaterThan(0);

    window.dispatchEvent(createPointerEvent("pointercancel", { pointerId: 1, clientY: 60 }));
    window.dispatchEvent(createPointerEvent("pointermove", { pointerId: 1, clientY: 120 }));

    expect(grid.getScroll().y).toBe(afterMove);

    grid.destroy();
    container.remove();
  });

  it("ignores pointercancel from other pointers during thumb drags", () => {
    const { grid, container, scrollbars } = createGrid({
      rowCount: 100,
      colCount: 10,
      defaultRowHeight: 10,
      defaultColWidth: 10
    });

    grid.resize(200, 200, 1);
    expect(grid.renderer.getViewportState().maxScrollY).toBeGreaterThan(0);

    vi.spyOn(scrollbars.vTrack, "getBoundingClientRect").mockReturnValue({
      left: 0,
      top: 0,
      right: 10,
      bottom: 200,
      width: 10,
      height: 200,
      x: 0,
      y: 0,
      toJSON: () => {}
    } as DOMRect);

    vi.spyOn(scrollbars.vThumb, "getBoundingClientRect").mockReturnValue({
      left: 0,
      top: 0,
      right: 10,
      bottom: 20,
      width: 10,
      height: 20,
      x: 0,
      y: 0,
      toJSON: () => {}
    } as DOMRect);

    scrollbars.vThumb.dispatchEvent(createPointerEvent("pointerdown", { pointerId: 1, clientY: 0 }));
    window.dispatchEvent(createPointerEvent("pointermove", { pointerId: 1, clientY: 60 }));

    const afterMove = grid.getScroll().y;
    expect(afterMove).toBeGreaterThan(0);

    window.dispatchEvent(createPointerEvent("pointercancel", { pointerId: 2, clientY: 60 }));
    window.dispatchEvent(createPointerEvent("pointermove", { pointerId: 1, clientY: 120 }));

    expect(grid.getScroll().y).toBeGreaterThan(afterMove);

    // End the drag to avoid leaving window listeners installed if this test fails.
    window.dispatchEvent(createPointerEvent("pointerup", { pointerId: 1, clientY: 120 }));

    grid.destroy();
    container.remove();
  });

  it("cleans up horizontal thumb drag listeners on pointercancel", () => {
    const { grid, container, scrollbars } = createGrid({
      rowCount: 10,
      colCount: 100,
      defaultRowHeight: 10,
      defaultColWidth: 10
    });

    grid.resize(200, 200, 1);
    expect(grid.renderer.getViewportState().maxScrollX).toBeGreaterThan(0);

    vi.spyOn(scrollbars.hTrack, "getBoundingClientRect").mockReturnValue({
      left: 0,
      top: 0,
      right: 200,
      bottom: 10,
      width: 200,
      height: 10,
      x: 0,
      y: 0,
      toJSON: () => {}
    } as DOMRect);

    vi.spyOn(scrollbars.hThumb, "getBoundingClientRect").mockReturnValue({
      left: 0,
      top: 0,
      right: 20,
      bottom: 10,
      width: 20,
      height: 10,
      x: 0,
      y: 0,
      toJSON: () => {}
    } as DOMRect);

    scrollbars.hThumb.dispatchEvent(createPointerEvent("pointerdown", { pointerId: 1, clientX: 0 }));
    window.dispatchEvent(createPointerEvent("pointermove", { pointerId: 1, clientX: 60 }));

    const afterMove = grid.getScroll().x;
    expect(afterMove).toBeGreaterThan(0);

    window.dispatchEvent(createPointerEvent("pointercancel", { pointerId: 1, clientX: 60 }));
    window.dispatchEvent(createPointerEvent("pointermove", { pointerId: 1, clientX: 120 }));

    expect(grid.getScroll().x).toBe(afterMove);

    grid.destroy();
    container.remove();
  });
});

