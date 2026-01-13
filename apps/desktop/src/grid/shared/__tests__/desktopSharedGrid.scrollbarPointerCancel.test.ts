// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { MockCellProvider } from "@formula/grid";
import { DesktopSharedGrid } from "../desktopSharedGrid";

function createPointerEvent(type: string, options: { clientX: number; clientY: number; pointerId: number }): Event {
  const PointerEventCtor = (window as unknown as { PointerEvent?: typeof PointerEvent }).PointerEvent;
  if (PointerEventCtor) {
    return new PointerEventCtor(type, {
      bubbles: true,
      cancelable: true,
      clientX: options.clientX,
      clientY: options.clientY,
      buttons: 1,
      pointerId: options.pointerId
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
  return event;
}

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
  }): { grid: DesktopSharedGrid; container: HTMLDivElement } {
    const { rowCount, colCount } = options;
    const provider = new MockCellProvider({ rowCount, colCount });

    const container = document.createElement("div");
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

    return { grid, container };
  }

  it("stops scrolling the vertical thumb drag after pointercancel", () => {
    const { grid, container } = createGrid({
      rowCount: 100,
      colCount: 10,
      defaultRowHeight: 10,
      defaultColWidth: 10
    });

    grid.resize(200, 200, 1);

    const trackRect = {
      left: 0,
      top: 0,
      right: 10,
      bottom: 200,
      width: 10,
      height: 200,
      x: 0,
      y: 0,
      toJSON: () => ({})
    } as unknown as DOMRect;

    const thumbRect = {
      left: 0,
      top: 0,
      right: 10,
      bottom: 40,
      width: 10,
      height: 40,
      x: 0,
      y: 0,
      toJSON: () => ({})
    } as unknown as DOMRect;

    const gridInternals = grid as unknown as { vTrack: HTMLDivElement; vThumb: HTMLDivElement };
    gridInternals.vTrack.getBoundingClientRect = vi.fn(() => trackRect);
    gridInternals.vThumb.getBoundingClientRect = vi.fn(() => thumbRect);

    expect(grid.getScroll().y).toBe(0);

    gridInternals.vThumb.dispatchEvent(createPointerEvent("pointerdown", { clientX: 0, clientY: 10, pointerId: 1 }));
    window.dispatchEvent(createPointerEvent("pointermove", { clientX: 0, clientY: 60, pointerId: 1 }));

    const duringDrag = grid.getScroll().y;
    expect(duringDrag).toBeGreaterThan(0);

    window.dispatchEvent(createPointerEvent("pointercancel", { clientX: 0, clientY: 60, pointerId: 1 }));
    const afterCancel = grid.getScroll().y;

    window.dispatchEvent(createPointerEvent("pointermove", { clientX: 0, clientY: 160, pointerId: 1 }));
    expect(grid.getScroll().y).toBe(afterCancel);

    grid.destroy();
    container.remove();
  });
});

