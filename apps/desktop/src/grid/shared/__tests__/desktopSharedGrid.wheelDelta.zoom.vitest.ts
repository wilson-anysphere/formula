// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { MockCellProvider } from "@formula/grid";
import { DesktopSharedGrid } from "../desktopSharedGrid";

function createMock2dContext(canvas: HTMLCanvasElement): CanvasRenderingContext2D {
  const noop = () => {};

  // DesktopSharedGrid relies on CanvasGridRenderer, which touches a broad surface
  // area of the 2D canvas context. For wheel/scroll unit tests, a no-op context is
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

describe("DesktopSharedGrid wheel delta normalization", () => {
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

  it("scales DOM_DELTA_LINE deltas by zoom (matching React CanvasGrid)", () => {
    const rowCount = 100;
    const colCount = 20;
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
      enableWheel: true,
      enableKeyboard: false,
      enableResize: false
    });

    grid.resize(200, 200, 1);
    grid.setZoom(2);

    container.dispatchEvent(new WheelEvent("wheel", { deltaY: 1, deltaMode: 1, bubbles: true, cancelable: true }));

    expect(grid.getScroll().y).toBeCloseTo(32, 6);

    grid.destroy();
    container.remove();
  });
});

