// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { MockCellProvider } from "@formula/grid";
import { DesktopSharedGrid } from "../desktopSharedGrid";

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
        actualBoundingBoxDescent: 2,
      }) as TextMetrics,
  } as unknown as CanvasRenderingContext2D;
}

describe("DesktopSharedGrid destroy()", () => {
  const originalGetContext = HTMLCanvasElement.prototype.getContext;

  beforeEach(() => {
    vi.stubGlobal("requestAnimationFrame", (cb: FrameRequestCallback) => {
      cb(0);
      return 0;
    });
    vi.stubGlobal("cancelAnimationFrame", () => {});

    HTMLCanvasElement.prototype.getContext = vi.fn(function (this: HTMLCanvasElement) {
      return createMock2dContext(this);
    }) as unknown as typeof HTMLCanvasElement.prototype.getContext;
  });

  afterEach(() => {
    HTMLCanvasElement.prototype.getContext = originalGetContext;
    vi.unstubAllGlobals();
    document.body.innerHTML = "";
  });

  it("resets canvas backing stores to 0x0 on destroy", () => {
    const container = document.createElement("div");
    document.body.appendChild(container);

    const provider = new MockCellProvider({ rowCount: 10, colCount: 10 });
    const canvases = {
      grid: document.createElement("canvas"),
      content: document.createElement("canvas"),
      selection: document.createElement("canvas"),
    };
    const scrollbars = {
      vTrack: document.createElement("div"),
      vThumb: document.createElement("div"),
      hTrack: document.createElement("div"),
      hThumb: document.createElement("div"),
    };

    const grid = new DesktopSharedGrid({
      container,
      provider,
      rowCount: 10,
      colCount: 10,
      canvases,
      scrollbars,
    });

    grid.resize(300, 200, 1);
    expect(canvases.grid.width).toBeGreaterThan(0);
    expect(canvases.grid.height).toBeGreaterThan(0);
    expect(canvases.content.width).toBeGreaterThan(0);
    expect(canvases.content.height).toBeGreaterThan(0);
    expect(canvases.selection.width).toBeGreaterThan(0);
    expect(canvases.selection.height).toBeGreaterThan(0);

    grid.destroy();
    expect(canvases.grid.width).toBe(0);
    expect(canvases.grid.height).toBe(0);
    expect(canvases.content.width).toBe(0);
    expect(canvases.content.height).toBe(0);
    expect(canvases.selection.width).toBe(0);
    expect(canvases.selection.height).toBe(0);
  });
});

