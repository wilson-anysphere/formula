// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { MockCellProvider } from "@formula/grid";
import * as gridExports from "@formula/grid";
import { DesktopSharedGrid } from "../desktopSharedGrid";

function createMock2dContext(canvas: HTMLCanvasElement): CanvasRenderingContext2D {
  const noop = () => {};

  // DesktopSharedGrid relies on CanvasGridRenderer, which touches a broad surface
  // area of the 2D canvas context. For scroll/selection unit tests, a no-op
  // context is sufficient as long as the used methods/properties exist.
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

describe("DesktopSharedGrid scrollbars", () => {
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
    frozenRows?: number;
    frozenCols?: number;
    defaultRowHeight?: number;
    defaultColWidth?: number;
    enableWheel?: boolean;
  }): { grid: DesktopSharedGrid; container: HTMLDivElement; selectionCanvas: HTMLCanvasElement } {
    const { rowCount, colCount } = options;
    const provider = new MockCellProvider({ rowCount, colCount });

    const container = document.createElement("div");
    document.body.appendChild(container);

    const canvases = {
      grid: document.createElement("canvas"),
      content: document.createElement("canvas"),
      selection: document.createElement("canvas")
    };

    // Mirror production DOM structure: the grid canvases live inside the container so wheel events
    // bubble to the container listener.
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
      frozenRows: options.frozenRows,
      frozenCols: options.frozenCols,
      defaultRowHeight: options.defaultRowHeight,
      defaultColWidth: options.defaultColWidth,
      enableWheel: options.enableWheel ?? false,
      enableKeyboard: false,
      enableResize: false
    });

    return { grid, container, selectionCanvas: canvases.selection };
  }

  it("avoids getBoundingClientRect during scroll-driven scrollbar sync", () => {
    const rectSpy = vi.spyOn(HTMLElement.prototype, "getBoundingClientRect");
    const { grid, container } = createGrid({
      rowCount: 100,
      colCount: 100,
      defaultRowHeight: 10,
      defaultColWidth: 10
    });

    grid.resize(300, 200, 1);
    rectSpy.mockClear();

    grid.scrollBy(10, 10);

    expect(rectSpy).not.toHaveBeenCalled();

    grid.destroy();
    container.remove();
  });

  it("does not compute thumbs when no scrollbars are needed", () => {
    const thumbSpy = vi.spyOn(gridExports, "computeScrollbarThumb");
    const { grid, container } = createGrid({
      rowCount: 10,
      colCount: 10,
      defaultRowHeight: 10,
      defaultColWidth: 10
    });

    grid.resize(300, 200, 1);
    thumbSpy.mockClear();

    // Trigger a sync after resize.
    grid.scrollBy(0, 0);

    expect(thumbSpy).not.toHaveBeenCalled();

    grid.destroy();
    container.remove();
  });

  it("passes correct thumb inputs for vertical-only scroll", () => {
    const thumbSpy = vi.spyOn(gridExports, "computeScrollbarThumb");
    const { grid, container } = createGrid({
      rowCount: 100,
      colCount: 10,
      defaultRowHeight: 10,
      defaultColWidth: 10
    });

    grid.resize(300, 200, 1);
    const viewport = grid.renderer.getViewportState();
    expect(viewport.maxScrollY).toBeGreaterThan(0);
    expect(viewport.maxScrollX).toBe(0);

    thumbSpy.mockClear();
    grid.scrollBy(0, 0);

    expect(thumbSpy).toHaveBeenCalledTimes(1);
    const args = thumbSpy.mock.calls[0]?.[0];
    expect(args).toBeTruthy();

    const padding = 2;
    const thickness = 10;
    const showH = viewport.maxScrollX > 0;
    const frozenHeight = Math.min(viewport.frozenHeight, viewport.height);

    expect(args?.trackSize).toBeCloseTo(Math.max(0, viewport.height - frozenHeight - (showH ? thickness : 0) - 2 * padding), 6);
    expect(args?.viewportSize).toBeCloseTo(Math.max(0, viewport.height - frozenHeight), 6);
    expect(args?.contentSize).toBeCloseTo(Math.max(0, viewport.totalHeight - frozenHeight), 6);

    grid.destroy();
    container.remove();
  });

  it("passes correct thumb inputs for both-axis scroll", () => {
    const thumbSpy = vi.spyOn(gridExports, "computeScrollbarThumb");
    const { grid, container } = createGrid({
      rowCount: 100,
      colCount: 100,
      defaultRowHeight: 10,
      defaultColWidth: 10
    });

    grid.resize(300, 200, 1);
    const viewport = grid.renderer.getViewportState();
    expect(viewport.maxScrollY).toBeGreaterThan(0);
    expect(viewport.maxScrollX).toBeGreaterThan(0);

    thumbSpy.mockClear();
    grid.scrollBy(0, 0);

    expect(thumbSpy).toHaveBeenCalledTimes(2);
    const vArgs = thumbSpy.mock.calls[0]?.[0];
    const hArgs = thumbSpy.mock.calls[1]?.[0];

    const padding = 2;
    const thickness = 10;
    const showV = viewport.maxScrollY > 0;
    const showH = viewport.maxScrollX > 0;
    const frozenHeight = Math.min(viewport.frozenHeight, viewport.height);
    const frozenWidth = Math.min(viewport.frozenWidth, viewport.width);

    expect(vArgs?.trackSize).toBeCloseTo(Math.max(0, viewport.height - frozenHeight - (showH ? thickness : 0) - 2 * padding), 6);
    expect(vArgs?.viewportSize).toBeCloseTo(Math.max(0, viewport.height - frozenHeight), 6);
    expect(vArgs?.contentSize).toBeCloseTo(Math.max(0, viewport.totalHeight - frozenHeight), 6);

    expect(hArgs?.trackSize).toBeCloseTo(Math.max(0, viewport.width - frozenWidth - (showV ? thickness : 0) - 2 * padding), 6);
    expect(hArgs?.viewportSize).toBeCloseTo(Math.max(0, viewport.width - frozenWidth), 6);
    expect(hArgs?.contentSize).toBeCloseTo(Math.max(0, viewport.totalWidth - frozenWidth), 6);

    grid.destroy();
    container.remove();
  });

  it("passes correct thumb inputs with frozen panes", () => {
    const thumbSpy = vi.spyOn(gridExports, "computeScrollbarThumb");
    const { grid, container } = createGrid({
      rowCount: 100,
      colCount: 100,
      frozenRows: 1,
      frozenCols: 1,
      defaultRowHeight: 10,
      defaultColWidth: 10
    });

    grid.resize(300, 200, 1);
    const viewport = grid.renderer.getViewportState();
    expect(viewport.frozenHeight).toBeGreaterThan(0);
    expect(viewport.frozenWidth).toBeGreaterThan(0);

    thumbSpy.mockClear();
    grid.scrollBy(0, 0);

    expect(thumbSpy).toHaveBeenCalledTimes(2);
    const vArgs = thumbSpy.mock.calls[0]?.[0];
    const hArgs = thumbSpy.mock.calls[1]?.[0];

    const padding = 2;
    const thickness = 10;
    const showV = viewport.maxScrollY > 0;
    const showH = viewport.maxScrollX > 0;
    const frozenHeight = Math.min(viewport.frozenHeight, viewport.height);
    const frozenWidth = Math.min(viewport.frozenWidth, viewport.width);

    expect(vArgs?.trackSize).toBeCloseTo(Math.max(0, viewport.height - frozenHeight - (showH ? thickness : 0) - 2 * padding), 6);
    expect(vArgs?.viewportSize).toBeCloseTo(Math.max(0, viewport.height - frozenHeight), 6);
    expect(vArgs?.contentSize).toBeCloseTo(Math.max(0, viewport.totalHeight - frozenHeight), 6);

    expect(hArgs?.trackSize).toBeCloseTo(Math.max(0, viewport.width - frozenWidth - (showV ? thickness : 0) - 2 * padding), 6);
    expect(hArgs?.viewportSize).toBeCloseTo(Math.max(0, viewport.width - frozenWidth), 6);
    expect(hArgs?.contentSize).toBeCloseTo(Math.max(0, viewport.totalWidth - frozenWidth), 6);

    grid.destroy();
    container.remove();
  });

  it("uses offsetX/offsetY for ctrl+wheel zoom without reading layout", () => {
    const { grid, container, selectionCanvas } = createGrid({
      rowCount: 100,
      colCount: 100,
      defaultRowHeight: 10,
      defaultColWidth: 10,
      enableWheel: true
    });

    grid.resize(300, 200, 1);

    const rectSpy = vi.spyOn(selectionCanvas, "getBoundingClientRect");
    rectSpy.mockClear();

    const event = new WheelEvent("wheel", { deltaY: -100, ctrlKey: true, bubbles: true, cancelable: true });
    Object.defineProperty(event, "offsetX", { value: 100 });
    Object.defineProperty(event, "offsetY", { value: 50 });
    selectionCanvas.dispatchEvent(event);

    expect(grid.getZoom()).toBeGreaterThan(1);
    expect(rectSpy).not.toHaveBeenCalled();

    grid.destroy();
    container.remove();
  });
});
