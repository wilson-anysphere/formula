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
  }): {
    grid: DesktopSharedGrid;
    container: HTMLDivElement;
    selectionCanvas: HTMLCanvasElement;
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

    return { grid, container, selectionCanvas: canvases.selection, scrollbars };
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

  it("avoids scroll/scrollbar sync work when ctrl+wheel zoom is clamped (no zoom change)", () => {
    const { grid, container, selectionCanvas } = createGrid({
      rowCount: 100,
      colCount: 100,
      defaultRowHeight: 10,
      defaultColWidth: 10,
      enableWheel: true
    });

    grid.resize(300, 200, 1);
    grid.setZoom(4);
    expect(grid.getZoom()).toBe(4);

    const syncSpy = vi.spyOn(grid, "syncScrollbars");
    syncSpy.mockClear();

    const event = new WheelEvent("wheel", { deltaY: -100, ctrlKey: true, bubbles: true, cancelable: true });
    Object.defineProperty(event, "offsetX", { value: 120 });
    Object.defineProperty(event, "offsetY", { value: 60 });
    selectionCanvas.dispatchEvent(event);

    // Zoom remains clamped; wheel handler should bail out before doing scrollbar/scroll sync work.
    expect(grid.getZoom()).toBe(4);
    expect(syncSpy).not.toHaveBeenCalled();

    grid.destroy();
    container.remove();
  });

  it("avoids calling renderer.setScroll when scrollTo aligns to the current device-pixel scroll", () => {
    const { grid, container } = createGrid({
      rowCount: 100,
      colCount: 100,
      defaultRowHeight: 10,
      defaultColWidth: 10,
    });

    // Use a HiDPI device pixel ratio so scroll alignment happens in 0.5px steps.
    grid.resize(300, 200, 2);

    grid.scrollTo(0.5, 0.5);
    expect(grid.getScroll()).toEqual({ x: 0.5, y: 0.5 });

    const spy = vi.spyOn(grid.renderer, "setScroll");
    spy.mockClear();

    // 0.6px rounds back to 0.5px at dpr=2, so this should be a no-op (and should not trigger
    // an unnecessary renderer scroll invalidation).
    grid.scrollTo(0.6, 0.6);
    expect(grid.getScroll()).toEqual({ x: 0.5, y: 0.5 });
    expect(spy).not.toHaveBeenCalled();

    grid.destroy();
    container.remove();
  });

  it("uses zoom-scaled minimum thumb size when dragging scrollbars", () => {
    const { grid, container, scrollbars } = createGrid({
      rowCount: 100,
      colCount: 10,
      defaultRowHeight: 10,
      defaultColWidth: 10
    });

    grid.resize(200, 200, 1);
    grid.setZoom(2);

    // Thumb size should respect the zoom-scaled minimum (24 * zoom).
    expect(scrollbars.vThumb.style.height).toBe("48px");

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
      bottom: 48,
      width: 10,
      height: 48,
      x: 0,
      y: 0,
      toJSON: () => {}
    } as DOMRect);

    const createPointerEvent = (type: string, options: { pointerId: number; clientY: number; clientX?: number }) => {
      const event = new MouseEvent(type, {
        bubbles: true,
        cancelable: true,
        clientX: options.clientX ?? 0,
        clientY: options.clientY
      });
      Object.defineProperty(event, "pointerId", { value: options.pointerId });
      return event;
    };

    scrollbars.vThumb.dispatchEvent(createPointerEvent("pointerdown", { pointerId: 1, clientY: 0 }));
    window.dispatchEvent(createPointerEvent("pointermove", { pointerId: 1, clientY: 152 }));
    window.dispatchEvent(createPointerEvent("pointerup", { pointerId: 1, clientY: 152 }));

    expect(grid.getScroll().y).toBe(1800);

    grid.destroy();
    container.remove();
  });

  it("updates scrollbar thumb sizes after renderer.applyAxisSizeOverrides without requiring scrollTo/scrollBy", () => {
    const { grid, container, scrollbars } = createGrid({
      rowCount: 100,
      colCount: 100,
      defaultRowHeight: 10,
      defaultColWidth: 10
    });

    grid.resize(300, 200, 1);

    const beforeThumbHeight = parseFloat(scrollbars.vThumb.style.height);
    expect(Number.isFinite(beforeThumbHeight)).toBe(true);

    const scrollToSpy = vi.spyOn(grid, "scrollTo");
    const scrollBySpy = vi.spyOn(grid, "scrollBy");
    scrollToSpy.mockClear();
    scrollBySpy.mockClear();

    const rowOverrides = new Map<number, number>();
    for (let row = 0; row < 50; row++) rowOverrides.set(row, 20);
    grid.renderer.applyAxisSizeOverrides({ rows: rowOverrides });

    expect(scrollToSpy).not.toHaveBeenCalled();
    expect(scrollBySpy).not.toHaveBeenCalled();

    const afterThumbHeight = parseFloat(scrollbars.vThumb.style.height);
    expect(afterThumbHeight).toBeLessThan(beforeThumbHeight);

    grid.destroy();
    container.remove();
  });
});
