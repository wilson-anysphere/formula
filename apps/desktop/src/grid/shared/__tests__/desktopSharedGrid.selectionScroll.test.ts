// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { MockCellProvider, type CellRange } from "@formula/grid";
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

describe("DesktopSharedGrid selection scrollIntoView option", () => {
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

  it("keeps scroll stable when scrollIntoView is false, but scrolls by default", () => {
    const rowCount = 200;
    const colCount = 200;
    const provider = new MockCellProvider({ rowCount, colCount });

    const onScroll = vi.fn();
    const onSelectionChange = vi.fn();
    const onSelectionRangeChange = vi.fn();

    const container = document.createElement("div");
    document.body.appendChild(container);

    const canvases = {
      grid: document.createElement("canvas"),
      content: document.createElement("canvas"),
      selection: document.createElement("canvas")
    };

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
      callbacks: { onScroll, onSelectionChange, onSelectionRangeChange }
    });

    grid.resize(300, 200, 1);
    grid.scrollTo(50, 120);
    onScroll.mockClear();

    const ranges: CellRange[] = [{ startRow: 80, endRow: 81, startCol: 40, endCol: 41 }];

    const scrollToCellSpy = vi.spyOn(grid.renderer, "scrollToCell");
    grid.setSelectionRanges(ranges, { scrollIntoView: false });

    expect(scrollToCellSpy).not.toHaveBeenCalled();
    expect(onScroll).not.toHaveBeenCalled();
    expect(grid.getScroll()).toEqual({ x: 50, y: 120 });

    expect(onSelectionChange).toHaveBeenCalledWith({ row: 80, col: 40 });
    expect(onSelectionRangeChange).toHaveBeenCalledWith(ranges[0]);

    const status = container.querySelector('[data-testid="canvas-grid-a11y-status"]');
    expect(status?.textContent).toContain("Active cell");

    onScroll.mockClear();
    grid.setSelectionRanges(ranges);

    expect(scrollToCellSpy).toHaveBeenCalled();
    const nextScroll = grid.getScroll();
    expect(nextScroll.x).toBeGreaterThan(50);
    expect(nextScroll.y).toBeGreaterThan(120);
    expect(onScroll).toHaveBeenCalled();

    grid.destroy();
    container.remove();
  });

  it("does not emit onScroll when scroll operations are a no-op (but does when viewport changes)", () => {
    const rowCount = 50;
    const colCount = 50;
    const provider = new MockCellProvider({ rowCount, colCount });

    const onScroll = vi.fn();

    const container = document.createElement("div");
    document.body.appendChild(container);

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
      callbacks: { onScroll },
    });

    grid.resize(300, 200, 1);
    onScroll.mockClear();

    // Cell 0,0 is already visible at scroll origin; ensure no onScroll is emitted.
    grid.scrollToCell(0, 0, { align: "auto", padding: 0 });
    expect(onScroll).not.toHaveBeenCalled();

    // Changing viewport state (e.g. frozen panes) should still emit onScroll even when scroll offsets remain unchanged.
    grid.renderer.setFrozen(1, 1);
    grid.scrollTo(0, 0);
    expect(onScroll).toHaveBeenCalled();

    grid.destroy();
    container.remove();
  });

  it("emits onScroll when axis size changes without scrolling", () => {
    const rowCount = 50;
    const colCount = 50;
    const provider = new MockCellProvider({ rowCount, colCount });

    const onScroll = vi.fn();

    const container = document.createElement("div");
    document.body.appendChild(container);

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
      callbacks: { onScroll },
    });

    grid.resize(300, 200, 1);
    onScroll.mockClear();

    const before = grid.getScroll();
    const beforeViewport = grid.renderer.getViewportState();
    const beforeMaxScrollY = beforeViewport.maxScrollY;

    grid.renderer.setRowHeight(0, grid.renderer.getRowHeight(0) + 10);

    expect(grid.getScroll()).toEqual(before);
    const afterMaxScrollY = grid.renderer.getViewportState().maxScrollY;
    expect(afterMaxScrollY).not.toBe(beforeMaxScrollY);
    expect(onScroll).toHaveBeenCalled();

    grid.destroy();
    container.remove();
  });

  it("does not re-sync scrollbars when scrollToCell is a no-op", () => {
    const rowCount = 100;
    const colCount = 100;
    const provider = new MockCellProvider({ rowCount, colCount });

    const onScroll = vi.fn();

    const container = document.createElement("div");
    document.body.appendChild(container);

    const canvases = {
      grid: document.createElement("canvas"),
      content: document.createElement("canvas"),
      selection: document.createElement("canvas")
    };

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
      callbacks: { onScroll },
    });

    const syncSpy = vi.spyOn(grid, "syncScrollbars");

    grid.resize(300, 200, 1);
    grid.scrollTo(50, 120);
    syncSpy.mockClear();
    onScroll.mockClear();

    const viewport = grid.renderer.getViewportState();
    const row = Math.max(0, Math.min(rowCount - 1, viewport.main.rows.start + 1));
    const col = Math.max(0, Math.min(colCount - 1, viewport.main.cols.start + 1));

    grid.scrollToCell(row, col, { align: "auto", padding: 0 });

    expect(syncSpy).not.toHaveBeenCalled();
    expect(onScroll).not.toHaveBeenCalled();

    grid.destroy();
    container.remove();
  });

  it("avoids scrollbar sync work when scrollBy hits a scroll boundary", () => {
    const rowCount = 100;
    const colCount = 100;
    const provider = new MockCellProvider({ rowCount, colCount });

    const onScroll = vi.fn();

    const container = document.createElement("div");
    document.body.appendChild(container);

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
      callbacks: { onScroll },
    });

    grid.resize(300, 200, 1);
    const viewport = grid.renderer.getViewportState();
    grid.scrollTo(viewport.maxScrollX, viewport.maxScrollY);

    const syncSpy = vi.spyOn(grid, "syncScrollbars");
    syncSpy.mockClear();
    onScroll.mockClear();

    const before = grid.getScroll();
    grid.scrollBy(10_000, 10_000);

    expect(grid.getScroll()).toEqual(before);
    expect(syncSpy).not.toHaveBeenCalled();
    expect(onScroll).not.toHaveBeenCalled();

    grid.destroy();
    container.remove();
  });
});
