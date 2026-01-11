// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { CanvasGridRenderer } from "../CanvasGridRenderer";
import type { CellProvider } from "../../model/CellProvider";

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
    save: noop,
    restore: noop,
    drawImage: noop,
    translate: noop,
    rotate: noop,
    fillText: noop,
    setLineDash: noop,
    measureText: (text: string) =>
      ({
        width: text.length * 6,
        actualBoundingBoxAscent: 8,
        actualBoundingBoxDescent: 2
      }) as TextMetrics
  } as unknown as CanvasRenderingContext2D;
}

describe("CanvasGridRenderer zoom", () => {
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

  it("scales default sizes, preserves overrides, and anchors scroll", () => {
    const provider: CellProvider = {
      getCell: (row, col) => ({ row, col, value: null })
    };

    const renderer = new CanvasGridRenderer({ provider, rowCount: 100, colCount: 100, defaultRowHeight: 10, defaultColWidth: 20 });
    const grid = document.createElement("canvas");
    const content = document.createElement("canvas");
    const selection = document.createElement("canvas");
    renderer.attach({ grid, content, selection });
    renderer.resize(200, 200, 1);

    const before = renderer.getCellRect(0, 0);
    expect(before).not.toBeNull();
    expect(before!.width).toBe(20);
    expect(before!.height).toBe(10);

    renderer.setColWidth(1, 50);
    expect(renderer.getColWidth(1)).toBe(50);

    renderer.setScroll(40, 0);

    renderer.setZoom(2, { anchorX: 100, anchorY: 0 });
    expect(renderer.getZoom()).toBe(2);
    expect(renderer.getColWidth(1)).toBe(100);
    expect(renderer.scroll.getScroll().x).toBe(180);

    const after = renderer.getCellRect(0, 0);
    expect(after).not.toBeNull();
    expect(after!.width).toBe(40);
    expect(after!.height).toBe(20);
  });

  it("does not anchor scroll when zooming with an anchor inside the frozen quadrant", () => {
    const provider: CellProvider = { getCell: (row, col) => ({ row, col, value: null }) };

    const renderer = new CanvasGridRenderer({ provider, rowCount: 100, colCount: 100, defaultRowHeight: 10, defaultColWidth: 100 });
    const grid = document.createElement("canvas");
    const content = document.createElement("canvas");
    const selection = document.createElement("canvas");
    renderer.attach({ grid, content, selection });
    renderer.resize(200, 200, 1);
    renderer.setFrozen(0, 1);

    renderer.setScroll(200, 0);
    expect(renderer.scroll.getScroll().x).toBe(200);

    // Anchor inside the frozen column (frozenWidth=100 at zoom=1). Zooming out shrinks the frozen
    // width to 50; we still should not "anchor" because frozen quadrants do not scroll.
    renderer.setZoom(0.5, { anchorX: 50, anchorY: 0 });

    // Base scroll is 200; at 0.5 zoom it should scale to 100.
    expect(renderer.scroll.getScroll().x).toBe(100);
  });

  it("renders the selection fill handle scaled with zoom", () => {
    const provider: CellProvider = { getCell: () => null };

    const fillRects: Array<{ fillStyle: string; x: number; y: number; width: number; height: number }> = [];

    const gridCanvas = document.createElement("canvas");
    const contentCanvas = document.createElement("canvas");
    const selectionCanvas = document.createElement("canvas");

    const contexts = new Map<HTMLCanvasElement, CanvasRenderingContext2D>();
    contexts.set(gridCanvas, createMock2dContext(gridCanvas));
    contexts.set(contentCanvas, createMock2dContext(contentCanvas));

    const selectionCtx = createMock2dContext(selectionCanvas) as unknown as CanvasRenderingContext2D & {
      _fillStyle?: unknown;
    };
    let fillStyle: unknown = "#000";
    Object.defineProperty(selectionCtx, "fillStyle", {
      get: () => fillStyle,
      set: (value: unknown) => {
        fillStyle = value;
      }
    });
    selectionCtx.fillRect = (x: number, y: number, width: number, height: number) => {
      if (typeof fillStyle === "string") fillRects.push({ fillStyle, x, y, width, height });
    };
    contexts.set(selectionCanvas, selectionCtx);

    HTMLCanvasElement.prototype.getContext = vi.fn(function (this: HTMLCanvasElement) {
      return contexts.get(this) ?? createMock2dContext(this);
    }) as unknown as typeof HTMLCanvasElement.prototype.getContext;

    const renderer = new CanvasGridRenderer({ provider, rowCount: 50, colCount: 50, defaultRowHeight: 10, defaultColWidth: 10 });
    renderer.attach({ grid: gridCanvas, content: contentCanvas, selection: selectionCanvas });
    renderer.resize(200, 200, 1);

    renderer.setSelection({ row: 1, col: 1 });
    fillRects.length = 0;

    renderer.setZoom(2);

    const handleStyle = renderer.getTheme().selectionHandle;
    const handleRects = fillRects.filter((rect) => rect.fillStyle === handleStyle);
    expect(handleRects).toHaveLength(1);
    expect(handleRects[0]!.width).toBeCloseTo(16, 5);
    expect(handleRects[0]!.height).toBeCloseTo(16, 5);
  });

  it("hides the fill handle when the selection corner is offscreen", () => {
    const provider: CellProvider = { getCell: () => null };

    const renderer = new CanvasGridRenderer({ provider, rowCount: 50, colCount: 50, defaultRowHeight: 10, defaultColWidth: 100 });
    const gridCanvas = document.createElement("canvas");
    const contentCanvas = document.createElement("canvas");
    const selectionCanvas = document.createElement("canvas");
    renderer.attach({ grid: gridCanvas, content: contentCanvas, selection: selectionCanvas });
    renderer.resize(150, 150, 1);

    // With scrollX=0, col=1 spans x=[100,200) and its bottom-right corner is
    // outside the viewport (x=150). The fill handle should not be exposed.
    renderer.setSelection({ row: 0, col: 1 });
    expect(renderer.getFillHandleRect()).toBeNull();

    // Scroll so the corner is at the viewport edge; the handle becomes visible
    // (potentially clipped) and should report a viewport-contained rect.
    renderer.setScroll(50, 0);
    const handle = renderer.getFillHandleRect();
    expect(handle).not.toBeNull();
    expect(handle!.x).toBeGreaterThanOrEqual(0);
    expect(handle!.y).toBeGreaterThanOrEqual(0);
    expect(handle!.x + handle!.width).toBeLessThanOrEqual(150);
    expect(handle!.y + handle!.height).toBeLessThanOrEqual(150);
  });

  it("scales remote presence badge geometry with zoom", () => {
    const provider: CellProvider = { getCell: () => null };

    const fillRects: Array<{ fillStyle: string; x: number; y: number; width: number; height: number }> = [];

    const gridCanvas = document.createElement("canvas");
    const contentCanvas = document.createElement("canvas");
    const selectionCanvas = document.createElement("canvas");

    const contexts = new Map<HTMLCanvasElement, CanvasRenderingContext2D>();
    contexts.set(gridCanvas, createMock2dContext(gridCanvas));
    contexts.set(contentCanvas, createMock2dContext(contentCanvas));

    const selectionCtx = createMock2dContext(selectionCanvas) as unknown as CanvasRenderingContext2D;
    let fillStyle: unknown = "#000";
    Object.defineProperty(selectionCtx, "fillStyle", {
      get: () => fillStyle,
      set: (value: unknown) => {
        fillStyle = value;
      }
    });
    selectionCtx.fillRect = (x: number, y: number, width: number, height: number) => {
      if (typeof fillStyle === "string") fillRects.push({ fillStyle, x, y, width, height });
    };
    contexts.set(selectionCanvas, selectionCtx);

    HTMLCanvasElement.prototype.getContext = vi.fn(function (this: HTMLCanvasElement) {
      return contexts.get(this) ?? createMock2dContext(this);
    }) as unknown as typeof HTMLCanvasElement.prototype.getContext;

    const renderer = new CanvasGridRenderer({ provider, rowCount: 50, colCount: 50, defaultRowHeight: 10, defaultColWidth: 10 });
    renderer.attach({ grid: gridCanvas, content: contentCanvas, selection: selectionCanvas });
    renderer.resize(200, 200, 1);

    const color = "#ff0000";
    renderer.setRemotePresences([
      { id: "ada", name: "Ada", color, cursor: { row: 1, col: 1 }, selections: [] }
    ]);
    fillRects.length = 0;

    renderer.setZoom(2);

    const badgeRects = fillRects.filter((rect) => rect.fillStyle === color);
    expect(badgeRects).toHaveLength(1);
    expect(badgeRects[0]!.height).toBeCloseTo(40, 5);
    expect(badgeRects[0]!.width).toBeCloseTo(42, 5);
    expect(badgeRects[0]!.y).toBeCloseTo(-16, 5);
  });
});
