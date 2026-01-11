// @vitest-environment jsdom
import { beforeEach, afterEach, describe, expect, it, vi } from "vitest";

import type { CellData, CellProvider, CellRange } from "../../model/CellProvider";
import { CanvasGridRenderer } from "../../rendering/CanvasGridRenderer";

function createMock2dContext(options: {
  canvas: HTMLCanvasElement;
  onFillText?: (text: string, x: number, y: number) => void;
  onStrokeSegment?: (segment: { x1: number; y1: number; x2: number; y2: number }) => void;
}): CanvasRenderingContext2D {
  const noop = () => {};

  let lastMove: { x: number; y: number } | null = null;

  const moveTo = (x: number, y: number) => {
    lastMove = { x, y };
  };

  const lineTo = (x: number, y: number) => {
    if (lastMove && options.onStrokeSegment) {
      options.onStrokeSegment({ x1: lastMove.x, y1: lastMove.y, x2: x, y2: y });
    }
    lastMove = null;
  };

  const fillText = (text: string, x: number, y: number) => {
    options.onFillText?.(text, x, y);
  };

  return {
    canvas: options.canvas,
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
    moveTo,
    lineTo,
    closePath: noop,
    save: noop,
    restore: noop,
    drawImage: noop,
    translate: noop,
    rotate: noop,
    fillText,
    measureText: (text: string) =>
      ({
        width: text.length * 6,
        actualBoundingBoxAscent: 8,
        actualBoundingBoxDescent: 2
      }) as TextMetrics
  } as unknown as CanvasRenderingContext2D;
}

function createMergedProvider(options: { rowCount: number; colCount: number; merged: CellRange }): CellProvider {
  const { rowCount, colCount, merged } = options;

  const contains = (range: CellRange, row: number, col: number) =>
    row >= range.startRow && row < range.endRow && col >= range.startCol && col < range.endCol;

  const intersects = (a: CellRange, b: CellRange) =>
    a.startRow < b.endRow && a.endRow > b.startRow && a.startCol < b.endCol && a.endCol > b.startCol;

  return {
    getCell: (row: number, col: number): CellData | null => {
      if (row < 0 || col < 0 || row >= rowCount || col >= colCount) return null;
      return { row, col, value: `R${row}C${col}` };
    },
    getMergedRangeAt: (row: number, col: number) => (contains(merged, row, col) ? merged : null),
    getMergedRangesInRange: (range: CellRange) => (intersects(range, merged) ? [merged] : [])
  };
}

describe("merged cells + overflow", () => {
  let ctxByCanvas: Map<HTMLCanvasElement, CanvasRenderingContext2D>;

  beforeEach(() => {
    vi.stubGlobal("requestAnimationFrame", (cb: FrameRequestCallback) => {
      cb(0);
      return 0;
    });

    ctxByCanvas = new Map();
    vi.spyOn(HTMLCanvasElement.prototype, "getContext").mockImplementation(function (this: HTMLCanvasElement) {
      return (ctxByCanvas.get(this) ?? createMock2dContext({ canvas: this })) as any;
    });
  });

  afterEach(() => {
    vi.restoreAllMocks();
    vi.unstubAllGlobals();
  });

  it("renders text only for merged anchors", () => {
    const merged: CellRange = { startRow: 0, endRow: 2, startCol: 0, endCol: 2 };
    const provider = createMergedProvider({ rowCount: 2, colCount: 2, merged });

    const gridCanvas = document.createElement("canvas");
    const contentCanvas = document.createElement("canvas");
    const selectionCanvas = document.createElement("canvas");

    const fillTextCalls: Array<{ text: string; x: number; y: number }> = [];

    ctxByCanvas.set(gridCanvas, createMock2dContext({ canvas: gridCanvas }));
    ctxByCanvas.set(
      contentCanvas,
      createMock2dContext({
        canvas: contentCanvas,
        onFillText: (text, x, y) => fillTextCalls.push({ text, x, y })
      })
    );
    ctxByCanvas.set(selectionCanvas, createMock2dContext({ canvas: selectionCanvas }));

    const renderer = new CanvasGridRenderer({ provider, rowCount: 2, colCount: 2, defaultRowHeight: 10, defaultColWidth: 10 });
    renderer.attach({ grid: gridCanvas, content: contentCanvas, selection: selectionCanvas });
    renderer.resize(50, 50, 1);
    renderer.renderImmediately();

    expect(fillTextCalls).toHaveLength(1);
    expect(fillTextCalls[0]?.text).toBe("R0C0");
  });

  it("suppresses interior gridlines inside merged regions", () => {
    const merged: CellRange = { startRow: 0, endRow: 2, startCol: 0, endCol: 2 };
    const provider = createMergedProvider({ rowCount: 2, colCount: 2, merged });

    const gridCanvas = document.createElement("canvas");
    const contentCanvas = document.createElement("canvas");
    const selectionCanvas = document.createElement("canvas");

    const segments: Array<{ x1: number; y1: number; x2: number; y2: number }> = [];

    ctxByCanvas.set(
      gridCanvas,
      createMock2dContext({
        canvas: gridCanvas,
        onStrokeSegment: (segment) => segments.push(segment)
      })
    );
    ctxByCanvas.set(contentCanvas, createMock2dContext({ canvas: contentCanvas }));
    ctxByCanvas.set(selectionCanvas, createMock2dContext({ canvas: selectionCanvas }));

    const renderer = new CanvasGridRenderer({ provider, rowCount: 2, colCount: 2, defaultRowHeight: 10, defaultColWidth: 10 });
    renderer.attach({ grid: gridCanvas, content: contentCanvas, selection: selectionCanvas });
    renderer.resize(50, 50, 1);
    renderer.renderImmediately();

    // The gridline between col 0 and col 1 would be at x=10.5 (crispLine(10)).
    const hasInteriorVertical = segments.some(
      (seg) => seg.x1 === 10.5 && seg.x2 === 10.5 && Math.min(seg.y1, seg.y2) < 20 && Math.max(seg.y1, seg.y2) > 0
    );
    expect(hasInteriorVertical).toBe(false);

    // The gridline between row 0 and row 1 would be at y=10.5 (crispLine(10)).
    const hasInteriorHorizontal = segments.some(
      (seg) => seg.y1 === 10.5 && seg.y2 === 10.5 && Math.min(seg.x1, seg.x2) < 20 && Math.max(seg.x1, seg.x2) > 0
    );
    expect(hasInteriorHorizontal).toBe(false);
  });

  it("pickCellAt resolves to the merged anchor cell", () => {
    const merged: CellRange = { startRow: 0, endRow: 2, startCol: 0, endCol: 2 };
    const provider = createMergedProvider({ rowCount: 3, colCount: 3, merged });

    const gridCanvas = document.createElement("canvas");
    const contentCanvas = document.createElement("canvas");
    const selectionCanvas = document.createElement("canvas");

    ctxByCanvas.set(gridCanvas, createMock2dContext({ canvas: gridCanvas }));
    ctxByCanvas.set(contentCanvas, createMock2dContext({ canvas: contentCanvas }));
    ctxByCanvas.set(selectionCanvas, createMock2dContext({ canvas: selectionCanvas }));

    const renderer = new CanvasGridRenderer({ provider, rowCount: 3, colCount: 3, defaultRowHeight: 10, defaultColWidth: 10 });
    renderer.attach({ grid: gridCanvas, content: contentCanvas, selection: selectionCanvas });
    renderer.resize(100, 100, 1);

    // Point (15,5) is inside row 0, col 1, which is part of the merged range.
    expect(renderer.pickCellAt(15, 5)).toEqual({ row: 0, col: 0 });
  });

  it("getCellRect returns merged bounds for merged cells", () => {
    const merged: CellRange = { startRow: 0, endRow: 2, startCol: 0, endCol: 2 };
    const provider = createMergedProvider({ rowCount: 3, colCount: 3, merged });

    const gridCanvas = document.createElement("canvas");
    const contentCanvas = document.createElement("canvas");
    const selectionCanvas = document.createElement("canvas");

    ctxByCanvas.set(gridCanvas, createMock2dContext({ canvas: gridCanvas }));
    ctxByCanvas.set(contentCanvas, createMock2dContext({ canvas: contentCanvas }));
    ctxByCanvas.set(selectionCanvas, createMock2dContext({ canvas: selectionCanvas }));

    const renderer = new CanvasGridRenderer({ provider, rowCount: 3, colCount: 3, defaultRowHeight: 10, defaultColWidth: 10 });
    renderer.attach({ grid: gridCanvas, content: contentCanvas, selection: selectionCanvas });
    renderer.resize(100, 100, 1);

    expect(renderer.getCellRect(0, 0)).toEqual({ x: 0, y: 0, width: 20, height: 20 });
    // Interior merged cells should report the same merged bounds.
    expect(renderer.getCellRect(1, 1)).toEqual({ x: 0, y: 0, width: 20, height: 20 });
  });
});
