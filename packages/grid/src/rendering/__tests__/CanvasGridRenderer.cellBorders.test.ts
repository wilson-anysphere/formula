// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type { CellProvider, CellProviderUpdate, CellRange } from "../../model/CellProvider";
import { CanvasGridRenderer } from "../CanvasGridRenderer";

type Segment = { x1: number; y1: number; x2: number; y2: number };
type StrokeRecord = { strokeStyle: unknown; lineWidth: number; lineDash: number[]; segments: Segment[] };
type RectCall = { x: number; y: number; width: number; height: number };

function createRecording2dContext(canvas: HTMLCanvasElement): {
  ctx: CanvasRenderingContext2D;
  strokes: StrokeRecord[];
  rectCalls: RectCall[];
} {
  const noop = () => {};

  let fillStyle: unknown = "#000";
  let strokeStyle: unknown = "#000";
  let lineWidth = 1;
  let lineDash: number[] = [];

  let cursor: { x: number; y: number } | null = null;
  let segments: Segment[] = [];
  const strokes: StrokeRecord[] = [];
  const rectCalls: RectCall[] = [];

  const ctx = {
    canvas,
    get fillStyle() {
      return fillStyle;
    },
    set fillStyle(value: unknown) {
      fillStyle = value;
    },
    get strokeStyle() {
      return strokeStyle;
    },
    set strokeStyle(value: unknown) {
      strokeStyle = value;
    },
    get lineWidth() {
      return lineWidth;
    },
    set lineWidth(value: number) {
      lineWidth = value;
    },
    font: "",
    textAlign: "left",
    textBaseline: "alphabetic",
    globalAlpha: 1,
    imageSmoothingEnabled: false,
    setTransform: noop,
    clearRect: noop,
    fillRect: noop,
    strokeRect: noop,
    beginPath: () => {
      cursor = null;
      segments = [];
    },
    rect: (x: number, y: number, width: number, height: number) => {
      rectCalls.push({ x, y, width, height });
    },
    clip: noop,
    fill: noop,
    stroke: () => {
      strokes.push({ strokeStyle, lineWidth, lineDash: [...lineDash], segments: [...segments] });
    },
    moveTo: (x: number, y: number) => {
      cursor = { x, y };
    },
    lineTo: (x: number, y: number) => {
      if (cursor) {
        segments.push({ x1: cursor.x, y1: cursor.y, x2: x, y2: y });
      }
      cursor = { x, y };
    },
    closePath: noop,
    save: noop,
    restore: noop,
    drawImage: noop,
    translate: noop,
    rotate: noop,
    fillText: noop,
    setLineDash: (dash: number[]) => {
      lineDash = dash;
    },
    measureText: (text: string) =>
      ({
        width: text.length * 6,
        actualBoundingBoxAscent: 8,
        actualBoundingBoxDescent: 2
      }) as TextMetrics
  } as unknown as CanvasRenderingContext2D;

  return { ctx, strokes, rectCalls };
}

function normalizeSegment(seg: Segment): Segment {
  // Normalize direction so tests can match segments regardless of moveTo/lineTo ordering.
  if (seg.y1 === seg.y2) {
    // horizontal
    if (seg.x1 <= seg.x2) return seg;
    return { x1: seg.x2, y1: seg.y2, x2: seg.x1, y2: seg.y1 };
  }
  if (seg.x1 === seg.x2) {
    // vertical
    if (seg.y1 <= seg.y2) return seg;
    return { x1: seg.x2, y1: seg.y2, x2: seg.x1, y2: seg.y1 };
  }
  return seg;
}

function hasNormalizedSegment(segments: Segment[], expected: Segment): boolean {
  const e = normalizeSegment(expected);
  return segments.some((seg) => {
    const s = normalizeSegment(seg);
    return (
      Math.abs(s.x1 - e.x1) < 1e-6 &&
      Math.abs(s.y1 - e.y1) < 1e-6 &&
      Math.abs(s.x2 - e.x2) < 1e-6 &&
      Math.abs(s.y2 - e.y2) < 1e-6
    );
  });
}

describe("CanvasGridRenderer side border rendering (Excel-like)", () => {
  const originalGetContext = HTMLCanvasElement.prototype.getContext;
  const originalRaf = globalThis.requestAnimationFrame;

  beforeEach(() => {
    vi.stubGlobal("requestAnimationFrame", (cb: FrameRequestCallback) => {
      cb(0);
      return 0;
    });
  });

  afterEach(() => {
    HTMLCanvasElement.prototype.getContext = originalGetContext;
    if (originalRaf) {
      vi.stubGlobal("requestAnimationFrame", originalRaf);
    } else {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      delete (globalThis as any).requestAnimationFrame;
    }
    vi.unstubAllGlobals();
  });

  it("renders double borders as two parallel strokes (horizontal + vertical)", () => {
    const provider: CellProvider = {
      getCell: (row, col) => {
        if (row === 0 && col === 0) {
          return {
            row,
            col,
            value: "A",
            style: {
              borders: {
                top: { width: 3, style: "double", color: "#ff00ff" },
                right: { width: 3, style: "double", color: "#ff00ff" },
                bottom: { width: 3, style: "double", color: "#ff00ff" },
                left: { width: 3, style: "double", color: "#ff00ff" }
              }
            }
          };
        }
        return null;
      }
    };

    const gridCanvas = document.createElement("canvas");
    const contentCanvas = document.createElement("canvas");
    const selectionCanvas = document.createElement("canvas");

    const { ctx: gridCtx, strokes: gridStrokes } = createRecording2dContext(gridCanvas);
    const contentCtx = createRecording2dContext(contentCanvas).ctx;
    const selectionCtx = createRecording2dContext(selectionCanvas).ctx;

    const contexts = new Map<HTMLCanvasElement, CanvasRenderingContext2D>([
      [gridCanvas, gridCtx],
      [contentCanvas, contentCtx],
      [selectionCanvas, selectionCtx]
    ]);

    HTMLCanvasElement.prototype.getContext = vi.fn(function (this: HTMLCanvasElement) {
      return contexts.get(this) ?? createRecording2dContext(this).ctx;
    }) as unknown as typeof HTMLCanvasElement.prototype.getContext;

    const renderer = new CanvasGridRenderer({
      provider,
      rowCount: 1,
      colCount: 1,
      defaultRowHeight: 20,
      defaultColWidth: 20
    });
    renderer.attach({ grid: gridCanvas, content: contentCanvas, selection: selectionCanvas });
    renderer.resize(100, 100, 1);
    renderer.renderImmediately();

    const purpleStrokes = gridStrokes.filter((stroke) => stroke.strokeStyle === "#ff00ff" && stroke.lineWidth === 1);
    expect(purpleStrokes).toHaveLength(1);

    const allSegments = purpleStrokes[0]!.segments;
    const horizontalYs = new Set<number>(
      allSegments.filter((seg) => seg.y1 === seg.y2).map((seg) => seg.y1)
    );
    const verticalXs = new Set<number>(
      allSegments.filter((seg) => seg.x1 === seg.x2).map((seg) => seg.x1)
    );

    // For a 20x20 cell with totalWidth=3 (per-line width=1, offset=1):
    // lines are drawn at boundary +/- 1, with crisp alignment for 1px strokes.
    expect([...horizontalYs].sort((a, b) => a - b)).toEqual([-0.5, 1.5, 19.5, 21.5]);
    expect([...verticalXs].sort((a, b) => a - b)).toEqual([-0.5, 1.5, 19.5, 21.5]);
  });

  it("scales double borders with zoom (per-line width + offsets)", () => {
    const provider: CellProvider = {
      getCell: (row, col) => {
        if (row === 0 && col === 0) {
          return {
            row,
            col,
            value: "A",
            style: {
              borders: {
                top: { width: 3, style: "double", color: "#ff00ff" },
                right: { width: 3, style: "double", color: "#ff00ff" },
                bottom: { width: 3, style: "double", color: "#ff00ff" },
                left: { width: 3, style: "double", color: "#ff00ff" }
              }
            }
          };
        }
        return null;
      }
    };

    const gridCanvas = document.createElement("canvas");
    const contentCanvas = document.createElement("canvas");
    const selectionCanvas = document.createElement("canvas");

    const { ctx: gridCtx, strokes: gridStrokes } = createRecording2dContext(gridCanvas);
    const contentCtx = createRecording2dContext(contentCanvas).ctx;
    const selectionCtx = createRecording2dContext(selectionCanvas).ctx;

    const contexts = new Map<HTMLCanvasElement, CanvasRenderingContext2D>([
      [gridCanvas, gridCtx],
      [contentCanvas, contentCtx],
      [selectionCanvas, selectionCtx]
    ]);

    HTMLCanvasElement.prototype.getContext = vi.fn(function (this: HTMLCanvasElement) {
      return contexts.get(this) ?? createRecording2dContext(this).ctx;
    }) as unknown as typeof HTMLCanvasElement.prototype.getContext;

    const renderer = new CanvasGridRenderer({
      provider,
      rowCount: 1,
      colCount: 1,
      defaultRowHeight: 20,
      defaultColWidth: 20
    });
    renderer.attach({ grid: gridCanvas, content: contentCanvas, selection: selectionCanvas });
    renderer.resize(100, 100, 1);

    // Clear strokes from the initial render triggered by `resize` (requestAnimationFrame is stubbed to run immediately).
    gridStrokes.length = 0;
    renderer.setZoom(2);

    // For zoom=2, totalWidth=6 and per-line width is 2 (totalWidth/3).
    const purpleStrokes = gridStrokes.filter(
      (stroke) => stroke.strokeStyle === "#ff00ff" && Math.abs(stroke.lineWidth - 2) < 1e-6
    );
    expect(purpleStrokes).toHaveLength(1);

    const allSegments = purpleStrokes[0]!.segments;
    const horizontalYs = new Set<number>(
      allSegments.filter((seg) => seg.y1 === seg.y2).map((seg) => seg.y1)
    );
    const verticalXs = new Set<number>(
      allSegments.filter((seg) => seg.x1 === seg.x2).map((seg) => seg.x1)
    );

    // For a 40x40 cell with per-line width=2, offsets are +/-2 and no half-pixel snapping is applied.
    expect([...horizontalYs].sort((a, b) => a - b)).toEqual([-2, 2, 38, 42]);
    expect([...verticalXs].sort((a, b) => a - b)).toEqual([-2, 2, 38, 42]);
  });

  it("renders merged range side borders around the merged rect (anchor cell styles)", () => {
    const merged: CellRange = { startRow: 0, endRow: 2, startCol: 0, endCol: 2 };
    const contains = (range: CellRange, row: number, col: number) =>
      row >= range.startRow && row < range.endRow && col >= range.startCol && col < range.endCol;
    const intersects = (a: CellRange, b: CellRange) =>
      a.startRow < b.endRow && a.endRow > b.startRow && a.startCol < b.endCol && a.endCol > b.startCol;

    const provider: CellProvider = {
      getCell: (row, col) => {
        if (row === 0 && col === 0) {
          return {
            row,
            col,
            value: "Merged",
            style: {
              borders: {
                top: { width: 1, style: "solid", color: "#22c55e" },
                right: { width: 1, style: "solid", color: "#22c55e" },
                bottom: { width: 1, style: "solid", color: "#22c55e" },
                left: { width: 1, style: "solid", color: "#22c55e" }
              }
            }
          };
        }
        return { row, col, value: null };
      },
      getMergedRangeAt: (row, col) => (contains(merged, row, col) ? merged : null),
      getMergedRangesInRange: (range) => (intersects(range, merged) ? [merged] : [])
    };

    const gridCanvas = document.createElement("canvas");
    const contentCanvas = document.createElement("canvas");
    const selectionCanvas = document.createElement("canvas");

    const { ctx: gridCtx, strokes: gridStrokes } = createRecording2dContext(gridCanvas);
    const contentCtx = createRecording2dContext(contentCanvas).ctx;
    const selectionCtx = createRecording2dContext(selectionCanvas).ctx;

    const contexts = new Map<HTMLCanvasElement, CanvasRenderingContext2D>([
      [gridCanvas, gridCtx],
      [contentCanvas, contentCtx],
      [selectionCanvas, selectionCtx]
    ]);

    HTMLCanvasElement.prototype.getContext = vi.fn(function (this: HTMLCanvasElement) {
      return contexts.get(this) ?? createRecording2dContext(this).ctx;
    }) as unknown as typeof HTMLCanvasElement.prototype.getContext;

    const renderer = new CanvasGridRenderer({
      provider,
      rowCount: 2,
      colCount: 2,
      defaultRowHeight: 10,
      defaultColWidth: 10
    });
    renderer.attach({ grid: gridCanvas, content: contentCanvas, selection: selectionCanvas });
    renderer.resize(100, 100, 1);
    renderer.renderImmediately();

    const greenStroke = gridStrokes.find((stroke) => stroke.strokeStyle === "#22c55e" && stroke.lineWidth === 1);
    expect(greenStroke).toBeTruthy();

    const segments = greenStroke!.segments;

    // Border should be drawn around the merged rect: (0,0) → (20,20), with crisp alignment for 1px strokes.
    const expected: Segment[] = [
      // Top (split into 2 cols)
      { x1: 0, y1: 0.5, x2: 10, y2: 0.5 },
      { x1: 10, y1: 0.5, x2: 20, y2: 0.5 },
      // Bottom (split into 2 cols)
      { x1: 0, y1: 20.5, x2: 10, y2: 20.5 },
      { x1: 10, y1: 20.5, x2: 20, y2: 20.5 },
      // Left (split into 2 rows)
      { x1: 0.5, y1: 0, x2: 0.5, y2: 10 },
      { x1: 0.5, y1: 10, x2: 0.5, y2: 20 },
      // Right (split into 2 rows)
      { x1: 20.5, y1: 0, x2: 20.5, y2: 10 },
      { x1: 20.5, y1: 10, x2: 20.5, y2: 20 }
    ];

    for (const seg of expected) {
      expect(hasNormalizedSegment(segments, seg)).toBe(true);
    }
  });

  it("resolves border conflicts across frozen row boundaries (collapsed borders across quadrants)", () => {
    const provider: CellProvider = {
      getCell: (row, col) => {
        if (row === 0 && col === 0) {
          return { row, col, value: null, style: { borders: { bottom: { width: 3, style: "solid", color: "#0000ff" } } } };
        }
        if (row === 1 && col === 0) {
          return { row, col, value: null, style: { borders: { top: { width: 1, style: "solid", color: "#ff0000" } } } };
        }
        return { row, col, value: null };
      }
    };

    const gridCanvas = document.createElement("canvas");
    const contentCanvas = document.createElement("canvas");
    const selectionCanvas = document.createElement("canvas");

    const { ctx: gridCtx, strokes: gridStrokes } = createRecording2dContext(gridCanvas);
    const contentCtx = createRecording2dContext(contentCanvas).ctx;
    const selectionCtx = createRecording2dContext(selectionCanvas).ctx;

    const contexts = new Map<HTMLCanvasElement, CanvasRenderingContext2D>([
      [gridCanvas, gridCtx],
      [contentCanvas, contentCtx],
      [selectionCanvas, selectionCtx]
    ]);

    HTMLCanvasElement.prototype.getContext = vi.fn(function (this: HTMLCanvasElement) {
      return contexts.get(this) ?? createRecording2dContext(this).ctx;
    }) as unknown as typeof HTMLCanvasElement.prototype.getContext;

    const renderer = new CanvasGridRenderer({
      provider,
      rowCount: 2,
      colCount: 1,
      defaultRowHeight: 10,
      defaultColWidth: 10
    });
    renderer.attach({ grid: gridCanvas, content: contentCanvas, selection: selectionCanvas });
    renderer.resize(100, 50, 1);
    gridStrokes.length = 0;

    // Freeze the first row so the shared edge between row 0 and row 1 is rendered across quadrants.
    renderer.setFrozen(1, 0);

    const blueStrokes = gridStrokes.filter((stroke) => stroke.strokeStyle === "#0000ff" && stroke.lineWidth === 3);
    expect(blueStrokes.length).toBeGreaterThan(0);
    expect(
      blueStrokes.some((stroke) => hasNormalizedSegment(stroke.segments, { x1: 0, y1: 10.5, x2: 10, y2: 10.5 }))
    ).toBe(true);

    // Ensure the thinner red border never wins the shared edge.
    expect(gridStrokes.some((stroke) => stroke.strokeStyle === "#ff0000")).toBe(false);
  });

  it("resolves border conflicts across frozen column boundaries (collapsed borders across quadrants)", () => {
    const provider: CellProvider = {
      getCell: (row, col) => {
        if (row === 0 && col === 0) {
          return { row, col, value: null, style: { borders: { right: { width: 3, style: "solid", color: "#0000ff" } } } };
        }
        if (row === 0 && col === 1) {
          return { row, col, value: null, style: { borders: { left: { width: 1, style: "solid", color: "#ff0000" } } } };
        }
        return { row, col, value: null };
      }
    };

    const gridCanvas = document.createElement("canvas");
    const contentCanvas = document.createElement("canvas");
    const selectionCanvas = document.createElement("canvas");

    const { ctx: gridCtx, strokes: gridStrokes } = createRecording2dContext(gridCanvas);
    const contentCtx = createRecording2dContext(contentCanvas).ctx;
    const selectionCtx = createRecording2dContext(selectionCanvas).ctx;

    const contexts = new Map<HTMLCanvasElement, CanvasRenderingContext2D>([
      [gridCanvas, gridCtx],
      [contentCanvas, contentCtx],
      [selectionCanvas, selectionCtx]
    ]);

    HTMLCanvasElement.prototype.getContext = vi.fn(function (this: HTMLCanvasElement) {
      return contexts.get(this) ?? createRecording2dContext(this).ctx;
    }) as unknown as typeof HTMLCanvasElement.prototype.getContext;

    const renderer = new CanvasGridRenderer({
      provider,
      rowCount: 1,
      colCount: 2,
      defaultRowHeight: 10,
      defaultColWidth: 10
    });
    renderer.attach({ grid: gridCanvas, content: contentCanvas, selection: selectionCanvas });
    renderer.resize(100, 50, 1);
    gridStrokes.length = 0;

    // Freeze the first column so the shared edge between col 0 and col 1 is rendered across quadrants.
    renderer.setFrozen(0, 1);

    const blueStrokes = gridStrokes.filter((stroke) => stroke.strokeStyle === "#0000ff" && stroke.lineWidth === 3);
    expect(blueStrokes.length).toBeGreaterThan(0);
    expect(
      blueStrokes.some((stroke) => hasNormalizedSegment(stroke.segments, { x1: 10.5, y1: 0, x2: 10.5, y2: 10 }))
    ).toBe(true);

    expect(gridStrokes.some((stroke) => stroke.strokeStyle === "#ff0000")).toBe(false);
  });

  it("resolves merged border conflicts across frozen row boundaries (collapsed borders across quadrants)", () => {
    const merged: CellRange = { startRow: 0, endRow: 1, startCol: 0, endCol: 2 };
    const contains = (range: CellRange, row: number, col: number) =>
      row >= range.startRow && row < range.endRow && col >= range.startCol && col < range.endCol;
    const intersects = (a: CellRange, b: CellRange) =>
      a.startRow < b.endRow && a.endRow > b.startRow && a.startCol < b.endCol && a.endCol > b.startCol;

    const provider: CellProvider = {
      getCell: (row, col) => {
        // Merged anchor in the frozen row has a thick blue bottom border on the merged perimeter.
        if (row === 0 && col === 0) {
          return { row, col, value: null, style: { borders: { bottom: { width: 3, style: "solid", color: "#0000ff" } } } };
        }
        // Cells below (in the scrollable pane) try to draw a thin red top border, but should lose.
        if (row === 1 && (col === 0 || col === 1)) {
          return { row, col, value: null, style: { borders: { top: { width: 1, style: "solid", color: "#ff0000" } } } };
        }
        return { row, col, value: null };
      },
      getMergedRangeAt: (row, col) => (contains(merged, row, col) ? merged : null),
      getMergedRangesInRange: (range) => (intersects(range, merged) ? [merged] : [])
    };

    const gridCanvas = document.createElement("canvas");
    const contentCanvas = document.createElement("canvas");
    const selectionCanvas = document.createElement("canvas");

    const { ctx: gridCtx, strokes: gridStrokes } = createRecording2dContext(gridCanvas);
    const contentCtx = createRecording2dContext(contentCanvas).ctx;
    const selectionCtx = createRecording2dContext(selectionCanvas).ctx;

    const contexts = new Map<HTMLCanvasElement, CanvasRenderingContext2D>([
      [gridCanvas, gridCtx],
      [contentCanvas, contentCtx],
      [selectionCanvas, selectionCtx]
    ]);

    HTMLCanvasElement.prototype.getContext = vi.fn(function (this: HTMLCanvasElement) {
      return contexts.get(this) ?? createRecording2dContext(this).ctx;
    }) as unknown as typeof HTMLCanvasElement.prototype.getContext;

    const renderer = new CanvasGridRenderer({
      provider,
      rowCount: 2,
      colCount: 2,
      defaultRowHeight: 10,
      defaultColWidth: 10
    });
    renderer.attach({ grid: gridCanvas, content: contentCanvas, selection: selectionCanvas });
    renderer.resize(100, 50, 1);
    gridStrokes.length = 0;
    renderer.setFrozen(1, 0);

    // Thick blue border should win on the shared boundary y=10 (odd width=3 => crisp at y=10.5).
    const blueStroke = gridStrokes.find((stroke) => stroke.strokeStyle === "#0000ff" && stroke.lineWidth === 3);
    expect(blueStroke).toBeTruthy();
    expect(hasNormalizedSegment(blueStroke!.segments, { x1: 0, y1: 10.5, x2: 10, y2: 10.5 })).toBe(true);
    expect(hasNormalizedSegment(blueStroke!.segments, { x1: 10, y1: 10.5, x2: 20, y2: 10.5 })).toBe(true);

    // Ensure the red border never wins the shared edge.
    expect(gridStrokes.some((stroke) => stroke.strokeStyle === "#ff0000")).toBe(false);
  });

  it("resolves merged border conflicts across frozen column boundaries (collapsed borders across quadrants)", () => {
    const merged: CellRange = { startRow: 0, endRow: 2, startCol: 0, endCol: 1 };
    const contains = (range: CellRange, row: number, col: number) =>
      row >= range.startRow && row < range.endRow && col >= range.startCol && col < range.endCol;
    const intersects = (a: CellRange, b: CellRange) =>
      a.startRow < b.endRow && a.endRow > b.startRow && a.startCol < b.endCol && a.endCol > b.startCol;

    const provider: CellProvider = {
      getCell: (row, col) => {
        // Merged anchor in the frozen col has a thick blue right border on the merged perimeter.
        if (row === 0 && col === 0) {
          return { row, col, value: null, style: { borders: { right: { width: 3, style: "solid", color: "#0000ff" } } } };
        }
        // Cells to the right (in the scrollable pane) try to draw a thin red left border, but should lose.
        if ((row === 0 || row === 1) && col === 1) {
          return { row, col, value: null, style: { borders: { left: { width: 1, style: "solid", color: "#ff0000" } } } };
        }
        return { row, col, value: null };
      },
      getMergedRangeAt: (row, col) => (contains(merged, row, col) ? merged : null),
      getMergedRangesInRange: (range) => (intersects(range, merged) ? [merged] : [])
    };

    const gridCanvas = document.createElement("canvas");
    const contentCanvas = document.createElement("canvas");
    const selectionCanvas = document.createElement("canvas");

    const { ctx: gridCtx, strokes: gridStrokes } = createRecording2dContext(gridCanvas);
    const contentCtx = createRecording2dContext(contentCanvas).ctx;
    const selectionCtx = createRecording2dContext(selectionCanvas).ctx;

    const contexts = new Map<HTMLCanvasElement, CanvasRenderingContext2D>([
      [gridCanvas, gridCtx],
      [contentCanvas, contentCtx],
      [selectionCanvas, selectionCtx]
    ]);

    HTMLCanvasElement.prototype.getContext = vi.fn(function (this: HTMLCanvasElement) {
      return contexts.get(this) ?? createRecording2dContext(this).ctx;
    }) as unknown as typeof HTMLCanvasElement.prototype.getContext;

    const renderer = new CanvasGridRenderer({
      provider,
      rowCount: 2,
      colCount: 2,
      defaultRowHeight: 10,
      defaultColWidth: 10
    });
    renderer.attach({ grid: gridCanvas, content: contentCanvas, selection: selectionCanvas });
    renderer.resize(100, 50, 1);
    gridStrokes.length = 0;
    renderer.setFrozen(0, 1);

    // Thick blue border should win on the shared boundary x=10 (odd width=3 => crisp at x=10.5).
    const blueStroke = gridStrokes.find((stroke) => stroke.strokeStyle === "#0000ff" && stroke.lineWidth === 3);
    expect(blueStroke).toBeTruthy();
    expect(hasNormalizedSegment(blueStroke!.segments, { x1: 10.5, y1: 0, x2: 10.5, y2: 10 })).toBe(true);
    expect(hasNormalizedSegment(blueStroke!.segments, { x1: 10.5, y1: 10, x2: 10.5, y2: 20 })).toBe(true);

    expect(gridStrokes.some((stroke) => stroke.strokeStyle === "#ff0000")).toBe(false);
  });

  it("resolves shared-edge border conflicts deterministically (width, style, tie-break)", () => {
    const provider: CellProvider = {
      getCell: (row, col) => {
        if (row !== 0) return null;

        // 4 columns → 3 shared edges.
        if (col === 0) {
          return { row, col, value: null, style: { borders: { right: { width: 2, style: "solid", color: "#ff0000" } } } };
        }
        if (col === 1) {
          return {
            row,
            col,
            value: null,
            style: {
              borders: {
                left: { width: 1, style: "solid", color: "#0000ff" }, // loses to thicker border
                right: { width: 3, style: "dotted", color: "#00ff00" } // loses to double when width ties
              }
            }
          };
        }
        if (col === 2) {
          return {
            row,
            col,
            value: null,
            style: {
              borders: {
                left: { width: 3, style: "double", color: "#ff00ff" }, // wins on style rank
                right: { width: 1, style: "solid", color: "#123456" } // loses tie-break to right cell
              }
            }
          };
        }
        if (col === 3) {
          return { row, col, value: null, style: { borders: { left: { width: 1, style: "solid", color: "#abcdef" } } } };
        }
        return null;
      }
    };

    const gridCanvas = document.createElement("canvas");
    const contentCanvas = document.createElement("canvas");
    const selectionCanvas = document.createElement("canvas");

    const { ctx: gridCtx, strokes: gridStrokes } = createRecording2dContext(gridCanvas);
    const contentCtx = createRecording2dContext(contentCanvas).ctx;
    const selectionCtx = createRecording2dContext(selectionCanvas).ctx;

    const contexts = new Map<HTMLCanvasElement, CanvasRenderingContext2D>([
      [gridCanvas, gridCtx],
      [contentCanvas, contentCtx],
      [selectionCanvas, selectionCtx]
    ]);

    HTMLCanvasElement.prototype.getContext = vi.fn(function (this: HTMLCanvasElement) {
      return contexts.get(this) ?? createRecording2dContext(this).ctx;
    }) as unknown as typeof HTMLCanvasElement.prototype.getContext;

    const renderer = new CanvasGridRenderer({
      provider,
      rowCount: 1,
      colCount: 4,
      defaultRowHeight: 10,
      defaultColWidth: 10
    });
    renderer.attach({ grid: gridCanvas, content: contentCanvas, selection: selectionCanvas });
    renderer.resize(200, 50, 1);
    renderer.renderImmediately();

    // Width winner (col0 right vs col1 left) → x=10, width=2, color=#ff0000.
    const redStroke = gridStrokes.find((stroke) => stroke.strokeStyle === "#ff0000" && stroke.lineWidth === 2);
    expect(redStroke).toBeTruthy();
    expect(hasNormalizedSegment(redStroke!.segments, { x1: 10, y1: 0, x2: 10, y2: 10 })).toBe(true);
    expect(gridStrokes.some((stroke) => stroke.strokeStyle === "#0000ff")).toBe(false);

    // Style winner when width ties (dotted vs double) → double wins; draw two parallel strokes at x=20±1.
    const purpleStrokes = gridStrokes.filter((stroke) => stroke.strokeStyle === "#ff00ff" && stroke.lineWidth === 1);
    expect(purpleStrokes).toHaveLength(1);
    const purpleSegments = purpleStrokes[0]!.segments;
    expect(hasNormalizedSegment(purpleSegments, { x1: 19.5, y1: 0, x2: 19.5, y2: 10 })).toBe(true);
    expect(hasNormalizedSegment(purpleSegments, { x1: 21.5, y1: 0, x2: 21.5, y2: 10 })).toBe(true);
    expect(gridStrokes.some((stroke) => stroke.strokeStyle === "#00ff00")).toBe(false);

    // Tie-break winner (solid width=1 vs solid width=1) → right cell wins (col3), x=30.5, color=#abcdef.
    const tieStroke = gridStrokes.find((stroke) => stroke.strokeStyle === "#abcdef" && stroke.lineWidth === 1);
    expect(tieStroke).toBeTruthy();
    expect(hasNormalizedSegment(tieStroke!.segments, { x1: 30.5, y1: 0, x2: 30.5, y2: 10 })).toBe(true);
    expect(gridStrokes.some((stroke) => stroke.strokeStyle === "#123456")).toBe(false);
  });

  it("resolves conflicts between merged perimeter borders and adjacent cell borders (thicker wins)", () => {
    const merged: CellRange = { startRow: 0, endRow: 2, startCol: 0, endCol: 2 };
    const contains = (range: CellRange, row: number, col: number) =>
      row >= range.startRow && row < range.endRow && col >= range.startCol && col < range.endCol;
    const intersects = (a: CellRange, b: CellRange) =>
      a.startRow < b.endRow && a.endRow > b.startRow && a.startCol < b.endCol && a.endCol > b.startCol;

    const provider: CellProvider = {
      getCell: (row, col) => {
        // Merged anchor has a thin red right border on the merged perimeter.
        if (row === 0 && col === 0) {
          return { row, col, value: null, style: { borders: { right: { width: 1, style: "solid", color: "#ff0000" } } } };
        }
        // Adjacent cell to the right has a thicker blue left border (should win).
        if ((row === 0 || row === 1) && col === 2) {
          return { row, col, value: null, style: { borders: { left: { width: 3, style: "solid", color: "#0000ff" } } } };
        }
        return { row, col, value: null };
      },
      getMergedRangeAt: (row, col) => (contains(merged, row, col) ? merged : null),
      getMergedRangesInRange: (range) => (intersects(range, merged) ? [merged] : [])
    };

    const gridCanvas = document.createElement("canvas");
    const contentCanvas = document.createElement("canvas");
    const selectionCanvas = document.createElement("canvas");

    const { ctx: gridCtx, strokes: gridStrokes } = createRecording2dContext(gridCanvas);
    const contentCtx = createRecording2dContext(contentCanvas).ctx;
    const selectionCtx = createRecording2dContext(selectionCanvas).ctx;

    const contexts = new Map<HTMLCanvasElement, CanvasRenderingContext2D>([
      [gridCanvas, gridCtx],
      [contentCanvas, contentCtx],
      [selectionCanvas, selectionCtx]
    ]);

    HTMLCanvasElement.prototype.getContext = vi.fn(function (this: HTMLCanvasElement) {
      return contexts.get(this) ?? createRecording2dContext(this).ctx;
    }) as unknown as typeof HTMLCanvasElement.prototype.getContext;

    const renderer = new CanvasGridRenderer({
      provider,
      rowCount: 2,
      colCount: 3,
      defaultRowHeight: 10,
      defaultColWidth: 10
    });
    renderer.attach({ grid: gridCanvas, content: contentCanvas, selection: selectionCanvas });
    renderer.resize(200, 100, 1);
    renderer.renderImmediately();

    // Thick blue border wins, drawn at x=20 with crisp alignment for odd width=3 -> 20.5.
    const blueStroke = gridStrokes.find((stroke) => stroke.strokeStyle === "#0000ff" && stroke.lineWidth === 3);
    expect(blueStroke).toBeTruthy();
    expect(hasNormalizedSegment(blueStroke!.segments, { x1: 20.5, y1: 0, x2: 20.5, y2: 10 })).toBe(true);
    expect(hasNormalizedSegment(blueStroke!.segments, { x1: 20.5, y1: 10, x2: 20.5, y2: 20 })).toBe(true);

    // Ensure the merged anchor's thin red border did not win.
    expect(gridStrokes.some((stroke) => stroke.strokeStyle === "#ff0000")).toBe(false);
  });

  it("applies right/bottom tie-break rules when merged perimeter borders tie with adjacent cells", () => {
    const merged: CellRange = { startRow: 0, endRow: 2, startCol: 0, endCol: 2 };
    const contains = (range: CellRange, row: number, col: number) =>
      row >= range.startRow && row < range.endRow && col >= range.startCol && col < range.endCol;
    const intersects = (a: CellRange, b: CellRange) =>
      a.startRow < b.endRow && a.endRow > b.startRow && a.startCol < b.endCol && a.endCol > b.startCol;

    const provider: CellProvider = {
      getCell: (row, col) => {
        // Merged anchor defines a solid red right border.
        if (row === 0 && col === 0) {
          return { row, col, value: null, style: { borders: { right: { width: 1, style: "solid", color: "#ff0000" } } } };
        }
        // Adjacent right cell defines an equal-width solid blue left border; tie-break prefers the right cell.
        if ((row === 0 || row === 1) && col === 2) {
          return { row, col, value: null, style: { borders: { left: { width: 1, style: "solid", color: "#0000ff" } } } };
        }
        return { row, col, value: null };
      },
      getMergedRangeAt: (row, col) => (contains(merged, row, col) ? merged : null),
      getMergedRangesInRange: (range) => (intersects(range, merged) ? [merged] : [])
    };

    const gridCanvas = document.createElement("canvas");
    const contentCanvas = document.createElement("canvas");
    const selectionCanvas = document.createElement("canvas");

    const { ctx: gridCtx, strokes: gridStrokes } = createRecording2dContext(gridCanvas);
    const contentCtx = createRecording2dContext(contentCanvas).ctx;
    const selectionCtx = createRecording2dContext(selectionCanvas).ctx;

    const contexts = new Map<HTMLCanvasElement, CanvasRenderingContext2D>([
      [gridCanvas, gridCtx],
      [contentCanvas, contentCtx],
      [selectionCanvas, selectionCtx]
    ]);

    HTMLCanvasElement.prototype.getContext = vi.fn(function (this: HTMLCanvasElement) {
      return contexts.get(this) ?? createRecording2dContext(this).ctx;
    }) as unknown as typeof HTMLCanvasElement.prototype.getContext;

    const renderer = new CanvasGridRenderer({
      provider,
      rowCount: 2,
      colCount: 3,
      defaultRowHeight: 10,
      defaultColWidth: 10
    });
    renderer.attach({ grid: gridCanvas, content: contentCanvas, selection: selectionCanvas });
    renderer.resize(200, 100, 1);
    renderer.renderImmediately();

    const blueStroke = gridStrokes.find((stroke) => stroke.strokeStyle === "#0000ff" && stroke.lineWidth === 1);
    expect(blueStroke).toBeTruthy();
    expect(hasNormalizedSegment(blueStroke!.segments, { x1: 20.5, y1: 0, x2: 20.5, y2: 10 })).toBe(true);
    expect(hasNormalizedSegment(blueStroke!.segments, { x1: 20.5, y1: 10, x2: 20.5, y2: 20 })).toBe(true);

    expect(gridStrokes.some((stroke) => stroke.strokeStyle === "#ff0000")).toBe(false);
  });

  it("prefers merged perimeter border styles over adjacent cells when widths tie (style rank)", () => {
    const merged: CellRange = { startRow: 0, endRow: 2, startCol: 0, endCol: 2 };
    const contains = (range: CellRange, row: number, col: number) =>
      row >= range.startRow && row < range.endRow && col >= range.startCol && col < range.endCol;
    const intersects = (a: CellRange, b: CellRange) =>
      a.startRow < b.endRow && a.endRow > b.startRow && a.startCol < b.endCol && a.endCol > b.startCol;

    const provider: CellProvider = {
      getCell: (row, col) => {
        // Merged anchor defines a *double* border on the merged perimeter.
        if (row === 0 && col === 0) {
          return { row, col, value: null, style: { borders: { right: { width: 3, style: "double", color: "#ff00ff" } } } };
        }
        // Adjacent right cell defines an equal-width *solid* border (should lose to double style rank).
        if ((row === 0 || row === 1) && col === 2) {
          return { row, col, value: null, style: { borders: { left: { width: 3, style: "solid", color: "#0000ff" } } } };
        }
        return { row, col, value: null };
      },
      getMergedRangeAt: (row, col) => (contains(merged, row, col) ? merged : null),
      getMergedRangesInRange: (range) => (intersects(range, merged) ? [merged] : [])
    };

    const gridCanvas = document.createElement("canvas");
    const contentCanvas = document.createElement("canvas");
    const selectionCanvas = document.createElement("canvas");

    const { ctx: gridCtx, strokes: gridStrokes } = createRecording2dContext(gridCanvas);
    const contentCtx = createRecording2dContext(contentCanvas).ctx;
    const selectionCtx = createRecording2dContext(selectionCanvas).ctx;

    const contexts = new Map<HTMLCanvasElement, CanvasRenderingContext2D>([
      [gridCanvas, gridCtx],
      [contentCanvas, contentCtx],
      [selectionCanvas, selectionCtx]
    ]);

    HTMLCanvasElement.prototype.getContext = vi.fn(function (this: HTMLCanvasElement) {
      return contexts.get(this) ?? createRecording2dContext(this).ctx;
    }) as unknown as typeof HTMLCanvasElement.prototype.getContext;

    const renderer = new CanvasGridRenderer({
      provider,
      rowCount: 2,
      colCount: 3,
      defaultRowHeight: 10,
      defaultColWidth: 10
    });
    renderer.attach({ grid: gridCanvas, content: contentCanvas, selection: selectionCanvas });
    renderer.resize(200, 100, 1);
    renderer.renderImmediately();

    // Double border width=3 -> per-line width=1, offset=1 -> x=20±1 with crisp alignment -> 19.5 / 21.5.
    const purpleStroke = gridStrokes.find((stroke) => stroke.strokeStyle === "#ff00ff" && stroke.lineWidth === 1);
    expect(purpleStroke).toBeTruthy();
    expect(hasNormalizedSegment(purpleStroke!.segments, { x1: 19.5, y1: 0, x2: 19.5, y2: 10 })).toBe(true);
    expect(hasNormalizedSegment(purpleStroke!.segments, { x1: 19.5, y1: 10, x2: 19.5, y2: 20 })).toBe(true);
    expect(hasNormalizedSegment(purpleStroke!.segments, { x1: 21.5, y1: 0, x2: 21.5, y2: 10 })).toBe(true);
    expect(hasNormalizedSegment(purpleStroke!.segments, { x1: 21.5, y1: 10, x2: 21.5, y2: 20 })).toBe(true);

    expect(gridStrokes.some((stroke) => stroke.strokeStyle === "#0000ff")).toBe(false);
  });

  it("pads provider-update dirty regions so bottom borders spanning into the next row are not clipped", () => {
    let listener: ((update: CellProviderUpdate) => void) | undefined;
    let enabled = false;

    const provider: CellProvider = {
      getCell: (row, col) => {
        if (row === 0 && col === 0) {
          return enabled
            ? { row, col, value: null, style: { borders: { bottom: { width: 3, style: "double", color: "#ff00ff" } } } }
            : { row, col, value: null };
        }
        if (row === 1 && col === 0) return { row, col, value: null };
        return null;
      },
      subscribe: (cb) => {
        listener = cb;
        return () => {
          listener = undefined;
        };
      }
    };

    const gridCanvas = document.createElement("canvas");
    const contentCanvas = document.createElement("canvas");
    const selectionCanvas = document.createElement("canvas");

    const { ctx: gridCtx, strokes: gridStrokes, rectCalls } = createRecording2dContext(gridCanvas);
    const contentCtx = createRecording2dContext(contentCanvas).ctx;
    const selectionCtx = createRecording2dContext(selectionCanvas).ctx;

    const contexts = new Map<HTMLCanvasElement, CanvasRenderingContext2D>([
      [gridCanvas, gridCtx],
      [contentCanvas, contentCtx],
      [selectionCanvas, selectionCtx]
    ]);

    HTMLCanvasElement.prototype.getContext = vi.fn(function (this: HTMLCanvasElement) {
      return contexts.get(this) ?? createRecording2dContext(this).ctx;
    }) as unknown as typeof HTMLCanvasElement.prototype.getContext;

    const renderer = new CanvasGridRenderer({
      provider,
      rowCount: 2,
      colCount: 1,
      defaultRowHeight: 10,
      defaultColWidth: 10
    });
    renderer.attach({ grid: gridCanvas, content: contentCanvas, selection: selectionCanvas });
    renderer.resize(100, 100, 1);

    // Clear initial render calls.
    rectCalls.length = 0;
    gridStrokes.length = 0;

    enabled = true;
    if (!listener) {
      throw new Error("Expected provider subscriber to be registered.");
    }
    listener({ type: "cells", range: { startRow: 0, endRow: 1, startCol: 0, endCol: 1 } });

    // The dirty-region clip rect should be padded beyond the 10px row height so the second double-border
    // stroke at y=11.5 is not clipped.
    expect(rectCalls[0]).toMatchObject({ x: 0, y: 0, width: 12, height: 12 });

    const purpleStroke = gridStrokes.find((stroke) => stroke.strokeStyle === "#ff00ff" && stroke.lineWidth === 1);
    expect(purpleStroke).toBeTruthy();
    expect(hasNormalizedSegment(purpleStroke!.segments, { x1: 0, y1: 9.5, x2: 10, y2: 9.5 })).toBe(true);
    expect(hasNormalizedSegment(purpleStroke!.segments, { x1: 0, y1: 11.5, x2: 10, y2: 11.5 })).toBe(true);
  });

  it("pads blit-scroll dirty regions so border strokes are not clipped", () => {
    const provider: CellProvider = {
      getCell: () => null
    };

    const gridCanvas = document.createElement("canvas");
    const contentCanvas = document.createElement("canvas");
    const selectionCanvas = document.createElement("canvas");

    const { ctx: gridCtx, rectCalls } = createRecording2dContext(gridCanvas);
    const contentCtx = createRecording2dContext(contentCanvas).ctx;
    const selectionCtx = createRecording2dContext(selectionCanvas).ctx;

    const contexts = new Map<HTMLCanvasElement, CanvasRenderingContext2D>([
      [gridCanvas, gridCtx],
      [contentCanvas, contentCtx],
      [selectionCanvas, selectionCtx]
    ]);

    HTMLCanvasElement.prototype.getContext = vi.fn(function (this: HTMLCanvasElement) {
      return contexts.get(this) ?? createRecording2dContext(this).ctx;
    }) as unknown as typeof HTMLCanvasElement.prototype.getContext;

    const renderer = new CanvasGridRenderer({
      provider,
      rowCount: 100,
      colCount: 100,
      defaultRowHeight: 10,
      defaultColWidth: 10
    });
    renderer.setPerfStatsEnabled(true);
    renderer.attach({ grid: gridCanvas, content: contentCanvas, selection: selectionCanvas });
    renderer.resize(50, 20, 1);

    // Clear initial render calls so we only inspect the scroll render.
    rectCalls.length = 0;

    // Scroll by 1px; with DPR=1 this should use blit, producing a narrow dirty stripe.
    renderer.setScroll(1, 0);
    expect(renderer.getPerfStats().blitUsed).toBe(true);

    // Find the clip rect used for the dirty stripe render (should be narrower than the full viewport).
    const stripeClip = rectCalls.find((rect) => rect.width < 50 && rect.height === 20);
    expect(stripeClip).toEqual({ x: 47, y: 0, width: 3, height: 20 });
  });
});
