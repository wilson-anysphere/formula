// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type { CellBorderSpec, CellProvider } from "../../model/CellProvider";
import { CanvasGridRenderer } from "../CanvasGridRenderer";

type StrokeRecord = {
  strokeStyle: unknown;
  lineWidth: number;
  lineDash: number[];
  lineCap: CanvasLineCap;
  segments: Array<{ x1: number; y1: number; x2: number; y2: number }>;
};

function createRecording2dContext(canvas: HTMLCanvasElement): { ctx: CanvasRenderingContext2D; strokes: StrokeRecord[] } {
  const noop = () => {};
  let fillStyle: unknown = "#000";
  let strokeStyle: unknown = "#000";
  let lineWidth = 1;
  let lineDash: number[] = [];
  let lineCap: CanvasLineCap = "butt";
  let cursor: { x: number; y: number } | null = null;
  let segments: StrokeRecord["segments"] = [];
  const strokes: StrokeRecord[] = [];

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
    get lineCap() {
      return lineCap;
    },
    set lineCap(value: CanvasLineCap) {
      lineCap = value;
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
    rect: noop,
    clip: noop,
    fill: noop,
    stroke: () => {
      strokes.push({ strokeStyle, lineWidth, lineDash: [...lineDash], lineCap, segments: [...segments] });
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
    setLineDash: (value: number[]) => {
      lineDash = [...value];
    },
    measureText: (text: string) =>
      ({
        width: text.length * 6,
        actualBoundingBoxAscent: 8,
        actualBoundingBoxDescent: 2
      }) as TextMetrics
  } as unknown as CanvasRenderingContext2D;

  return { ctx, strokes };
}

describe("CanvasGridRenderer diagonal borders", () => {
  const originalGetContext = HTMLCanvasElement.prototype.getContext;

  beforeEach(() => {
    vi.stubGlobal("requestAnimationFrame", (cb: FrameRequestCallback) => {
      cb(0);
      return 0;
    });
  });

  afterEach(() => {
    HTMLCanvasElement.prototype.getContext = originalGetContext;
    vi.unstubAllGlobals();
  });

  it("draws diagonal up/down borders with expected stroke style and corner-to-corner segments", () => {
    const up: CellBorderSpec = { width: 2, style: "solid", color: "#ff0000" };
    const down: CellBorderSpec = { width: 2, style: "solid", color: "#0000ff" };

    const provider: CellProvider = {
      getCell: (row, col) => {
        if (row === 0 && col === 0) {
          return { row, col, value: "A", style: { diagonalBorders: { up } } };
        }
        if (row === 0 && col === 1) {
          return { row, col, value: "B", style: { diagonalBorders: { down } } };
        }
        return { row, col, value: null };
      }
    };

    const gridCanvas = document.createElement("canvas");
    const contentCanvas = document.createElement("canvas");
    const selectionCanvas = document.createElement("canvas");

    const { ctx: gridCtx, strokes } = createRecording2dContext(gridCanvas);
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
      defaultRowHeight: 20,
      defaultColWidth: 20
    });

    renderer.attach({ grid: gridCanvas, content: contentCanvas, selection: selectionCanvas });
    renderer.resize(200, 200, 1);
    renderer.renderImmediately();

    expect(strokes.length).toBeGreaterThan(0);

    const upStroke = strokes.find((stroke) => stroke.strokeStyle === "#ff0000" && stroke.lineWidth === 2);
    expect(upStroke).toBeTruthy();
    expect(
      upStroke!.segments.some(
        (seg) => Math.abs(seg.x1 - 0) < 1e-6 && Math.abs(seg.y1 - 20) < 1e-6 && Math.abs(seg.x2 - 20) < 1e-6 && Math.abs(seg.y2 - 0) < 1e-6
      )
    ).toBe(true);

    const downStroke = strokes.find((stroke) => stroke.strokeStyle === "#0000ff" && stroke.lineWidth === 2);
    expect(downStroke).toBeTruthy();
    expect(
      downStroke!.segments.some(
        (seg) =>
          Math.abs(seg.x1 - 20) < 1e-6 &&
          Math.abs(seg.y1 - 0) < 1e-6 &&
          Math.abs(seg.x2 - 40) < 1e-6 &&
          Math.abs(seg.y2 - 20) < 1e-6
      )
    ).toBe(true);
  });

  it("uses round caps + dotted dash patterns for dotted diagonal borders", () => {
    const dotted: CellBorderSpec = { width: 1, style: "dotted", color: "#ff00ff" };

    const provider: CellProvider = {
      getCell: (row, col) => {
        if (row === 0 && col === 0) {
          return { row, col, value: "A", style: { diagonalBorders: { down: dotted } } };
        }
        return { row, col, value: null };
      }
    };

    const gridCanvas = document.createElement("canvas");
    const contentCanvas = document.createElement("canvas");
    const selectionCanvas = document.createElement("canvas");

    const { ctx: gridCtx, strokes } = createRecording2dContext(gridCanvas);
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
      defaultRowHeight: 20,
      defaultColWidth: 20
    });

    renderer.attach({ grid: gridCanvas, content: contentCanvas, selection: selectionCanvas });
    renderer.resize(200, 200, 1);
    renderer.renderImmediately();

    const stroke = strokes.find((entry) => entry.strokeStyle === "#ff00ff" && entry.lineWidth === 1);
    expect(stroke).toBeTruthy();
    expect(stroke!.lineDash.join(",")).toBe("1,2");
    expect(stroke!.lineCap).toBe("round");
    expect(stroke!.segments.length).toBeGreaterThan(0);
  });
});
