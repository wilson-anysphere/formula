// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type { CellProvider } from "../../model/CellProvider";
import { CanvasGridRenderer } from "../CanvasGridRenderer";

type Recording = {
  fonts: string[];
  strokes: number;
  lineWidths: number[];
  strokeStyles: string[];
  lineDashes: number[][];
};

function createRecordingContext(canvas: HTMLCanvasElement): { ctx: CanvasRenderingContext2D; rec: Recording } {
  const rec: Recording = {
    fonts: [],
    strokes: 0,
    lineWidths: [],
    strokeStyles: [],
    lineDashes: []
  };

  let font = "";
  let strokeStyle: string | CanvasGradient | CanvasPattern = "#000";
  let fillStyle: string | CanvasGradient | CanvasPattern = "#000";
  let lineWidth = 1;

  const ctx: Partial<CanvasRenderingContext2D> = {
    canvas,
    get font() {
      return font;
    },
    set font(value: string) {
      font = value;
      rec.fonts.push(value);
    },
    get strokeStyle() {
      return strokeStyle;
    },
    set strokeStyle(value: string | CanvasGradient | CanvasPattern) {
      strokeStyle = value;
      if (typeof value === "string") rec.strokeStyles.push(value);
    },
    get fillStyle() {
      return fillStyle;
    },
    set fillStyle(value: string | CanvasGradient | CanvasPattern) {
      fillStyle = value;
    },
    get lineWidth() {
      return lineWidth;
    },
    set lineWidth(value: number) {
      lineWidth = value;
      rec.lineWidths.push(value);
    },
    textAlign: "left",
    textBaseline: "alphabetic",
    globalAlpha: 1,
    imageSmoothingEnabled: false,
    setTransform: vi.fn(),
    clearRect: vi.fn(),
    fillRect: vi.fn(),
    strokeRect: vi.fn(),
    beginPath: vi.fn(),
    rect: vi.fn(),
    clip: vi.fn(),
    fill: vi.fn(),
    stroke: vi.fn(() => {
      rec.strokes += 1;
    }),
    moveTo: vi.fn(),
    lineTo: vi.fn(),
    closePath: vi.fn(),
    save: vi.fn(),
    restore: vi.fn(),
    drawImage: vi.fn(),
    translate: vi.fn(),
    rotate: vi.fn(),
    fillText: vi.fn(),
    setLineDash: vi.fn((segments: number[]) => {
      rec.lineDashes.push([...segments]);
    }),
    measureText: (text: string) =>
      ({
        width: text.length * 6,
        actualBoundingBoxAscent: 8,
        actualBoundingBoxDescent: 2
      }) as TextMetrics
  };

  return { ctx: ctx as unknown as CanvasRenderingContext2D, rec };
}

describe("CanvasGridRenderer CellStyle primitives", () => {
  const originalGetContext = HTMLCanvasElement.prototype.getContext;

  beforeEach(() => {
    vi.stubGlobal("requestAnimationFrame", (cb: FrameRequestCallback) => {
      cb(0);
      return 0;
    });
  });

  afterEach(() => {
    HTMLCanvasElement.prototype.getContext = originalGetContext;
    vi.restoreAllMocks();
    vi.unstubAllGlobals();
  });

  it("renders italic fonts via style.fontStyle", () => {
    const provider: CellProvider = {
      getCell: (row, col) => (row === 0 && col === 0 ? { row, col, value: "A", style: { fontStyle: "italic" } } : null)
    };

    const gridCanvas = document.createElement("canvas");
    const contentCanvas = document.createElement("canvas");
    const selectionCanvas = document.createElement("canvas");

    const grid = createRecordingContext(gridCanvas);
    const content = createRecordingContext(contentCanvas);
    const selection = createRecordingContext(selectionCanvas);

    const contexts = new Map<HTMLCanvasElement, CanvasRenderingContext2D>([
      [gridCanvas, grid.ctx],
      [contentCanvas, content.ctx],
      [selectionCanvas, selection.ctx]
    ]);

    HTMLCanvasElement.prototype.getContext = vi.fn(function (this: HTMLCanvasElement) {
      const existing = contexts.get(this);
      if (existing) return existing;
      const created = createRecordingContext(this).ctx;
      contexts.set(this, created);
      return created;
    }) as unknown as typeof HTMLCanvasElement.prototype.getContext;

    const renderer = new CanvasGridRenderer({ provider, rowCount: 2, colCount: 2 });
    renderer.attach({ grid: gridCanvas, content: contentCanvas, selection: selectionCanvas });
    renderer.resize(200, 80, 1);
    renderer.renderImmediately();

    expect(content.rec.fonts.some((f) => f.startsWith("italic "))).toBe(true);
  });

  it("draws underlines (stroke) in wrapped layout mode", () => {
    const provider: CellProvider = {
      getCell: (row, col) =>
        row === 0 && col === 0
          ? { row, col, value: "Hello world", style: { underline: true, wrapMode: "word" } }
          : null
    };

    const gridCanvas = document.createElement("canvas");
    const contentCanvas = document.createElement("canvas");
    const selectionCanvas = document.createElement("canvas");

    const grid = createRecordingContext(gridCanvas);
    const content = createRecordingContext(contentCanvas);
    const selection = createRecordingContext(selectionCanvas);

    const contexts = new Map<HTMLCanvasElement, CanvasRenderingContext2D>([
      [gridCanvas, grid.ctx],
      [contentCanvas, content.ctx],
      [selectionCanvas, selection.ctx]
    ]);

    HTMLCanvasElement.prototype.getContext = vi.fn(function (this: HTMLCanvasElement) {
      const existing = contexts.get(this);
      if (existing) return existing;
      const created = createRecordingContext(this).ctx;
      contexts.set(this, created);
      return created;
    }) as unknown as typeof HTMLCanvasElement.prototype.getContext;

    const renderer = new CanvasGridRenderer({ provider, rowCount: 2, colCount: 2, defaultColWidth: 50, defaultRowHeight: 20 });
    renderer.attach({ grid: gridCanvas, content: contentCanvas, selection: selectionCanvas });
    renderer.resize(200, 80, 1);

    // Use a deterministic layout stub so the test doesn't depend on HarfBuzz availability.
    (renderer as unknown as { textLayoutEngine?: unknown }).textLayoutEngine = {
      layout: () => ({
        lines: [{ text: "Hello", runs: [], width: 30, ascent: 8, descent: 2, x: 0 }],
        width: 30,
        height: 10,
        lineHeight: 10,
        direction: "ltr",
        maxWidth: 100,
        resolvedAlign: "left"
      })
    };

    renderer.renderImmediately();

    expect(content.rec.strokes).toBeGreaterThan(0);
  });

  it("renders per-cell borders with zoom-scaled widths", () => {
    const borderColor = "#ff0000";
    const provider: CellProvider = {
      getCell: (row, col) =>
        row === 0 && col === 0
          ? {
              row,
              col,
              value: null,
              style: {
                borders: {
                  bottom: { width: 2, style: "solid", color: borderColor }
                }
              }
            }
          : null
    };

    const gridCanvas = document.createElement("canvas");
    const contentCanvas = document.createElement("canvas");
    const selectionCanvas = document.createElement("canvas");

    const grid = createRecordingContext(gridCanvas);
    const content = createRecordingContext(contentCanvas);
    const selection = createRecordingContext(selectionCanvas);

    const contexts = new Map<HTMLCanvasElement, CanvasRenderingContext2D>([
      [gridCanvas, grid.ctx],
      [contentCanvas, content.ctx],
      [selectionCanvas, selection.ctx]
    ]);

    HTMLCanvasElement.prototype.getContext = vi.fn(function (this: HTMLCanvasElement) {
      const existing = contexts.get(this);
      if (existing) return existing;
      const created = createRecordingContext(this).ctx;
      contexts.set(this, created);
      return created;
    }) as unknown as typeof HTMLCanvasElement.prototype.getContext;

    const renderer = new CanvasGridRenderer({ provider, rowCount: 2, colCount: 2, defaultColWidth: 50, defaultRowHeight: 20 });
    renderer.attach({ grid: gridCanvas, content: contentCanvas, selection: selectionCanvas });
    renderer.resize(200, 80, 1);

    renderer.setZoom(2);
    renderer.renderImmediately();

    expect(grid.rec.strokeStyles).toContain(borderColor);
    // width=2 at zoom=2 => effective 4px.
    expect(grid.rec.lineWidths).toContain(4);
  });

  it("resolves shared-edge border conflicts in favor of thicker widths", () => {
    const thinColor = "#ff0000";
    const thickColor = "#0000ff";

    const provider: CellProvider = {
      getCell: (row, col) => {
        if (row !== 0) return null;
        if (col === 0) {
          return { row, col, value: null, style: { borders: { right: { width: 1, style: "solid", color: thinColor } } } };
        }
        if (col === 1) {
          return { row, col, value: null, style: { borders: { left: { width: 2, style: "solid", color: thickColor } } } };
        }
        return null;
      }
    };

    const gridCanvas = document.createElement("canvas");
    const contentCanvas = document.createElement("canvas");
    const selectionCanvas = document.createElement("canvas");

    const grid = createRecordingContext(gridCanvas);
    const content = createRecordingContext(contentCanvas);
    const selection = createRecordingContext(selectionCanvas);

    const contexts = new Map<HTMLCanvasElement, CanvasRenderingContext2D>([
      [gridCanvas, grid.ctx],
      [contentCanvas, content.ctx],
      [selectionCanvas, selection.ctx]
    ]);

    HTMLCanvasElement.prototype.getContext = vi.fn(function (this: HTMLCanvasElement) {
      const existing = contexts.get(this);
      if (existing) return existing;
      const created = createRecordingContext(this).ctx;
      contexts.set(this, created);
      return created;
    }) as unknown as typeof HTMLCanvasElement.prototype.getContext;

    const renderer = new CanvasGridRenderer({ provider, rowCount: 1, colCount: 2, defaultColWidth: 50, defaultRowHeight: 20 });
    renderer.attach({ grid: gridCanvas, content: contentCanvas, selection: selectionCanvas });
    renderer.resize(200, 80, 1);
    renderer.renderImmediately();

    expect(grid.rec.strokeStyles).toContain(thickColor);
    expect(grid.rec.lineWidths).toContain(2);
    expect(grid.rec.strokeStyles).not.toContain(thinColor);
  });
});
