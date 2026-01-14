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
  lineCaps: CanvasLineCap[];
};

function createRecordingContext(canvas: HTMLCanvasElement): { ctx: CanvasRenderingContext2D; rec: Recording } {
  const rec: Recording = {
    fonts: [],
    strokes: 0,
    lineWidths: [],
    strokeStyles: [],
    lineDashes: [],
    lineCaps: []
  };

  let font = "";
  let strokeStyle: string | CanvasGradient | CanvasPattern = "#000";
  let fillStyle: string | CanvasGradient | CanvasPattern = "#000";
  let lineWidth = 1;
  let lineCap: CanvasLineCap = "butt";

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
    get lineCap() {
      return lineCap;
    },
    set lineCap(value: CanvasLineCap) {
      lineCap = value;
      rec.lineCaps.push(value);
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

  it("renders double underline in wrapped layout mode for plain text", () => {
    const provider: CellProvider = {
      getCell: (row, col) =>
        row === 0 && col === 0
          ? { row, col, value: "Hello world", style: { underlineStyle: "double", wrapMode: "word" } }
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

    // Deterministic layout stub (same shape used by other tests).
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

    const lineToCalls = ((content.ctx as unknown as { lineTo: unknown }).lineTo as any).mock?.calls?.length ?? 0;
    expect(lineToCalls).toBeGreaterThanOrEqual(2);
  });

  it("applies baseline shift + smaller font size for style.fontVariantPosition (sub/superscript)", () => {
    const measureBaselineY = (fontVariantPosition?: "subscript" | "superscript") => {
      const provider: CellProvider = {
        getCell: (row, col) =>
          row === 0 && col === 0
            ? { row, col, value: "A", style: fontVariantPosition ? { fontVariantPosition } : {} }
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

      const renderer = new CanvasGridRenderer({ provider, rowCount: 1, colCount: 1, defaultColWidth: 100, defaultRowHeight: 40 });
      renderer.attach({ grid: gridCanvas, content: contentCanvas, selection: selectionCanvas });
      renderer.resize(200, 80, 1);
      renderer.renderImmediately();

      const calls = (content.ctx as any).fillText.mock.calls as any[];
      const call = calls.find((args) => args?.[0] === "A");
      expect(call, `Expected fillText("A", ...) to be called (variant=${String(fontVariantPosition)})`).toBeTruthy();
      const y = Number(call?.[2]);
      expect(Number.isFinite(y)).toBe(true);
      return { y, fonts: content.rec.fonts };
    };

    const normal = measureBaselineY(undefined);
    const superscript = measureBaselineY("superscript");
    const subscript = measureBaselineY("subscript");

    expect(superscript.y).toBeLessThan(normal.y);
    expect(subscript.y).toBeGreaterThan(normal.y);

    // Font size should be smaller for sub/superscript.
    expect(superscript.fonts.join("\n")).not.toEqual(normal.fonts.join("\n"));
  });

  it("renders double underline for plain text when style.underlineStyle='double'", () => {
    const provider: CellProvider = {
      getCell: (row, col) =>
        row === 0 && col === 0
          ? { row, col, value: "Hi", style: { underline: true, underlineStyle: "double" } }
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

    const renderer = new CanvasGridRenderer({ provider, rowCount: 2, colCount: 2 });
    renderer.attach({ grid: gridCanvas, content: contentCanvas, selection: selectionCanvas });
    renderer.resize(200, 80, 1);
    renderer.renderImmediately();

    const lineToCalls = ((content.ctx as unknown as { lineTo: unknown }).lineTo as any).mock?.calls?.length ?? 0;
    expect(lineToCalls).toBeGreaterThanOrEqual(2);
  });

  it("draws strike-through for rich text when style.strike=true", () => {
    const provider: CellProvider = {
      getCell: (row, col) =>
        row === 0 && col === 0
          ? {
              row,
              col,
              value: "Hi",
              richText: { text: "Hi" },
              style: { strike: true }
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

    const renderer = new CanvasGridRenderer({ provider, rowCount: 2, colCount: 2 });
    renderer.attach({ grid: gridCanvas, content: contentCanvas, selection: selectionCanvas });
    renderer.resize(200, 80, 1);
    renderer.renderImmediately();

    expect(content.rec.strokes).toBeGreaterThan(0);
  });

  it("renders double underline for rich text runs", () => {
    const provider: CellProvider = {
      getCell: (row, col) =>
        row === 0 && col === 0
          ? {
              row,
              col,
              value: "Hi",
              richText: { text: "Hi", runs: [{ start: 0, end: 2, style: { underline: "double" } }] }
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

    const renderer = new CanvasGridRenderer({ provider, rowCount: 2, colCount: 2 });
    renderer.attach({ grid: gridCanvas, content: contentCanvas, selection: selectionCanvas });
    renderer.resize(200, 80, 1);
    renderer.renderImmediately();

    const lineToCalls = ((content.ctx as unknown as { lineTo: unknown }).lineTo as any).mock?.calls?.length ?? 0;
    expect(lineToCalls).toBeGreaterThanOrEqual(2);
  });

  it("renders double underline for rich text in wrapped layout mode", () => {
    const provider: CellProvider = {
      getCell: (row, col) =>
        row === 0 && col === 0
          ? {
              row,
              col,
              value: "Hello world",
              richText: { text: "Hello world", runs: [{ start: 0, end: 5, style: { underline: "double" } }] },
              style: { wrapMode: "word" }
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

    const renderer = new CanvasGridRenderer({ provider, rowCount: 2, colCount: 2, defaultColWidth: 60, defaultRowHeight: 20 });
    renderer.attach({ grid: gridCanvas, content: contentCanvas, selection: selectionCanvas });
    renderer.resize(200, 80, 1);

    (renderer as unknown as { textLayoutEngine?: unknown }).textLayoutEngine = {
      measure: (text: string) => ({ width: text.length * 6, ascent: 8, descent: 2 }),
      layout: (options: any) => ({
        lines: [
          {
            text: "Hello",
            runs: options.runs ?? [],
            width: 30,
            ascent: 8,
            descent: 2,
            x: 0
          }
        ],
        width: 30,
        height: 10,
        lineHeight: 10,
        direction: "ltr",
        maxWidth: options.maxWidth ?? 100,
        resolvedAlign: "left"
      })
    };

    renderer.renderImmediately();

    const lineToCalls = ((content.ctx as unknown as { lineTo: unknown }).lineTo as any).mock?.calls?.length ?? 0;
    expect(lineToCalls).toBeGreaterThanOrEqual(2);
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

  it("renders dotted borders with round line caps (Excel-like)", () => {
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
                  bottom: { width: 1, style: "dotted", color: borderColor }
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
    renderer.renderImmediately();

    expect(grid.rec.strokeStyles).toContain(borderColor);
    expect(grid.rec.lineCaps).toContain("round");
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

  it("resolves shared-edge border conflicts deterministically on ties (prefer right/bottom)", () => {
    const leftColor = "#ff0000";
    const rightColor = "#0000ff";

    const provider: CellProvider = {
      getCell: (row, col) => {
        if (row !== 0) return null;
        if (col === 0) {
          return { row, col, value: null, style: { borders: { right: { width: 1, style: "solid", color: leftColor } } } };
        }
        if (col === 1) {
          return { row, col, value: null, style: { borders: { left: { width: 1, style: "solid", color: rightColor } } } };
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

    expect(grid.rec.strokeStyles).toContain(rightColor);
    expect(grid.rec.strokeStyles).not.toContain(leftColor);
  });

  it("resolves shared-edge border conflicts by style rank when widths tie", () => {
    const solidColor = "#ff0000";
    const dottedColor = "#0000ff";

    const provider: CellProvider = {
      getCell: (row, col) => {
        if (row !== 0) return null;
        if (col === 0) {
          return { row, col, value: null, style: { borders: { right: { width: 1, style: "solid", color: solidColor } } } };
        }
        if (col === 1) {
          return { row, col, value: null, style: { borders: { left: { width: 1, style: "dotted", color: dottedColor } } } };
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

    expect(grid.rec.strokeStyles).toContain(solidColor);
    expect(grid.rec.strokeStyles).not.toContain(dottedColor);
    expect(grid.rec.lineCaps).not.toContain("round");
  });
});
