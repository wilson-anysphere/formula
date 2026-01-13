// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type { CellProvider } from "../../model/CellProvider";
import { CanvasGridRenderer } from "../CanvasGridRenderer";

type Recording = {
  fonts: string[];
};

function createRecordingContext(canvas: HTMLCanvasElement): { ctx: CanvasRenderingContext2D; rec: Recording } {
  const rec: Recording = { fonts: [] };

  let font = "";
  let fillStyle: string | CanvasGradient | CanvasPattern = "#000";
  let strokeStyle: string | CanvasGradient | CanvasPattern = "#000";
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
    get fillStyle() {
      return fillStyle;
    },
    set fillStyle(value: string | CanvasGradient | CanvasPattern) {
      fillStyle = value;
    },
    get strokeStyle() {
      return strokeStyle;
    },
    set strokeStyle(value: string | CanvasGradient | CanvasPattern) {
      strokeStyle = value;
    },
    get lineWidth() {
      return lineWidth;
    },
    set lineWidth(value: number) {
      lineWidth = value;
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
    stroke: vi.fn(),
    moveTo: vi.fn(),
    lineTo: vi.fn(),
    closePath: vi.fn(),
    save: vi.fn(),
    restore: vi.fn(),
    drawImage: vi.fn(),
    translate: vi.fn(),
    rotate: vi.fn(),
    setLineDash: vi.fn(),
    fillText: vi.fn(),
    measureText: (text: string) =>
      ({
        width: text.length * 6,
        actualBoundingBoxAscent: 8,
        actualBoundingBoxDescent: 2
      }) as TextMetrics
  };

  return { ctx: ctx as unknown as CanvasRenderingContext2D, rec };
}

describe("CanvasGridRenderer default font family", () => {
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

  it("uses defaultCellFontFamily when CellStyle.fontFamily is unset", () => {
    const provider: CellProvider = {
      getCell: (row, col) => (row === 0 && col === 0 ? { row, col, value: "A" } : null)
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

    const renderer = new CanvasGridRenderer({
      provider,
      rowCount: 2,
      colCount: 2,
      defaultCellFontFamily: "monospace"
    });
    renderer.attach({ grid: gridCanvas, content: contentCanvas, selection: selectionCanvas });
    renderer.resize(200, 80, 1);
    renderer.renderImmediately();

    expect(content.rec.fonts.some((f) => f.includes("monospace"))).toBe(true);
  });

  it("allows CellStyle.fontFamily to override defaultCellFontFamily", () => {
    const provider: CellProvider = {
      getCell: (row, col) =>
        row === 0 && col === 0
          ? { row, col, value: "A", style: { fontFamily: "serif" } }
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

    const renderer = new CanvasGridRenderer({
      provider,
      rowCount: 2,
      colCount: 2,
      defaultCellFontFamily: "monospace"
    });
    renderer.attach({ grid: gridCanvas, content: contentCanvas, selection: selectionCanvas });
    renderer.resize(200, 80, 1);
    renderer.renderImmediately();

    expect(content.rec.fonts.some((f) => f.includes("serif"))).toBe(true);
    expect(content.rec.fonts.some((f) => f.includes("monospace"))).toBe(false);
  });
});

