// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type { CellProvider } from "../../model/CellProvider";
import { CanvasGridRenderer } from "../CanvasGridRenderer";

type FillStyle = string | CanvasGradient | CanvasPattern;

function createMock2dContext(options: {
  canvas: HTMLCanvasElement;
  onFillText?: (args: { text: string; x: number; y: number; font: string; fillStyle: FillStyle }) => void;
  onStroke?: () => void;
}): CanvasRenderingContext2D {
  const noop = () => {};
  let fillStyle: FillStyle = "#000";
  let strokeStyle: FillStyle = "#000";
  let font = "";

  return {
    canvas: options.canvas,
    get fillStyle() {
      return fillStyle;
    },
    set fillStyle(value: FillStyle) {
      fillStyle = value;
    },
    get strokeStyle() {
      return strokeStyle;
    },
    set strokeStyle(value: FillStyle) {
      strokeStyle = value;
    },
    get font() {
      return font;
    },
    set font(value: string) {
      font = value;
    },
    lineWidth: 1,
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
    stroke: () => options.onStroke?.(),
    moveTo: noop,
    lineTo: noop,
    closePath: noop,
    save: noop,
    restore: noop,
    drawImage: noop,
    translate: noop,
    rotate: noop,
    setLineDash: noop,
    fillText: (text: string, x: number, y: number) => {
      options.onFillText?.({ text, x, y, font, fillStyle });
    },
    measureText: (text: string) =>
      ({
        width: text.length * 6,
        actualBoundingBoxAscent: 8,
        actualBoundingBoxDescent: 2
      }) as TextMetrics
  } as unknown as CanvasRenderingContext2D;
}

describe("CanvasGridRenderer rich text rendering", () => {
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
    vi.restoreAllMocks();
  });

  it("renders rich text runs with per-run fonts and underlines", () => {
    const richText = {
      text: "Hello world",
      runs: [
        // Only style the first run; the renderer should fill the remaining range with defaults.
        // Also explicitly disable bold/italic so we exercise the "false overrides defaults" behavior.
        { start: 0, end: 5, style: { italic: false, bold: false, underline: true, color: "#80FF0000", size_100pt: 1200 } }
      ]
    };

    const provider: CellProvider = {
      getCell: (row, col) => {
        if (row !== 0 || col !== 0) return null;
        return { row, col, value: richText.text, richText, style: { fontWeight: "700", fontStyle: "italic" } };
      }
    };

    const gridCanvas = document.createElement("canvas");
    const contentCanvas = document.createElement("canvas");
    const selectionCanvas = document.createElement("canvas");

    const fillTextCalls: Array<{ text: string; x: number; y: number; font: string; fillStyle: FillStyle }> = [];
    const strokeCalls: number[] = [];

    const contexts = new Map<HTMLCanvasElement, CanvasRenderingContext2D>();
    contexts.set(gridCanvas, createMock2dContext({ canvas: gridCanvas }));
    contexts.set(
      contentCanvas,
      createMock2dContext({
        canvas: contentCanvas,
        onFillText: (args) => fillTextCalls.push(args),
        onStroke: () => strokeCalls.push(1)
      })
    );
    contexts.set(selectionCanvas, createMock2dContext({ canvas: selectionCanvas }));

    HTMLCanvasElement.prototype.getContext = vi.fn(function (this: HTMLCanvasElement) {
      const ctx = contexts.get(this);
      return ctx ?? createMock2dContext({ canvas: this });
    }) as unknown as typeof HTMLCanvasElement.prototype.getContext;

    const renderer = new CanvasGridRenderer({ provider, rowCount: 1, colCount: 1 });
    renderer.attach({ grid: gridCanvas, content: contentCanvas, selection: selectionCanvas });
    renderer.resize(200, 80, 1);
    renderer.renderImmediately();

    expect(fillTextCalls.length).toBeGreaterThanOrEqual(2);
    const uniqueFonts = new Set(fillTextCalls.map((c) => c.font));
    expect(uniqueFonts.size).toBeGreaterThanOrEqual(2);
    expect(fillTextCalls.some((c) => c.font.startsWith("normal normal"))).toBe(true);
    expect(fillTextCalls.some((c) => c.font.startsWith("italic 700"))).toBe(true);
    expect(fillTextCalls.some((c) => c.font.includes("16px"))).toBe(true);

    const styledRun = fillTextCalls.find((c) => c.text === "Hello");
    expect(styledRun).toBeTruthy();
    expect(typeof styledRun?.fillStyle === "string" ? styledRun.fillStyle : "").toMatch(/^rgba\(255,\s*0,\s*0,/);
    // We expect at least one underline stroke from the italic+underline run.
    expect(strokeCalls.length).toBeGreaterThan(0);
  });

  it("interprets run start/end as Unicode code point indices (surrogate pairs)", () => {
    const richText = {
      text: "A😀B",
      // 😀 is a surrogate pair in UTF-16, but occupies 1 code point.
      // If we treat start/end as code points, (1,2) should select only the emoji.
      runs: [{ start: 1, end: 2, style: { underline: true } }]
    };

    const provider: CellProvider = {
      getCell: (row, col) => {
        if (row !== 0 || col !== 0) return null;
        return { row, col, value: richText.text, richText };
      }
    };

    const gridCanvas = document.createElement("canvas");
    const contentCanvas = document.createElement("canvas");
    const selectionCanvas = document.createElement("canvas");

    const fillTextCalls: Array<{ text: string; x: number; y: number; font: string; fillStyle: FillStyle }> = [];

    const contexts = new Map<HTMLCanvasElement, CanvasRenderingContext2D>();
    contexts.set(gridCanvas, createMock2dContext({ canvas: gridCanvas }));
    contexts.set(
      contentCanvas,
      createMock2dContext({
        canvas: contentCanvas,
        onFillText: (args) => fillTextCalls.push(args)
      })
    );
    contexts.set(selectionCanvas, createMock2dContext({ canvas: selectionCanvas }));

    HTMLCanvasElement.prototype.getContext = vi.fn(function (this: HTMLCanvasElement) {
      const ctx = contexts.get(this);
      return ctx ?? createMock2dContext({ canvas: this });
    }) as unknown as typeof HTMLCanvasElement.prototype.getContext;

    const renderer = new CanvasGridRenderer({ provider, rowCount: 1, colCount: 1 });
    renderer.attach({ grid: gridCanvas, content: contentCanvas, selection: selectionCanvas });
    renderer.resize(200, 80, 1);
    renderer.renderImmediately();

    expect(fillTextCalls.map((c) => c.text)).toContain("😀");
  });
});
