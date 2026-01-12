// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type { CellProvider } from "../../model/CellProvider";
import { CanvasGridRenderer } from "../CanvasGridRenderer";

type FillStyle = string | CanvasGradient | CanvasPattern;

function createMock2dContext(options: {
  canvas: HTMLCanvasElement;
  onFillText?: (args: { text: string; x: number; y: number }) => void;
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
    stroke: noop,
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
      options.onFillText?.({ text, x, y });
    },
    measureText: (text: string) =>
      ({
        width: text.length * 6,
        actualBoundingBoxAscent: 8,
        actualBoundingBoxDescent: 2
      }) as TextMetrics
  } as unknown as CanvasRenderingContext2D;
}

describe("CanvasGridRenderer horizontalAlign=justify", () => {
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

  it("draws justified wrapped lines by drawing words separately (not a single fillText per line)", () => {
    const provider: CellProvider = {
      getCell: (row, col) => {
        if (row !== 0 || col !== 0) return null;
        return { row, col, value: "aa bb cc dd", style: { horizontalAlign: "justify", wrapMode: "word" } };
      }
    };

    const gridCanvas = document.createElement("canvas");
    const contentCanvas = document.createElement("canvas");
    const selectionCanvas = document.createElement("canvas");

    const fillTextCalls: Array<{ text: string; x: number; y: number }> = [];

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
      const existing = contexts.get(this);
      if (existing) return existing;
      // CanvasGridRenderer creates an internal measurer canvas; ensure it can acquire a 2D context.
      const fallback = createMock2dContext({ canvas: this });
      contexts.set(this, fallback);
      return fallback;
    }) as unknown as typeof HTMLCanvasElement.prototype.getContext;

    const renderer = new CanvasGridRenderer({ provider, rowCount: 1, colCount: 1 });
    renderer.attach({ grid: gridCanvas, content: contentCanvas, selection: selectionCanvas });

    // Ensure the row is tall enough to render multiple lines (wrap + justify only triggers for 2+ lines).
    renderer.applyAxisSizeOverrides({ cols: new Map([[0, 60]]), rows: new Map([[0, 50]]) });
    renderer.resize(200, 120, 1);
    renderer.renderImmediately();

    // Normal wrapped rendering uses one fillText call per line (with spaces intact).
    // The justify implementation draws per-word fragments to inject additional spacing.
    expect(fillTextCalls.length).toBeGreaterThan(2);
    expect(fillTextCalls.some((c) => c.text === "aa bb cc")).toBe(false);
    expect(fillTextCalls.some((c) => c.text === "aa")).toBe(true);
    expect(fillTextCalls.some((c) => c.text === "bb")).toBe(true);
    expect(fillTextCalls.some((c) => c.text === "cc")).toBe(true);
  });

  it("draws justified wrapped rich text by splitting word tokens within a line", () => {
    const provider: CellProvider = {
      getCell: (row, col) => {
        if (row !== 0 || col !== 0) return null;
        return {
          row,
          col,
          value: null,
          richText: { text: "aa bb cc dd" },
          style: { horizontalAlign: "justify", wrapMode: "word" }
        };
      }
    };

    const gridCanvas = document.createElement("canvas");
    const contentCanvas = document.createElement("canvas");
    const selectionCanvas = document.createElement("canvas");

    const fillTextCalls: Array<{ text: string; x: number; y: number }> = [];

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
      const existing = contexts.get(this);
      if (existing) return existing;
      // CanvasGridRenderer creates an internal measurer canvas; ensure it can acquire a 2D context.
      const fallback = createMock2dContext({ canvas: this });
      contexts.set(this, fallback);
      return fallback;
    }) as unknown as typeof HTMLCanvasElement.prototype.getContext;

    const renderer = new CanvasGridRenderer({ provider, rowCount: 1, colCount: 1 });
    renderer.attach({ grid: gridCanvas, content: contentCanvas, selection: selectionCanvas });

    // Configure a width that wraps into multiple lines, but keeps multiple words on the first line.
    // With our mock measurer (6px/char) and paddingX=4, col width 48 => maxWidth ~40.
    renderer.applyAxisSizeOverrides({ cols: new Map([[0, 48]]), rows: new Map([[0, 40]]) });
    renderer.resize(200, 120, 1);
    renderer.renderImmediately();

    // Without justify, the first line would be drawn as a single run "aa bb".
    // With justify, we expect per-word fillText calls for that non-final line.
    expect(fillTextCalls.some((c) => c.text === "aa bb")).toBe(false);
    expect(fillTextCalls.some((c) => c.text === "aa")).toBe(true);
    expect(fillTextCalls.some((c) => c.text === "bb")).toBe(true);
  });
});
