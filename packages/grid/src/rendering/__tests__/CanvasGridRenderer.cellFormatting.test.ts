// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type { CellProvider } from "../../model/CellProvider";
import { CanvasGridRenderer } from "../CanvasGridRenderer";

type FillStyle = string | CanvasGradient | CanvasPattern;

function createRecording2dContext(options: {
  canvas: HTMLCanvasElement;
  onFont?: (value: string) => void;
  onStrokeStyle?: (value: FillStyle) => void;
  calls: Array<[string, ...any[]]>;
}): CanvasRenderingContext2D {
  const noop = () => {};
  let fillStyle: FillStyle = "#000";
  let strokeStyle: FillStyle = "#000";
  let font = "";

  const record = (name: string, ...args: any[]) => options.calls.push([name, ...args]);

  const ctx: any = {
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
      options.onStrokeStyle?.(value);
    },
    get font() {
      return font;
    },
    set font(value: string) {
      font = value;
      options.onFont?.(value);
    },
    lineWidth: 1,
    textAlign: "left",
    textBaseline: "alphabetic",
    globalAlpha: 1,
    imageSmoothingEnabled: false,
    setTransform: noop,
    clearRect: (...args: any[]) => record("clearRect", ...args),
    fillRect: (...args: any[]) => record("fillRect", ...args),
    strokeRect: (...args: any[]) => record("strokeRect", ...args),
    beginPath: (...args: any[]) => record("beginPath", ...args),
    rect: (...args: any[]) => record("rect", ...args),
    clip: (...args: any[]) => record("clip", ...args),
    fill: (...args: any[]) => record("fill", ...args),
    stroke: (...args: any[]) => record("stroke", ...args),
    moveTo: (...args: any[]) => record("moveTo", ...args),
    lineTo: (...args: any[]) => record("lineTo", ...args),
    closePath: (...args: any[]) => record("closePath", ...args),
    save: (...args: any[]) => record("save", ...args),
    restore: (...args: any[]) => record("restore", ...args),
    drawImage: (...args: any[]) => record("drawImage", ...args),
    translate: (...args: any[]) => record("translate", ...args),
    rotate: (...args: any[]) => record("rotate", ...args),
    fillText: (...args: any[]) => record("fillText", ...args),
    measureText: (text: string) =>
      ({
        width: text.length * 6,
        actualBoundingBoxAscent: 8,
        actualBoundingBoxDescent: 2
      }) as TextMetrics
  };

  // Optional API used by border rendering.
  ctx.setLineDash = (...args: any[]) => record("setLineDash", ...args);

  return ctx as CanvasRenderingContext2D;
}

describe("CanvasGridRenderer cell formatting primitives", () => {
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

  it("applies italic fontStyle and draws underline for wrapped text", () => {
    const provider: CellProvider = {
      getCell: (row, col) => {
        if (row === 0 && col === 0) {
          return {
            row,
            col,
            value: "Styled",
            style: { fontStyle: "italic", wrapMode: "word", underline: true }
          };
        }
        return null;
      }
    };

    const contentFonts: string[] = [];
    const gridStrokeStyles: FillStyle[] = [];
    const gridCalls: Array<[string, ...any[]]> = [];
    const contentCalls: Array<[string, ...any[]]> = [];
    const selectionCalls: Array<[string, ...any[]]> = [];

    const gridCanvas = document.createElement("canvas");
    const contentCanvas = document.createElement("canvas");
    const selectionCanvas = document.createElement("canvas");

    const contexts = new Map<HTMLCanvasElement, CanvasRenderingContext2D>();
    contexts.set(
      gridCanvas,
      createRecording2dContext({
        canvas: gridCanvas,
        calls: gridCalls,
        onStrokeStyle: (value) => gridStrokeStyles.push(value)
      })
    );
    contexts.set(
      contentCanvas,
      createRecording2dContext({
        canvas: contentCanvas,
        calls: contentCalls,
        onFont: (value) => contentFonts.push(value)
      })
    );
    contexts.set(selectionCanvas, createRecording2dContext({ canvas: selectionCanvas, calls: selectionCalls }));

    HTMLCanvasElement.prototype.getContext = vi.fn(function (this: HTMLCanvasElement) {
      const existing = contexts.get(this);
      if (existing) return existing;
      // CanvasGridRenderer creates an internal measurer canvas; ensure it can acquire a 2D context.
      const fallback = createRecording2dContext({ canvas: this, calls: [] });
      contexts.set(this, fallback);
      return fallback;
    }) as unknown as typeof HTMLCanvasElement.prototype.getContext;

    const renderer = new CanvasGridRenderer({ provider, rowCount: 2, colCount: 2 });
    renderer.attach({ grid: gridCanvas, content: contentCanvas, selection: selectionCanvas });
    renderer.resize(200, 80, 1);
    renderer.renderImmediately();

    expect(contentFonts.some((f) => f.startsWith("italic "))).toBe(true);
    expect(contentCalls.some((c) => c[0] === "stroke")).toBe(true);
    // Gridline strokes should still occur.
    expect(gridStrokeStyles.length).toBeGreaterThan(0);
  });

  it("draws per-cell borders on the grid layer", () => {
    const provider: CellProvider = {
      getCell: (row, col) => {
        if (row === 0 && col === 0) {
          return {
            row,
            col,
            value: "x",
            style: {
              borders: {
                top: { width: 1, style: "solid", color: "rgb(255,0,0)" },
                right: { width: 1, style: "solid", color: "rgb(255,0,0)" },
                bottom: { width: 1, style: "solid", color: "rgb(255,0,0)" },
                left: { width: 1, style: "solid", color: "rgb(255,0,0)" }
              }
            }
          };
        }
        return null;
      }
    };

    const gridStrokeStyles: FillStyle[] = [];
    const gridCalls: Array<[string, ...any[]]> = [];
    const contentCalls: Array<[string, ...any[]]> = [];
    const selectionCalls: Array<[string, ...any[]]> = [];

    const gridCanvas = document.createElement("canvas");
    const contentCanvas = document.createElement("canvas");
    const selectionCanvas = document.createElement("canvas");

    const contexts = new Map<HTMLCanvasElement, CanvasRenderingContext2D>();
    contexts.set(
      gridCanvas,
      createRecording2dContext({
        canvas: gridCanvas,
        calls: gridCalls,
        onStrokeStyle: (value) => gridStrokeStyles.push(value)
      })
    );
    contexts.set(contentCanvas, createRecording2dContext({ canvas: contentCanvas, calls: contentCalls }));
    contexts.set(selectionCanvas, createRecording2dContext({ canvas: selectionCanvas, calls: selectionCalls }));

    HTMLCanvasElement.prototype.getContext = vi.fn(function (this: HTMLCanvasElement) {
      const existing = contexts.get(this);
      if (existing) return existing;
      const fallback = createRecording2dContext({ canvas: this, calls: [] });
      contexts.set(this, fallback);
      return fallback;
    }) as unknown as typeof HTMLCanvasElement.prototype.getContext;

    const renderer = new CanvasGridRenderer({ provider, rowCount: 2, colCount: 2 });
    renderer.attach({ grid: gridCanvas, content: contentCanvas, selection: selectionCanvas });
    renderer.resize(200, 80, 1);
    renderer.renderImmediately();

    expect(gridStrokeStyles).toContain("rgb(255,0,0)");
    expect(gridCalls.some((c) => c[0] === "stroke")).toBe(true);
  });
});
