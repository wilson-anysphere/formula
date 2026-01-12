// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type { CellProvider, CellStyle } from "../../model/CellProvider";
import { CanvasGridRenderer } from "../CanvasGridRenderer";

type FillStyle = string | CanvasGradient | CanvasPattern;

function createRecording2dContext(options: {
  canvas: HTMLCanvasElement;
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

function renderAndGetFillTextX(options: { style: CellStyle }): number {
  const provider: CellProvider = {
    getCell: (row, col) => {
      if (row !== 0 || col !== 0) return null;
      return {
        row,
        col,
        value: "Indent",
        style: options.style
      };
    }
  };

  const gridCalls: Array<[string, ...any[]]> = [];
  const contentCalls: Array<[string, ...any[]]> = [];
  const selectionCalls: Array<[string, ...any[]]> = [];

  const gridCanvas = document.createElement("canvas");
  const contentCanvas = document.createElement("canvas");
  const selectionCanvas = document.createElement("canvas");

  const contexts = new Map<HTMLCanvasElement, CanvasRenderingContext2D>();
  contexts.set(gridCanvas, createRecording2dContext({ canvas: gridCanvas, calls: gridCalls }));
  contexts.set(contentCanvas, createRecording2dContext({ canvas: contentCanvas, calls: contentCalls }));
  contexts.set(selectionCanvas, createRecording2dContext({ canvas: selectionCanvas, calls: selectionCalls }));

  HTMLCanvasElement.prototype.getContext = vi.fn(function (this: HTMLCanvasElement) {
    const existing = contexts.get(this);
    if (existing) return existing;
    // CanvasGridRenderer creates an internal measurer canvas; ensure it can acquire a 2D context.
    const fallback = createRecording2dContext({ canvas: this, calls: [] });
    contexts.set(this, fallback);
    return fallback;
  }) as unknown as typeof HTMLCanvasElement.prototype.getContext;

  const renderer = new CanvasGridRenderer({ provider, rowCount: 1, colCount: 1 });
  renderer.attach({ grid: gridCanvas, content: contentCanvas, selection: selectionCanvas });
  renderer.resize(200, 80, 1);
  renderer.renderImmediately();

  const fillTextCalls = contentCalls.filter((c) => c[0] === "fillText");
  expect(fillTextCalls).toHaveLength(1);
  const [, , x] = fillTextCalls[0];
  return x as number;
}

describe("CanvasGridRenderer textIndentPx", () => {
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

  it("indents start-aligned text by shifting the draw X coordinate", () => {
    const baselineX = renderAndGetFillTextX({ style: { textAlign: "start", direction: "ltr" } });
    const indentedX = renderAndGetFillTextX({ style: { textAlign: "start", direction: "ltr", textIndentPx: 12 } });
    expect(indentedX - baselineX).toBeCloseTo(12, 5);
  });

  it("indents end-aligned text from the right edge by shifting the draw X coordinate left", () => {
    const baselineX = renderAndGetFillTextX({ style: { textAlign: "end", direction: "ltr" } });
    const indentedX = renderAndGetFillTextX({ style: { textAlign: "end", direction: "ltr", textIndentPx: 12 } });
    expect(baselineX - indentedX).toBeCloseTo(12, 5);
  });
});
