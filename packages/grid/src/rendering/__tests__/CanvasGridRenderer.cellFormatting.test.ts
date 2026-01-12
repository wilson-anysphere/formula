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
  let lineWidth = 1;
  let lineDash: number[] = [];

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
    setTransform: noop,
    clearRect: (...args: any[]) => record("clearRect", ...args),
    fillRect: (...args: any[]) => record("fillRect", ...args),
    strokeRect: (...args: any[]) => record("strokeRect", ...args),
    beginPath: (...args: any[]) => record("beginPath", ...args),
    rect: (...args: any[]) => record("rect", ...args),
    clip: (...args: any[]) => record("clip", ...args),
    fill: (...args: any[]) => record("fill", ...args),
    stroke: (...args: any[]) => record("stroke", { strokeStyle, lineWidth, lineDash: [...lineDash] }, ...args),
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
  ctx.setLineDash = (...args: any[]) => {
    const next = Array.isArray(args[0]) ? args[0] : [];
    lineDash = [...next];
    record("setLineDash", ...args);
  };

  return ctx as CanvasRenderingContext2D;
}

describe("CanvasGridRenderer cell formatting primitives", () => {
  const originalGetContext = HTMLCanvasElement.prototype.getContext;
  const originalRaf = globalThis.requestAnimationFrame;

  const segmentsForStroke = (
    calls: Array<[string, ...any[]]>,
    predicate: (strokeState: { strokeStyle: FillStyle; lineWidth: number; lineDash: number[] }) => boolean
  ): Array<{ x1: number; y1: number; x2: number; y2: number }> => {
    const strokeIndex = calls.findIndex((call) => call[0] === "stroke" && predicate(call[1]));
    expect(strokeIndex).toBeGreaterThanOrEqual(0);

    let beginIndex = strokeIndex - 1;
    while (beginIndex >= 0 && calls[beginIndex][0] !== "beginPath") beginIndex -= 1;
    expect(beginIndex).toBeGreaterThanOrEqual(0);

    const segments: Array<{ x1: number; y1: number; x2: number; y2: number }> = [];
    for (let i = beginIndex + 1; i < strokeIndex; i++) {
      const call = calls[i];
      if (call[0] !== "moveTo") continue;
      const next = calls[i + 1];
      if (!next || next[0] !== "lineTo") continue;
      segments.push({ x1: call[1], y1: call[2], x2: next[1], y2: next[2] });
      i += 1;
    }
    return segments;
  };

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

  it("renders double borders as two parallel strokes", () => {
    const provider: CellProvider = {
      getCell: (row, col) => {
        if (row === 0 && col === 0) {
          return {
            row,
            col,
            value: "x",
            style: {
              borders: {
                bottom: { width: 3, style: "double", color: "rgb(0,0,255)" }
              }
            }
          };
        }
        return null;
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
      const fallback = createRecording2dContext({ canvas: this, calls: [] });
      contexts.set(this, fallback);
      return fallback;
    }) as unknown as typeof HTMLCanvasElement.prototype.getContext;

    const renderer = new CanvasGridRenderer({ provider, rowCount: 2, colCount: 2 });
    renderer.attach({ grid: gridCanvas, content: contentCanvas, selection: selectionCanvas });
    renderer.resize(400, 120, 1);
    renderer.renderImmediately();

    const borderSegments = segmentsForStroke(
      gridCalls,
      (state) => state.strokeStyle === "rgb(0,0,255)" && Math.abs(state.lineWidth - 1) < 1e-6
    );
    expect(borderSegments).toHaveLength(2);

    const uniqueYs = new Set(borderSegments.map((s) => s.y1));
    expect(uniqueYs.size).toBe(2);
  });

  it("draws merged anchor borders around the merged perimeter", () => {
    const merge = { startRow: 0, endRow: 2, startCol: 0, endCol: 2 };
    const provider: CellProvider = {
      getMergedRangesInRange: () => [merge],
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
      const fallback = createRecording2dContext({ canvas: this, calls: [] });
      contexts.set(this, fallback);
      return fallback;
    }) as unknown as typeof HTMLCanvasElement.prototype.getContext;

    const renderer = new CanvasGridRenderer({ provider, rowCount: 3, colCount: 3 });
    renderer.attach({ grid: gridCanvas, content: contentCanvas, selection: selectionCanvas });
    renderer.resize(400, 160, 1);
    renderer.renderImmediately();

    const borderSegments = segmentsForStroke(gridCalls, (state) => state.strokeStyle === "rgb(255,0,0)" && state.lineWidth === 1);

    const crispStrokePos = (pos: number, lineWidth: number): number => {
      const roundedPos = Math.round(pos);
      const roundedWidth = Math.round(lineWidth);
      return roundedWidth % 2 === 1 ? roundedPos + 0.5 : roundedPos;
    };

    const maxX = Math.max(...borderSegments.flatMap((s) => [s.x1, s.x2]));
    const maxY = Math.max(...borderSegments.flatMap((s) => [s.y1, s.y2]));

    // The merged rect is 2 columns (100px each) by 2 rows (21px each), so the perimeter
    // should reach beyond the anchor cell's 1x1 bounds.
    expect(maxX).toBeGreaterThan(150);
    expect(maxY).toBeGreaterThan(30);

    // Ensure we did NOT draw borders on the interior merged gridlines (x=100px, y=21px).
    const interiorX = crispStrokePos(100, 1);
    const interiorY = crispStrokePos(21, 1);
    expect(borderSegments.some((s) => s.x1 === interiorX && s.x2 === interiorX)).toBe(false);
    expect(borderSegments.some((s) => s.y1 === interiorY && s.y2 === interiorY)).toBe(false);
  });

  it("resolves shared-edge border conflicts by preferring the thicker border", () => {
    const provider: CellProvider = {
      getCell: (row, col) => {
        if (row === 0 && col === 0) {
          return {
            row,
            col,
            value: null,
            style: {
              borders: {
                right: { width: 1, style: "solid", color: "rgb(255,0,0)" }
              }
            }
          };
        }
        if (row === 0 && col === 1) {
          return {
            row,
            col,
            value: null,
            style: {
              borders: {
                left: { width: 3, style: "solid", color: "rgb(0,0,255)" }
              }
            }
          };
        }
        return null;
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
      const fallback = createRecording2dContext({ canvas: this, calls: [] });
      contexts.set(this, fallback);
      return fallback;
    }) as unknown as typeof HTMLCanvasElement.prototype.getContext;

    const renderer = new CanvasGridRenderer({ provider, rowCount: 2, colCount: 2 });
    renderer.attach({ grid: gridCanvas, content: contentCanvas, selection: selectionCanvas });
    renderer.resize(400, 120, 1);
    renderer.renderImmediately();

    const borderSegments = segmentsForStroke(
      gridCalls,
      (state) => state.strokeStyle === "rgb(0,0,255)" && Math.abs(state.lineWidth - 3) < 1e-6
    );
    expect(borderSegments).toHaveLength(1);
    // Shared edge between col 0 and col 1 is at x=100px (default col width).
    const crispStrokePos = (pos: number, lineWidth: number): number => {
      const roundedPos = Math.round(pos);
      const roundedWidth = Math.round(lineWidth);
      return roundedWidth % 2 === 1 ? roundedPos + 0.5 : roundedPos;
    };
    const expectedX = crispStrokePos(100, 3);
    expect(borderSegments[0].x1).toBeCloseTo(expectedX, 5);
    expect(borderSegments[0].x2).toBeCloseTo(expectedX, 5);
  });

  it("resolves equal-width conflicts deterministically (prefers right/bottom borders)", () => {
    const provider: CellProvider = {
      getCell: (row, col) => {
        if (row === 0 && col === 0) {
          return {
            row,
            col,
            value: null,
            style: {
              borders: {
                right: { width: 1, style: "solid", color: "rgb(255,0,0)" }
              }
            }
          };
        }
        if (row === 0 && col === 1) {
          return {
            row,
            col,
            value: null,
            style: {
              borders: {
                left: { width: 1, style: "solid", color: "rgb(0,0,255)" }
              }
            }
          };
        }
        return null;
      }
    };

    const gridCalls: Array<[string, ...any[]]> = [];
    const gridCanvas = document.createElement("canvas");
    const contentCanvas = document.createElement("canvas");
    const selectionCanvas = document.createElement("canvas");

    const contexts = new Map<HTMLCanvasElement, CanvasRenderingContext2D>();
    contexts.set(gridCanvas, createRecording2dContext({ canvas: gridCanvas, calls: gridCalls }));
    contexts.set(contentCanvas, createRecording2dContext({ canvas: contentCanvas, calls: [] }));
    contexts.set(selectionCanvas, createRecording2dContext({ canvas: selectionCanvas, calls: [] }));

    HTMLCanvasElement.prototype.getContext = vi.fn(function (this: HTMLCanvasElement) {
      const existing = contexts.get(this);
      if (existing) return existing;
      const fallback = createRecording2dContext({ canvas: this, calls: [] });
      contexts.set(this, fallback);
      return fallback;
    }) as unknown as typeof HTMLCanvasElement.prototype.getContext;

    const renderer = new CanvasGridRenderer({ provider, rowCount: 2, colCount: 2 });
    renderer.attach({ grid: gridCanvas, content: contentCanvas, selection: selectionCanvas });
    renderer.resize(400, 120, 1);
    renderer.renderImmediately();

    // The shared edge should be rendered with the right-hand cell's blue border.
    const borderSegments = segmentsForStroke(gridCalls, (state) => state.strokeStyle === "rgb(0,0,255)" && state.lineWidth === 1);
    expect(borderSegments).toHaveLength(1);
    expect(gridCalls.some((c) => c[0] === "stroke" && c[1]?.strokeStyle === "rgb(255,0,0)")).toBe(false);
  });

  it("resolves equal-width conflicts deterministically on horizontal edges (prefers bottom borders)", () => {
    const provider: CellProvider = {
      getCell: (row, col) => {
        if (row === 0 && col === 0) {
          return {
            row,
            col,
            value: null,
            style: {
              borders: {
                bottom: { width: 1, style: "solid", color: "rgb(255,0,0)" }
              }
            }
          };
        }
        if (row === 1 && col === 0) {
          return {
            row,
            col,
            value: null,
            style: {
              borders: {
                top: { width: 1, style: "solid", color: "rgb(0,0,255)" }
              }
            }
          };
        }
        return null;
      }
    };

    const gridCalls: Array<[string, ...any[]]> = [];
    const gridCanvas = document.createElement("canvas");
    const contentCanvas = document.createElement("canvas");
    const selectionCanvas = document.createElement("canvas");

    const contexts = new Map<HTMLCanvasElement, CanvasRenderingContext2D>();
    contexts.set(gridCanvas, createRecording2dContext({ canvas: gridCanvas, calls: gridCalls }));
    contexts.set(contentCanvas, createRecording2dContext({ canvas: contentCanvas, calls: [] }));
    contexts.set(selectionCanvas, createRecording2dContext({ canvas: selectionCanvas, calls: [] }));

    HTMLCanvasElement.prototype.getContext = vi.fn(function (this: HTMLCanvasElement) {
      const existing = contexts.get(this);
      if (existing) return existing;
      const fallback = createRecording2dContext({ canvas: this, calls: [] });
      contexts.set(this, fallback);
      return fallback;
    }) as unknown as typeof HTMLCanvasElement.prototype.getContext;

    const renderer = new CanvasGridRenderer({ provider, rowCount: 3, colCount: 2 });
    renderer.attach({ grid: gridCanvas, content: contentCanvas, selection: selectionCanvas });
    renderer.resize(400, 160, 1);
    renderer.renderImmediately();

    // Shared edge between row 0 and row 1 is at y=21px (default row height).
    const crispStrokePos = (pos: number, lineWidth: number): number => {
      const roundedPos = Math.round(pos);
      const roundedWidth = Math.round(lineWidth);
      return roundedWidth % 2 === 1 ? roundedPos + 0.5 : roundedPos;
    };
    const expectedY = crispStrokePos(21, 1);

    const borderSegments = segmentsForStroke(gridCalls, (state) => state.strokeStyle === "rgb(0,0,255)" && state.lineWidth === 1);
    expect(borderSegments).toHaveLength(1);
    expect(borderSegments[0].y1).toBeCloseTo(expectedY, 5);
    expect(borderSegments[0].y2).toBeCloseTo(expectedY, 5);
    // Ensure the top cell's red border did not win the tie.
    expect(gridCalls.some((c) => c[0] === "stroke" && c[1]?.strokeStyle === "rgb(255,0,0)" && c[1]?.lineWidth === 1)).toBe(false);
  });

  it("batches border segments by stroke config (single stroke for many segments)", () => {
    const border = { width: 1, style: "solid" as const, color: "rgb(255,0,0)" };
    const provider: CellProvider = {
      getCell: (row, col) => {
        if (row < 2 && col < 2) {
          return {
            row,
            col,
            value: null,
            style: {
              borders: { top: border, right: border, bottom: border, left: border }
            }
          };
        }
        return null;
      }
    };

    const gridCalls: Array<[string, ...any[]]> = [];
    const gridCanvas = document.createElement("canvas");
    const contentCanvas = document.createElement("canvas");
    const selectionCanvas = document.createElement("canvas");

    const contexts = new Map<HTMLCanvasElement, CanvasRenderingContext2D>();
    contexts.set(gridCanvas, createRecording2dContext({ canvas: gridCanvas, calls: gridCalls }));
    contexts.set(contentCanvas, createRecording2dContext({ canvas: contentCanvas, calls: [] }));
    contexts.set(selectionCanvas, createRecording2dContext({ canvas: selectionCanvas, calls: [] }));

    HTMLCanvasElement.prototype.getContext = vi.fn(function (this: HTMLCanvasElement) {
      const existing = contexts.get(this);
      if (existing) return existing;
      const fallback = createRecording2dContext({ canvas: this, calls: [] });
      contexts.set(this, fallback);
      return fallback;
    }) as unknown as typeof HTMLCanvasElement.prototype.getContext;

    const renderer = new CanvasGridRenderer({ provider, rowCount: 3, colCount: 3 });
    renderer.attach({ grid: gridCanvas, content: contentCanvas, selection: selectionCanvas });
    renderer.resize(400, 160, 1);
    renderer.renderImmediately();

    const borderStrokes = gridCalls.filter(
      (call) => call[0] === "stroke" && call[1]?.strokeStyle === "rgb(255,0,0)" && call[1]?.lineWidth === 1
    );
    expect(borderStrokes).toHaveLength(1);
  });

  it("batches double borders across orientations into a single stroke call", () => {
    const border = { width: 3, style: "double" as const, color: "rgb(0,0,255)" };
    const provider: CellProvider = {
      getCell: (row, col) => {
        if (row === 0 && col === 0) {
          return {
            row,
            col,
            value: null,
            style: {
              borders: { top: border, right: border, bottom: border, left: border }
            }
          };
        }
        return null;
      }
    };

    const gridCalls: Array<[string, ...any[]]> = [];
    const gridCanvas = document.createElement("canvas");
    const contentCanvas = document.createElement("canvas");
    const selectionCanvas = document.createElement("canvas");

    const contexts = new Map<HTMLCanvasElement, CanvasRenderingContext2D>();
    contexts.set(gridCanvas, createRecording2dContext({ canvas: gridCanvas, calls: gridCalls }));
    contexts.set(contentCanvas, createRecording2dContext({ canvas: contentCanvas, calls: [] }));
    contexts.set(selectionCanvas, createRecording2dContext({ canvas: selectionCanvas, calls: [] }));

    HTMLCanvasElement.prototype.getContext = vi.fn(function (this: HTMLCanvasElement) {
      const existing = contexts.get(this);
      if (existing) return existing;
      const fallback = createRecording2dContext({ canvas: this, calls: [] });
      contexts.set(this, fallback);
      return fallback;
    }) as unknown as typeof HTMLCanvasElement.prototype.getContext;

    const renderer = new CanvasGridRenderer({ provider, rowCount: 2, colCount: 2 });
    renderer.attach({ grid: gridCanvas, content: contentCanvas, selection: selectionCanvas });
    renderer.resize(400, 160, 1);
    renderer.renderImmediately();

    // width=3 double => effective width 3, rendered as 2 strokes at width=1 each (3/3).
    const borderStrokes = gridCalls.filter(
      (call) => call[0] === "stroke" && call[1]?.strokeStyle === "rgb(0,0,255)" && Math.abs(call[1]?.lineWidth - 1) < 1e-6
    );
    expect(borderStrokes).toHaveLength(1);

    const segments = segmentsForStroke(
      gridCalls,
      (state) => state.strokeStyle === "rgb(0,0,255)" && Math.abs(state.lineWidth - 1) < 1e-6
    );
    // 4 edges × 2 parallel lines per edge.
    expect(segments).toHaveLength(8);
  });
});
