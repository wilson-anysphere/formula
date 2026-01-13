// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type { CellProvider, CellRange } from "../../model/CellProvider";
import type { GridPresence } from "../../presence/types";
import { CanvasGridRenderer } from "../CanvasGridRenderer";

type FillStyle = string | CanvasGradient | CanvasPattern;

type RecordedCall = [op: string, ...args: any[]];

type RecordingState = {
  fillStyle: FillStyle;
  strokeStyle: FillStyle;
  lineWidth: number;
  lineDash: number[];
  lineCap: CanvasLineCap;
  font: string;
  globalAlpha: number;
  textAlign: CanvasTextAlign;
  textBaseline: CanvasTextBaseline;
};

function createRecording2dContext(options: { canvas: HTMLCanvasElement; calls: RecordedCall[] }): CanvasRenderingContext2D {
  const noop = () => {};

  const state: RecordingState = {
    fillStyle: "#000",
    strokeStyle: "#000",
    lineWidth: 1,
    lineDash: [],
    lineCap: "butt",
    font: "",
    globalAlpha: 1,
    textAlign: "left",
    textBaseline: "alphabetic"
  };

  const stateStack: RecordingState[] = [];

  const snapshotStrokeState = () => ({
    strokeStyle: state.strokeStyle,
    lineWidth: state.lineWidth,
    lineDash: [...state.lineDash],
    lineCap: state.lineCap,
    globalAlpha: state.globalAlpha
  });

  const snapshotFillState = () => ({
    fillStyle: state.fillStyle,
    globalAlpha: state.globalAlpha
  });

  const snapshotTextState = () => ({
    fillStyle: state.fillStyle,
    font: state.font,
    globalAlpha: state.globalAlpha,
    textAlign: state.textAlign,
    textBaseline: state.textBaseline
  });

  const record = (op: string, ...args: any[]) => options.calls.push([op, ...args]);

  const ctx: any = {
    canvas: options.canvas,
    get fillStyle() {
      return state.fillStyle;
    },
    set fillStyle(value: FillStyle) {
      state.fillStyle = value;
    },
    get strokeStyle() {
      return state.strokeStyle;
    },
    set strokeStyle(value: FillStyle) {
      state.strokeStyle = value;
    },
    get lineWidth() {
      return state.lineWidth;
    },
    set lineWidth(value: number) {
      state.lineWidth = value;
    },
    get lineCap() {
      return state.lineCap;
    },
    set lineCap(value: CanvasLineCap) {
      state.lineCap = value;
    },
    get font() {
      return state.font;
    },
    set font(value: string) {
      state.font = value;
    },
    get globalAlpha() {
      return state.globalAlpha;
    },
    set globalAlpha(value: number) {
      state.globalAlpha = value;
    },
    get textAlign() {
      return state.textAlign;
    },
    set textAlign(value: CanvasTextAlign) {
      state.textAlign = value;
    },
    get textBaseline() {
      return state.textBaseline;
    },
    set textBaseline(value: CanvasTextBaseline) {
      state.textBaseline = value;
    },
    imageSmoothingEnabled: false,
    setTransform: noop,
    clearRect: (...args: any[]) => record("clearRect", ...args),
    fillRect: (...args: any[]) => record("fillRect", snapshotFillState(), ...args),
    strokeRect: (...args: any[]) => record("strokeRect", snapshotStrokeState(), ...args),
    beginPath: (...args: any[]) => record("beginPath", ...args),
    rect: (...args: any[]) => record("rect", ...args),
    clip: (...args: any[]) => record("clip", ...args),
    fill: (...args: any[]) => record("fill", snapshotFillState(), ...args),
    stroke: (...args: any[]) => record("stroke", snapshotStrokeState(), ...args),
    moveTo: (...args: any[]) => record("moveTo", ...args),
    lineTo: (...args: any[]) => record("lineTo", ...args),
    closePath: (...args: any[]) => record("closePath", ...args),
    save: (...args: any[]) => {
      stateStack.push({
        fillStyle: state.fillStyle,
        strokeStyle: state.strokeStyle,
        lineWidth: state.lineWidth,
        lineDash: [...state.lineDash],
        lineCap: state.lineCap,
        font: state.font,
        globalAlpha: state.globalAlpha,
        textAlign: state.textAlign,
        textBaseline: state.textBaseline
      });
      record("save", ...args);
    },
    restore: (...args: any[]) => {
      const restored = stateStack.pop();
      if (restored) {
        state.fillStyle = restored.fillStyle;
        state.strokeStyle = restored.strokeStyle;
        state.lineWidth = restored.lineWidth;
        state.lineDash = [...restored.lineDash];
        state.lineCap = restored.lineCap;
        state.font = restored.font;
        state.globalAlpha = restored.globalAlpha;
        state.textAlign = restored.textAlign;
        state.textBaseline = restored.textBaseline;
      }
      record("restore", ...args);
    },
    drawImage: (...args: any[]) => record("drawImage", ...args),
    translate: (...args: any[]) => record("translate", ...args),
    rotate: (...args: any[]) => record("rotate", ...args),
    fillText: (...args: any[]) => record("fillText", snapshotTextState(), ...args),
    measureText: (text: string) =>
      ({
        width: text.length * 6,
        actualBoundingBoxAscent: 8,
        actualBoundingBoxDescent: 2
      }) as TextMetrics
  };

  // Optional API used by borders and selection previews.
  ctx.setLineDash = (dash: number[]) => {
    state.lineDash = [...dash];
    record("setLineDash", dash);
  };

  return ctx as CanvasRenderingContext2D;
}

function extractCallWindow(calls: RecordedCall[], predicate: (call: RecordedCall) => boolean): RecordedCall[] {
  const index = calls.findIndex(predicate);
  expect(index).toBeGreaterThanOrEqual(0);

  let start = index;
  while (start >= 0 && calls[start]![0] !== "save") start -= 1;
  expect(start).toBeGreaterThanOrEqual(0);

  let end = index;
  while (end < calls.length && calls[end]![0] !== "restore") end += 1;
  expect(end).toBeLessThan(calls.length);

  return calls.slice(start, end + 1);
}

describe("CanvasGridRenderer visual primitive snapshots (recorded draw calls)", () => {
  const originalGetContext = HTMLCanvasElement.prototype.getContext;
  const originalRaf = globalThis.requestAnimationFrame;

  beforeEach(() => {
    // Prevent `requestRender()` from auto-running during setup so we can control when rendering happens.
    vi.stubGlobal("requestAnimationFrame", () => 0);
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

  it("renders a selection rectangle + fill handle with stable geometry/order", () => {
    const provider: CellProvider = { getCell: () => null };

    const gridCalls: RecordedCall[] = [];
    const contentCalls: RecordedCall[] = [];
    const selectionCalls: RecordedCall[] = [];

    const gridCanvas = document.createElement("canvas");
    const contentCanvas = document.createElement("canvas");
    const selectionCanvas = document.createElement("canvas");

    const contexts = new Map<HTMLCanvasElement, CanvasRenderingContext2D>([
      [gridCanvas, createRecording2dContext({ canvas: gridCanvas, calls: gridCalls })],
      [contentCanvas, createRecording2dContext({ canvas: contentCanvas, calls: contentCalls })],
      [selectionCanvas, createRecording2dContext({ canvas: selectionCanvas, calls: selectionCalls })]
    ]);

    HTMLCanvasElement.prototype.getContext = vi.fn(function (this: HTMLCanvasElement) {
      const existing = contexts.get(this);
      if (existing) return existing;
      const fallback = createRecording2dContext({ canvas: this, calls: [] });
      contexts.set(this, fallback);
      return fallback;
    }) as unknown as typeof HTMLCanvasElement.prototype.getContext;

    const renderer = new CanvasGridRenderer({
      provider,
      rowCount: 3,
      colCount: 3,
      defaultRowHeight: 20,
      defaultColWidth: 50,
      theme: {
        selectionFill: "rgba(255,0,0,0.25)",
        selectionBorder: "#ff0000",
        selectionHandle: "#00ff00"
      }
    });
    renderer.attach({ grid: gridCanvas, content: contentCanvas, selection: selectionCanvas });
    renderer.resize(200, 80, 1);
    // Use a multi-cell range so the fill-handle computation is exercised (bottom-right cell != start cell).
    renderer.setSelectionRange({ startRow: 0, endRow: 2, startCol: 0, endCol: 2 });
    renderer.renderImmediately();

    const drawCalls = selectionCalls.filter((c) => c[0] === "fillRect" || c[0] === "strokeRect");
    expect(drawCalls).toMatchInlineSnapshot(`
      [
        [
          "fillRect",
          {
            "fillStyle": "rgba(255,0,0,0.25)",
            "globalAlpha": 1,
          },
          0,
          0,
          100,
          40,
        ],
        [
          "strokeRect",
          {
            "globalAlpha": 1,
            "lineCap": "butt",
            "lineDash": [],
            "lineWidth": 2,
            "strokeStyle": "#ff0000",
          },
          1,
          1,
          98,
          38,
        ],
        [
          "fillRect",
          {
            "fillStyle": "#00ff00",
            "globalAlpha": 1,
          },
          96,
          36,
          8,
          8,
        ],
        [
          "strokeRect",
          {
            "globalAlpha": 1,
            "lineCap": "butt",
            "lineDash": [],
            "lineWidth": 2,
            "strokeStyle": "#ff0000",
          },
          1,
          1,
          48,
          18,
        ],
      ]
    `);
  });

  it("renders comment indicators (resolved/unresolved) as stable top-right triangles", () => {
    const provider: CellProvider = {
      getCell: (row, col) => {
        if (row !== 0) return null;
        if (col === 0) return { row, col, value: null, comment: { resolved: false } };
        if (col === 1) return { row, col, value: null, comment: { resolved: true } };
        return null;
      }
    };

    const gridCalls: RecordedCall[] = [];
    const contentCalls: RecordedCall[] = [];
    const selectionCalls: RecordedCall[] = [];

    const gridCanvas = document.createElement("canvas");
    const contentCanvas = document.createElement("canvas");
    const selectionCanvas = document.createElement("canvas");

    const contexts = new Map<HTMLCanvasElement, CanvasRenderingContext2D>([
      [gridCanvas, createRecording2dContext({ canvas: gridCanvas, calls: gridCalls })],
      [contentCanvas, createRecording2dContext({ canvas: contentCanvas, calls: contentCalls })],
      [selectionCanvas, createRecording2dContext({ canvas: selectionCanvas, calls: selectionCalls })]
    ]);

    HTMLCanvasElement.prototype.getContext = vi.fn(function (this: HTMLCanvasElement) {
      const existing = contexts.get(this);
      if (existing) return existing;
      const fallback = createRecording2dContext({ canvas: this, calls: [] });
      contexts.set(this, fallback);
      return fallback;
    }) as unknown as typeof HTMLCanvasElement.prototype.getContext;

    const renderer = new CanvasGridRenderer({
      provider,
      rowCount: 1,
      colCount: 2,
      defaultRowHeight: 20,
      defaultColWidth: 50,
      theme: {
        commentIndicator: "#ff0000",
        commentIndicatorResolved: "#00ff00"
      }
    });
    renderer.attach({ grid: gridCanvas, content: contentCanvas, selection: selectionCanvas });
    renderer.resize(120, 40, 1);
    renderer.renderImmediately();

    const triangles: RecordedCall[][] = [];
    for (let i = 0; i + 7 < contentCalls.length; i++) {
      const window = contentCalls.slice(i, i + 8);
      const ops = window.map((c) => c[0]);
      const fill = window[6];
      const fillStyle = (fill?.[1] as { fillStyle?: unknown } | undefined)?.fillStyle;
      if (
        ops.join(",") === "save,beginPath,moveTo,lineTo,lineTo,closePath,fill,restore" &&
        (fillStyle === "#ff0000" || fillStyle === "#00ff00")
      ) {
        triangles.push(window);
      }
    }

    expect(triangles).toMatchInlineSnapshot(`
      [
        [
          [
            "save",
          ],
          [
            "beginPath",
          ],
          [
            "moveTo",
            50,
            0,
          ],
          [
            "lineTo",
            44,
            0,
          ],
          [
            "lineTo",
            50,
            6,
          ],
          [
            "closePath",
          ],
          [
            "fill",
            {
              "fillStyle": "#ff0000",
              "globalAlpha": 1,
            },
          ],
          [
            "restore",
          ],
        ],
        [
          [
            "save",
          ],
          [
            "beginPath",
          ],
          [
            "moveTo",
            100,
            0,
          ],
          [
            "lineTo",
            94,
            0,
          ],
          [
            "lineTo",
            100,
            6,
          ],
          [
            "closePath",
          ],
          [
            "fill",
            {
              "fillStyle": "#00ff00",
              "globalAlpha": 1,
            },
          ],
          [
            "restore",
          ],
        ],
      ]
    `);
  });

  it("renders remote presence selection + cursor + name badge with stable draw calls", () => {
    const provider: CellProvider = { getCell: () => null };

    const gridCalls: RecordedCall[] = [];
    const contentCalls: RecordedCall[] = [];
    const selectionCalls: RecordedCall[] = [];

    const gridCanvas = document.createElement("canvas");
    const contentCanvas = document.createElement("canvas");
    const selectionCanvas = document.createElement("canvas");

    const contexts = new Map<HTMLCanvasElement, CanvasRenderingContext2D>([
      [gridCanvas, createRecording2dContext({ canvas: gridCanvas, calls: gridCalls })],
      [contentCanvas, createRecording2dContext({ canvas: contentCanvas, calls: contentCalls })],
      [selectionCanvas, createRecording2dContext({ canvas: selectionCanvas, calls: selectionCalls })]
    ]);

    HTMLCanvasElement.prototype.getContext = vi.fn(function (this: HTMLCanvasElement) {
      const existing = contexts.get(this);
      if (existing) return existing;
      const fallback = createRecording2dContext({ canvas: this, calls: [] });
      contexts.set(this, fallback);
      return fallback;
    }) as unknown as typeof HTMLCanvasElement.prototype.getContext;

    const renderer = new CanvasGridRenderer({
      provider,
      rowCount: 3,
      colCount: 3,
      defaultRowHeight: 20,
      defaultColWidth: 50
    });
    renderer.attach({ grid: gridCanvas, content: contentCanvas, selection: selectionCanvas });
    renderer.resize(240, 120, 1);

    const presences: GridPresence[] = [
      {
        id: "u1",
        name: "Ada",
        color: "#ff00ff",
        cursor: { row: 0, col: 0 },
        selections: [{ startRow: 0, startCol: 0, endRow: 1, endCol: 1 }]
      }
    ];

    renderer.setRemotePresences(presences);
    renderer.renderImmediately();

    const drawCalls = selectionCalls.filter((c) => c[0] === "fillRect" || c[0] === "strokeRect" || c[0] === "fillText");
    expect(drawCalls).toMatchInlineSnapshot(`
      [
        [
          "fillRect",
          {
            "fillStyle": "#ff00ff",
            "globalAlpha": 0.12,
          },
          0,
          0,
          100,
          40,
        ],
        [
          "strokeRect",
          {
            "globalAlpha": 0.9,
            "lineCap": "butt",
            "lineDash": [],
            "lineWidth": 2,
            "strokeStyle": "#ff00ff",
          },
          1,
          1,
          98,
          38,
        ],
        [
          "strokeRect",
          {
            "globalAlpha": 1,
            "lineCap": "butt",
            "lineDash": [],
            "lineWidth": 2,
            "strokeStyle": "#ff00ff",
          },
          1,
          1,
          48,
          18,
        ],
        [
          "fillRect",
          {
            "fillStyle": "#ff00ff",
            "globalAlpha": 1,
          },
          58,
          -18,
          30,
          20,
        ],
        [
          "fillText",
          {
            "fillStyle": "#ffffff",
            "font": "12px system-ui, sans-serif",
            "globalAlpha": 1,
            "textAlign": "left",
            "textBaseline": "top",
          },
          "Ada",
          64,
          -15,
        ],
      ]
    `);
  });

  it("expands merged-cell overflow clip rects into adjacent empty columns (draw-call snapshot)", () => {
    const merge: CellRange = { startRow: 0, endRow: 1, startCol: 0, endCol: 2 };
    const text = "X".repeat(30);

    const provider: CellProvider = {
      getCell: (row, col) => {
        if (row === 0 && col === 0) return { row, col, value: text };
        return null;
      },
      getMergedRangesInRange: (range) => {
        const intersects =
          range.startRow < merge.endRow &&
          range.endRow > merge.startRow &&
          range.startCol < merge.endCol &&
          range.endCol > merge.startCol;
        return intersects ? [merge] : [];
      }
    };

    const gridCalls: RecordedCall[] = [];
    const contentCalls: RecordedCall[] = [];
    const selectionCalls: RecordedCall[] = [];

    const gridCanvas = document.createElement("canvas");
    const contentCanvas = document.createElement("canvas");
    const selectionCanvas = document.createElement("canvas");

    const contexts = new Map<HTMLCanvasElement, CanvasRenderingContext2D>([
      [gridCanvas, createRecording2dContext({ canvas: gridCanvas, calls: gridCalls })],
      [contentCanvas, createRecording2dContext({ canvas: contentCanvas, calls: contentCalls })],
      [selectionCanvas, createRecording2dContext({ canvas: selectionCanvas, calls: selectionCalls })]
    ]);

    HTMLCanvasElement.prototype.getContext = vi.fn(function (this: HTMLCanvasElement) {
      const existing = contexts.get(this);
      if (existing) return existing;
      const fallback = createRecording2dContext({ canvas: this, calls: [] });
      contexts.set(this, fallback);
      return fallback;
    }) as unknown as typeof HTMLCanvasElement.prototype.getContext;

    const renderer = new CanvasGridRenderer({
      provider,
      rowCount: 1,
      colCount: 4,
      defaultRowHeight: 20,
      defaultColWidth: 50
    });
    renderer.attach({ grid: gridCanvas, content: contentCanvas, selection: selectionCanvas });
    // Wider than the sheet to make the expanded overflow clip rect stand out.
    renderer.resize(260, 40, 1);
    renderer.renderImmediately();

    const window = extractCallWindow(contentCalls, (call) => call[0] === "fillText" && call[2] === text);
    // Strip font from the snapshot to keep this focused on clip geometry rather than typography.
    const normalized = window.map((call) => {
      if (call[0] !== "fillText") return call;
      const [op, state, ...rest] = call;
      const { font: _font, ...stateWithoutFont } = state as { font?: string };
      return [op, stateWithoutFont, ...rest] as RecordedCall;
    });

    expect(normalized).toMatchInlineSnapshot(`
      [
        [
          "save",
        ],
        [
          "beginPath",
        ],
        [
          "rect",
          0,
          0,
          200,
          20,
        ],
        [
          "clip",
        ],
        [
          "fillText",
          {
            "fillStyle": "#111111",
            "globalAlpha": 1,
            "textAlign": "left",
            "textBaseline": "alphabetic",
          },
          "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXX",
          4,
          13,
        ],
        [
          "restore",
        ],
      ]
    `);
  });

  it("uses a stable default font string for unstyled cell text (monospace adoption snapshot)", () => {
    const provider: CellProvider = {
      getCell: (row, col) => {
        if (row === 0 && col === 0) return { row, col, value: "A" };
        return null;
      }
    };

    const gridCalls: RecordedCall[] = [];
    const contentCalls: RecordedCall[] = [];
    const selectionCalls: RecordedCall[] = [];

    const gridCanvas = document.createElement("canvas");
    const contentCanvas = document.createElement("canvas");
    const selectionCanvas = document.createElement("canvas");

    const contexts = new Map<HTMLCanvasElement, CanvasRenderingContext2D>([
      [gridCanvas, createRecording2dContext({ canvas: gridCanvas, calls: gridCalls })],
      [contentCanvas, createRecording2dContext({ canvas: contentCanvas, calls: contentCalls })],
      [selectionCanvas, createRecording2dContext({ canvas: selectionCanvas, calls: selectionCalls })]
    ]);

    HTMLCanvasElement.prototype.getContext = vi.fn(function (this: HTMLCanvasElement) {
      const existing = contexts.get(this);
      if (existing) return existing;
      const fallback = createRecording2dContext({ canvas: this, calls: [] });
      contexts.set(this, fallback);
      return fallback;
    }) as unknown as typeof HTMLCanvasElement.prototype.getContext;

    const renderer = new CanvasGridRenderer({
      provider,
      rowCount: 1,
      colCount: 1,
      defaultRowHeight: 20,
      defaultColWidth: 50
    });
    renderer.attach({ grid: gridCanvas, content: contentCanvas, selection: selectionCanvas });
    renderer.resize(80, 40, 1);
    renderer.renderImmediately();

    const fillTextCall = contentCalls.find((call) => call[0] === "fillText");
    expect(fillTextCall).toMatchInlineSnapshot(`
      [
        "fillText",
        {
          "fillStyle": "#111111",
          "font": "normal 400 12px system-ui",
          "globalAlpha": 1,
          "textAlign": "left",
          "textBaseline": "alphabetic",
        },
        "A",
        4,
        13,
      ]
    `);
  });
});
