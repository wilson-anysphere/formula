// @vitest-environment jsdom
import { describe, expect, it, vi, beforeEach, afterEach } from "vitest";
import type { CellProvider } from "../../model/CellProvider";
import { CanvasGridRenderer } from "../CanvasGridRenderer";

type FillStyle = string | CanvasGradient | CanvasPattern;

function createMock2dContext(options: { canvas: HTMLCanvasElement; onFillStyle?: (value: FillStyle) => void }): CanvasRenderingContext2D {
  const noop = () => {};
  let fillStyle: FillStyle = "#000";
  let strokeStyle: FillStyle = "#000";

  return {
    canvas: options.canvas,
    get fillStyle() {
      return fillStyle;
    },
    set fillStyle(value: FillStyle) {
      fillStyle = value;
      options.onFillStyle?.(value);
    },
    get strokeStyle() {
      return strokeStyle;
    },
    set strokeStyle(value: FillStyle) {
      strokeStyle = value;
    },
    lineWidth: 1,
    font: "",
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
    fillText: noop,
    measureText: (text: string) =>
      ({
        width: text.length * 6,
        actualBoundingBoxAscent: 8,
        actualBoundingBoxDescent: 2
      }) as TextMetrics
  } as unknown as CanvasRenderingContext2D;
}

describe("CanvasGridRenderer theme usage (render path)", () => {
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

  it("uses theme cell/header/error text colors when no explicit style.color is provided", () => {
    const provider: CellProvider = {
      getCell: (row, col) => {
        if (row === 0 && col === 1) return { row, col, value: "A" }; // header row
        if (row === 1 && col === 1) return { row, col, value: "hello" }; // normal cell
        if (row === 1 && col === 2) return { row, col, value: "#DIV/0!" }; // error cell
        return null;
      }
    };

    const contentFillStyles: string[] = [];

    const gridCanvas = document.createElement("canvas");
    const contentCanvas = document.createElement("canvas");
    const selectionCanvas = document.createElement("canvas");

    const contexts = new Map<HTMLCanvasElement, CanvasRenderingContext2D>();
    contexts.set(gridCanvas, createMock2dContext({ canvas: gridCanvas }));
    contexts.set(
      contentCanvas,
      createMock2dContext({
        canvas: contentCanvas,
        onFillStyle: (value) => {
          if (typeof value === "string") contentFillStyles.push(value);
        }
      })
    );
    contexts.set(selectionCanvas, createMock2dContext({ canvas: selectionCanvas }));

    HTMLCanvasElement.prototype.getContext = vi.fn(function (this: HTMLCanvasElement) {
      const ctx = contexts.get(this);
      return ctx ?? createMock2dContext({ canvas: this });
    }) as unknown as typeof HTMLCanvasElement.prototype.getContext;

    const renderer = new CanvasGridRenderer({
      provider,
      rowCount: 10,
      colCount: 10,
      theme: {
        headerText: "magenta",
        cellText: "cyan",
        errorText: "yellow",
        headerBg: "#111111",
        gridBg: "#000000",
        gridLine: "#222222"
      }
    });

    renderer.attach({ grid: gridCanvas, content: contentCanvas, selection: selectionCanvas });
    renderer.setFrozen(1, 1);
    renderer.resize(400, 200, 1);
    renderer.renderImmediately();

    expect(contentFillStyles).toContain("magenta");
    expect(contentFillStyles).toContain("cyan");
    expect(contentFillStyles).toContain("yellow");
  });

  it("treats explicit headerRows/headerCols as headers even when no rows/cols are frozen", () => {
    const provider: CellProvider = {
      getCell: (row, col) => {
        if (row === 0 && col === 1) return { row, col, value: "A" }; // header row
        if (row === 1 && col === 1) return { row, col, value: "hello" }; // normal cell
        if (row === 1 && col === 2) return { row, col, value: "#DIV/0!" }; // error cell
        return null;
      }
    };

    const contentFillStyles: string[] = [];

    const gridCanvas = document.createElement("canvas");
    const contentCanvas = document.createElement("canvas");
    const selectionCanvas = document.createElement("canvas");

    const contexts = new Map<HTMLCanvasElement, CanvasRenderingContext2D>();
    contexts.set(gridCanvas, createMock2dContext({ canvas: gridCanvas }));
    contexts.set(
      contentCanvas,
      createMock2dContext({
        canvas: contentCanvas,
        onFillStyle: (value) => {
          if (typeof value === "string") contentFillStyles.push(value);
        }
      })
    );
    contexts.set(selectionCanvas, createMock2dContext({ canvas: selectionCanvas }));

    HTMLCanvasElement.prototype.getContext = vi.fn(function (this: HTMLCanvasElement) {
      const ctx = contexts.get(this);
      return ctx ?? createMock2dContext({ canvas: this });
    }) as unknown as typeof HTMLCanvasElement.prototype.getContext;

    const renderer = new CanvasGridRenderer({
      provider,
      rowCount: 10,
      colCount: 10,
      headerRows: 1,
      headerCols: 1,
      theme: {
        headerText: "magenta",
        cellText: "cyan",
        errorText: "yellow",
        headerBg: "#111111",
        gridBg: "#000000",
        gridLine: "#222222"
      }
    });

    renderer.attach({ grid: gridCanvas, content: contentCanvas, selection: selectionCanvas });
    renderer.resize(400, 200, 1);
    renderer.renderImmediately();

    expect(contentFillStyles).toContain("magenta");
    expect(contentFillStyles).toContain("cyan");
    expect(contentFillStyles).toContain("yellow");
  });
});
