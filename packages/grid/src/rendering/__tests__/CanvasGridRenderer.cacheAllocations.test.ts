// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { CanvasGridRenderer } from "../CanvasGridRenderer";
import type { CellProvider } from "../../model/CellProvider";

function createMock2dContext(canvas: HTMLCanvasElement): CanvasRenderingContext2D {
  const noop = () => {};

  return {
    canvas,
    fillStyle: "#000",
    strokeStyle: "#000",
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

describe("CanvasGridRenderer cache allocations", () => {
  const originalGetContext = HTMLCanvasElement.prototype.getContext;

  beforeEach(() => {
    vi.stubGlobal("requestAnimationFrame", (cb: FrameRequestCallback) => {
      cb(0);
      return 0;
    });

    HTMLCanvasElement.prototype.getContext = vi.fn(function (this: HTMLCanvasElement) {
      return createMock2dContext(this);
    }) as unknown as typeof HTMLCanvasElement.prototype.getContext;
  });

  afterEach(() => {
    HTMLCanvasElement.prototype.getContext = originalGetContext;
    vi.unstubAllGlobals();
  });

  it("does not allocate per-row Maps for cell/blocked caches when rendering ~50x20 visible cells", () => {
    const longText = "x".repeat(50);

    const provider: CellProvider = {
      getCell: (row, col) => ({
        row,
        col,
        // Force some overflow probing so `blockedCache` is exercised too.
        value: col === 0 ? longText : null
      })
    };

    const renderer = new CanvasGridRenderer({ provider, rowCount: 1_000, colCount: 1_000 });

    const grid = document.createElement("canvas");
    const content = document.createElement("canvas");
    const selection = document.createElement("canvas");
    renderer.attach({ grid, content, selection });

    renderer.resize(2_000, 1_050, 1);

    const viewport = renderer.scroll.getViewportState();
    const visibleRows = viewport.main.rows.end - viewport.main.rows.start;
    const visibleCols = viewport.main.cols.end - viewport.main.cols.start;
    expect(visibleRows).toBeGreaterThan(40);
    expect(visibleCols).toBeGreaterThan(15);

    expect((renderer as any).__testOnly_rowCacheMapAllocs).toBe(0);
  });
});

