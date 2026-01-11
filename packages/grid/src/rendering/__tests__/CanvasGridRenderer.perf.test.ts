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

describe("CanvasGridRenderer perf characteristics", () => {
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

  it("fetches each cell once per frame for combined background+content rendering", () => {
    let cellFetches = 0;
    const provider: CellProvider = {
      getCell: (row, col) => {
        cellFetches += 1;
        return { row, col, value: null };
      }
    };

    const renderer = new CanvasGridRenderer({ provider, rowCount: 1_000, colCount: 1_000 });
    renderer.setPerfStatsEnabled(true);

    const grid = document.createElement("canvas");
    const content = document.createElement("canvas");
    const selection = document.createElement("canvas");
    renderer.attach({ grid, content, selection });

    cellFetches = 0;
    renderer.resize(321, 106, 1);

    const viewport = renderer.scroll.getViewportState();
    const expectedCells =
      (viewport.main.rows.end - viewport.main.rows.start) * (viewport.main.cols.end - viewport.main.cols.start);

    expect(cellFetches).toBe(expectedCells);
    expect(renderer.getPerfStats().cellFetches).toBe(expectedCells);
  });
});

