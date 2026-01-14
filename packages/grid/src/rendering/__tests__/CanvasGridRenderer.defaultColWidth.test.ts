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
    setLineDash: noop,
    measureText: (text: string) =>
      ({
        width: text.length * 6,
        actualBoundingBoxAscent: 8,
        actualBoundingBoxDescent: 2,
      }) as TextMetrics,
  } as unknown as CanvasRenderingContext2D;
}

describe("CanvasGridRenderer default column width", () => {
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

  it("updates the default col width and preserves explicit overrides", () => {
    const provider: CellProvider = { getCell: () => null };
    const renderer = new CanvasGridRenderer({ provider, rowCount: 10, colCount: 10, defaultRowHeight: 10, defaultColWidth: 20 });

    const grid = document.createElement("canvas");
    const content = document.createElement("canvas");
    const selection = document.createElement("canvas");
    renderer.attach({ grid, content, selection });
    renderer.resize(200, 200, 1);

    // Column 2 uses the default before the update.
    expect(renderer.getColWidth(2)).toBe(20);

    // Apply an explicit override that should persist across default changes.
    renderer.setColWidth(1, 50);
    expect(renderer.getColWidth(1)).toBe(50);

    renderer.setDefaultColWidth(40);
    expect(renderer.scroll.cols.defaultSize).toBe(40);
    expect(renderer.getColWidth(2)).toBe(40);
    expect(renderer.getColWidth(1)).toBe(50);
  });

  it("preserves the main visible column start index when changing defaults", () => {
    const provider: CellProvider = {
      getCell: (row, col) => ({ row, col, value: null }),
    };
    const renderer = new CanvasGridRenderer({ provider, rowCount: 100, colCount: 100, defaultRowHeight: 10, defaultColWidth: 20 });

    const grid = document.createElement("canvas");
    const content = document.createElement("canvas");
    const selection = document.createElement("canvas");
    renderer.attach({ grid, content, selection });
    renderer.resize(200, 200, 1);

    // Scroll to the start of col 10 (offset=0) so the visible start index should be stable across
    // the default-size change.
    renderer.setScroll(renderer.scroll.cols.positionOf(10), 0);
    const before = renderer.scroll.getViewportState();
    expect(before.main.cols.start).toBe(10);
    expect(before.main.cols.offset).toBe(0);

    renderer.setDefaultColWidth(40);
    const after = renderer.scroll.getViewportState();
    expect(after.main.cols.start).toBe(10);
    expect(after.main.cols.offset).toBe(0);
  });
});

