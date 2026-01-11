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

describe("CanvasGridRenderer zoom", () => {
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

  it("scales default sizes, preserves overrides, and anchors scroll", () => {
    const provider: CellProvider = {
      getCell: (row, col) => ({ row, col, value: null })
    };

    const renderer = new CanvasGridRenderer({ provider, rowCount: 100, colCount: 100, defaultRowHeight: 10, defaultColWidth: 20 });
    const grid = document.createElement("canvas");
    const content = document.createElement("canvas");
    const selection = document.createElement("canvas");
    renderer.attach({ grid, content, selection });
    renderer.resize(200, 200, 1);

    const before = renderer.getCellRect(0, 0);
    expect(before).not.toBeNull();
    expect(before!.width).toBe(20);
    expect(before!.height).toBe(10);

    renderer.setColWidth(1, 50);
    expect(renderer.getColWidth(1)).toBe(50);

    renderer.setScroll(40, 0);

    renderer.setZoom(2, { anchorX: 100, anchorY: 0 });
    expect(renderer.getZoom()).toBe(2);
    expect(renderer.getColWidth(1)).toBe(100);
    expect(renderer.scroll.getScroll().x).toBe(180);

    const after = renderer.getCellRect(0, 0);
    expect(after).not.toBeNull();
    expect(after!.width).toBe(40);
    expect(after!.height).toBe(20);
  });
});

