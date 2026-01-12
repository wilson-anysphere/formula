// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type { CellProvider, CellRange } from "../../model/CellProvider";
import { CanvasGridRenderer } from "../CanvasGridRenderer";

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
    // Used by some selection/overlay paths; safe to include even if not exercised.
    setLineDash: noop,
    measureText: (text: string) =>
      ({
        width: text.length * 6,
        actualBoundingBoxAscent: 8,
        actualBoundingBoxDescent: 2
      }) as TextMetrics
  } as unknown as CanvasRenderingContext2D;
}

describe("CanvasGridRenderer merged text overflow probing", () => {
  const originalGetContext = HTMLCanvasElement.prototype.getContext;

  beforeEach(() => {
    // Prevent `requestRender()` from auto-running during setup so we can control when rendering happens.
    vi.stubGlobal("requestAnimationFrame", () => 0);

    HTMLCanvasElement.prototype.getContext = vi.fn(function (this: HTMLCanvasElement) {
      return createMock2dContext(this);
    }) as unknown as typeof HTMLCanvasElement.prototype.getContext;
  });

  afterEach(() => {
    HTMLCanvasElement.prototype.getContext = originalGetContext;
    vi.unstubAllGlobals();
  });

  it("does not scan the full merged height when probing horizontal overflow for very tall merges", () => {
    const rowCount = 100_000;
    const tallMerge: CellRange = { startRow: 0, endRow: rowCount, startCol: 0, endCol: 1 };

    let getCellCalls = 0;
    const provider: CellProvider = {
      getCell: (row, col) => {
        getCellCalls += 1;
        if (row === 0 && col === 0) {
          // Long string should overflow beyond the merged cell width and trigger probing into the next column.
          return { row, col, value: "X".repeat(25) };
        }
        return null;
      },
      getMergedRangesInRange: (range) => {
        const intersects =
          range.startRow < tallMerge.endRow &&
          range.endRow > tallMerge.startRow &&
          range.startCol < tallMerge.endCol &&
          range.endCol > tallMerge.startCol;
        return intersects ? [tallMerge] : [];
      }
    };

    const renderer = new CanvasGridRenderer({ provider, rowCount, colCount: 2 });
    renderer.scroll.setScroll(0, renderer.scroll.rows.positionOf(50_000));

    const grid = document.createElement("canvas");
    const content = document.createElement("canvas");
    const selection = document.createElement("canvas");
    renderer.attach({ grid, content, selection });
    renderer.resize(320, 200, 1);

    // Execute the frame explicitly (setup RAF is a noop).
    renderer.renderImmediately();

    const viewport = renderer.scroll.getViewportState();
    const visibleRows = viewport.main.rows.end - viewport.main.rows.start;

    // Without viewport-bounded probing, this would approach `rowCount` because we'd scan every row in the merge
    // while searching for a blocking cell in the next column.
    expect(getCellCalls).toBeLessThan(visibleRows + 50);
  });
});

