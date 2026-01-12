// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type { CellProvider } from "../../model/CellProvider";
import { CanvasGridRenderer } from "../CanvasGridRenderer";

function createMock2dContext(): Partial<CanvasRenderingContext2D> {
  return {
    setTransform: vi.fn(),
    // Used by `createCanvasTextMeasurer` + a handful of renderer helpers.
    measureText: (text: string) =>
      ({
        width: text.length * 6,
        actualBoundingBoxAscent: 8,
        actualBoundingBoxDescent: 2
      }) as TextMetrics
  };
}

describe("CanvasGridRenderer auto-fit", () => {
  const originalGetContext = HTMLCanvasElement.prototype.getContext;

  beforeEach(() => {
    // Avoid running full render frames; these tests only validate sizing math.
    vi.stubGlobal("requestAnimationFrame", vi.fn(() => 0));

    const ctxStub = createMock2dContext();
    vi.spyOn(HTMLCanvasElement.prototype, "getContext").mockImplementation(
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      () => ctxStub as any
    );
  });

  afterEach(() => {
    HTMLCanvasElement.prototype.getContext = originalGetContext;
    vi.restoreAllMocks();
    vi.unstubAllGlobals();
  });

  it("autoFitCol scans a small overscan window beyond the visible viewport", () => {
    const provider: CellProvider = {
      getCell: (row, col) => {
        if (col !== 1) return null;
        if (row === 0) return { row, col, value: "A" };
        if (row === 5) return { row, col, value: "LONGTEXT" }; // outside visible range, inside overscan
        if (row === 1) return { row, col, value: "x" };
        return null;
      }
    };

    const renderer = new CanvasGridRenderer({ provider, rowCount: 20, colCount: 3, defaultRowHeight: 10, defaultColWidth: 50 });

    renderer.attach({
      grid: document.createElement("canvas"),
      content: document.createElement("canvas"),
      selection: document.createElement("canvas")
    });
    renderer.setFrozen(1, 0);
    renderer.resize(200, 40, 1);

    // Replace the layout engine with a deterministic stub.
    (renderer as unknown as { textLayoutEngine?: unknown }).textLayoutEngine = {
      measure: (text: string) => ({ width: text.length * 10 })
    };

    const next = renderer.autoFitCol(1);
    // "LONGTEXT" => 8 * 10px + padding (4px*2) + extra (8px) = 96px
    expect(next).toBe(96);
    expect(renderer.getColWidth(1)).toBe(96);
  });

  it("autoFitCol respects maxWidth caps", () => {
    const provider: CellProvider = {
      getCell: (row, col) => {
        if (row === 0 && col === 1) return { row, col, value: "X".repeat(200) };
        return null;
      }
    };

    const renderer = new CanvasGridRenderer({ provider, rowCount: 10, colCount: 3, defaultRowHeight: 10, defaultColWidth: 50 });
    renderer.attach({
      grid: document.createElement("canvas"),
      content: document.createElement("canvas"),
      selection: document.createElement("canvas")
    });
    renderer.setFrozen(1, 0);
    renderer.resize(200, 40, 1);

    (renderer as unknown as { textLayoutEngine?: unknown }).textLayoutEngine = {
      measure: (text: string) => ({ width: text.length * 10 })
    };

    const next = renderer.autoFitCol(1, { maxWidth: 120 });
    expect(next).toBe(120);
    expect(renderer.getColWidth(1)).toBe(120);
  });

  it("autoFitRow uses font metrics even when wrapping is disabled", () => {
    const provider: CellProvider = {
      getCell: (row, col) => {
        if (row === 1 && col === 1) return { row, col, value: "A", style: { fontSize: 30 } };
        return null;
      }
    };

    const renderer = new CanvasGridRenderer({ provider, rowCount: 10, colCount: 3, defaultRowHeight: 21, defaultColWidth: 80 });
    renderer.attach({
      grid: document.createElement("canvas"),
      content: document.createElement("canvas"),
      selection: document.createElement("canvas")
    });
    renderer.setFrozen(1, 0);
    renderer.resize(200, 80, 1);

    const next = renderer.autoFitRow(1);
    // lineHeight = ceil(30 * 1.2) = 36, + paddingY*2 (2*2) = 40.
    expect(next).toBe(40);
    expect(renderer.getRowHeight(1)).toBe(40);
  });

  it("autoFitRow respects maxHeight caps when wrapping expands height", () => {
    const provider: CellProvider = {
      getCell: (row, col) => {
        if (row === 1 && col === 1) return { row, col, value: "wrapped", style: { wrapMode: "word", fontSize: 12 } };
        return null;
      }
    };

    const renderer = new CanvasGridRenderer({ provider, rowCount: 10, colCount: 3, defaultRowHeight: 21, defaultColWidth: 30 });
    renderer.attach({
      grid: document.createElement("canvas"),
      content: document.createElement("canvas"),
      selection: document.createElement("canvas")
    });
    renderer.setFrozen(1, 0);
    renderer.resize(200, 80, 1);

    (renderer as unknown as { textLayoutEngine?: unknown }).textLayoutEngine = {
      layout: () => ({ height: 200 })
    };

    const next = renderer.autoFitRow(1, { maxHeight: 100 });
    expect(next).toBe(100);
    expect(renderer.getRowHeight(1)).toBe(100);
  });
});

