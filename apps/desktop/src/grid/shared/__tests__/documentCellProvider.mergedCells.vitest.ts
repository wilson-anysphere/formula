/**
 * @vitest-environment jsdom
 */

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { CanvasGridRenderer, type CellProviderUpdate } from "@formula/grid";

import { DocumentController } from "../../../document/documentController.js";
import { DocumentCellProvider } from "../documentCellProvider.js";

function createMockCanvasContext(options: { onFillText?: (text: string, x: number, y: number) => void }): CanvasRenderingContext2D {
  const noop = () => {};
  const gradient = { addColorStop: noop } as any;

  const base: any = {
    canvas: document.createElement("canvas"),
    measureText: (text: string) =>
      ({
        width: text.length * 6,
        actualBoundingBoxAscent: 8,
        actualBoundingBoxDescent: 2,
      }) as TextMetrics,
    createLinearGradient: () => gradient,
    createPattern: () => null,
    getImageData: () => ({ data: new Uint8ClampedArray(), width: 0, height: 0 }),
    putImageData: noop,
    fillText: (text: string, x: number, y: number) => options.onFillText?.(text, x, y),
  };

  return new Proxy(base, {
    get(target, prop) {
      if (prop in target) return (target as any)[prop];
      return noop;
    },
    set(target, prop, value) {
      (target as any)[prop] = value;
      return true;
    },
  }) as any;
}

describe("DocumentCellProvider merged cells (shared grid)", () => {
  let ctxByCanvas: Map<HTMLCanvasElement, CanvasRenderingContext2D>;

  beforeEach(() => {
    ctxByCanvas = new Map();

    vi.stubGlobal("requestAnimationFrame", (cb: FrameRequestCallback) => {
      cb(0);
      return 0;
    });
    vi.stubGlobal("cancelAnimationFrame", () => {});

    vi.spyOn(HTMLCanvasElement.prototype, "getContext").mockImplementation(function (this: HTMLCanvasElement) {
      return (ctxByCanvas.get(this) ?? createMockCanvasContext({})) as any;
    });
  });

  afterEach(() => {
    vi.restoreAllMocks();
    vi.unstubAllGlobals();
  });

  it("reports merged ranges and the renderer treats merged regions as a single cell", () => {
    const doc = new DocumentController();

    doc.setCellValue("Sheet1", { row: 0, col: 0 }, "A");
    doc.setCellValue("Sheet1", { row: 0, col: 1 }, "B");
    doc.setCellValue("Sheet1", { row: 1, col: 0 }, "C");
    doc.setCellValue("Sheet1", { row: 1, col: 1 }, "D");

    doc.mergeCells("Sheet1", { startRow: 0, endRow: 1, startCol: 0, endCol: 1 });

    // Restore interior values to ensure merged rendering suppresses them even when present.
    doc.setCellValue("Sheet1", { row: 0, col: 1 }, "B");
    doc.setCellValue("Sheet1", { row: 1, col: 0 }, "C");
    doc.setCellValue("Sheet1", { row: 1, col: 1 }, "D");

    const provider = new DocumentCellProvider({
      document: doc,
      getSheetId: () => "Sheet1",
      headerRows: 0,
      headerCols: 0,
      rowCount: 2,
      colCount: 2,
      showFormulas: () => false,
      getComputedValue: () => null,
    });

    expect(provider.getMergedRangeAt?.(0, 0)).toEqual({ startRow: 0, endRow: 2, startCol: 0, endCol: 2 });
    expect(provider.getMergedRangeAt?.(1, 1)).toEqual({ startRow: 0, endRow: 2, startCol: 0, endCol: 2 });
    expect(provider.getMergedRangesInRange?.({ startRow: 1, endRow: 2, startCol: 1, endCol: 2 })).toEqual([
      { startRow: 0, endRow: 2, startCol: 0, endCol: 2 },
    ]);

    const gridCanvas = document.createElement("canvas");
    const contentCanvas = document.createElement("canvas");
    const selectionCanvas = document.createElement("canvas");

    const fillTextCalls: Array<{ text: string; x: number; y: number }> = [];

    ctxByCanvas.set(gridCanvas, createMockCanvasContext({}));
    ctxByCanvas.set(contentCanvas, createMockCanvasContext({ onFillText: (text, x, y) => fillTextCalls.push({ text, x, y }) }));
    ctxByCanvas.set(selectionCanvas, createMockCanvasContext({}));

    const renderer = new CanvasGridRenderer({ provider, rowCount: 2, colCount: 2, defaultRowHeight: 10, defaultColWidth: 10 });
    renderer.attach({ grid: gridCanvas, content: contentCanvas, selection: selectionCanvas });
    renderer.resize(50, 50, 1);
    renderer.renderImmediately();

    // Only the anchor cell should render text.
    expect(fillTextCalls.map((c) => c.text)).toEqual(["A"]);

    // Picking inside the merged region should resolve to the anchor cell.
    expect(renderer.pickCellAt(15, 15)).toEqual({ row: 0, col: 0 });
  });

  it("emits invalidation updates when unmerging (sheet view deltas)", () => {
    const doc = new DocumentController();
    doc.mergeCells("Sheet1", { startRow: 0, endRow: 1, startCol: 0, endCol: 1 });

    const provider = new DocumentCellProvider({
      document: doc,
      getSheetId: () => "Sheet1",
      headerRows: 0,
      headerCols: 0,
      rowCount: 10,
      colCount: 10,
      showFormulas: () => false,
      getComputedValue: () => null,
    });

    const updates: CellProviderUpdate[] = [];
    const unsubscribe = provider.subscribe((update) => updates.push(update));

    doc.unmergeCells("Sheet1", { row: 0, col: 0 });

    unsubscribe();

    // We expect at least one redraw signal for the merged region.
    //
    // The provider may emit either:
    // - a targeted "cells" invalidation for the affected region, or
    // - an "invalidateAll" redraw hint (without flushing caches).
    const invalidateAll = updates.some((u) => u.type === "invalidateAll");
    const cellUpdates = updates.filter((u) => u.type === "cells") as Array<{ type: "cells"; range: any }>;
    const hasTargetedCellsInvalidation = cellUpdates.some(
      (u) => u.range.startRow === 0 && u.range.endRow === 2 && u.range.startCol === 0 && u.range.endCol === 2,
    );
    expect(invalidateAll || hasTargetedCellsInvalidation).toBe(true);
  });
});
