/**
 * @vitest-environment jsdom
 */

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { SecondaryGridView } from "../secondaryGridView";
import { DocumentController } from "../../../document/documentController.js";
import type { ImageStore } from "../../../drawings/types";

function createMockCanvasContext(): CanvasRenderingContext2D {
  const noop = () => {};
  const gradient = { addColorStop: noop } as any;
  const context = new Proxy(
    {
      canvas: document.createElement("canvas"),
      measureText: (text: string) => ({ width: text.length * 8 }),
      createLinearGradient: () => gradient,
      createPattern: () => null,
      getImageData: () => ({ data: new Uint8ClampedArray(), width: 0, height: 0 }),
      putImageData: noop,
    },
    {
      get(target, prop) {
        if (prop in target) return (target as any)[prop];
        return noop;
      },
      set(target, prop, value) {
        (target as any)[prop] = value;
        return true;
      },
    },
  );
  return context as any;
}

describe("SecondaryGridView fill large selections", () => {
  afterEach(() => {
    delete (globalThis as any).__formulaSpreadsheetIsEditing;
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
  });

  beforeEach(() => {
    document.body.innerHTML = "";

    Object.defineProperty(globalThis, "requestAnimationFrame", {
      configurable: true,
      value: () => 0,
    });
    Object.defineProperty(globalThis, "cancelAnimationFrame", { configurable: true, value: () => {} });

    Object.defineProperty(HTMLCanvasElement.prototype, "getContext", {
      configurable: true,
      value: () => createMockCanvasContext(),
    });

    (globalThis as any).ResizeObserver = class {
      observe() {}
      disconnect() {}
    };
  });

  const images: ImageStore = { get: () => undefined, set: () => {}, delete: () => {}, clear: () => {} };

  it("does not apply extremely large fill commits", async () => {
    const container = document.createElement("div");
    // Keep viewport 0-sized so the renderer doesn't do any expensive work in jsdom.
    Object.defineProperty(container, "clientWidth", { configurable: true, value: 0 });
    Object.defineProperty(container, "clientHeight", { configurable: true, value: 0 });
    document.body.appendChild(container);

    const doc = new DocumentController();
    const selectionChange = vi.fn();

    const gridView = new SecondaryGridView({
      container,
      document: doc,
      getSheetId: () => "Sheet1",
      rowCount: 10_001,
      colCount: 201,
      showFormulas: () => false,
      getComputedValue: () => null,
      getDrawingObjects: () => [],
      images,
      onSelectionChange: selectionChange,
      onSelectionRangeChange: selectionChange,
    });

    // Ensure a stable pre-fill selection exists.
    gridView.grid.setSelectionRanges(
      [{ startRow: 1, endRow: 2, startCol: 1, endCol: 2 }],
      { activeIndex: 0, activeCell: { row: 1, col: 1 }, scrollIntoView: false },
    );
    selectionChange.mockClear();

    const beginBatch = vi.spyOn(doc, "beginBatch");
    const setCellInput = vi.spyOn(doc, "setCellInput");

    // Grid ranges include a 1-row/1-col header at index 0.
    // Target delta: A2:A200002 => 200,001 cells (exceeds MAX_FILL_CELLS=200,000).
    (gridView as any).onFillCommit({
      sourceRange: { startRow: 1, endRow: 2, startCol: 1, endCol: 2 },
      targetRange: { startRow: 2, endRow: 200_003, startCol: 1, endCol: 2 },
      mode: "formulas",
    });

    // Flush queued microtasks (selection restore).
    await new Promise<void>((resolve) => queueMicrotask(resolve));

    expect(beginBatch).not.toHaveBeenCalled();
    expect(setCellInput).not.toHaveBeenCalled();
    expect(selectionChange).not.toHaveBeenCalled();
    expect((gridView as any).suppressSelectionCallbacks).toBe(false);

    gridView.destroy();
    container.remove();
  });

  it("does not apply fill commits while global spreadsheet editing is active", async () => {
    const container = document.createElement("div");
    // Keep viewport 0-sized so the renderer doesn't do any expensive work in jsdom.
    Object.defineProperty(container, "clientWidth", { configurable: true, value: 0 });
    Object.defineProperty(container, "clientHeight", { configurable: true, value: 0 });
    document.body.appendChild(container);

    const doc = new DocumentController();
    const selectionChange = vi.fn();

    const gridView = new SecondaryGridView({
      container,
      document: doc,
      getSheetId: () => "Sheet1",
      rowCount: 100,
      colCount: 50,
      showFormulas: () => false,
      getComputedValue: () => null,
      getDrawingObjects: () => [],
      images,
      onSelectionChange: selectionChange,
      onSelectionRangeChange: selectionChange,
    });

    // Ensure a stable pre-fill selection exists.
    gridView.grid.setSelectionRanges(
      [{ startRow: 1, endRow: 2, startCol: 1, endCol: 2 }],
      { activeIndex: 0, activeCell: { row: 1, col: 1 }, scrollIntoView: false },
    );
    selectionChange.mockClear();

    const beginBatch = vi.spyOn(doc, "beginBatch");
    const setCellInput = vi.spyOn(doc, "setCellInput");

    // Simulate the desktop shell reporting that an edit session is already active (e.g. primary formula bar).
    (globalThis as any).__formulaSpreadsheetIsEditing = true;

    (gridView as any).onFillCommit({
      sourceRange: { startRow: 1, endRow: 2, startCol: 1, endCol: 2 },
      targetRange: { startRow: 2, endRow: 3, startCol: 1, endCol: 2 },
      mode: "formulas",
    });

    // Flush queued microtasks (selection restore).
    await new Promise<void>((resolve) => queueMicrotask(resolve));

    expect(beginBatch).not.toHaveBeenCalled();
    expect(setCellInput).not.toHaveBeenCalled();
    expect(selectionChange).not.toHaveBeenCalled();
    expect((gridView as any).suppressSelectionCallbacks).toBe(false);

    gridView.destroy();
    container.remove();
  });
});
