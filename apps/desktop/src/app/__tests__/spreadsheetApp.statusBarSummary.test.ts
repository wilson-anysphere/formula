/**
 * @vitest-environment jsdom
 */

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { SpreadsheetApp } from "../spreadsheetApp";
import { buildSelection } from "../../selection/selection";

function createInMemoryLocalStorage(): Storage {
  const store = new Map<string, string>();
  return {
    getItem: (key: string) => (store.has(key) ? store.get(key)! : null),
    setItem: (key: string, value: string) => {
      store.set(String(key), String(value));
    },
    removeItem: (key: string) => {
      store.delete(String(key));
    },
    clear: () => {
      store.clear();
    },
    key: (index: number) => Array.from(store.keys())[index] ?? null,
    get length() {
      return store.size;
    },
  } as Storage;
}

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

function createRoot(): HTMLElement {
  const root = document.createElement("div");
  root.tabIndex = 0;
  root.getBoundingClientRect = () =>
    ({
      width: 800,
      height: 600,
      left: 0,
      top: 0,
      right: 800,
      bottom: 600,
      x: 0,
      y: 0,
      toJSON: () => {},
    }) as any;
  document.body.appendChild(root);
  return root;
}

function clearSeededCells(app: SpreadsheetApp): void {
  const doc = app.getDocument();
  const sheetId = app.getCurrentSheetId();
  // SpreadsheetApp seeds demo/navigation data in A1:D5; clear it so tests can reason about a blank sheet.
  doc.clearRange(sheetId, "A1:D5", { label: "Clear seeded cells" });
}

describe("SpreadsheetApp selection summary (status bar)", () => {
  afterEach(() => {
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
  });

  beforeEach(() => {
    document.body.innerHTML = "";

    const storage = createInMemoryLocalStorage();
    Object.defineProperty(globalThis, "localStorage", { configurable: true, value: storage });
    Object.defineProperty(window, "localStorage", { configurable: true, value: storage });
    storage.clear();

    Object.defineProperty(globalThis, "requestAnimationFrame", {
      configurable: true,
      value: (cb: FrameRequestCallback) => {
        cb(0);
        return 0;
      },
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

  it("computes stats for a single numeric cell selection", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    clearSeededCells(app);
    const sheetId = app.getCurrentSheetId();
    const doc = app.getDocument();

    doc.setCellValue(sheetId, { row: 0, col: 0 }, 42);
    app.selectRange({ range: { startRow: 0, endRow: 0, startCol: 0, endCol: 0 } });

    const summary = app.getSelectionSummary();
    expect(summary).toEqual({
      sum: 42,
      average: 42,
      count: 1,
      numericCount: 1,
      countNonEmpty: 1,
    });

    app.destroy();
    root.remove();
  });

  it("computes stats for a single text cell selection", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    clearSeededCells(app);
    const sheetId = app.getCurrentSheetId();
    const doc = app.getDocument();

    doc.setCellValue(sheetId, { row: 0, col: 0 }, "hello");
    app.selectRange({ range: { startRow: 0, endRow: 0, startCol: 0, endCol: 0 } });

    const summary = app.getSelectionSummary();
    expect(summary).toEqual({
      sum: null,
      average: null,
      count: 1,
      numericCount: 0,
      countNonEmpty: 1,
    });

    app.destroy();
    root.remove();
  });

  it("computes stats for a single blank cell selection", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    clearSeededCells(app);

    app.selectRange({ range: { startRow: 1, endRow: 1, startCol: 1, endCol: 1 } });
    const summary = app.getSelectionSummary();
    expect(summary).toEqual({
      sum: null,
      average: null,
      count: 0,
      numericCount: 0,
      countNonEmpty: 0,
    });

    app.destroy();
    root.remove();
  });

  it("aggregates across a rectangular range with numbers, text, and formulas", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    clearSeededCells(app);
    const sheetId = app.getCurrentSheetId();
    const doc = app.getDocument();

    doc.setCellValue(sheetId, "A1", 1);
    doc.setCellValue(sheetId, "A2", 2);
    doc.setCellFormula(sheetId, "A3", "=SUM(A1:A2)");
    doc.setCellValue(sheetId, "B1", "x");
    doc.setCellValue(sheetId, "B2", 10);

    app.selectRange({ range: { startRow: 0, endRow: 2, startCol: 0, endCol: 1 } }); // A1:B3
    const summary = app.getSelectionSummary();

    expect(summary).toEqual({
      sum: 16,
      average: 4,
      count: 5,
      numericCount: 4,
      countNonEmpty: 5,
    });

    app.destroy();
    root.remove();
  });

  it("aggregates across multi-range selection", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    clearSeededCells(app);
    const sheetId = app.getCurrentSheetId();
    const doc = app.getDocument();

    doc.setCellValue(sheetId, "A1", 1);
    doc.setCellValue(sheetId, "B2", 2);
    doc.setCellValue(sheetId, "C3", "hello");

    // A1 plus the rectangular range B2:C3.
    (app as any).selection = buildSelection(
      {
        ranges: [
          { startRow: 0, endRow: 0, startCol: 0, endCol: 0 },
          { startRow: 1, endRow: 2, startCol: 1, endCol: 2 },
        ],
        active: { row: 0, col: 0 },
        anchor: { row: 0, col: 0 },
        activeRangeIndex: 0,
      },
      (app as any).limits,
    );

    const summary = app.getSelectionSummary();
    expect(summary).toEqual({
      sum: 3,
      average: 1.5,
      count: 3,
      numericCount: 2,
      countNonEmpty: 3,
    });

    app.destroy();
    root.remove();
  });

  it("does not double-count overlapping multi-range selection", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    clearSeededCells(app);
    const sheetId = app.getCurrentSheetId();
    const doc = app.getDocument();

    doc.setCellValue(sheetId, "A1", 1);
    doc.setCellValue(sheetId, "B2", 2);
    doc.setCellValue(sheetId, "C3", 3);

    // A1:B2 overlaps B2:C3 at B2. Only three cells contain content; B2 should
    // not be double-counted.
    (app as any).selection = buildSelection(
      {
        ranges: [
          { startRow: 0, endRow: 1, startCol: 0, endCol: 1 }, // A1:B2
          { startRow: 1, endRow: 2, startCol: 1, endCol: 2 }, // B2:C3
        ],
        active: { row: 0, col: 0 },
        anchor: { row: 0, col: 0 },
        activeRangeIndex: 0,
      },
      (app as any).limits,
    );

    const summary = app.getSelectionSummary();
    expect(summary).toEqual({
      sum: 6,
      average: 2,
      count: 3,
      numericCount: 3,
      countNonEmpty: 3,
    });

    app.destroy();
    root.remove();
  });

  it("includes computed values for formula cells", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    clearSeededCells(app);
    const sheetId = app.getCurrentSheetId();
    const doc = app.getDocument();

    doc.setCellValue(sheetId, "A1", 1);
    doc.setCellValue(sheetId, "A2", 2);
    doc.setCellFormula(sheetId, "A3", "=SUM(A1:A2)");

    app.selectRange({ range: { startRow: 0, endRow: 2, startCol: 0, endCol: 0 } }); // A1:A3
    const summary = app.getSelectionSummary();
    expect(summary).toEqual({
      sum: 6,
      average: 2,
      count: 3,
      numericCount: 3,
      countNonEmpty: 3,
    });

    app.destroy();
    root.remove();
  });

  it("ignores error formula results for sum/average but still counts the cell", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    clearSeededCells(app);
    const sheetId = app.getCurrentSheetId();
    const doc = app.getDocument();

    doc.setCellFormula(sheetId, "A1", "=1/0");
    doc.setCellValue(sheetId, "B1", 5);

    app.selectRange({ range: { startRow: 0, endRow: 0, startCol: 0, endCol: 1 } }); // A1:B1
    const summary = app.getSelectionSummary();
    expect(summary).toEqual({
      sum: 5,
      average: 5,
      count: 2,
      numericCount: 1,
      countNonEmpty: 2,
    });

    app.destroy();
    root.remove();
  });

  it("does not count format-only cells", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    clearSeededCells(app);
    const sheetId = app.getCurrentSheetId();
    const doc = app.getDocument();

    doc.setRangeFormat(sheetId, "A1", { foo: "bar" }, { label: "Format" });
    app.selectRange({ range: { startRow: 0, endRow: 0, startCol: 0, endCol: 0 } });

    const summary = app.getSelectionSummary();
    expect(summary).toEqual({
      sum: null,
      average: null,
      count: 0,
      numericCount: 0,
      countNonEmpty: 0,
    });

    app.destroy();
    root.remove();
  });

  it("avoids scanning every coordinate for large selections (sparse iteration)", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    clearSeededCells(app);
    const sheetId = app.getCurrentSheetId();
    const doc = app.getDocument();

    doc.setCellValue(sheetId, "A1", 1);
    doc.setCellValue(sheetId, "B2", 2);

    // Build a large selection that includes A1 and B2 (area > 10k cells) so the
    // implementation should use sparse iteration instead of scanning every coordinate.
    (app as any).selection = buildSelection(
      {
        ranges: [{ startRow: 0, endRow: 999, startCol: 0, endCol: 25 }], // A1:Z1000
        active: { row: 0, col: 0 },
        anchor: { row: 0, col: 0 },
        activeRangeIndex: 0,
      },
      (app as any).limits,
    );

    const getCellSpy = vi.spyOn(doc, "getCell");
    const sparseSpy = vi.spyOn(doc as any, "forEachCellInSheet");

    const summary = app.getSelectionSummary();
    expect(summary).toEqual({
      sum: 3,
      average: 1.5,
      count: 2,
      numericCount: 2,
      countNonEmpty: 2,
    });

    expect(sparseSpy).toHaveBeenCalledTimes(1);
    expect(getCellSpy).not.toHaveBeenCalled();

    app.destroy();
    root.remove();
  });

  it("caches selection summary when selection and sheet content are unchanged", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    clearSeededCells(app);
    const sheetId = app.getCurrentSheetId();
    const doc = app.getDocument();

    doc.setCellValue(sheetId, "A1", 1);
    doc.setCellValue(sheetId, "B1", 2);
    app.selectRange({ range: { startRow: 0, endRow: 0, startCol: 0, endCol: 1 } }); // A1:B1

    const getCellSpy = vi.spyOn(doc, "getCell");
    const first = app.getSelectionSummary();
    expect(getCellSpy).toHaveBeenCalled();

    getCellSpy.mockClear();
    const second = app.getSelectionSummary();
    expect(second).toEqual(first);
    expect(getCellSpy).not.toHaveBeenCalled();

    app.destroy();
    root.remove();
  });

  it("recomputes selection summary when computed values change without a content-version bump", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    clearSeededCells(app);
    const sheetId = app.getCurrentSheetId();
    const doc = app.getDocument();

    doc.setCellFormula(sheetId, "A1", "=1+1");
    app.selectRange({ range: { startRow: 0, endRow: 0, startCol: 0, endCol: 0 } }); // A1

    const before = app.getSelectionSummary();
    expect(before).toEqual({
      sum: 2,
      average: 2,
      count: 1,
      numericCount: 1,
      countNonEmpty: 1,
    });

    // Simulate the engine delivering a computed-value update without a document content change.
    // `applyComputedChanges` bumps an internal version counter that should invalidate the cached
    // selection summary.
    (app as any).applyComputedChanges([{ sheetId, row: 0, col: 0, value: 3 }]);

    const after = app.getSelectionSummary();
    expect(after).toEqual({
      sum: 3,
      average: 3,
      count: 1,
      numericCount: 1,
      countNonEmpty: 1,
    });

    app.destroy();
    root.remove();
  });
});
