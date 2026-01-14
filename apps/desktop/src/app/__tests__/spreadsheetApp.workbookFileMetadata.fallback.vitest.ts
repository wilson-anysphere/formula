/**
 * @vitest-environment jsdom
 */

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { SpreadsheetApp } from "../spreadsheetApp";

let priorGridMode: string | undefined;

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

describe("SpreadsheetApp workbook file metadata fallback evaluation", () => {
  beforeEach(() => {
    priorGridMode = process.env.DESKTOP_GRID_MODE;
    process.env.DESKTOP_GRID_MODE = "shared";

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

  afterEach(() => {
    if (priorGridMode === undefined) delete process.env.DESKTOP_GRID_MODE;
    else process.env.DESKTOP_GRID_MODE = priorGridMode;
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
  });

  it("recomputes CELL/INFO results from workbook metadata for multi-sheet workbooks (JS fallback)", async () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };
    const app = new SpreadsheetApp(root, status);
    const doc = app.getDocument();

    // Seed formulas that depend on workbook metadata.
    doc.setCellFormula("Sheet1", { row: 0, col: 0 }, '=CELL("filename")');
    doc.setCellFormula("Sheet1", { row: 1, col: 0 }, '=INFO("directory")');
    // Use a numeric wrapper around `CELL("filename")` so the selection summary (Sum/Avg) changes
    // when workbook metadata becomes available.
    doc.setCellFormula("Sheet1", { row: 0, col: 1 }, '=IF(CELL("filename")="",0,1)');

    // Force multi-sheet mode so SpreadsheetApp uses the in-process evaluator path
    // (engine computed-value cache is gated to single-sheet workbooks).
    doc.addSheet({ sheetId: "Sheet2", name: "Sheet2" });

    await app.whenIdle();

    expect(app.getCellComputedValueForSheet("Sheet1", { row: 0, col: 0 })).toBe("");
    expect(app.getCellComputedValueForSheet("Sheet1", { row: 1, col: 0 })).toBe("#N/A");

    const provider = (app as any).sharedProvider;
    expect(provider).toBeTruthy();
    const invalidateSpy = vi.spyOn(provider, "invalidateAll");

    // Prime the selection summary cache, then confirm it is invalidated on workbook metadata changes.
    app.activateCell({ sheetId: "Sheet1", row: 0, col: 1 }, { focus: false, scrollIntoView: false });
    expect(app.getSelectionSummary().sum).toBe(0);

    // Simulate Save As (directory + filename become known).
    await app.setWorkbookFileMetadata("/tmp/", "Book.xlsx");
    expect(invalidateSpy).toHaveBeenCalled();

    expect(app.getCellComputedValueForSheet("Sheet1", { row: 0, col: 0 })).toBe("/tmp/[Book.xlsx]Sheet1");
    expect(app.getCellComputedValueForSheet("Sheet1", { row: 1, col: 0 })).toBe("/tmp/");
    expect(app.getSelectionSummary().sum).toBe(1);

    // Update identity again and ensure trailing separators match Excel semantics.
    await app.setWorkbookFileMetadata("/other", "New.xlsx");
    expect(app.getCellComputedValueForSheet("Sheet1", { row: 0, col: 0 })).toBe("/other/[New.xlsx]Sheet1");
    expect(app.getCellComputedValueForSheet("Sheet1", { row: 1, col: 0 })).toBe("/other/");

    // Simulate returning to an unsaved workbook (metadata cleared).
    await app.setWorkbookFileMetadata(null, null);
    expect(app.getCellComputedValueForSheet("Sheet1", { row: 0, col: 0 })).toBe("");
    expect(app.getCellComputedValueForSheet("Sheet1", { row: 1, col: 0 })).toBe("#N/A");

    app.destroy();
    root.remove();
  });
});
