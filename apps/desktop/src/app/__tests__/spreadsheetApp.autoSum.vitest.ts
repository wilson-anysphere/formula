/**
 * @vitest-environment jsdom
 */

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { SpreadsheetApp } from "../spreadsheetApp";

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

describe("SpreadsheetApp AutoSum (Alt+=)", () => {
  afterEach(() => {
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
    delete process.env.DESKTOP_GRID_MODE;
  });

  beforeEach(() => {
    document.body.innerHTML = "";
    process.env.DESKTOP_GRID_MODE = "legacy";

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

  it("inserts a SUM formula below a selected vertical range and moves the active cell", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    const sheetId = app.getCurrentSheetId();
    const doc = app.getDocument();

    // SpreadsheetApp seeds some demo/navigation data for non-collab sessions. Clear it so
    // AutoSum tests start from an empty grid.
    doc.clearRange(sheetId, "A1:E5");

    doc.setCellValue(sheetId, "A1", 1);
    doc.setCellValue(sheetId, "A2", 2);
    doc.setCellValue(sheetId, "A3", 3);

    // Select A1:A3 (Excel-style AutoSum should insert into A4).
    app.selectRange({ range: { startRow: 0, endRow: 2, startCol: 0, endCol: 0 } }, { scrollIntoView: false, focus: true });

    const event = new KeyboardEvent("keydown", { altKey: true, code: "Equal", key: "=", cancelable: true });
    root.dispatchEvent(event);
    expect(event.defaultPrevented).toBe(true);

    expect(status.activeCell.textContent).toBe("A4");
    expect(doc.getCell(sheetId, "A4").formula).toBe("=SUM(A1:A3)");
    expect(doc.undoLabel).toBe("AutoSum");

    app.destroy();
    root.remove();
  });

  it("does not resurrect a deleted sheet when AutoSum is invoked while the app holds a stale sheet id", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    try {
      const doc = app.getDocument();

      // Ensure the default sheet exists and create a second sheet we can delete.
      doc.getCell("Sheet1", { row: 0, col: 0 });
      doc.addSheet({ sheetId: "Sheet2", name: "Sheet2", insertAfterId: "Sheet1" });
      expect(doc.getSheetIds()).toEqual(["Sheet1", "Sheet2"]);

      doc.deleteSheet("Sheet2");
      expect(doc.getSheetIds()).toEqual(["Sheet1"]);

      // Simulate a stale active sheet id in UI state.
      (app as any).sheetId = "Sheet2";

      app.autoSum();

      // AutoSum should be a no-op and must not recreate Sheet2.
      expect(doc.getSheetIds()).toEqual(["Sheet1"]);
    } finally {
      app.destroy();
      root.remove();
    }
  });

  it.each([
    { localeId: "de-DE", expectedFn: "SUMME" },
    { localeId: "fr-FR", expectedFn: "SOMME" },
    { localeId: "es-ES", expectedFn: "SUMA" },
  ])("inserts a localized SUM formula when document.lang is $localeId ($expectedFn)", ({ localeId, expectedFn }) => {
    const prevLang = document.documentElement.lang;
    document.documentElement.lang = localeId;

    try {
      const root = createRoot();
      const status = {
        activeCell: document.createElement("div"),
        selectionRange: document.createElement("div"),
        activeValue: document.createElement("div"),
      };
      const app = new SpreadsheetApp(root, status);
      const sheetId = app.getCurrentSheetId();
      const doc = app.getDocument();

      doc.clearRange(sheetId, "A1:E5");

      doc.setCellValue(sheetId, "A1", 1);
      doc.setCellValue(sheetId, "A2", 2);
      doc.setCellValue(sheetId, "A3", 3);

      app.selectRange({ range: { startRow: 0, endRow: 2, startCol: 0, endCol: 0 } }, { scrollIntoView: false, focus: true });
      app.autoSum();

      expect(status.activeCell.textContent).toBe("A4");
      expect(doc.getCell(sheetId, "A4").formula).toBe(`=${expectedFn}(A1:A3)`);

      app.destroy();
      root.remove();
    } finally {
      document.documentElement.lang = prevLang;
    }
  });

  it("localizes AutoSum variants like AVERAGE in de-DE (MITTELWERT)", () => {
    const prevLang = document.documentElement.lang;
    document.documentElement.lang = "de-DE";

    try {
      const root = createRoot();
      const status = {
        activeCell: document.createElement("div"),
        selectionRange: document.createElement("div"),
        activeValue: document.createElement("div"),
      };
      const app = new SpreadsheetApp(root, status);
      const sheetId = app.getCurrentSheetId();
      const doc = app.getDocument();

      doc.clearRange(sheetId, "A1:E5");

      doc.setCellValue(sheetId, "A1", 1);
      doc.setCellValue(sheetId, "A2", 2);
      doc.setCellValue(sheetId, "A3", 3);

      app.selectRange({ range: { startRow: 0, endRow: 2, startCol: 0, endCol: 0 } }, { scrollIntoView: false, focus: true });
      app.autoSumAverage();

      expect(status.activeCell.textContent).toBe("A4");
      expect(doc.getCell(sheetId, "A4").formula).toBe("=MITTELWERT(A1:A3)");

      app.destroy();
      root.remove();
    } finally {
      document.documentElement.lang = prevLang;
    }
  });

  const variants: Array<{
    name: string;
    run: (app: SpreadsheetApp) => void;
    fn: "AVERAGE" | "COUNT" | "MAX" | "MIN";
  }> = [
    { name: "Average", run: (app) => app.autoSumAverage(), fn: "AVERAGE" },
    { name: "Count Numbers", run: (app) => app.autoSumCountNumbers(), fn: "COUNT" },
    { name: "Max", run: (app) => app.autoSumMax(), fn: "MAX" },
    { name: "Min", run: (app) => app.autoSumMin(), fn: "MIN" },
  ];

  it.each(variants)(
    "inserts a $fn formula below a selected vertical range and moves the active cell ($name)",
    ({ run, fn }) => {
      const root = createRoot();
      const status = {
        activeCell: document.createElement("div"),
        selectionRange: document.createElement("div"),
        activeValue: document.createElement("div"),
      };

      const app = new SpreadsheetApp(root, status);
      const sheetId = app.getCurrentSheetId();
      const doc = app.getDocument();

      doc.clearRange(sheetId, "A1:E5");

      doc.setCellValue(sheetId, "A1", 1);
      doc.setCellValue(sheetId, "A2", 2);
      doc.setCellValue(sheetId, "A3", 3);

      app.selectRange({ range: { startRow: 0, endRow: 2, startCol: 0, endCol: 0 } }, { scrollIntoView: false, focus: true });

      run(app);

      expect(status.activeCell.textContent).toBe("A4");
      expect(doc.getCell(sheetId, "A4").formula).toBe(`=${fn}(A1:A3)`);

      app.destroy();
      root.remove();
    },
  );

  it.each(variants)(
    "uses the last selected cell when it's empty for a vertical range ($name)",
    ({ run, fn }) => {
      const root = createRoot();
      const status = {
        activeCell: document.createElement("div"),
        selectionRange: document.createElement("div"),
        activeValue: document.createElement("div"),
      };

      const app = new SpreadsheetApp(root, status);
      const sheetId = app.getCurrentSheetId();
      const doc = app.getDocument();

      doc.clearRange(sheetId, "A1:E5");

      doc.setCellValue(sheetId, "A1", 1);
      doc.setCellValue(sheetId, "A2", 2);
      doc.setCellValue(sheetId, "A3", 3);
      // A4 intentionally empty.
      expect(doc.getCell(sheetId, "A4")).toMatchObject({ value: null, formula: null });

      // Select A1:A4 where the last cell (A4) is empty. Excel-style AutoSum should use A4 and sum A1:A3.
      app.selectRange({ range: { startRow: 0, endRow: 3, startCol: 0, endCol: 0 } }, { scrollIntoView: false, focus: true });

      run(app);

      expect(status.activeCell.textContent).toBe("A4");
      expect(doc.getCell(sheetId, "A4").formula).toBe(`=${fn}(A1:A3)`);
      expect(doc.getCell(sheetId, "A5").formula).toBeNull();

      app.destroy();
      root.remove();
    },
  );

  it.each(variants)(
    "inserts a $fn formula to the right of a selected horizontal range when the last cell is non-empty ($name)",
    ({ run, fn }) => {
      const root = createRoot();
      const status = {
        activeCell: document.createElement("div"),
        selectionRange: document.createElement("div"),
        activeValue: document.createElement("div"),
      };

      const app = new SpreadsheetApp(root, status);
      const sheetId = app.getCurrentSheetId();
      const doc = app.getDocument();

      doc.clearRange(sheetId, "A1:E5");

      doc.setCellValue(sheetId, "A1", 1);
      doc.setCellValue(sheetId, "B1", 2);
      doc.setCellValue(sheetId, "C1", 3);

      // Select A1:C1 (Excel-style AutoSum should insert into D1).
      app.selectRange({ range: { startRow: 0, endRow: 0, startCol: 0, endCol: 2 } }, { scrollIntoView: false, focus: true });

      run(app);

      expect(status.activeCell.textContent).toBe("D1");
      expect(doc.getCell(sheetId, "D1").formula).toBe(`=${fn}(A1:C1)`);

      app.destroy();
      root.remove();
    },
  );

  it.each(variants)(
    "uses the last selected cell when it's empty for a horizontal range ($name)",
    ({ run, fn }) => {
      const root = createRoot();
      const status = {
        activeCell: document.createElement("div"),
        selectionRange: document.createElement("div"),
        activeValue: document.createElement("div"),
      };

      const app = new SpreadsheetApp(root, status);
      const sheetId = app.getCurrentSheetId();
      const doc = app.getDocument();

      doc.clearRange(sheetId, "A1:E5");

      doc.setCellValue(sheetId, "A1", 1);
      doc.setCellValue(sheetId, "B1", 2);
      doc.setCellValue(sheetId, "C1", 3);
      // D1 intentionally empty.

      // Select A1:D1 where the last cell (D1) is empty. Excel-style AutoSum should use D1 and sum A1:C1.
      app.selectRange({ range: { startRow: 0, endRow: 0, startCol: 0, endCol: 3 } }, { scrollIntoView: false, focus: true });

      run(app);

      expect(status.activeCell.textContent).toBe("D1");
      expect(doc.getCell(sheetId, "D1").formula).toBe(`=${fn}(A1:C1)`);
      expect(doc.getCell(sheetId, "E1").formula).toBeNull();

      app.destroy();
      root.remove();
    },
  );

  it.each(variants)(
    "uses the contiguous numeric block above the active cell when selection is not a 1D range ($name)",
    ({ run, fn }) => {
      const root = createRoot();
      const status = {
        activeCell: document.createElement("div"),
        selectionRange: document.createElement("div"),
        activeValue: document.createElement("div"),
      };

      const app = new SpreadsheetApp(root, status);
      const sheetId = app.getCurrentSheetId();
      const doc = app.getDocument();

      doc.clearRange(sheetId, "A1:E5");

      doc.setCellValue(sheetId, "A1", 1);
      doc.setCellValue(sheetId, "A2", 2);
      doc.setCellValue(sheetId, "A3", 3);

      // Active cell A4 with a single-cell selection (not a 1D multi-cell range).
      app.selectRange({ range: { startRow: 3, endRow: 3, startCol: 0, endCol: 0 } }, { scrollIntoView: false, focus: true });

      run(app);

      expect(status.activeCell.textContent).toBe("A4");
      expect(doc.getCell(sheetId, "A4").formula).toBe(`=${fn}(A1:A3)`);

      app.destroy();
      root.remove();
    },
  );

  it.each(variants)(
    "uses the contiguous numeric block to the left of the active cell when there is no numeric block above ($name)",
    ({ run, fn }) => {
      const root = createRoot();
      const status = {
        activeCell: document.createElement("div"),
        selectionRange: document.createElement("div"),
        activeValue: document.createElement("div"),
      };

      const app = new SpreadsheetApp(root, status);
      const sheetId = app.getCurrentSheetId();
      const doc = app.getDocument();

      doc.clearRange(sheetId, "A1:E5");

      doc.setCellValue(sheetId, "A1", 1);
      doc.setCellValue(sheetId, "B1", 2);
      doc.setCellValue(sheetId, "C1", 3);

      // Active cell D1: no numeric cells above, but a numeric block to the left (A1:C1).
      app.selectRange({ range: { startRow: 0, endRow: 0, startCol: 3, endCol: 3 } }, { scrollIntoView: false, focus: true });

      run(app);

      expect(status.activeCell.textContent).toBe("D1");
      expect(doc.getCell(sheetId, "D1").formula).toBe(`=${fn}(A1:C1)`);

      app.destroy();
      root.remove();
    },
  );

  it.each(variants)(
    "treats formula cells with numeric computed values as numeric for range detection ($name)",
    async ({ run, fn }) => {
      const root = createRoot();
      const status = {
        activeCell: document.createElement("div"),
        selectionRange: document.createElement("div"),
        activeValue: document.createElement("div"),
      };

      const app = new SpreadsheetApp(root, status);
      const sheetId = app.getCurrentSheetId();
      const doc = app.getDocument();

      doc.clearRange(sheetId, "A1:E5");

      // A1 is a formula cell whose *computed* value is numeric. AutoSum should use
      // the computed value when deciding whether it's part of the contiguous numeric block.
      doc.setCellInput(sheetId, "A1", "=2");
      doc.setCellValue(sheetId, "A2", 3);
      await app.whenIdle();

      // Active cell A3: contiguous numeric block above should be A1:A2.
      app.selectRange({ range: { startRow: 2, endRow: 2, startCol: 0, endCol: 0 } }, { scrollIntoView: false, focus: true });

      run(app);

      expect(status.activeCell.textContent).toBe("A3");
      expect(doc.getCell(sheetId, "A3").formula).toBe(`=${fn}(A1:A2)`);

      app.destroy();
      root.remove();
    },
  );
});
