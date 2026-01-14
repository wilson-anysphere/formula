/**
 * @vitest-environment jsdom
 */

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { SpreadsheetApp } from "../spreadsheetApp";
import { createSheetNameResolverFromIdToNameMap } from "../../sheet/sheetNameResolver";

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
        // Default all unknown properties to no-op functions so rendering code can execute.
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

describe("SpreadsheetApp fallback evaluator", () => {
  afterEach(() => {
    if (priorGridMode === undefined) delete process.env.DESKTOP_GRID_MODE;
    else process.env.DESKTOP_GRID_MODE = priorGridMode;
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
  });

  beforeEach(() => {
    priorGridMode = process.env.DESKTOP_GRID_MODE;
    process.env.DESKTOP_GRID_MODE = "legacy";
    document.body.innerHTML = "";

    // Node 22 ships an experimental `localStorage` global that errors unless configured via flags.
    // Provide a stable in-memory implementation for unit tests.
    const storage = createInMemoryLocalStorage();
    Object.defineProperty(globalThis, "localStorage", { configurable: true, value: storage });
    Object.defineProperty(window, "localStorage", { configurable: true, value: storage });
    storage.clear();

    // jsdom lacks a real canvas implementation; SpreadsheetApp expects a 2D context.
    Object.defineProperty(HTMLCanvasElement.prototype, "getContext", {
      configurable: true,
      value: () => createMockCanvasContext(),
    });

    // jsdom doesn't ship ResizeObserver by default.
    (globalThis as any).ResizeObserver = class {
      observe() {}
      disconnect() {}
    };
  });

  it("resolves sheet-qualified references case-insensitively", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    const doc = app.getDocument();
    doc.setCellValue("Sheet2", { row: 0, col: 0 }, 123);
    doc.setCellFormula("Sheet1", { row: 0, col: 1 }, "=sheet2!A1+1");

    const computed = app.getCellComputedValueForSheet(app.getCurrentSheetId(), { row: 0, col: 1 });
    expect(computed).toBe(124);
    expect(doc.getSheetIds()).not.toContain("sheet2");

    app.destroy();
    root.remove();
  });

  it("returns #REF! for unknown sheet references without creating a sheet", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    const doc = app.getDocument();
    doc.setCellFormula("Sheet1", { row: 0, col: 1 }, "=MissingSheet!A1");

    const computed = app.getCellComputedValueForSheet(app.getCurrentSheetId(), { row: 0, col: 1 });
    expect(computed).toBe("#REF!");
    expect(doc.getSheetIds()).not.toContain("MissingSheet");

    app.destroy();
    root.remove();
  });

  it("returns #REF! for stale sheetNameResolver mappings after a sheet is deleted (without recreating it)", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const sheetIdToName = new Map<string, string>([
      ["Sheet1", "Sheet1"],
      ["Sheet2", "Sheet2"],
    ]);
    const resolver = createSheetNameResolverFromIdToNameMap(sheetIdToName);

    const app = new SpreadsheetApp(root, status, { sheetNameResolver: resolver });
    const doc = app.getDocument();

    doc.setCellValue("Sheet2", { row: 0, col: 0 }, 123);
    doc.setCellFormula("Sheet1", { row: 0, col: 1 }, "=Sheet2!A1");
    expect(app.getCellComputedValueForSheet("Sheet1", { row: 0, col: 1 })).toBe(123);

    // Delete the referenced sheet but keep the resolver mapping stale.
    doc.deleteSheet("Sheet2");
    expect(doc.getSheetIds()).toEqual(["Sheet1"]);

    // The computed-value evaluator must not resurrect the deleted sheet via lazy reads.
    expect(app.getCellComputedValueForSheet("Sheet1", { row: 0, col: 1 })).toBe("#REF!");
    expect(doc.getSheetIds()).toEqual(["Sheet1"]);

    app.destroy();
    root.remove();
  });

  it("resolves sheet-qualified references using display names (supports rename where id !== name)", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const namesById = new Map<string, string>([
      ["Sheet1", "Sheet1"],
      ["Sheet2", "Sheet2"],
    ]);
    const resolver = {
      getSheetNameById: (id: string) => namesById.get(id) ?? null,
      getSheetIdByName: (name: string) => {
        const needle = name.trim().toLowerCase();
        if (!needle) return null;
        for (const [id, sheetName] of namesById.entries()) {
          if (sheetName.toLowerCase() === needle) return id;
        }
        return null;
      },
    };

    const app = new SpreadsheetApp(root, status, { sheetNameResolver: resolver });
    const doc = app.getDocument();

    doc.setCellValue("Sheet2", { row: 0, col: 0 }, 123);
    doc.setCellFormula("Sheet1", { row: 0, col: 1 }, "=Sheet2!A1+1");
    expect(app.getCellComputedValueForSheet(app.getCurrentSheetId(), { row: 0, col: 1 })).toBe(124);

    // Rename Sheet2 (id stays "Sheet2", display name becomes "Budget").
    namesById.set("Sheet2", "Budget");

    // New formulas should resolve display names back to sheet ids.
    doc.setCellFormula("Sheet1", { row: 0, col: 2 }, "=Budget!A1+1");
    expect(app.getCellComputedValueForSheet(app.getCurrentSheetId(), { row: 0, col: 2 })).toBe(124);

    // Must not create a new sheet for the display name.
    expect(doc.getSheetIds()).not.toContain("Budget");

    app.destroy();
    root.remove();
  });

  it("can compute a cell value for a non-active sheet via getCellComputedValueForSheet", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    const doc = app.getDocument();

    // Active sheet remains Sheet1 (default).
    doc.setCellValue("Sheet1", { row: 0, col: 0 }, 10);
    doc.setCellFormula("Sheet2", { row: 0, col: 1 }, "=Sheet1!A1+1");

    expect(app.getCurrentSheetId()).toBe("Sheet1");
    expect(app.getCellComputedValueForSheet("Sheet2", { row: 0, col: 1 })).toBe(11);

    app.destroy();
    root.remove();
  });

  it("resolves quoted sheet-qualified references (supports spaces and escaped apostrophes)", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const namesById = new Map<string, string>([
      ["Sheet1", "Sheet1"],
      ["Sheet2", "My Sheet"],
    ]);
    const resolver = {
      getSheetNameById: (id: string) => namesById.get(id) ?? null,
      getSheetIdByName: (name: string) => {
        const needle = name.trim().toLowerCase();
        if (!needle) return null;
        for (const [id, sheetName] of namesById.entries()) {
          if (sheetName.toLowerCase() === needle) return id;
        }
        return null;
      },
    };

    const app = new SpreadsheetApp(root, status, { sheetNameResolver: resolver });
    const doc = app.getDocument();

    doc.setCellValue("Sheet2", { row: 0, col: 0 }, 123);
    doc.setCellFormula("Sheet1", { row: 0, col: 1 }, "='My Sheet'!A1+1");
    expect(app.getCellComputedValueForSheet(app.getCurrentSheetId(), { row: 0, col: 1 })).toBe(124);

    // Also support Excel-style escaping of apostrophes inside the quoted sheet name.
    namesById.set("Sheet2", "O'Brien");
    doc.setCellFormula("Sheet1", { row: 0, col: 2 }, "='O''Brien'!A1+1");
    expect(app.getCellComputedValueForSheet(app.getCurrentSheetId(), { row: 0, col: 2 })).toBe(124);

    // Must not create new sheets for display names.
    expect(doc.getSheetIds()).not.toContain("My Sheet");
    expect(doc.getSheetIds()).not.toContain("O'Brien");

    app.destroy();
    root.remove();
  });
});
