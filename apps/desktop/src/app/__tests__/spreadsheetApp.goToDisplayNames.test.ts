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

describe("SpreadsheetApp goTo", () => {
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

  it("navigates sheet-qualified references using sheet display names without creating phantom sheets", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const sheetNames = new Map<string, string>();
    // SpreadsheetApp consults the sheet name resolver in several subsystems (including charts when
    // present). Provide a baseline id->name mapping for the default sheet so any sheet-qualified
    // references resolve consistently during initialization.
    sheetNames.set("Sheet1", "Sheet1");
    const app = new SpreadsheetApp(root, status, {
      sheetNameResolver: {
        getSheetNameById: (sheetId) => sheetNames.get(sheetId) ?? null,
        getSheetIdByName: (sheetName) => {
          const trimmed = String(sheetName).trim();
          if (!trimmed) return null;
          for (const [id, name] of sheetNames.entries()) {
            if (name.toLowerCase() === trimmed.toLowerCase()) return id;
          }
          return null;
        },
      },
    });

    const doc = app.getDocument();
    // Create a sheet with a stable id.
    doc.setCellValue("Sheet2", { row: 0, col: 0 }, 123);
    // Rename the sheet (display name) without changing its stable id.
    sheetNames.set("Sheet2", "Budget");

    // Move away from A1 so we can assert selection changes.
    app.activateCell({ row: 2, col: 3 });
    expect(app.getActiveCell()).toEqual({ row: 2, col: 3 });

    app.goTo("Budget!A1");

    expect(app.getCurrentSheetId()).toBe("Sheet2");
    expect(app.getActiveCell()).toEqual({ row: 0, col: 0 });
    expect(doc.getSheetIds()).not.toContain("Budget");

    app.destroy();
    root.remove();
  });

  it("does not recreate a deleted sheet when navigating with a stale sheetNameResolver mapping", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const sheetNames = new Map<string, string>();
    sheetNames.set("Sheet1", "Sheet1");

    const app = new SpreadsheetApp(root, status, {
      sheetNameResolver: {
        getSheetNameById: (sheetId) => sheetNames.get(sheetId) ?? null,
        getSheetIdByName: (sheetName) => {
          const trimmed = String(sheetName).trim();
          if (!trimmed) return null;
          for (const [id, name] of sheetNames.entries()) {
            if (name.toLowerCase() === trimmed.toLowerCase()) return id;
          }
          return null;
        },
      },
    });

    const doc = app.getDocument();
    // Ensure Sheet1 exists so deleting Sheet2 is allowed.
    doc.getCell("Sheet1", { row: 0, col: 0 });
    doc.setCellValue("Sheet2", { row: 0, col: 0 }, 123);
    sheetNames.set("Sheet2", "Budget");

    // Delete the sheet but keep the resolver stale (it still maps "Budget" -> "Sheet2").
    doc.deleteSheet("Sheet2");
    expect(doc.getSheetIds()).toEqual(["Sheet1"]);

    expect(app.goTo("Budget!A1")).toBe(false);
    expect(app.getCurrentSheetId()).toBe("Sheet1");
    expect(doc.getSheetIds()).toEqual(["Sheet1"]);
    expect(doc.getSheetMeta("Sheet2")).toBeNull();

    app.destroy();
    root.remove();
  });

  it("resolves named range sheetName display names to stable ids", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const sheetNames = new Map<string, string>();
    const app = new SpreadsheetApp(root, status, {
      sheetNameResolver: {
        getSheetNameById: (sheetId) => sheetNames.get(sheetId) ?? null,
        getSheetIdByName: (sheetName) => {
          const trimmed = String(sheetName).trim().toLowerCase();
          if (!trimmed) return null;
          for (const [id, name] of sheetNames.entries()) {
            if (name.toLowerCase() === trimmed) return id;
          }
          return null;
        },
      },
    });

    const doc = app.getDocument();
    doc.setCellValue("Sheet2", { row: 0, col: 0 }, "NamedRangeTarget");
    sheetNames.set("Sheet2", "Budget");

    // Define a name scoped to the renamed sheet using its *display* name.
    app.getSearchWorkbook().defineName("MyName", {
      sheetName: "Budget",
      range: { startRow: 0, endRow: 0, startCol: 0, endCol: 0 },
    });

    app.activateSheet("Sheet1");
    app.goTo("MyName");

    expect(app.getCurrentSheetId()).toBe("Sheet2");
    expect(app.getActiveCell()).toEqual({ row: 0, col: 0 });
    expect(doc.getSheetIds()).not.toContain("Budget");

    app.destroy();
    root.remove();
  });

  it("resolves table structured ref sheetName display names to stable ids", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const sheetNames = new Map<string, string>();
    const app = new SpreadsheetApp(root, status, {
      sheetNameResolver: {
        getSheetNameById: (sheetId) => sheetNames.get(sheetId) ?? null,
        getSheetIdByName: (sheetName) => {
          const trimmed = String(sheetName).trim().toLowerCase();
          if (!trimmed) return null;
          for (const [id, name] of sheetNames.entries()) {
            if (name.toLowerCase() === trimmed) return id;
          }
          return null;
        },
      },
    });

    const doc = app.getDocument();
    doc.setCellValue("Sheet2", { row: 0, col: 0 }, "TableTarget");
    sheetNames.set("Sheet2", "Budget");

    // Define a table with its sheetName set to the *display* name.
    app.getSearchWorkbook().addTable({
      name: "Table1",
      sheetName: "Budget",
      startRow: 0,
      startCol: 0,
      endRow: 1,
      endCol: 1,
      columns: ["Col1", "Col2"],
    });

    app.activateSheet("Sheet1");
    app.goTo("Table1[#All]");

    expect(app.getCurrentSheetId()).toBe("Sheet2");
    expect(app.getActiveCell()).toEqual({ row: 0, col: 0 });
    expect(app.getSelectionRanges()[0]).toEqual({ startRow: 0, endRow: 1, startCol: 0, endCol: 1 });
    expect(doc.getSheetIds()).not.toContain("Budget");

    app.destroy();
    root.remove();
  });

  it("passes through stable sheet ids returned by named ranges/tables when resolver recognizes the id", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const sheetNames = new Map<string, string>();
    const app = new SpreadsheetApp(root, status, {
      sheetNameResolver: {
        getSheetNameById: (sheetId) => sheetNames.get(sheetId) ?? null,
        // Resolver intentionally does *not* treat ids as names. We rely on pass-through behavior.
        getSheetIdByName: (sheetName) => {
          const trimmed = String(sheetName).trim().toLowerCase();
          if (!trimmed) return null;
          for (const [id, name] of sheetNames.entries()) {
            if (name.toLowerCase() === trimmed) return id;
          }
          return null;
        },
      },
    });

    const doc = app.getDocument();
    doc.setCellValue("Sheet2", { row: 0, col: 0 }, "IdTarget");
    sheetNames.set("Sheet2", "Budget");

    // Name defined with stable id instead of display name.
    app.getSearchWorkbook().defineName("MyIdName", {
      sheetName: "Sheet2",
      range: { startRow: 0, endRow: 0, startCol: 0, endCol: 0 },
    });

    app.activateSheet("Sheet1");
    app.goTo("MyIdName");
    expect(app.getCurrentSheetId()).toBe("Sheet2");

    app.destroy();
    root.remove();
  });

  it("treats unqualified A1 references as relative to the current sheet id (even if display->id lookup fails)", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status, {
      sheetNameResolver: {
        // Resolver can format the sheet name, but cannot map it back to an id (simulating missing metadata).
        getSheetNameById: (sheetId) => (String(sheetId).trim() === "Sheet2" ? "Budget" : null),
        getSheetIdByName: (_sheetName) => null,
      },
    });

    const doc = app.getDocument();
    doc.setCellValue("Sheet2", { row: 0, col: 0 }, "BudgetCell");

    app.activateCell({ sheetId: "Sheet2", row: 5, col: 5 });
    expect(app.getCurrentSheetId()).toBe("Sheet2");

    app.goTo("A1");

    expect(app.getCurrentSheetId()).toBe("Sheet2");
    expect(app.getActiveCell()).toEqual({ row: 0, col: 0 });
    expect(doc.getSheetIds()).not.toContain("Budget");

    app.destroy();
    root.remove();
  });
});
