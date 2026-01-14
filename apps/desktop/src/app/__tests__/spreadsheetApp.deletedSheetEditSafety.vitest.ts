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

function createStatus() {
  return {
    activeCell: document.createElement("div"),
    selectionRange: document.createElement("div"),
    activeValue: document.createElement("div"),
  };
}

function seedThreeSheets(app: SpreadsheetApp): void {
  const doc = app.getDocument();
  // Ensure the default sheet is materialized (DocumentController creates sheets lazily).
  doc.getCell("Sheet1", { row: 0, col: 0 });
  doc.addSheet({ sheetId: "Sheet2", name: "Sheet2", insertAfterId: "Sheet1" });
  doc.addSheet({ sheetId: "Sheet3", name: "Sheet3", insertAfterId: "Sheet2" });
}

describe("SpreadsheetApp sheet deletion edit safety", () => {
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

  it("closes the cell editor when the active sheet is deleted", () => {
    const app = new SpreadsheetApp(createRoot(), createStatus());
    try {
      seedThreeSheets(app);
      const doc = app.getDocument();

      app.activateSheet("Sheet2");
      app.openCellEditorAtActiveCell();
      expect(app.isCellEditorOpen()).toBe(true);

      doc.deleteSheet("Sheet2");

      expect(app.isCellEditorOpen()).toBe(false);
      expect(app.getCurrentSheetId()).toBe("Sheet3");
      expect(doc.getSheetIds()).toEqual(["Sheet1", "Sheet3"]);
    } finally {
      app.destroy();
    }
  });

  it("does not recreate a deleted sheet when committing a formula-bar edit targeting it", () => {
    const app = new SpreadsheetApp(createRoot(), createStatus());
    try {
      seedThreeSheets(app);
      const doc = app.getDocument();

      // Simulate: user started editing in Sheet2, then navigated away before commit.
      app.activateSheet("Sheet3");
      (app as any).formulaEditCell = { sheetId: "Sheet2", cell: { row: 0, col: 0 } };

      doc.deleteSheet("Sheet2");
      expect(doc.getSheetIds()).toEqual(["Sheet1", "Sheet3"]);

      (app as any).commitFormulaBar("Hello", { reason: "command", shift: false });

      // The edit should be cancelled rather than lazily recreating Sheet2.
      expect(doc.getSheetIds()).toEqual(["Sheet1", "Sheet3"]);
      expect(app.getCurrentSheetId()).toBe("Sheet3");
    } finally {
      app.destroy();
    }
  });

  it("does not recreate a deleted sheet when cancelling a formula-bar edit targeting it", () => {
    const app = new SpreadsheetApp(createRoot(), createStatus());
    try {
      seedThreeSheets(app);
      const doc = app.getDocument();

      app.activateSheet("Sheet3");
      (app as any).formulaEditCell = { sheetId: "Sheet2", cell: { row: 0, col: 0 } };

      doc.deleteSheet("Sheet2");
      expect(doc.getSheetIds()).toEqual(["Sheet1", "Sheet3"]);

      (app as any).cancelFormulaBar();

      expect(doc.getSheetIds()).toEqual(["Sheet1", "Sheet3"]);
      expect(app.getCurrentSheetId()).toBe("Sheet3");
    } finally {
      app.destroy();
    }
  });
});

