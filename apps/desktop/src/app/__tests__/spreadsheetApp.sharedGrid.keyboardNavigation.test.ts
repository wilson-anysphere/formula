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

describe("SpreadsheetApp shared-grid keyboard navigation", () => {
  afterEach(() => {
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
  });

  beforeEach(() => {
    document.body.innerHTML = "";

    // Node 22 ships an experimental `localStorage` global that errors unless configured via flags.
    // Provide a stable in-memory implementation for unit tests.
    const storage = createInMemoryLocalStorage();
    Object.defineProperty(globalThis, "localStorage", { configurable: true, value: storage });
    Object.defineProperty(window, "localStorage", { configurable: true, value: storage });
    storage.clear();

    // CanvasGridRenderer schedules renders via requestAnimationFrame; ensure it exists in jsdom.
    Object.defineProperty(globalThis, "requestAnimationFrame", {
      configurable: true,
      value: (cb: FrameRequestCallback) => {
        cb(0);
        return 0;
      },
    });
    Object.defineProperty(globalThis, "cancelAnimationFrame", { configurable: true, value: () => {} });

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

  it("supports PageDown/Home/End without building legacy visibility caches (1M rows)", () => {
    const prior = process.env.DESKTOP_GRID_MODE;
    process.env.DESKTOP_GRID_MODE = "shared";
    try {
      const rebuildSpy = vi.spyOn(SpreadsheetApp.prototype as any, "rebuildAxisVisibilityCache");

      const root = createRoot();
      const status = {
        activeCell: document.createElement("div"),
        selectionRange: document.createElement("div"),
        activeValue: document.createElement("div"),
      };

      const limits = { maxRows: 1_000_000, maxCols: 200 };
      const app = new SpreadsheetApp(root, status, { limits });

      expect(app.getGridMode()).toBe("shared");
      expect(rebuildSpy).not.toHaveBeenCalled();

      // The legacy visibility caches should remain empty in shared-grid mode.
      expect((app as any).rowIndexByVisual).toHaveLength(0);
      expect((app as any).colIndexByVisual).toHaveLength(0);
      expect(((app as any).rowToVisual as Map<number, number>).size).toBe(0);
      expect(((app as any).colToVisual as Map<number, number>).size).toBe(0);

      // Make any accidental cache access fail loudly.
      const fail = (name: string) => () => {
        throw new Error(`legacy visibility cache accessed in shared-grid mode: ${name}`);
      };

      const rowToVisual = (app as any).rowToVisual as Map<number, number>;
      const colToVisual = (app as any).colToVisual as Map<number, number>;
      (rowToVisual as any).get = fail("rowToVisual.get");
      (rowToVisual as any).set = fail("rowToVisual.set");
      (rowToVisual as any).clear = fail("rowToVisual.clear");
      (colToVisual as any).get = fail("colToVisual.get");
      (colToVisual as any).set = fail("colToVisual.set");
      (colToVisual as any).clear = fail("colToVisual.clear");

      Object.defineProperty(app as any, "rowIndexByVisual", {
        configurable: true,
        get: fail("rowIndexByVisual"),
        set: fail("rowIndexByVisual"),
      });
      Object.defineProperty(app as any, "colIndexByVisual", {
        configurable: true,
        get: fail("colIndexByVisual"),
        set: fail("colIndexByVisual"),
      });

      // Compute expected page size using shared grid viewport + default sizes.
      const sharedGrid = (app as any).sharedGrid as any;
      const viewport = sharedGrid.renderer.scroll.getViewportState();
      const pageRows = Math.max(
        1,
        Math.floor((viewport.height - viewport.frozenHeight) / sharedGrid.renderer.scroll.rows.defaultSize),
      );
      const pageCols = Math.max(
        1,
        Math.floor((viewport.width - viewport.frozenWidth) / sharedGrid.renderer.scroll.cols.defaultSize),
      );

      expect(app.getActiveCell()).toEqual({ row: 0, col: 0 });

      root.dispatchEvent(new KeyboardEvent("keydown", { key: "PageDown", bubbles: true, cancelable: true }));
      expect(app.getActiveCell()).toEqual({ row: pageRows, col: 0 });
      expect(app.getScroll().y).toBeGreaterThan(0);

      // PageUp should move back up by the same amount.
      root.dispatchEvent(new KeyboardEvent("keydown", { key: "PageUp", bubbles: true, cancelable: true }));
      expect(app.getActiveCell()).toEqual({ row: 0, col: 0 });
      expect(app.getScroll().y).toBe(0);

      // Shift+PageDown should extend the selection while keeping the anchor at A1.
      root.dispatchEvent(new KeyboardEvent("keydown", { key: "PageDown", shiftKey: true, bubbles: true, cancelable: true }));
      expect(app.getActiveCell()).toEqual({ row: pageRows, col: 0 });
      expect((app as any).selection.ranges).toEqual([{ startRow: 0, endRow: pageRows, startCol: 0, endCol: 0 }]);

      // Shift+PageDown again should extend further from the same anchor.
      root.dispatchEvent(new KeyboardEvent("keydown", { key: "PageDown", shiftKey: true, bubbles: true, cancelable: true }));
      const row2 = Math.min(limits.maxRows - 1, pageRows * 2);
      expect(app.getActiveCell()).toEqual({ row: row2, col: 0 });
      expect((app as any).selection.ranges).toEqual([{ startRow: 0, endRow: row2, startCol: 0, endCol: 0 }]);

      // Shift+PageUp should move back up while still extending from the same anchor.
      root.dispatchEvent(new KeyboardEvent("keydown", { key: "PageUp", shiftKey: true, bubbles: true, cancelable: true }));
      expect(app.getActiveCell()).toEqual({ row: pageRows, col: 0 });
      expect((app as any).selection.ranges).toEqual([{ startRow: 0, endRow: pageRows, startCol: 0, endCol: 0 }]);

      // Collapse selection back to a single cell before continuing with other assertions.
      root.dispatchEvent(new KeyboardEvent("keydown", { key: "PageUp", bubbles: true, cancelable: true }));
      expect(app.getActiveCell()).toEqual({ row: 0, col: 0 });
      expect(app.getScroll().y).toBe(0);

      // Alt+PageDown should page horizontally by approx one viewport.
      root.dispatchEvent(new KeyboardEvent("keydown", { key: "PageDown", altKey: true, bubbles: true, cancelable: true }));
      expect(app.getActiveCell()).toEqual({ row: 0, col: pageCols });

      // Shift+End should extend to the last column from the current anchor.
      root.dispatchEvent(new KeyboardEvent("keydown", { key: "End", shiftKey: true, bubbles: true, cancelable: true }));
      expect(app.getActiveCell()).toEqual({ row: 0, col: limits.maxCols - 1 });
      expect((app as any).selection.anchor).toEqual({ row: 0, col: pageCols });
      expect((app as any).selection.ranges).toEqual([
        { startRow: 0, endRow: 0, startCol: pageCols, endCol: limits.maxCols - 1 },
      ]);

      // Reset back to a single cell at column A so subsequent paging assertions are stable.
      root.dispatchEvent(new KeyboardEvent("keydown", { key: "Home", bubbles: true, cancelable: true }));
      expect(app.getActiveCell()).toEqual({ row: 0, col: 0 });

      // Return to the original horizontal PageDown target.
      root.dispatchEvent(new KeyboardEvent("keydown", { key: "PageDown", altKey: true, bubbles: true, cancelable: true }));
      expect(app.getActiveCell()).toEqual({ row: 0, col: pageCols });

      // Shift+Home should extend selection back to the first column (keeping the anchor fixed).
      root.dispatchEvent(new KeyboardEvent("keydown", { key: "Home", shiftKey: true, bubbles: true, cancelable: true }));
      expect(app.getActiveCell()).toEqual({ row: 0, col: 0 });
      expect((app as any).selection.anchor).toEqual({ row: 0, col: pageCols });
      expect((app as any).selection.ranges).toEqual([{ startRow: 0, endRow: 0, startCol: 0, endCol: pageCols }]);

      // Collapse back to a single cell before paging further.
      root.dispatchEvent(new KeyboardEvent("keydown", { key: "PageDown", altKey: true, bubbles: true, cancelable: true }));
      expect(app.getActiveCell()).toEqual({ row: 0, col: pageCols });

      root.dispatchEvent(new KeyboardEvent("keydown", { key: "PageDown", altKey: true, bubbles: true, cancelable: true }));
      const col2 = Math.min(limits.maxCols - 1, pageCols * 2);
      expect(app.getActiveCell()).toEqual({ row: 0, col: col2 });

      // Alt+PageUp should page back left by the same amount.
      root.dispatchEvent(new KeyboardEvent("keydown", { key: "PageUp", altKey: true, bubbles: true, cancelable: true }));
      expect(app.getActiveCell()).toEqual({ row: 0, col: Math.max(0, col2 - pageCols) });

      // Home should return to the first column (row unchanged).
      root.dispatchEvent(new KeyboardEvent("keydown", { key: "Home", bubbles: true, cancelable: true }));
      expect(app.getActiveCell()).toEqual({ row: 0, col: 0 });

      // Return to the original PageDown target so End/Home assertions remain stable.
      root.dispatchEvent(new KeyboardEvent("keydown", { key: "PageDown", bubbles: true, cancelable: true }));
      expect(app.getActiveCell()).toEqual({ row: pageRows, col: 0 });

      root.dispatchEvent(new KeyboardEvent("keydown", { key: "End", bubbles: true, cancelable: true }));
      expect(app.getActiveCell()).toEqual({ row: pageRows, col: limits.maxCols - 1 });

      root.dispatchEvent(new KeyboardEvent("keydown", { key: "Home", bubbles: true, cancelable: true }));
      expect(app.getActiveCell()).toEqual({ row: pageRows, col: 0 });

      // Ctrl/Cmd+Home/End should still work via the generic navigation helper.
      // Validate this doesn't accidentally depend on the legacy visibility caches either.
      root.dispatchEvent(new KeyboardEvent("keydown", { key: "Home", ctrlKey: true, bubbles: true, cancelable: true }));
      expect(app.getActiveCell()).toEqual({ row: 0, col: 0 });

      const used = (app as any).document.getUsedRange((app as any).sheetId) as
        | { startRow: number; endRow: number; startCol: number; endCol: number }
        | null;
      expect(used).not.toBeNull();
      const usedEnd = used ? { row: used.endRow, col: used.endCol } : { row: 0, col: 0 };

      root.dispatchEvent(
        new KeyboardEvent("keydown", { key: "End", ctrlKey: true, shiftKey: true, bubbles: true, cancelable: true }),
      );
      expect(app.getActiveCell()).toEqual(usedEnd);
      expect((app as any).selection.ranges).toEqual([
        { startRow: 0, endRow: usedEnd.row, startCol: 0, endCol: usedEnd.col },
      ]);

      // Collapse back to a single cell before jumping to End without shift.
      root.dispatchEvent(new KeyboardEvent("keydown", { key: "Home", ctrlKey: true, bubbles: true, cancelable: true }));
      expect(app.getActiveCell()).toEqual({ row: 0, col: 0 });

      root.dispatchEvent(new KeyboardEvent("keydown", { key: "End", ctrlKey: true, bubbles: true, cancelable: true }));
      expect(app.getActiveCell()).toEqual(usedEnd);

      app.destroy();
      root.remove();
    } finally {
      if (prior === undefined) delete process.env.DESKTOP_GRID_MODE;
      else process.env.DESKTOP_GRID_MODE = prior;
    }
  });
});
