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
    }
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
      putImageData: noop
    },
    {
      get(target, prop) {
        if (prop in target) return (target as any)[prop];
        return noop;
      },
      set(target, prop, value) {
        (target as any)[prop] = value;
        return true;
      }
    }
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
      toJSON: () => {}
    }) as any;
  document.body.appendChild(root);
  return root;
}

describe("SpreadsheetApp shared-grid axis size sync", () => {
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
      }
    });
    Object.defineProperty(globalThis, "cancelAnimationFrame", { configurable: true, value: () => {} });

    Object.defineProperty(HTMLCanvasElement.prototype, "getContext", {
      configurable: true,
      value: () => createMockCanvasContext()
    });

    (globalThis as any).ResizeObserver = class {
      observe() {}
      disconnect() {}
    };
  });

  it("uses CanvasGridRenderer.applyAxisSizeOverrides (no per-index setters)", () => {
    const prior = process.env.DESKTOP_GRID_MODE;
    process.env.DESKTOP_GRID_MODE = "shared";
    try {
      const root = createRoot();
      const status = {
        activeCell: document.createElement("div"),
        selectionRange: document.createElement("div"),
        activeValue: document.createElement("div")
      };

      const app = new SpreadsheetApp(root, status);
      expect(app.getGridMode()).toBe("shared");

      const sharedGrid = (app as any).sharedGrid;
      expect(sharedGrid).toBeTruthy();
      const renderer = sharedGrid.renderer;

      // Seed a sheet view with many overrides, but bypass DocumentController deltas so we can
      // call the sync method deterministically (without extra intermediate sync calls).
      const sheetId = app.getCurrentSheetId();
      const doc = app.getDocument() as any;
      const view = doc.getSheetView(sheetId) ?? {};

      const rowHeights: Record<string, number> = {};
      for (let i = 0; i < 1_000; i += 1) rowHeights[String(i)] = 30 + (i % 3);
      // Include some invalid/out-of-range indices; the sync should ignore them (and not throw).
      rowHeights["999999"] = 42;

      const colWidths: Record<string, number> = {};
      for (let i = 0; i < 150; i += 1) colWidths[String(i)] = 120 + (i % 5);
      colWidths["999999"] = 321;

      (view as any).rowHeights = rowHeights;
      (view as any).colWidths = colWidths;
      doc.model.setSheetView(sheetId, view);

      const batchSpy = vi.spyOn(renderer, "applyAxisSizeOverrides");
      const setRowSpy = vi.spyOn(renderer, "setRowHeight");
      const setColSpy = vi.spyOn(renderer, "setColWidth");
      const resetRowSpy = vi.spyOn(renderer, "resetRowHeight");
      const resetColSpy = vi.spyOn(renderer, "resetColWidth");

      (app as any).syncSharedGridAxisSizesFromDocument();

      expect(batchSpy).toHaveBeenCalledTimes(1);
      expect(setRowSpy).not.toHaveBeenCalled();
      expect(setColSpy).not.toHaveBeenCalled();
      expect(resetRowSpy).not.toHaveBeenCalled();
      expect(resetColSpy).not.toHaveBeenCalled();

      app.destroy();
      root.remove();
    } finally {
      if (prior === undefined) delete process.env.DESKTOP_GRID_MODE;
      else process.env.DESKTOP_GRID_MODE = prior;
    }
  });

  it("does not rebuild/apply sheet-view axis overrides on zoom changes", () => {
    const prior = process.env.DESKTOP_GRID_MODE;
    process.env.DESKTOP_GRID_MODE = "shared";
    try {
      const root = createRoot();
      const status = {
        activeCell: document.createElement("div"),
        selectionRange: document.createElement("div"),
        activeValue: document.createElement("div")
      };

      const app = new SpreadsheetApp(root, status);
      expect(app.getGridMode()).toBe("shared");

      const sharedGrid = (app as any).sharedGrid;
      expect(sharedGrid).toBeTruthy();
      const renderer = sharedGrid.renderer;

      const sheetId = app.getCurrentSheetId();
      const doc = app.getDocument() as any;
      const view = doc.getSheetView(sheetId) ?? {};

      (view as any).rowHeights = { "0": 30, "1": 31 };
      (view as any).colWidths = { "0": 120, "1": 121 };
      doc.model.setSheetView(sheetId, view);

      // Apply once at the current zoom so the renderer has persisted overrides.
      (app as any).syncSharedGridAxisSizesFromDocument();

      const batchSpy = vi.spyOn(renderer, "applyAxisSizeOverrides");
      batchSpy.mockClear();

      // Zooming should scale the existing overrides; it should not rebuild/apply them via
      // SpreadsheetApp's document sync path (which can be expensive with many overrides).
      app.setZoom(2);
      expect(batchSpy).not.toHaveBeenCalled();

      app.destroy();
      root.remove();
    } finally {
      if (prior === undefined) delete process.env.DESKTOP_GRID_MODE;
      else process.env.DESKTOP_GRID_MODE = prior;
    }
  });

  it("does not resync axis overrides back into the same shared grid after a local resize edit", () => {
    const prior = process.env.DESKTOP_GRID_MODE;
    process.env.DESKTOP_GRID_MODE = "shared";
    try {
      const root = createRoot();
      const status = {
        activeCell: document.createElement("div"),
        selectionRange: document.createElement("div"),
        activeValue: document.createElement("div")
      };

      const app = new SpreadsheetApp(root, status);
      expect(app.getGridMode()).toBe("shared");

      const sharedGrid = (app as any).sharedGrid;
      expect(sharedGrid).toBeTruthy();
      const renderer = sharedGrid.renderer;

      const batchSpy = vi.spyOn(renderer, "applyAxisSizeOverrides");
      batchSpy.mockClear();

      // Simulate the end-of-drag resize callback. The renderer would already be updated during the drag;
      // this should only mutate the document (source-tagged) and should not trigger a full re-sync of
      // all axis overrides back into the same renderer instance.
      (app as any).onSharedGridAxisSizeChange({
        kind: "col",
        index: 2,
        size: 130,
        previousSize: 120,
        defaultSize: 100,
        zoom: 1,
        source: "resize"
      });

      expect(batchSpy).not.toHaveBeenCalled();

      app.destroy();
      root.remove();
    } finally {
      if (prior === undefined) delete process.env.DESKTOP_GRID_MODE;
      else process.env.DESKTOP_GRID_MODE = prior;
    }
  });

  it("does not persist axis resize edits while the formula bar is actively editing", () => {
    const prior = process.env.DESKTOP_GRID_MODE;
    process.env.DESKTOP_GRID_MODE = "shared";
    try {
      const root = createRoot();
      const status = {
        activeCell: document.createElement("div"),
        selectionRange: document.createElement("div"),
        activeValue: document.createElement("div"),
      };

      const formulaBar = document.createElement("div");
      document.body.appendChild(formulaBar);

      const app = new SpreadsheetApp(root, status, { formulaBar });
      expect(app.getGridMode()).toBe("shared");

      const input = formulaBar.querySelector<HTMLTextAreaElement>('[data-testid="formula-input"]');
      expect(input).not.toBeNull();
      input!.focus();
      input!.value = "=SUM(";
      input!.dispatchEvent(new Event("input", { bubbles: true }));

      // Mimic selecting ranges while the formula bar is still editing.
      root.focus();
      expect(app.isEditing()).toBe(true);

      const sharedGrid = (app as any).sharedGrid;
      expect(sharedGrid).toBeTruthy();
      const renderer = sharedGrid.renderer;

      // Resize the first data column (grid index 1 => doc col 0) in the renderer to mimic an interactive drag.
      const index = 1;
      const prevSize = renderer.getColWidth(index);
      const nextSize = prevSize + 25;
      renderer.setColWidth(index, nextSize);

      const doc = app.getDocument() as any;
      const sheetId = app.getCurrentSheetId();

      // The resize callback should no-op while editing and restore the renderer size.
      (app as any).onSharedGridAxisSizeChange({
        kind: "col",
        index,
        size: nextSize,
        previousSize: prevSize,
        defaultSize: renderer.scroll.cols.defaultSize,
        zoom: renderer.getZoom(),
        source: "resize",
      });

      expect(renderer.getColWidth(index)).toBeCloseTo(prevSize, 6);

      const view = doc.getSheetView(sheetId) ?? {};
      expect((view as any).colWidths?.["0"]).toBeUndefined();

      // Focus should be restored to the formula bar so the user can continue typing.
      expect(document.activeElement).toBe(input);

      app.destroy();
      root.remove();
      formulaBar.remove();
    } finally {
      if (prior === undefined) delete process.env.DESKTOP_GRID_MODE;
      else process.env.DESKTOP_GRID_MODE = prior;
    }
  });
});
