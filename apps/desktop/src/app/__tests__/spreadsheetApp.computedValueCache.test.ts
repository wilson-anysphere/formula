/**
 * @vitest-environment jsdom
 */

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import * as selectionA1 from "../../selection/a1";
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

describe("SpreadsheetApp computed-value cache", () => {
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

  it("avoids cellToA1 calls on numeric cache hits", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    const doc = app.getDocument() as any;
    const sheetId = app.getCurrentSheetId();

    const row = 0;
    const col = 0;
    const key = row * 16_384 + col;

    const spy = vi.spyOn(selectionA1, "cellToA1");
    const sheetIdsSpy = vi.spyOn(doc, "getSheetIds");

    // Seed the numeric computed-value cache directly.
    const byCoord = new Map<number, any>();
    byCoord.set(key, 123);
    (app as any).computedValuesByCoord.set(sheetId, byCoord);
    // Ensure any prior negative cache entry is cleared so lookups see the seeded map.
    (app as any).lastComputedValuesSheetId = null;
    (app as any).lastComputedValuesSheetCache = null;

    spy.mockClear();
    sheetIdsSpy.mockClear();

    for (let i = 0; i < 10_000; i += 1) {
      expect((app as any).getCellComputedValue({ row, col })).toBe(123);
    }

    expect(spy).not.toHaveBeenCalled();
    expect(sheetIdsSpy).not.toHaveBeenCalled();

    // Sanity check: when the numeric cache entry is missing, we should fall back and
    // evaluate the formula.
    byCoord.delete(key);
    doc.setCellFormula(sheetId, { row, col }, "=1+1");
    spy.mockClear();
    expect((app as any).getCellComputedValue({ row, col })).toBe(2);
    // Non-AI formulas do not require cellAddress, so we should not need to allocate an A1 string.
    expect(spy).not.toHaveBeenCalled();

    app.destroy();
    root.remove();
  });

  it("parses A1 addresses in engine computed changes (including $ markers) without allocating A1 strings on lookup", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    const doc = app.getDocument() as any;
    const sheetId = app.getCurrentSheetId();

    // Make the target cell a formula cell so computed values are meaningful.
    doc.setCellFormula(sheetId, { row: 0, col: 0 }, "=1+1");

    // Apply computed changes using only an A1 address (with `$` absolute markers).
    (app as any).applyComputedChanges([{ sheetId, address: "$A$1", value: 123 }]);

    const spy = vi.spyOn(selectionA1, "cellToA1");
    spy.mockClear();

    for (let i = 0; i < 10_000; i += 1) {
      expect(app.getCellComputedValueForSheet(sheetId, { row: 0, col: 0 })).toBe(123);
    }
    expect(spy).not.toHaveBeenCalled();

    // Invalidate using the address-only change and ensure we fall back to local evaluation.
    (app as any).invalidateComputedValues([{ sheetId, address: "$A$1" }]);
    spy.mockClear();
    expect(app.getCellComputedValueForSheet(sheetId, { row: 0, col: 0 })).toBe(2);
    expect(spy).not.toHaveBeenCalled();

    app.destroy();
    root.remove();
  });
});
