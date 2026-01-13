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

describe("SpreadsheetApp outline state", () => {
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

  afterEach(() => {
    if (priorGridMode === undefined) delete process.env.DESKTOP_GRID_MODE;
    else process.env.DESKTOP_GRID_MODE = priorGridMode;
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
  });

  it("does not leak outline-collapsed hidden rows across sheets", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    expect(app.getGridMode()).toBe("legacy");

    // Ensure Sheet2 exists.
    const doc = app.getDocument();
    doc.setCellValue("Sheet2", { row: 0, col: 0 }, "X");

    const sheet1 = app.getCurrentSheetId();
    const sheet2 = "Sheet2";

    // Create a simple outline group on Sheet2 and collapse it (rows 2-4 under summary row 5).
    const outline2 = (app as any).getOutlineForSheet(sheet2) as any;
    outline2.groupRows(2, 4);
    outline2.toggleRowGroup(5);
    expect(outline2.rows.entry(2).hidden.outline).toBe(true);

    // Switching to Sheet2 should apply the collapsed outline hidden rows to the legacy caches.
    app.activateSheet(sheet2);
    const rowIndexByVisual2 = (app as any).rowIndexByVisual as number[];
    const rowToVisual2 = (app as any).rowToVisual as Map<number, number>;
    expect(rowIndexByVisual2[0]).toBe(0);
    expect(rowIndexByVisual2[1]).toBe(4); // rows 2-4 hidden => next visible is row 5 (0-based 4)
    expect(rowToVisual2.has(1)).toBe(false);
    expect(rowToVisual2.has(2)).toBe(false);
    expect(rowToVisual2.has(3)).toBe(false);

    // Switching back to Sheet1 should *not* inherit the outline hidden state from Sheet2.
    app.activateSheet(sheet1);
    const rowIndexByVisual1 = (app as any).rowIndexByVisual as number[];
    const rowToVisual1 = (app as any).rowToVisual as Map<number, number>;
    expect(rowIndexByVisual1[0]).toBe(0);
    expect(rowToVisual1.has(1)).toBe(true); // row 2 visible on Sheet1

    // And switching again to Sheet2 should retain its own collapsed state.
    app.activateSheet(sheet2);
    const rowIndexByVisual2b = (app as any).rowIndexByVisual as number[];
    expect(rowIndexByVisual2b[1]).toBe(4);

    app.destroy();
    root.remove();
  });
});

