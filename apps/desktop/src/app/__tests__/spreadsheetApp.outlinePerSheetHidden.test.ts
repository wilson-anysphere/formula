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

  it("does not leak hidden rows across sheets", () => {
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

    // Hide row 1 (0-based row 0) in Sheet1.
    app.hideRows([0]);

    const outline1 = (app as any).getOutlineForSheet(sheet1) as any;
    expect(outline1.rows.entry(1).hidden.user).toBe(true);

    const rowIndexByVisual = (app as any).rowIndexByVisual as number[];
    const rowToVisual = (app as any).rowToVisual as Map<number, number>;
    expect(rowToVisual.has(0)).toBe(false);
    expect(rowIndexByVisual[0]).toBe(1);

    // Switching to Sheet2 should not carry the hidden row with it.
    app.activateSheet(sheet2);

    const outline2 = (app as any).getOutlineForSheet(sheet2) as any;
    expect(outline2.rows.entry(1).hidden.user).toBe(false);

    const rowIndexByVisual2 = (app as any).rowIndexByVisual as number[];
    const rowToVisual2 = (app as any).rowToVisual as Map<number, number>;
    expect(rowToVisual2.has(0)).toBe(true);
    expect(rowIndexByVisual2[0]).toBe(0);

    // Switching back to Sheet1 should preserve its hidden state.
    app.activateSheet(sheet1);

    const rowIndexByVisual3 = (app as any).rowIndexByVisual as number[];
    const rowToVisual3 = (app as any).rowToVisual as Map<number, number>;
    expect(rowToVisual3.has(0)).toBe(false);
    expect(rowIndexByVisual3[0]).toBe(1);

    app.destroy();
    root.remove();
  });

  it("drops outline state when deleting a sheet", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    expect(app.getGridMode()).toBe("legacy");

    const doc = app.getDocument();
    doc.setCellValue("Sheet2", { row: 0, col: 0 }, "X");

    // Ensure we have outline state cached for Sheet2.
    const outline2 = (app as any).getOutlineForSheet("Sheet2") as any;
    outline2.rows.entryMut(1).hidden.user = true;

    const outlinesBySheet = (app as any).outlinesBySheet as Map<string, unknown>;
    expect(outlinesBySheet.has("Sheet2")).toBe(true);

    // Deleting the sheet should remove its outline entry so state doesn't leak or accumulate.
    doc.deleteSheet("Sheet2");

    expect(outlinesBySheet.has("Sheet2")).toBe(false);

    app.destroy();
    root.remove();
  });
});
