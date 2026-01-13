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

describe("SpreadsheetApp shared-grid hide/unhide", () => {
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

  it("collapses user-hidden rows/cols, updates navigation, and moves active cell off hidden indices", () => {
    const prior = process.env.DESKTOP_GRID_MODE;
    process.env.DESKTOP_GRID_MODE = "shared";
    try {
      const root = createRoot();
      const status = {
        activeCell: document.createElement("div"),
        selectionRange: document.createElement("div"),
        activeValue: document.createElement("div"),
      };

      const app = new SpreadsheetApp(root, status);
      expect(app.getGridMode()).toBe("shared");

      const sharedGrid = (app as any).sharedGrid;
      expect(sharedGrid).toBeTruthy();
      const renderer = sharedGrid.renderer;

      const headerRows = 1;
      const headerCols = 1;
      const hiddenSize = 2 * renderer.getZoom();

      const gridRow = headerRows + 0; // doc row 0
      const gridCol = headerCols + 0; // doc col 0

      const defaultRowHeight = renderer.getRowHeight(gridRow);
      const defaultColWidth = renderer.getColWidth(gridCol);

      // Sanity: initial selection starts at A1 (doc row/col 0).
      expect((app as any).selection.active).toEqual({ row: 0, col: 0 });

      app.hideRows([0]);

      // User-hidden rows should collapse in the CanvasGridRenderer.
      expect(renderer.getRowHeight(gridRow)).toBeCloseTo(hiddenSize, 6);

      // Navigation provider should report user-hidden rows/cols as hidden in shared-grid mode.
      const provider = (app as any).usedRangeProvider();
      expect(provider.isRowHidden(0)).toBe(true);

      // Hiding the active row should move the active cell to a visible row.
      expect((app as any).selection.active.row).toBe(1);

      app.unhideRows([0]);
      expect(renderer.getRowHeight(gridRow)).toBeCloseTo(defaultRowHeight, 6);
      expect(provider.isRowHidden(0)).toBe(false);

      app.hideCols([0]);
      expect(renderer.getColWidth(gridCol)).toBeCloseTo(hiddenSize, 6);
      expect(provider.isColHidden(0)).toBe(true);
      expect((app as any).selection.active.col).toBe(1);

      app.unhideCols([0]);
      expect(renderer.getColWidth(gridCol)).toBeCloseTo(defaultColWidth, 6);
      expect(provider.isColHidden(0)).toBe(false);

      app.destroy();
      root.remove();
    } finally {
      if (prior === undefined) delete process.env.DESKTOP_GRID_MODE;
      else process.env.DESKTOP_GRID_MODE = prior;
    }
  });
});

