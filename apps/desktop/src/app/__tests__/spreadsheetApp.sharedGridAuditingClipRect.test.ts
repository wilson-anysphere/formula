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

function createMockCanvasContext(canvas: HTMLCanvasElement): CanvasRenderingContext2D {
  const noop = () => {};
  const gradient = { addColorStop: noop } as any;
  const context = new Proxy(
    {
      canvas,
      rect: vi.fn(),
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

function createRoot(rect: { width: number; height: number } = { width: 800, height: 600 }): HTMLElement {
  const root = document.createElement("div");
  root.tabIndex = 0;
  root.getBoundingClientRect = () =>
    ({
      width: rect.width,
      height: rect.height,
      left: 0,
      top: 0,
      right: rect.width,
      bottom: rect.height,
      x: 0,
      y: 0,
      toJSON: () => {},
    }) as any;
  document.body.appendChild(root);
  return root;
}

describe("SpreadsheetApp auditing overlays clip rect (shared grid)", () => {
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

    // CanvasGridRenderer schedules renders via requestAnimationFrame; ensure it exists in jsdom.
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
      value(this: HTMLCanvasElement) {
        return createMockCanvasContext(this);
      },
    });

    (globalThis as any).ResizeObserver = class {
      observe() {}
      disconnect() {}
    };
  });

  it("clips auditing overlays to the full cell area (header only)", () => {
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
      const viewport = sharedGrid.renderer.scroll.getViewportState();
      const headerWidth = sharedGrid.renderer.getColWidth(0);
      const headerHeight = sharedGrid.renderer.getRowHeight(0);

      const ctx = (app as any).auditingCtx;
      ctx.rect.mockClear();
      (app as any).auditingMode = "precedents";
      (app as any).renderAuditing();

      expect(ctx.rect).toHaveBeenCalledTimes(1);
      const [x, y, width, height] = ctx.rect.mock.calls[0] as number[];
      expect(x).toBe(headerWidth);
      expect(y).toBe(headerHeight);
      expect(x + width).toBe(viewport.width);
      expect(y + height).toBe(viewport.height);

      app.destroy();
      root.remove();
    } finally {
      if (prior === undefined) delete process.env.DESKTOP_GRID_MODE;
      else process.env.DESKTOP_GRID_MODE = prior;
    }
  });

  it("clips auditing overlays across frozen panes (excluding headers)", () => {
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

      // Freeze more than just the shared-grid header row/col.
      app.activateCell({ row: 1, col: 2 });
      app.freezePanes(); // document frozenRows=1, frozenCols=2 (renderer frozenRows=2, frozenCols=3)

      const sharedGrid = (app as any).sharedGrid;
      const viewport = sharedGrid.renderer.scroll.getViewportState();
      expect(viewport.frozenRows).toBe(2);
      expect(viewport.frozenCols).toBe(3);

      const headerWidth = sharedGrid.renderer.getColWidth(0);
      const headerHeight = sharedGrid.renderer.getRowHeight(0);

      const ctx = (app as any).auditingCtx;
      ctx.rect.mockClear();
      (app as any).auditingMode = "precedents";
      (app as any).renderAuditing();

      expect(ctx.rect).toHaveBeenCalledTimes(1);
      const [x, y, width, height] = ctx.rect.mock.calls[0] as number[];
      expect(x).toBe(headerWidth);
      expect(y).toBe(headerHeight);
      expect(x + width).toBe(viewport.width);
      expect(y + height).toBe(viewport.height);

      // Regression guard: the clip rect should not shrink by the user-frozen extents.
      expect(width).toBe(viewport.width - headerWidth);
      expect(height).toBe(viewport.height - headerHeight);

      app.destroy();
      root.remove();
    } finally {
      if (prior === undefined) delete process.env.DESKTOP_GRID_MODE;
      else process.env.DESKTOP_GRID_MODE = prior;
    }
  });
});

