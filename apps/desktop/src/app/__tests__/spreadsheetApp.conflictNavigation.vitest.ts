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

function createRoot(options: { width?: number; height?: number } = {}): HTMLElement {
  const root = document.createElement("div");
  root.tabIndex = 0;
  const width = options.width ?? 800;
  const height = options.height ?? 600;
  root.getBoundingClientRect = () =>
    ({
      width,
      height,
      left: 0,
      top: 0,
      right: width,
      bottom: height,
      x: 0,
      y: 0,
      toJSON: () => {},
    }) as any;
  document.body.appendChild(root);
  return root;
}

describe("SpreadsheetApp conflict navigation callback", () => {
  beforeEach(() => {
    priorGridMode = process.env.DESKTOP_GRID_MODE;
    process.env.DESKTOP_GRID_MODE = "legacy";
    document.body.innerHTML = "";

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

    // Avoid leaving pending RAF work in unit tests.
    Object.defineProperty(globalThis, "requestAnimationFrame", {
      configurable: true,
      value: (cb: FrameRequestCallback) => {
        cb(0);
        return 0;
      },
    });
    Object.defineProperty(globalThis, "cancelAnimationFrame", { configurable: true, value: () => {} });
  });

  afterEach(() => {
    if (priorGridMode === undefined) delete process.env.DESKTOP_GRID_MODE;
    else process.env.DESKTOP_GRID_MODE = priorGridMode;
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
  });

  it("centers the grid on the requested cell and updates selection (legacy grid)", () => {
    const root = createRoot({ width: 800, height: 600 });
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    const sheetId = (app as any).sheetId as string;

    (app as any).navigateToConflictCell({ sheetId, row: 50, col: 30 });

    expect((app as any).selection.active).toEqual({ row: 50, col: 30 });

    const rect = (app as any).getCellRect({ row: 50, col: 30 }) as { x: number; y: number; width: number; height: number } | null;
    expect(rect).not.toBeNull();

    // Center of the scrollable viewport in legacy mode (no frozen panes by default).
    const expectedCenterX = 48 + (800 - 48) / 2;
    const expectedCenterY = 24 + (600 - 24) / 2;

    expect(rect!.x + rect!.width / 2).toBeCloseTo(expectedCenterX, 1);
    expect(rect!.y + rect!.height / 2).toBeCloseTo(expectedCenterY, 1);

    app.destroy();
    root.remove();
  });

  it("activates the requested sheet before navigating", () => {
    const root = createRoot({ width: 800, height: 600 });
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    const documentController = (app as any).document as { addSheet: () => string };
    const sheetId = documentController.addSheet();

    (app as any).navigateToConflictCell({ sheetId, row: 10, col: 5 });

    expect((app as any).sheetId).toBe(sheetId);
    expect((app as any).selection.active).toEqual({ row: 10, col: 5 });

    app.destroy();
    root.remove();
  });
});

