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

describe("SpreadsheetApp fill series shortcut", () => {
  beforeEach(() => {
    priorGridMode = process.env.DESKTOP_GRID_MODE;
    process.env.DESKTOP_GRID_MODE = "legacy";

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

  afterEach(() => {
    if (priorGridMode === undefined) delete process.env.DESKTOP_GRID_MODE;
    else process.env.DESKTOP_GRID_MODE = priorGridMode;
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
  });

  it("fills a numeric series down for a simple vertical range", async () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    const doc = app.getDocument();
    const sheetId = app.getCurrentSheetId();

    doc.setRangeValues(sheetId, "A1", [[1], [2], [null], [null]], { label: "Seed" });

    // Select A1:A4.
    (app as any).selection = {
      type: "range",
      ranges: [{ startRow: 0, endRow: 3, startCol: 0, endCol: 0 }],
      active: { row: 0, col: 0 },
      anchor: { row: 0, col: 0 },
      activeRangeIndex: 0,
    };

    app.fillSeries("down");
    await app.whenIdle();

    expect((doc.getCell(sheetId, { row: 0, col: 0 }) as any).value).toBe(1);
    expect((doc.getCell(sheetId, { row: 1, col: 0 }) as any).value).toBe(2);
    expect((doc.getCell(sheetId, { row: 2, col: 0 }) as any).value).toBe(3);
    expect((doc.getCell(sheetId, { row: 3, col: 0 }) as any).value).toBe(4);

    app.destroy();
    root.remove();
  });

  it("fills a numeric series right for a simple horizontal range", async () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    const doc = app.getDocument();
    const sheetId = app.getCurrentSheetId();

    doc.setRangeValues(sheetId, "A1", [[1, 2, null, null]], { label: "Seed" });

    // Select A1:D1.
    (app as any).selection = {
      type: "range",
      ranges: [{ startRow: 0, endRow: 0, startCol: 0, endCol: 3 }],
      active: { row: 0, col: 0 },
      anchor: { row: 0, col: 0 },
      activeRangeIndex: 0,
    };

    app.fillSeries("right");
    await app.whenIdle();

    expect((doc.getCell(sheetId, { row: 0, col: 0 }) as any).value).toBe(1);
    expect((doc.getCell(sheetId, { row: 0, col: 1 }) as any).value).toBe(2);
    expect((doc.getCell(sheetId, { row: 0, col: 2 }) as any).value).toBe(3);
    expect((doc.getCell(sheetId, { row: 0, col: 3 }) as any).value).toBe(4);

    app.destroy();
    root.remove();
  });
});

