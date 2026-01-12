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
      toJSON: () => {},
    }) as any;
  document.body.appendChild(root);
  return root;
}

describe("SpreadsheetApp fill shortcuts (multi-range) prefer engine formula rewrite when available", () => {
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

  it("rewrites row-range formulas when filling down across multiple disjoint ranges", async () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    expect(app.getGridMode()).toBe("legacy");

    const doc = app.getDocument();
    const sheetId = app.getCurrentSheetId();

    // Seed two disjoint 1x2 vertical ranges: A1:A2 and C1:C2.
    doc.setRangeValues(sheetId, "A1", [[{ formula: "=SUM(1:1)" }], [null]], { label: "Seed" });
    doc.setRangeValues(sheetId, "C1", [[{ formula: "=SUM(1:1)" }], [null]], { label: "Seed" });

    // Multi-range selection (two ranges).
    (app as any).selection = {
      type: "range",
      ranges: [
        { startRow: 0, endRow: 1, startCol: 0, endCol: 0 }, // A1:A2
        { startRow: 0, endRow: 1, startCol: 2, endCol: 2 }, // C1:C2
      ],
      active: { row: 0, col: 0 },
      anchor: { row: 0, col: 0 },
      activeRangeIndex: 0,
    };

    const rewrite = vi.fn(async (requests: Array<{ formula: string; deltaRow: number; deltaCol: number }>) => {
      return requests.map((r) => {
        if (r.formula === "=SUM(1:1)" && r.deltaRow === 1 && r.deltaCol === 0) {
          return "=SUM(2:2)";
        }
        return r.formula;
      });
    });

    (app as any).wasmEngine = { rewriteFormulasForCopyDelta: rewrite, terminate: () => {} };

    app.fillDown();
    await app.whenIdle();

    const a2 = doc.getCell(sheetId, { row: 1, col: 0 }) as any;
    const c2 = doc.getCell(sheetId, { row: 1, col: 2 }) as any;
    expect(a2.formula).toBe("=SUM(2:2)");
    expect(c2.formula).toBe("=SUM(2:2)");

    expect(rewrite).toHaveBeenCalled();

    app.destroy();
    root.remove();
  });
});

