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

describe("SpreadsheetApp fill commit prefers engine formula rewrite when available", () => {
  beforeEach(() => {
    priorGridMode = process.env.DESKTOP_GRID_MODE;
    process.env.DESKTOP_GRID_MODE = "shared";

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

  it("rewrites formulas like =SUM(A:A) when filling right (engine can shift column ranges; JS shifter cannot)", async () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    expect(app.getGridMode()).toBe("shared");

    const doc = app.getDocument();
    const sheetId = app.getCurrentSheetId();
    doc.setRangeValues(sheetId, "A1", [[{ formula: "=SUM(A:A)" }]], { label: "Seed" });

    const rewrite = vi.fn(async (requests: Array<{ formula: string; deltaRow: number; deltaCol: number }>) => {
      return requests.map((r) => {
        if (r.formula === "=SUM(A:A)" && r.deltaRow === 0 && r.deltaCol === 1) {
          return "=SUM(B:B)";
        }
        return r.formula;
      });
    });

    (app as any).wasmEngine = { rewriteFormulasForCopyDelta: rewrite, terminate: () => {} };

    const sharedGrid = (app as any).sharedGrid as any;
    const onFillCommit = sharedGrid?.callbacks?.onFillCommit as ((event: any) => void) | undefined;
    expect(typeof onFillCommit).toBe("function");

    // Shared grid ranges include 1-row/1-col headers at index 0.
    // Fill right: A1 -> B1
    onFillCommit!({
      sourceRange: { startRow: 1, endRow: 2, startCol: 1, endCol: 2 },
      targetRange: { startRow: 1, endRow: 2, startCol: 2, endCol: 3 },
      mode: "formulas",
    });

    await app.whenIdle();

    expect(rewrite).toHaveBeenCalledTimes(1);
    expect(rewrite).toHaveBeenCalledWith([{ formula: "=SUM(A:A)", deltaRow: 0, deltaCol: 1 }]);

    const filled = doc.getCell(sheetId, { row: 0, col: 1 }) as any; // B1
    expect(filled.formula).toBe("=SUM(B:B)");

    app.destroy();
    root.remove();
  });
});
