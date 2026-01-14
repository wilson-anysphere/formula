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

function createDeferred<T>(): { promise: Promise<T>; resolve: (value: T) => void; reject: (err: unknown) => void } {
  let resolve!: (value: T) => void;
  let reject!: (err: unknown) => void;
  const promise = new Promise<T>((res, rej) => {
    resolve = res;
    reject = rej;
  });
  return { promise, resolve, reject };
}

describe("SpreadsheetApp fill shortcuts sheet switching", () => {
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

  it("keeps multi-range fill sheet-scoped when switching sheets mid-fill (no focus steal)", async () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    expect(app.getGridMode()).toBe("legacy");

    const doc = app.getDocument();
    const sheet1 = app.getCurrentSheetId();

    // Ensure Sheet2 exists (and seed sentinel values so we can detect cross-sheet writes).
    doc.setCellValue("Sheet2", { row: 0, col: 0 }, "X");
    doc.setCellValue("Sheet2", { row: 1, col: 0 }, "S2_A2");
    doc.setCellValue("Sheet2", { row: 1, col: 2 }, "S2_C2");

    // Seed two disjoint 1x2 vertical ranges: A1:A2 and C1:C2.
    doc.setRangeValues(sheet1, "A1", [[{ formula: "=SUM(1:1)" }], [null]], { label: "Seed" });
    doc.setRangeValues(sheet1, "C1", [[{ formula: "=SUM(1:1)" }], [null]], { label: "Seed" });

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

    const deferred = createDeferred<string[]>();
    const rewrite = vi.fn(async () => deferred.promise);
    (app as any).wasmEngine = { rewriteFormulasForCopyDelta: rewrite, terminate: () => {} };

    const focusSpy = vi.spyOn(app, "focus");

    // Trigger the async fill path, then switch sheets before the formula rewrite resolves.
    app.fillDown();
    app.activateSheet("Sheet2");

    // Each range produces a single rewrite request (fill down by 1 row).
    deferred.resolve(["=SUM(2:2)"]);
    await app.whenIdle();

    const a2 = doc.getCell(sheet1, { row: 1, col: 0 }) as any;
    const c2 = doc.getCell(sheet1, { row: 1, col: 2 }) as any;
    expect(a2.formula).toBe("=SUM(2:2)");
    expect(c2.formula).toBe("=SUM(2:2)");

    const sheet2A2 = doc.getCell("Sheet2", { row: 1, col: 0 }) as any;
    const sheet2C2 = doc.getCell("Sheet2", { row: 1, col: 2 }) as any;
    expect(sheet2A2.formula).toBeNull();
    expect(sheet2A2.value).toBe("S2_A2");
    expect(sheet2C2.formula).toBeNull();
    expect(sheet2C2.value).toBe("S2_C2");

    // Completion should not steal focus back to the grid after the user navigated to a different sheet.
    expect(focusSpy).not.toHaveBeenCalled();

    app.destroy();
    root.remove();
  });
});

