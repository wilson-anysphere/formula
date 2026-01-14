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

describe("SpreadsheetApp shared-grid fill handle sheet switching", () => {
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

  it("keeps async shared-grid fill-handle commits sheet-scoped when switching sheets mid-fill (no focus steal)", async () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    expect(app.getGridMode()).toBe("shared");

    const doc = app.getDocument();
    const sheet1 = app.getCurrentSheetId();

    // Ensure Sheet2 exists (and seed sentinel values so we can detect cross-sheet writes).
    doc.setCellValue("Sheet2", { row: 0, col: 1 }, "S2_B1");

    // Seed Sheet1 A1 with a formula. We'll fill right to B1.
    doc.setRangeValues(sheet1, "A1", [[{ formula: "=1+1" }]], { label: "Seed" });

    const deferred = createDeferred<string[]>();
    const rewrite = vi.fn(() => deferred.promise);
    (app as any).wasmEngine = { rewriteFormulasForCopyDelta: rewrite, terminate: () => {} };

    const sharedGrid = (app as any).sharedGrid as any;
    const onFillCommit = sharedGrid?.callbacks?.onFillCommit as ((event: any) => void) | undefined;
    expect(typeof onFillCommit).toBe("function");

    const focusSpy = vi.spyOn(app, "focus");

    // Trigger the async fill path, then switch sheets before the formula rewrite fails.
    // Shared grid ranges include 1-row/1-col headers at index 0.
    onFillCommit!({
      sourceRange: { startRow: 1, endRow: 2, startCol: 1, endCol: 2 }, // A1
      targetRange: { startRow: 1, endRow: 2, startCol: 2, endCol: 3 }, // B1
      mode: "formulas",
    });
    app.activateSheet("Sheet2");

    // Reject to force the `.catch` fallback path, which must still apply to Sheet1 (not Sheet2).
    deferred.reject(new Error("rewrite unavailable"));
    await app.whenIdle();

    const sheet1B1 = doc.getCell(sheet1, { row: 0, col: 1 }) as any;
    expect(sheet1B1.formula).toBe("=1+1");

    const sheet2B1 = doc.getCell("Sheet2", { row: 0, col: 1 }) as any;
    expect(sheet2B1.formula).toBeNull();
    expect(sheet2B1.value).toBe("S2_B1");

    // Completion should not steal focus back to the grid after the user navigated to a different sheet.
    expect(focusSpy).not.toHaveBeenCalled();

    app.destroy();
    root.remove();
  });
});

