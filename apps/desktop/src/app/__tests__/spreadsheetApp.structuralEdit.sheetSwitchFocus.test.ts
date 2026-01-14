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

describe("SpreadsheetApp async structural edits sheet switching", () => {
  beforeEach(() => {
    priorGridMode = process.env.DESKTOP_GRID_MODE;
    process.env.DESKTOP_GRID_MODE = "shared";
    document.body.innerHTML = "";

    // Node 22 ships an experimental `localStorage` global that errors unless configured via flags.
    // Provide a stable in-memory implementation for unit tests.
    const storage = createInMemoryLocalStorage();
    Object.defineProperty(globalThis, "localStorage", { configurable: true, value: storage });
    Object.defineProperty(window, "localStorage", { configurable: true, value: storage });
    storage.clear();

    // jsdom lacks a real canvas implementation; SpreadsheetApp expects a 2D context.
    Object.defineProperty(HTMLCanvasElement.prototype, "getContext", {
      configurable: true,
      value: () => createMockCanvasContext(),
    });

    // CanvasGridRenderer schedules renders via requestAnimationFrame; ensure it exists in jsdom.
    Object.defineProperty(globalThis, "requestAnimationFrame", {
      configurable: true,
      value: (cb: FrameRequestCallback) => {
        cb(0);
        return 0;
      },
    });
    Object.defineProperty(globalThis, "cancelAnimationFrame", { configurable: true, value: () => {} });

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

  it("does not steal focus after switching sheets while insertRows is awaiting structural edits", async () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    app.getDocument().setCellValue("Sheet2", { row: 0, col: 0 }, "X");

    // Force the async (WASM-backed) code path without bringing up a real engine worker.
    (app as any).wasmEngine = { terminate: () => {} };
    const deferred = createDeferred<void>();
    (app as any).applyStructuralEdit = async () => {
      await deferred.promise;
    };

    const focusSpy = vi.spyOn(app, "focus");
    const op = app.insertRows(0, 1);

    // Switch sheets while the async structural edit is still in-flight.
    app.activateSheet("Sheet2");

    deferred.resolve();
    await op;

    expect(focusSpy).not.toHaveBeenCalled();

    app.destroy();
    root.remove();
  });
});
