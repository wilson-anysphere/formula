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

describe("SpreadsheetApp clipboard cut sheet switching", () => {
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
      writable: true,
      value: (cb: FrameRequestCallback) => {
        cb(0);
        return 0;
      },
    });
    Object.defineProperty(globalThis, "cancelAnimationFrame", { configurable: true, writable: true, value: () => {} });

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

  it("cuts cells from the originating sheet even if the user switches sheets mid-cut", async () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    const doc: any = app.getDocument();
    const sheet1 = app.getCurrentSheetId();

    // Ensure Sheet2 exists.
    doc.setCellValue("Sheet2", { row: 0, col: 0 }, "sheet2");
    doc.setCellValue(sheet1, { row: 0, col: 0 }, "sheet1");

    // Select A1.
    (app as any).selection = {
      type: "range",
      ranges: [{ startRow: 0, endRow: 0, startCol: 0, endCol: 0 }],
      active: { row: 0, col: 0 },
      anchor: { row: 0, col: 0 },
      activeRangeIndex: 0,
    };

    let resolveWrite: (() => void) | null = null;
    let notifyWriteCalled: (() => void) | null = null;
    const writeCalled = new Promise<void>((resolve) => {
      notifyWriteCalled = resolve;
    });
    const writePromise = new Promise<void>((resolve) => {
      resolveWrite = resolve;
    });

    const provider = {
      write: vi.fn(async () => {
        notifyWriteCalled?.();
        return writePromise;
      }),
      read: vi.fn(),
    };
    (app as any).clipboardProviderPromise = Promise.resolve(provider);

    const clearSpy = vi.spyOn(doc, "clearRange");
    const focusSpy = vi.spyOn(app, "focus");

    const cutPromise = (app as any).cutSelectionToClipboard();
    await writeCalled;

    // Switch sheets while the clipboard write is still pending.
    app.activateSheet("Sheet2");
    expect(app.getCurrentSheetId()).toBe("Sheet2");
    focusSpy.mockClear();

    resolveWrite?.();
    await cutPromise;

    // Cut should apply to the sheet where it started.
    expect(clearSpy).toHaveBeenCalled();
    expect(clearSpy.mock.calls[0]?.[0]).toBe(sheet1);
    expect(doc.getCell(sheet1, { row: 0, col: 0 }).value).not.toBe("sheet1");
    expect(doc.getCell("Sheet2", { row: 0, col: 0 }).value).toBe("sheet2");

    // Finishing the cut should not steal focus after switching sheets.
    expect(focusSpy).not.toHaveBeenCalled();

    app.destroy();
    root.remove();
  });
});
