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

describe("SpreadsheetApp drag/drop image file insertion", () => {
  afterEach(() => {
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
    delete process.env.DESKTOP_GRID_MODE;
  });

  beforeEach(() => {
    document.body.innerHTML = "";
    process.env.DESKTOP_GRID_MODE = "legacy";

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

  it("dispatches dropped image files to insertPicturesFromFiles at the drop cell", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    const insertPicturesFromFiles = vi.fn();
    (app as any).insertPicturesFromFiles = insertPicturesFromFiles;

    const file = new File([new Uint8Array([1, 2, 3])], "cat.png", { type: "image/png" });
    const dataTransfer = { files: [file], types: ["Files"], dropEffect: "none", items: [] } as any;

    const event = new Event("drop", { bubbles: true, cancelable: true }) as any;
    Object.defineProperty(event, "dataTransfer", { value: dataTransfer });
    Object.defineProperty(event, "clientX", { value: 60 });
    Object.defineProperty(event, "clientY", { value: 30 });
    root.dispatchEvent(event);

    expect(insertPicturesFromFiles).toHaveBeenCalledWith([file], { placeAt: { row: 0, col: 0 } });

    app.destroy();
    root.remove();
  });

  it("falls back to the active cell when the drop point cannot be resolved", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    const insertPicturesFromFiles = vi.fn();
    (app as any).insertPicturesFromFiles = insertPicturesFromFiles;

    app.selectRange({ range: { startRow: 2, endRow: 2, startCol: 2, endCol: 2 } }, { scrollIntoView: false, focus: false });

    const file = new File([new Uint8Array([4, 5, 6])], "dog.png", { type: "image/png" });
    const dataTransfer = { files: [file], types: ["Files"], dropEffect: "none", items: [] } as any;

    const event = new Event("drop", { bubbles: true, cancelable: true }) as any;
    Object.defineProperty(event, "dataTransfer", { value: dataTransfer });
    Object.defineProperty(event, "clientX", { value: 10 });
    Object.defineProperty(event, "clientY", { value: 10 });
    root.dispatchEvent(event);

    expect(insertPicturesFromFiles).toHaveBeenCalledWith([file], { placeAt: { row: 2, col: 2 } });

    app.destroy();
    root.remove();
  });
});

