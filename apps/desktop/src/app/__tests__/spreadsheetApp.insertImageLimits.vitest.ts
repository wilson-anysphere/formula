/**
 * @vitest-environment jsdom
 */

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { SpreadsheetApp } from "../spreadsheetApp";
import { MAX_INSERT_IMAGE_BYTES } from "../../drawings/insertImage";
import * as ui from "../../extensions/ui.js";

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

describe("SpreadsheetApp image insertion limits", () => {
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

  it("skips oversized image files and shows a toast", async () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const toastSpy = vi.spyOn(ui, "showToast").mockImplementation(() => {});

    const app = new SpreadsheetApp(root, status);
    const sheetId = app.getCurrentSheetId();

    const arrayBuffer = vi.fn(async () => {
      throw new Error("should not read oversized files");
    });
    const oversizedFile = {
      name: "big.png",
      type: "image/png",
      size: MAX_INSERT_IMAGE_BYTES + 1,
      arrayBuffer,
    } as any as File;

    await app.insertPicturesFromFiles([oversizedFile]);

    expect(arrayBuffer).not.toHaveBeenCalled();
    const doc: any = app.getDocument();
    expect(doc.getSheetDrawings(sheetId)).toHaveLength(0);
    expect(toastSpy).toHaveBeenCalledWith("Image too large (>10MB). Choose a smaller file.", "warning");

    app.destroy();
    root.remove();
  });

  it("still inserts normal-sized images", async () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const toastSpy = vi.spyOn(ui, "showToast").mockImplementation(() => {});

    const app = new SpreadsheetApp(root, status);
    const sheetId = app.getCurrentSheetId();

    const arrayBuffer = vi.fn(async () => new Uint8Array([1, 2, 3]).buffer);
    const smallFile = {
      name: "small.png",
      type: "image/png",
      size: 3,
      arrayBuffer,
    } as any as File;

    await app.insertPicturesFromFiles([smallFile]);

    expect(arrayBuffer).toHaveBeenCalledTimes(1);
    const doc: any = app.getDocument();
    expect(doc.getSheetDrawings(sheetId)).toHaveLength(1);
    expect(toastSpy).not.toHaveBeenCalled();

    app.destroy();
    root.remove();
  });
});
