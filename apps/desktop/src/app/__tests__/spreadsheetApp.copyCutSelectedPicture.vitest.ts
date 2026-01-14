/**
 * @vitest-environment jsdom
 */

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { SpreadsheetApp } from "../spreadsheetApp";
import type { DrawingObject } from "../../drawings/types";

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

describe("SpreadsheetApp copy/cut selected picture", () => {
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

  it("copies and cuts a selected picture instead of the cell range", async () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    root.focus();

    const write = vi.fn(async () => {});
    (app as any).clipboardProviderPromise = Promise.resolve({ write, read: vi.fn(async () => ({})) });

    const sheetId = app.getCurrentSheetId();
    const imageId = "img-1";
    const pngBytes = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);

    (app as any).drawingImages.set({ id: imageId, bytes: pngBytes, mimeType: "image/png" });

    const drawing: DrawingObject = {
      id: 1,
      kind: { type: "image", imageId },
      anchor: { type: "absolute", pos: { xEmu: 0, yEmu: 0 }, size: { cx: 0, cy: 0 } },
      zOrder: 0,
    };
    app.getDocument().setSheetDrawings(sheetId, [drawing]);
    (app as any).selectedDrawingId = drawing.id;

    app.copy();
    await app.whenIdle();
    expect(write).toHaveBeenCalledTimes(1);
    expect(write).toHaveBeenCalledWith({ text: "", imagePng: pngBytes });

    app.cut();
    await app.whenIdle();
    expect(write).toHaveBeenCalledTimes(2);
    expect(app.getDocument().getSheetDrawings(sheetId)).toEqual([]);

    app.destroy();
    root.remove();
  });

  it("cuts pictures whose underlying drawing id is a non-numeric string", async () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    root.focus();

    const write = vi.fn(async () => {});
    (app as any).clipboardProviderPromise = Promise.resolve({ write, read: vi.fn(async () => ({})) });

    const sheetId = app.getCurrentSheetId();
    const doc = app.getDocument() as any;

    const imageId = "img-2";
    const pngBytes = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
    doc.setImage(imageId, { bytes: pngBytes, mimeType: "image/png" });

    // Use a non-numeric drawing id (common in imported docs / some backends).
    doc.setSheetDrawings(sheetId, [
      {
        id: "drawing_1",
        kind: { type: "image", imageId },
        anchor: { type: "cell", row: 0, col: 0, size: { width: 10, height: 10 } },
        zOrder: 0,
      },
    ]);

    // Select via the UI-normalized id (which may be a stable hash).
    const objects = (app as any).listDrawingObjectsForSheet(sheetId) as DrawingObject[];
    expect(objects).toHaveLength(1);
    (app as any).selectedDrawingId = objects[0]!.id;

    app.cut();
    await app.whenIdle();

    expect(write).toHaveBeenCalledWith({ text: "", imagePng: pngBytes });
    expect(doc.getSheetDrawings(sheetId)).toEqual([]);

    app.destroy();
    root.remove();
  });
});
