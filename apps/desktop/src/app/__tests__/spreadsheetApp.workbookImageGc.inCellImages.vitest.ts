/**
 * @vitest-environment jsdom
 */

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { pxToEmu } from "../../drawings/overlay";
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

describe("SpreadsheetApp workbook image GC (in-cell images)", () => {
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

  it("treats in-cell image references as external refs so bytes are not GC'd", async () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    const sheetId = app.getCurrentSheetId();

    const imageId = "img_cell_only";
    const bytes = new Uint8Array([1, 2, 3]);
    app.getDocument().setImage(imageId, { bytes, mimeType: "image/png" });

    const manager = (app as any).workbookImageManager as { imageRefCount: Map<string, number> };
    expect(manager.imageRefCount.get(imageId)).toBeUndefined();

    app.getDocument().setCellValue(sheetId, { row: 0, col: 0 }, { type: "image", value: { imageId } });

    expect(manager.imageRefCount.get(imageId)).toBe(1);

    await app.runImageGcNow({ force: true });
    expect(app.getDocument().getImage(imageId)).not.toBeNull();

    app.destroy();
    root.remove();
  });

  it("does not GC shared image bytes while still referenced by a cell after deleting a drawing", async () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    const sheetId = app.getCurrentSheetId();

    const imageId = "img_shared_drawing_cell";
    const bytes = new Uint8Array([4, 5, 6, 7]);
    app.getDocument().setImage(imageId, { bytes, mimeType: "image/png" });

    const manager = (app as any).workbookImageManager as { imageRefCount: Map<string, number> };

    // Reference the image from both a drawing and an in-cell image value.
    app.getDocument().setCellValue(sheetId, { row: 0, col: 0 }, { type: "image", value: { imageId } });
    app.getDocument().setSheetDrawings(sheetId, [
      {
        id: "d1",
        kind: { type: "image", imageId },
        anchor: { type: "absolute", pos: { xEmu: 0, yEmu: 0 }, size: { cx: pxToEmu(10), cy: pxToEmu(10) } },
        zOrder: 0,
      },
    ]);

    expect(manager.imageRefCount.get(imageId)).toBe(2);

    // Delete the drawing; the cell reference should keep bytes alive.
    app.getDocument().setSheetDrawings(sheetId, []);
    expect(manager.imageRefCount.get(imageId)).toBe(1);
    await app.runImageGcNow({ force: true });
    expect(app.getDocument().getImage(imageId)).not.toBeNull();

    // Clear the cell value; bytes should become eligible for GC.
    app.getDocument().setCellValue(sheetId, { row: 0, col: 0 }, null);
    expect(manager.imageRefCount.get(imageId)).toBeUndefined();
    await app.runImageGcNow({ force: true });
    expect(app.getDocument().getImage(imageId)).toBeNull();

    app.destroy();
    root.remove();
  });
});

