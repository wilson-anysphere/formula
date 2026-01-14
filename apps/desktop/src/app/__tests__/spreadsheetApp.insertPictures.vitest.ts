/**
 * @vitest-environment jsdom
 */

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { SpreadsheetApp } from "../spreadsheetApp";
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

function createApp(root: HTMLElement, status: { activeCell: HTMLElement; selectionRange: HTMLElement; activeValue: HTMLElement }): SpreadsheetApp {
  const app = new SpreadsheetApp(root, status);
  // SpreadsheetApp seeds a demo ChartStore chart in non-collab mode. With canvas charts enabled by
  // default, that chart appears in `getDrawingObjects()` and would make these picture-focused tests
  // assert on the wrong object counts.
  for (const chart of app.listCharts()) {
    (app as any).chartStore.deleteChart(chart.id);
  }
  return app;
}

describe("SpreadsheetApp insertPicturesFromFiles", () => {
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

  it("inserts a DrawingObject + ImageStore entry and selects it", async () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = createApp(root, status);
    const file = new File([new Uint8Array([1, 2, 3])], "cat.png", { type: "image/png" });

    await app.insertPicturesFromFiles([file]);

    const sheetId = app.getCurrentSheetId();
    const objects = app.getDrawingObjects(sheetId).filter((obj) => obj.kind.type === "image");
    expect(objects).toHaveLength(1);
    const obj = objects[0]!;
    if (obj.kind.type !== "image") {
      throw new Error(`Expected inserted object kind to be image, got ${obj.kind.type}`);
    }

    const imageId = obj.kind.imageId;
    const entry = app.getDrawingImages().get(imageId);
    expect(entry).toBeDefined();
    expect(entry?.mimeType).toBe("image/png");
    expect(Array.from(entry?.bytes ?? [])).toEqual([1, 2, 3]);

    expect(app.getSelectedDrawingId()).toBe(obj.id);

    app.destroy();
    root.remove();
  });

  it("increments zOrder for multiple inserts", async () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = createApp(root, status);
    const file1 = new File([new Uint8Array([1])], "a.png", { type: "image/png" });
    const file2 = new File([new Uint8Array([2])], "b.png", { type: "image/png" });

    await app.insertPicturesFromFiles([file1, file2]);

    const sheetId = app.getCurrentSheetId();
    const objects = app.getDrawingObjects(sheetId).filter((obj) => obj.kind.type === "image");
    expect(objects).toHaveLength(2);
    expect(objects[0]!.zOrder).toBeLessThan(objects[1]!.zOrder);
    expect(app.getSelectedDrawingId()).toBe(objects[1]!.id);

    app.destroy();
    root.remove();
  });

  it("clears any active chart selection when inserting pictures", async () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = createApp(root, status);
    // Simulate a selected chart (e.g. user previously clicked a chart). Inserting a picture should
    // clear chart selection so ribbon/panels reflect the newly-selected picture.
    (app as any).selectedChartId = "chart_123";

    const file = new File([new Uint8Array([1, 2, 3])], "cat.png", { type: "image/png" });
    await app.insertPicturesFromFiles([file]);

    expect(app.getSelectedChartId()).toBeNull();

    app.destroy();
    root.remove();
  });

  it("skips PNGs with extremely large dimensions and shows a toast", async () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const toastSpy = vi.spyOn(ui, "showToast").mockImplementation(() => {});

    const app = createApp(root, status);
    const sheetId = app.getCurrentSheetId();

    // Construct a minimal PNG header with an oversized IHDR width.
    const png = new Uint8Array(24);
    png.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a], 0);
    png.set([0x49, 0x48, 0x44, 0x52], 12); // IHDR
    // width=10001, height=1 (big-endian)
    png.set([0x00, 0x00, 0x27, 0x11], 16);
    png.set([0x00, 0x00, 0x00, 0x01], 20);

    const file = new File([png], "huge.png", { type: "image/png" });

    const setSpy = vi.spyOn(app.getDrawingImages(), "set");
    await app.insertPicturesFromFiles([file]);

    expect(setSpy).not.toHaveBeenCalled();
    expect((app.getDocument() as any).getSheetDrawings(sheetId)).toHaveLength(0);
    expect(toastSpy).toHaveBeenCalledWith("Image dimensions too large (10001x1). Choose a smaller image.", "warning");

    app.destroy();
    root.remove();
  });

  it("does not persist image bytes when inserting pictures fails", async () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = createApp(root, status);
    const docAny = app.getDocument() as any;

    // Force the document mutation to fail so we can assert we don't persist image bytes.
    docAny.setSheetDrawings = vi.fn(() => {
      throw new Error("setSheetDrawings failed");
    });

    const setSpy = vi.spyOn(app.getDrawingImages(), "set");

    const file = new File([new Uint8Array([1, 2, 3])], "cat.png", { type: "image/png" });
    await expect(app.insertPicturesFromFiles([file])).rejects.toThrow(/setSheetDrawings failed/);

    expect(setSpy).not.toHaveBeenCalled();

    app.destroy();
    root.remove();
  });
});
