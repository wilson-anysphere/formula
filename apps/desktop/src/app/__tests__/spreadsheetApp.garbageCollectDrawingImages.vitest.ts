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

describe("SpreadsheetApp.garbageCollectDrawingImages", () => {
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

  it("includes image ids from document drawings and local caches", async () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    const sheetId = app.getCurrentSheetId();

    // 1) Document-backed drawings.
    app.getDocument().setSheetDrawings(sheetId, [
      {
        id: "drawing_doc",
        kind: { type: "image", imageId: "image_doc" },
        anchor: { type: "absolute", pos: { xEmu: 0, yEmu: 0 }, size: { cx: pxToEmu(10), cy: pxToEmu(10) } },
        zOrder: 0,
      },
    ]);

    // 2) In-memory drawings cache (used by tests/insert flows and pointer-move previews).
    (app as any).drawingObjectsCache = {
      sheetId,
      objects: [
        {
          id: 2,
          kind: { type: "image", imageId: "image_cache" },
          anchor: { type: "absolute", pos: { xEmu: 0, yEmu: 0 }, size: { cx: pxToEmu(10), cy: pxToEmu(10) } },
          zOrder: 0,
        },
      ],
      source: null,
    };

    const garbageCollectAsync = vi.fn(async (_keep: Iterable<string>) => {});
    // Stub the ImageStore used by SpreadsheetApp so we can assert on the "keep" set passed to GC.
    // Note: SpreadsheetApp teardown also clears the image store.
    (app as any).drawingImages = { get: () => undefined, set: () => {}, garbageCollectAsync, clear: () => {} };

    await app.garbageCollectDrawingImages();

    expect(garbageCollectAsync).toHaveBeenCalledTimes(1);
    const keep = garbageCollectAsync.mock.calls[0]?.[0] as Iterable<string>;
    expect(new Set(Array.from(keep))).toEqual(new Set(["image_doc", "image_cache"]));

    app.destroy();
    root.remove();
  });

  it("includes image ids from externally-tagged drawing kinds", async () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    const sheetId = app.getCurrentSheetId();

    // DocumentController can store drawing kinds as externally-tagged enums:
    // `{ kind: { Image: { image_id } } }`.
    app.getDocument().setSheetDrawings(sheetId, [
      {
        id: "drawing_doc_tagged",
        kind: { Image: { image_id: "image_doc_tagged.png" } },
        anchor: { type: "absolute", pos: { xEmu: 0, yEmu: 0 }, size: { cx: pxToEmu(10), cy: pxToEmu(10) } },
        zOrder: 0,
      },
    ]);

    const garbageCollectAsync = vi.fn(async (_keep: Iterable<string>) => {});
    (app as any).drawingImages = { get: () => undefined, set: () => {}, garbageCollectAsync, clear: () => {} };

    await app.garbageCollectDrawingImages();

    expect(garbageCollectAsync).toHaveBeenCalledTimes(1);
    const keep = garbageCollectAsync.mock.calls[0]?.[0] as Iterable<string>;
    const keepSet = new Set(Array.from(keep));
    expect(keepSet.has("image_doc_tagged.png")).toBe(true);

    app.destroy();
    root.remove();
  });

  it("handles wrapped image ids in drawing kinds", async () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    const sheetId = app.getCurrentSheetId();

    // Some model encodings wrap scalar ids as singleton tuples/structs.
    app.getDocument().setSheetDrawings(sheetId, [
      {
        id: "drawing_doc_wrapped_array",
        kind: { Image: { image_id: ["image_wrapped_array.png"] } },
        anchor: { type: "absolute", pos: { xEmu: 0, yEmu: 0 }, size: { cx: pxToEmu(10), cy: pxToEmu(10) } },
        zOrder: 0,
      },
      {
        id: "drawing_doc_wrapped_object",
        kind: { Image: { image_id: { 0: "image_wrapped_object.png" } } },
        anchor: { type: "absolute", pos: { xEmu: 0, yEmu: 0 }, size: { cx: pxToEmu(10), cy: pxToEmu(10) } },
        zOrder: 1,
      },
    ]);

    const garbageCollectAsync = vi.fn(async (_keep: Iterable<string>) => {});
    (app as any).drawingImages = { get: () => undefined, set: () => {}, garbageCollectAsync, clear: () => {} };

    await app.garbageCollectDrawingImages();

    expect(garbageCollectAsync).toHaveBeenCalledTimes(1);
    const keep = garbageCollectAsync.mock.calls[0]?.[0] as Iterable<string>;
    const keepSet = new Set(Array.from(keep));
    expect(keepSet.has("image_wrapped_array.png")).toBe(true);
    expect(keepSet.has("image_wrapped_object.png")).toBe(true);

    app.destroy();
    root.remove();
  });

  it("keeps image ids referenced by in-cell images", async () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    const sheetId = app.getCurrentSheetId();
    const doc: any = app.getDocument();

    // An image referenced only via a stored cell value (Excel "place in cell" picture / IMAGE()).
    doc.setCellValue(sheetId, { row: 0, col: 0 }, { type: "image", value: { imageId: "image_cell", altText: "Kitten" } });

    const garbageCollectAsync = vi.fn(async (_keep: Iterable<string>) => {});
    (app as any).drawingImages = { get: () => undefined, set: () => {}, garbageCollectAsync, clear: () => {} };

    await app.garbageCollectDrawingImages();

    expect(garbageCollectAsync).toHaveBeenCalledTimes(1);
    const keep = garbageCollectAsync.mock.calls[0]?.[0] as Iterable<string>;
    const keepSet = new Set(Array.from(keep));
    expect(keepSet.has("image_cell")).toBe(true);

    app.destroy();
    root.remove();
  });

  it("keeps image ids referenced by sheet background images", async () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    const sheetId = app.getCurrentSheetId();

    // An image referenced only as a sheet-level tiled background image.
    app.setSheetBackgroundImageId(sheetId, "image_bg");

    const garbageCollectAsync = vi.fn(async (_keep: Iterable<string>) => {});
    (app as any).drawingImages = { get: () => undefined, set: () => {}, garbageCollectAsync, clear: () => {} };

    await app.garbageCollectDrawingImages();

    expect(garbageCollectAsync).toHaveBeenCalledTimes(1);
    const keep = garbageCollectAsync.mock.calls[0]?.[0] as Iterable<string>;
    const keepSet = new Set(Array.from(keep));
    expect(keepSet.has("image_bg")).toBe(true);

    app.destroy();
    root.remove();
  });

  it("keeps image ids referenced by sheet backgrounds on non-active sheets", async () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    const doc: any = app.getDocument();
    const sheet2 = doc.addSheet({ sheetId: "Sheet2" });

    // Set a background image id on a sheet that is not the active one.
    app.setSheetBackgroundImageId(sheet2, "image_bg_other_sheet");

    const garbageCollectAsync = vi.fn(async (_keep: Iterable<string>) => {});
    (app as any).drawingImages = { get: () => undefined, set: () => {}, garbageCollectAsync, clear: () => {} };

    await app.garbageCollectDrawingImages();

    expect(garbageCollectAsync).toHaveBeenCalledTimes(1);
    const keep = garbageCollectAsync.mock.calls[0]?.[0] as Iterable<string>;
    const keepSet = new Set(Array.from(keep));
    expect(keepSet.has("image_bg_other_sheet")).toBe(true);

    app.destroy();
    root.remove();
  });

  it("keeps image ids referenced by in-cell images on non-active sheets", async () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    const doc: any = app.getDocument();
    const sheet2 = doc.addSheet({ sheetId: "Sheet2" });

    // An in-cell image stored on a non-active sheet should still be kept.
    doc.setCellValue(sheet2, { row: 0, col: 0 }, { type: "image", value: { imageId: "image_cell_other_sheet", altText: null } });

    const garbageCollectAsync = vi.fn(async (_keep: Iterable<string>) => {});
    (app as any).drawingImages = { get: () => undefined, set: () => {}, garbageCollectAsync, clear: () => {} };

    await app.garbageCollectDrawingImages();

    expect(garbageCollectAsync).toHaveBeenCalledTimes(1);
    const keep = garbageCollectAsync.mock.calls[0]?.[0] as Iterable<string>;
    const keepSet = new Set(Array.from(keep));
    expect(keepSet.has("image_cell_other_sheet")).toBe(true);

    app.destroy();
    root.remove();
  });
});
