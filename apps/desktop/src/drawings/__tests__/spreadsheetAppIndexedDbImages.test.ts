/**
 * @vitest-environment jsdom
 */

import "fake-indexeddb/auto";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { SpreadsheetApp } from "../../app/spreadsheetApp";
import { pxToEmu } from "../overlay";
import { IndexedDbImageStore } from "../persistence/indexedDbImageStore";

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

function createMockCanvasContext(): CanvasRenderingContext2D {
  const noop = () => {};
  const gradient = { addColorStop: noop } as any;
  const base: any = {
    canvas: document.createElement("canvas"),
    measureText: (text: string) => ({ width: text.length * 8 }),
    createLinearGradient: () => gradient,
    createPattern: () => null,
    getImageData: () => ({ data: new Uint8ClampedArray(), width: 0, height: 0 }),
    putImageData: noop,
    drawImage: vi.fn(),
    clearRect: vi.fn(),
    setTransform: vi.fn(),
    scale: vi.fn(),
  };
  return new Proxy(base, {
    get(target, prop) {
      if (prop in target) return (target as any)[prop];
      return noop;
    },
    set(target, prop, value) {
      (target as any)[prop] = value;
      return true;
    },
  }) as CanvasRenderingContext2D;
}

describe("SpreadsheetApp + IndexedDbImageStore", () => {
  const contexts = new WeakMap<HTMLCanvasElement, CanvasRenderingContext2D>();
  const priorGridMode = process.env.DESKTOP_GRID_MODE;

  afterEach(() => {
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
    if (priorGridMode === undefined) delete process.env.DESKTOP_GRID_MODE;
    else process.env.DESKTOP_GRID_MODE = priorGridMode;
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
      value: (cb: FrameRequestCallback) => {
        cb(0);
        return 0;
      },
    });
    Object.defineProperty(globalThis, "cancelAnimationFrame", { configurable: true, value: () => {} });

    Object.defineProperty(globalThis, "createImageBitmap", {
      configurable: true,
      value: vi.fn(async () => ({} as any)),
    });

    (globalThis as any).ResizeObserver = class {
      observe() {}
      disconnect() {}
    };

    Object.defineProperty(HTMLCanvasElement.prototype, "getContext", {
      configurable: true,
      value(this: HTMLCanvasElement, type: string) {
        if (type !== "2d") return null;
        const existing = contexts.get(this);
        if (existing) return existing;
        const created = createMockCanvasContext();
        contexts.set(this, created);
        return created;
      },
    });
  });

  it("can render an image after its in-memory store is cleared, as long as IndexedDB is populated", async () => {
    const workbookId = `wb_${Date.now()}_${Math.random().toString(16).slice(2)}`;
    const entry = { id: "image_1", mimeType: "image/png", bytes: new Uint8Array([9, 8, 7, 6]) };

    // Pre-seed IndexedDB (simulating a prior session / reload).
    const seedStore = new IndexedDbImageStore(workbookId);
    await seedStore.setAsync(entry);

    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status, { workbookId });

    // Seed a floating drawing that references our image id (bytes are intentionally not in memory).
    const sheetId = app.getCurrentSheetId();
    app.getDocument().setSheetDrawings(sheetId, [
      {
        id: "drawing_1",
        kind: { type: "image", imageId: entry.id },
        anchor: { type: "absolute", pos: { xEmu: 0, yEmu: 0 }, size: { cx: pxToEmu(10), cy: pxToEmu(10) } },
        zOrder: 0,
      },
    ]);

    const images = (app as any).drawingImages as { get: (id: string) => any };
    expect(images.get(entry.id)).toBeUndefined();

    const overlay = (app as any).drawingOverlay as { render: (...args: any[]) => Promise<void> };
    const viewport = (app as any).getDrawingRenderViewport();
    const objects = (app as any).listDrawingObjectsForSheet();
    await overlay.render(objects, viewport);

    const hydrated = images.get(entry.id);
    expect(hydrated).toBeTruthy();
    expect(Array.from(hydrated.bytes)).toEqual(Array.from(entry.bytes));

    const drawingCanvas = root.querySelector<HTMLCanvasElement>('[data-testid="drawing-layer-canvas"]')!;
    const ctx = contexts.get(drawingCanvas) as any;
    expect(ctx).toBeTruthy();
    expect(ctx.drawImage).toHaveBeenCalled();

    app.destroy();
    root.remove();
  });

  it("persists inserted picture bytes to IndexedDB without storing them in DocumentController snapshots", async () => {
    const workbookId = `wb_${Date.now()}_${Math.random().toString(16).slice(2)}`;
    const bytes = new Uint8Array([1, 2, 3, 4, 5, 6]);
    // JSDOM's `File` implementation can vary; provide a minimal `File`-like object that
    // matches what `insertPicturesFromFiles` needs (`name`, `type`, `arrayBuffer()`).
    const file = {
      name: "test.png",
      type: "image/png",
      arrayBuffer: async () => bytes.buffer.slice(0),
    } as unknown as File;

    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status, { workbookId });
    await app.insertPicturesFromFiles([file]);

    const sheetId = app.getCurrentSheetId();
    const drawings = app.getDocument().getSheetDrawings(sheetId) as any[];
    expect(drawings.length).toBeGreaterThan(0);
    const imageId = String(drawings[drawings.length - 1]?.kind?.imageId ?? "");
    expect(imageId).toBeTruthy();

    const snapshotBytes = app.getDocument().encodeState();
    const snapshotText = new TextDecoder().decode(snapshotBytes);
    const snapshot = JSON.parse(snapshotText);
    expect(snapshot.images).toBeUndefined();

    const store = new IndexedDbImageStore(workbookId);
    let loaded: any = null;
    // Persistence is best-effort and happens async; poll briefly for the record to appear.
    for (let attempt = 0; attempt < 20; attempt += 1) {
      loaded = await store.getAsync(imageId);
      if (loaded) break;
      await new Promise((r) => setTimeout(r, 0));
    }

    expect(loaded).toBeTruthy();
    expect(loaded?.mimeType).toBe("image/png");
    expect(Array.from(loaded!.bytes)).toEqual(Array.from(bytes));

    app.destroy();
    root.remove();
  });

  it("can reload and render a previously inserted picture using IndexedDB bytes (no snapshot bytes)", async () => {
    const workbookId = `wb_${Date.now()}_${Math.random().toString(16).slice(2)}`;
    const bytes = new Uint8Array([11, 22, 33, 44]);
    const file = {
      name: "reloaded.png",
      type: "image/png",
      arrayBuffer: async () => bytes.buffer.slice(0),
    } as unknown as File;

    const root1 = createRoot();
    const status1 = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app1 = new SpreadsheetApp(root1, status1, { workbookId });
    await app1.insertPicturesFromFiles([file]);

    const sheetId = app1.getCurrentSheetId();
    const drawings = app1.getDocument().getSheetDrawings(sheetId) as any[];
    expect(drawings.length).toBeGreaterThan(0);
    const imageId = String(drawings[drawings.length - 1]?.kind?.imageId ?? "");
    expect(imageId).toBeTruthy();

    const snapshotBytes = app1.getDocument().encodeState();
    const snapshot = JSON.parse(new TextDecoder().decode(snapshotBytes));
    expect(snapshot.images).toBeUndefined();

    // Ensure the bytes actually made it to IndexedDB before simulating reload.
    const seedStore = new IndexedDbImageStore(workbookId);
    let loaded: any = null;
    for (let attempt = 0; attempt < 20; attempt += 1) {
      loaded = await seedStore.getAsync(imageId);
      if (loaded) break;
      await new Promise((r) => setTimeout(r, 0));
    }
    expect(loaded).toBeTruthy();

    app1.destroy();
    root1.remove();

    const root2 = createRoot();
    const status2 = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app2 = new SpreadsheetApp(root2, status2, { workbookId });
    // Simulate reloading the workbook from a snapshot that does not include image bytes.
    app2.getDocument().applyState(snapshotBytes);

    const images2 = (app2 as any).drawingImages as { get: (id: string) => any };
    expect(images2.get(imageId)).toBeUndefined();

    const overlay = (app2 as any).drawingOverlay as { render: (...args: any[]) => Promise<void> };
    const viewport = (app2 as any).getDrawingRenderViewport();
    const objects = (app2 as any).listDrawingObjectsForSheet();
    await overlay.render(objects, viewport);

    const hydrated = images2.get(imageId);
    expect(hydrated).toBeTruthy();
    expect(Array.from(hydrated.bytes)).toEqual(Array.from(bytes));

    const drawingCanvas = root2.querySelector<HTMLCanvasElement>('[data-testid="drawing-layer-canvas"]')!;
    const ctx = contexts.get(drawingCanvas) as any;
    expect(ctx.drawImage).toHaveBeenCalled();

    app2.destroy();
    root2.remove();
  });
});
