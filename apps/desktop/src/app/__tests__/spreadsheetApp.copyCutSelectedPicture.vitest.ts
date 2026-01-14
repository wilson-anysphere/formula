/**
 * @vitest-environment jsdom
 */

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { SpreadsheetApp } from "../spreadsheetApp";
import type { DrawingObject } from "../../drawings/types";

let priorGridMode: string | undefined;

function createPngHeaderBytes(width = 1, height = 1): Uint8Array {
  const bytes = new Uint8Array(24);
  bytes.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a], 0);
  // 13-byte IHDR chunk length.
  bytes[8] = 0x00;
  bytes[9] = 0x00;
  bytes[10] = 0x00;
  bytes[11] = 0x0d;
  // IHDR chunk type.
  bytes[12] = 0x49;
  bytes[13] = 0x48;
  bytes[14] = 0x44;
  bytes[15] = 0x52;

  const view = new DataView(bytes.buffer);
  view.setUint32(16, width, false);
  view.setUint32(20, height, false);
  return bytes;
}

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

    const app = new SpreadsheetApp(root, status, { enableDrawingInteractions: true });
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
    app.selectDrawing(drawing.id);

    app.copy();
    await app.whenIdle();
    expect(write).toHaveBeenCalledTimes(1);
    expect(write).toHaveBeenCalledWith({ text: "", imagePng: pngBytes });

    app.cut();
    await app.whenIdle();
    expect(write).toHaveBeenCalledTimes(2);
    expect(app.getDocument().getSheetDrawings(sheetId)).toEqual([]);
    expect((app as any).selectedDrawingId).toBe(null);
    expect(((app as any).drawingOverlay as any).selectedId).toBe(null);
    expect(((app as any).drawingInteractionController as any).selectedId).toBe(null);

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

  it("does not delete a picture on a different sheet if the user switches sheets mid-cut", async () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status, { enableDrawingInteractions: true });
    root.focus();

    const doc: any = app.getDocument();
    const sheet1 = app.getCurrentSheetId();
    // Ensure Sheet2 exists so we can switch away while the clipboard write is in-flight.
    doc.setCellValue("Sheet2", { row: 0, col: 0 }, "X");

    const imageId1 = "img-sheet1";
    const imageId2 = "img-sheet2";
    const pngBytes = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
    (app as any).drawingImages.set({ id: imageId1, bytes: pngBytes, mimeType: "image/png" });
    (app as any).drawingImages.set({ id: imageId2, bytes: pngBytes, mimeType: "image/png" });

    // Intentionally collide UI ids across sheets: if the cut implementation uses the *current* sheet
    // after an async clipboard write, it could delete the wrong drawing on Sheet2.
    doc.setSheetDrawings(sheet1, [
      {
        id: 1,
        kind: { type: "image", imageId: imageId1 },
        anchor: { type: "absolute", pos: { xEmu: 0, yEmu: 0 }, size: { cx: 0, cy: 0 } },
        zOrder: 0,
      },
    ] satisfies DrawingObject[]);
    doc.setSheetDrawings("Sheet2", [
      {
        id: 1,
        kind: { type: "image", imageId: imageId2 },
        anchor: { type: "absolute", pos: { xEmu: 0, yEmu: 0 }, size: { cx: 0, cy: 0 } },
        zOrder: 0,
      },
    ] satisfies DrawingObject[]);

    app.selectDrawing(1);

    let resolveWrite: (() => void) | null = null;
    const writePromise = new Promise<void>((resolve) => {
      resolveWrite = resolve;
    });
    const write = vi.fn(async () => writePromise);
    (app as any).clipboardProviderPromise = Promise.resolve({ write, read: vi.fn(async () => ({})) });

    const focusSpy = vi.spyOn(app, "focus");

    // Start the cut, but keep the clipboard write pending.
    app.cut();
    await new Promise((resolve) => setTimeout(resolve, 0));
    expect(write).toHaveBeenCalledTimes(1);

    // Switch sheets while `cutSelectedDrawingToClipboard` is awaiting the clipboard write.
    app.activateSheet("Sheet2");
    app.selectDrawing(1);

    // Clear incidental focus calls before resolving the clipboard write.
    focusSpy.mockClear();

    resolveWrite?.();
    await app.whenIdle();

    // Cut should apply to the originating sheet, not the newly active sheet.
    expect(doc.getSheetDrawings(sheet1)).toEqual([]);
    expect(doc.getSheetDrawings("Sheet2")).toHaveLength(1);

    // Completing the cut should not clear the current sheet's selection or steal focus.
    expect(app.getCurrentSheetId()).toBe("Sheet2");
    expect(app.getSelectedDrawingId()).toBe(1);
    expect(focusSpy).not.toHaveBeenCalled();

    app.destroy();
    root.remove();
  });

  it("does not write oversized pictures to the clipboard", async () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status, { enableDrawingInteractions: true });
    root.focus();

    const write = vi.fn(async () => {});
    (app as any).clipboardProviderPromise = Promise.resolve({ write, read: vi.fn(async () => ({})) });

    const sheetId = app.getCurrentSheetId();
    const imageId = "img-oversized";
    const maxBytes = 5 * 1024 * 1024; // keep in sync with CLIPBOARD_LIMITS.maxImageBytes
    const oversized = new Uint8Array(maxBytes + 1);
    oversized[0] = 0x89;
    oversized[1] = 0x50;
    oversized[2] = 0x4e;
    oversized[3] = 0x47;
    oversized[4] = 0x0d;
    oversized[5] = 0x0a;
    oversized[6] = 0x1a;
    oversized[7] = 0x0a;

    (app as any).drawingImages.set({ id: imageId, bytes: oversized, mimeType: "image/png" });

    const drawing: DrawingObject = {
      id: 1,
      kind: { type: "image", imageId },
      anchor: { type: "absolute", pos: { xEmu: 0, yEmu: 0 }, size: { cx: 0, cy: 0 } },
      zOrder: 0,
    };
    app.getDocument().setSheetDrawings(sheetId, [drawing]);
    app.selectDrawing(drawing.id);

    app.copy();
    await app.whenIdle();
    expect(write).not.toHaveBeenCalled();
    expect(app.getDocument().getSheetDrawings(sheetId)).toHaveLength(1);

    app.destroy();
    root.remove();
  });

  it("copies PNG bytes even when mimeType metadata is incorrect", async () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status, { enableDrawingInteractions: true });
    root.focus();

    const write = vi.fn(async () => {});
    (app as any).clipboardProviderPromise = Promise.resolve({ write, read: vi.fn(async () => ({})) });

    const createImageBitmapMock = vi.fn(() => Promise.reject(new Error("should not be called")));
    vi.stubGlobal("createImageBitmap", createImageBitmapMock as unknown as typeof createImageBitmap);
    // The drawing overlay can try to decode bitmaps for rendering when `createImageBitmap` exists.
    // This test is scoped to the clipboard copy path, so keep overlay rendering disabled to avoid
    // unrelated decode attempts.
    vi.spyOn((app as any).drawingOverlay, "render").mockImplementation(() => {});
    const sheetId = app.getCurrentSheetId();
    const imageId = "img-wrong-mime";
    const pngBytes = createPngHeaderBytes(1, 1);

    // Intentionally wrong mimeType; bytes are actually a PNG.
    (app as any).drawingImages.set({ id: imageId, bytes: pngBytes, mimeType: "image/jpeg" });

    const drawing: DrawingObject = {
      id: 1,
      kind: { type: "image", imageId },
      anchor: { type: "absolute", pos: { xEmu: 0, yEmu: 0 }, size: { cx: 0, cy: 0 } },
      zOrder: 0,
    };
    app.getDocument().setSheetDrawings(sheetId, [drawing]);
    app.selectDrawing(drawing.id);
    await app.whenIdle();
    // The drawings overlay may attempt best-effort decoding when `createImageBitmap` is available.
    // This test specifically asserts that the clipboard copy path does not invoke `createImageBitmap`
    // when the stored bytes already look like a PNG, even if the mimeType metadata is wrong.
    // Clear any incidental decode attempts from selection/initialization; the assertions below
    // are scoped to the `copy()` path.
    createImageBitmapMock.mockClear();

    // Selecting a drawing may trigger an asynchronous render, which can preload images for display.
    // This test specifically asserts that the *copy* path does not invoke `createImageBitmap` when
    // the bytes already look like a PNG (even if the stored mimeType is wrong).
    createImageBitmapMock.mockClear();

    app.copy();
    await app.whenIdle();
    expect(write).toHaveBeenCalledWith({ text: "", imagePng: pngBytes });
    expect(createImageBitmapMock).not.toHaveBeenCalled();

    app.destroy();
    root.remove();
  });
});
