/**
 * @vitest-environment jsdom
 */

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { SpreadsheetApp } from "../spreadsheetApp";
import * as ui from "../../extensions/ui.js";

function createPngHeaderBytes(width: number, height: number): Uint8Array {
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

function createApp(root: HTMLElement, status: any): SpreadsheetApp {
  const app = new SpreadsheetApp(root, status);
  // Canvas charts are enabled by default, so any ChartStore charts appear in `getDrawingObjects()`
  // alongside pasted images. Remove any charts so these tests can assert on the number of pasted
  // image drawings deterministically.
  for (const chart of app.listCharts()) {
    (app as any).chartStore.deleteChart(chart.id);
  }
  return app;
}

describe("SpreadsheetApp paste image clipboard", () => {
  afterEach(() => {
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
  });

  beforeEach(() => {
    document.body.innerHTML = "";

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

    Object.defineProperty(HTMLCanvasElement.prototype, "getContext", {
      configurable: true,
      value: () => createMockCanvasContext(),
    });

    (globalThis as any).ResizeObserver = class {
      observe() {}
      disconnect() {}
    };
  });

  it("pastes an image-only clipboard payload as a drawing object and selects it", async () => {
    // 1x1 transparent PNG.
    const pngBytes = new Uint8Array(
      // eslint-disable-next-line no-undef
      Buffer.from(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO7+FeAAAAAASUVORK5CYII=",
        "base64",
      ),
    );

    Object.defineProperty(globalThis, "createImageBitmap", {
      configurable: true,
      value: vi.fn(async () => ({ width: 64, height: 32 })),
    });

    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = createApp(root, status);

    const provider = {
      read: vi.fn(async () => ({ imagePng: pngBytes })),
      write: vi.fn(async () => {}),
    };
    (app as any).clipboardProviderPromise = Promise.resolve(provider);

    await app.pasteClipboardToSelection();

    const images = app.getDrawingObjects().filter((obj) => obj.kind.type === "image");
    expect(images).toHaveLength(1);
    expect(images[0]!.kind.type).toBe("image");
    expect(app.getSelectedDrawingId()).toBe(images[0]!.id);

    const imageId = (images[0]!.kind as any).imageId;
    expect(typeof imageId).toBe("string");
    // Image bytes are stored out-of-band (IndexedDB + in-memory drawing image store)
    // rather than in DocumentController snapshots.
    const stored = app.getDrawingImages().get(imageId);
    expect(stored?.mimeType).toBe("image/png");
    expect(stored?.bytes).toBeInstanceOf(Uint8Array);

    app.destroy();
    root.remove();
  });

  it("pastes image clipboard payloads even when text/plain is an empty string", async () => {
    // 1x1 transparent PNG.
    const pngBytes = new Uint8Array(
      // eslint-disable-next-line no-undef
      Buffer.from(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO7+FeAAAAAASUVORK5CYII=",
        "base64",
      ),
    );

    Object.defineProperty(globalThis, "createImageBitmap", {
      configurable: true,
      value: vi.fn(async () => ({ width: 64, height: 32 })),
    });

    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = createApp(root, status);

    const provider = {
      // Some platforms include `text/plain=""` alongside the image bytes.
      read: vi.fn(async () => ({ imagePng: pngBytes, text: "" })),
      write: vi.fn(async () => {}),
    };
    (app as any).clipboardProviderPromise = Promise.resolve(provider);

    await app.pasteClipboardToSelection();

    const images = app.getDrawingObjects().filter((obj) => obj.kind.type === "image");
    expect(images).toHaveLength(1);
    expect(images[0]!.kind.type).toBe("image");

    app.destroy();
    root.remove();
  });

  it("pastes pngBase64 clipboard payloads as a drawing object (legacy fallback)", async () => {
    const base64 =
      "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO7+FeAAAAAASUVORK5CYII=";

    Object.defineProperty(globalThis, "createImageBitmap", {
      configurable: true,
      value: vi.fn(async () => ({ width: 64, height: 32 })),
    });

    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = createApp(root, status);

    const provider = {
      read: vi.fn(async () => ({ pngBase64: `data:image/png;base64,${base64}` })),
      write: vi.fn(async () => {}),
    };
    (app as any).clipboardProviderPromise = Promise.resolve(provider);

    await app.pasteClipboardToSelection();

    const images = app.getDrawingObjects().filter((obj) => obj.kind.type === "image");
    expect(images).toHaveLength(1);
    expect(images[0]!.kind.type).toBe("image");

    app.destroy();
    root.remove();
  });

  it("shows a toast and no-ops when the clipboard provider skips an oversized image", async () => {
    const toastSpy = vi.spyOn(ui, "showToast").mockImplementation(() => {});

    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = createApp(root, status);

    const content: any = {};
    Object.defineProperty(content, "skippedOversizedImagePng", { value: true });

    const provider = {
      read: vi.fn(async () => content),
      write: vi.fn(async () => {}),
    };
    (app as any).clipboardProviderPromise = Promise.resolve(provider);

    await app.pasteClipboardToSelection();

    expect(app.getDrawingObjects().filter((obj) => obj.kind.type === "image")).toHaveLength(0);
    expect(toastSpy).toHaveBeenCalledWith("Image too large (>5MB). Choose a smaller file.", "warning");

    app.destroy();
    root.remove();
  });

  it("shows a toast and no-ops when a PNG has extremely large dimensions", async () => {
    const toastSpy = vi.spyOn(ui, "showToast").mockImplementation(() => {});

    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = createApp(root, status);

    const pngBytes = createPngHeaderBytes(10_001, 1);
    const provider = {
      read: vi.fn(async () => ({ imagePng: pngBytes })),
      write: vi.fn(async () => {}),
    };
    (app as any).clipboardProviderPromise = Promise.resolve(provider);

    await app.pasteClipboardToSelection();

    expect(app.getDrawingObjects().filter((obj) => obj.kind.type === "image")).toHaveLength(0);
    expect(toastSpy).toHaveBeenCalledWith("Image dimensions too large (10001x1). Choose a smaller image.", "warning");

    app.destroy();
    root.remove();
  });

  it("does not persist image bytes when pasting an image fails to insert a drawing", async () => {
    // 1x1 transparent PNG.
    const pngBytes = new Uint8Array(
      // eslint-disable-next-line no-undef
      Buffer.from(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO7+FeAAAAAASUVORK5CYII=",
        "base64",
      ),
    );

    Object.defineProperty(globalThis, "createImageBitmap", {
      configurable: true,
      value: vi.fn(async () => ({ width: 64, height: 32 })),
    });

    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = createApp(root, status);
    const docAny = app.getDocument() as any;
    docAny.insertDrawing = vi.fn(() => {
      throw new Error("insertDrawing failed");
    });

    const setSpy = vi.spyOn(app.getDrawingImages(), "set");

    const provider = {
      read: vi.fn(async () => ({ imagePng: pngBytes })),
      write: vi.fn(async () => {}),
    };
    (app as any).clipboardProviderPromise = Promise.resolve(provider);

    await app.pasteClipboardToSelection();

    expect(setSpy).not.toHaveBeenCalled();
    expect(app.getDrawingObjects().filter((obj) => obj.kind.type === "image")).toHaveLength(0);

    app.destroy();
    root.remove();
  });

  it("pastes an image even when insertDrawing is unavailable (fallback to setSheetDrawings)", async () => {
    // 1x1 transparent PNG.
    const pngBytes = new Uint8Array(
      // eslint-disable-next-line no-undef
      Buffer.from(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO7+FeAAAAAASUVORK5CYII=",
        "base64",
      ),
    );

    Object.defineProperty(globalThis, "createImageBitmap", {
      configurable: true,
      value: vi.fn(async () => ({ width: 64, height: 32 })),
    });

    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = createApp(root, status);
    const docAny = app.getDocument() as any;
    // Simulate an older DocumentController surface without insertDrawing.
    docAny.insertDrawing = undefined;

    const provider = {
      read: vi.fn(async () => ({ imagePng: pngBytes })),
      write: vi.fn(async () => {}),
    };
    (app as any).clipboardProviderPromise = Promise.resolve(provider);

    await app.pasteClipboardToSelection();

    const images = app.getDrawingObjects().filter((obj) => obj.kind.type === "image");
    expect(images).toHaveLength(1);
    expect(images[0]!.kind.type).toBe("image");

    app.destroy();
    root.remove();
  });

  it("pastes images into the original sheet even if the user switches sheets mid-paste", async () => {
    // 1x1 transparent PNG.
    const pngBytes = new Uint8Array(
      // eslint-disable-next-line no-undef
      Buffer.from(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO7+FeAAAAAASUVORK5CYII=",
        "base64",
      ),
    );

    Object.defineProperty(globalThis, "createImageBitmap", {
      configurable: true,
      value: vi.fn(async () => ({ width: 64, height: 32 })),
    });

    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = createApp(root, status);
    const focusSpy = vi.spyOn(app, "focus");
    const doc: any = app.getDocument();
    const sheet1 = app.getCurrentSheetId();

    // Ensure Sheet2 exists so we can switch away while the clipboard read is still in-flight.
    doc.setCellValue("Sheet2", { row: 0, col: 0 }, "X");

    let resolveRead: ((value: any) => void) | null = null;
    const readPromise = new Promise<any>((resolve) => {
      resolveRead = resolve;
    });
    const provider = {
      read: vi.fn(async () => readPromise),
      write: vi.fn(async () => {}),
    };
    (app as any).clipboardProviderPromise = Promise.resolve(provider);

    const pastePromise = app.pasteClipboardToSelection();

    // Switch sheets while paste is waiting on clipboard bytes.
    app.activateSheet("Sheet2");
    expect(app.getCurrentSheetId()).toBe("Sheet2");
    focusSpy.mockClear();

    resolveRead?.({ imagePng: pngBytes });
    await pastePromise;

    const sheet1Drawings = Array.isArray(doc.getSheetDrawings?.(sheet1)) ? doc.getSheetDrawings(sheet1) : [];
    const sheet2Drawings = Array.isArray(doc.getSheetDrawings?.("Sheet2")) ? doc.getSheetDrawings("Sheet2") : [];
    expect(sheet1Drawings).toHaveLength(1);
    expect(sheet2Drawings).toHaveLength(0);

    // Pasting into a non-active sheet should not disrupt the active sheet's drawing selection.
    const state = app.getDrawingsDebugState();
    expect(state.sheetId).toBe("Sheet2");
    expect(state.drawings).toHaveLength(0);
    expect(state.selectedId).toBe(null);
    expect(app.getSelectedDrawingId()).toBe(null);
    expect(focusSpy).not.toHaveBeenCalled();

    app.destroy();
    root.remove();
  });
});
