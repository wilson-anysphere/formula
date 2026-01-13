/**
 * @vitest-environment jsdom
 */

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { convertDocumentSheetDrawingsToUiDrawingObjects } from "../../drawings/modelAdapters";
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

    const app = new SpreadsheetApp(root, status);

    const provider = {
      read: vi.fn(async () => ({ imagePng: pngBytes })),
      write: vi.fn(async () => {}),
    };
    (app as any).clipboardProviderPromise = Promise.resolve(provider);

    await app.pasteClipboardToSelection();

    const sheetId = app.getCurrentSheetId();
    const docAny = app.getDocument() as any;
    const drawings = docAny.getSheetDrawings?.(sheetId) ?? [];
    expect(drawings).toHaveLength(1);

    const raw = drawings[0] as any;
    expect(raw?.kind?.type).toBe("image");

    const uiObjects = convertDocumentSheetDrawingsToUiDrawingObjects(drawings);
    expect(uiObjects).toHaveLength(1);
    expect((app as any).selectedDrawingId).toBe(uiObjects[0]!.id);

    const imageId = raw?.kind?.imageId;
    expect(typeof imageId).toBe("string");
    const stored = docAny.getImage?.(imageId);
    expect(stored?.mimeType).toBe("image/png");
    expect(stored?.bytes).toBeInstanceOf(Uint8Array);

    app.destroy();
    root.remove();
  });
});
