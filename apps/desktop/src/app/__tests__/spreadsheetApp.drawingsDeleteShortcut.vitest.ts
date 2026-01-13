/**
 * @vitest-environment jsdom
 */

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { SpreadsheetApp } from "../spreadsheetApp";
import { pxToEmu } from "../../drawings/overlay";

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

describe("SpreadsheetApp drawings delete shortcut", () => {
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

    // DrawingOverlay uses createImageBitmap for image decoding; stub it for jsdom.
    Object.defineProperty(globalThis, "createImageBitmap", {
      configurable: true,
      value: async () => ({}) as any,
    });

    (globalThis as any).ResizeObserver = class {
      observe() {}
      disconnect() {}
    };
  });

  for (const mode of ["shared", "legacy"] as const) {
    it(`removes selected drawings on Delete (${mode} grid)`, () => {
      const prior = process.env.DESKTOP_GRID_MODE;
      process.env.DESKTOP_GRID_MODE = mode;
      try {
        const root = createRoot();
        const status = {
          activeCell: document.createElement("div"),
          selectionRange: document.createElement("div"),
          activeValue: document.createElement("div"),
        };

        const app = new SpreadsheetApp(root, status);
        const sheetId = app.getCurrentSheetId();
        const doc = app.getDocument() as any;

        const imageId = "img-1";
        doc.setImage(imageId, { bytes: new Uint8Array([1, 2, 3]), mimeType: "image/png" });

        doc.setSheetDrawings(sheetId, [
          {
            id: "1",
            kind: { type: "image", imageId },
            anchor: {
              type: "oneCell",
              from: { cell: { row: 0, col: 0 }, offset: { xEmu: 0, yEmu: 0 } },
              size: { cx: pxToEmu(50), cy: pxToEmu(50) },
            },
            zOrder: 0,
          },
        ]);

        (app as any).selectedDrawingId = 1;
        (app as any).drawingOverlay.setSelectedId(1);

        root.dispatchEvent(new KeyboardEvent("keydown", { key: "Delete", bubbles: true, cancelable: true }));

        expect(doc.getSheetDrawings(sheetId)).toHaveLength(0);
        expect((app as any).selectedDrawingId).toBe(null);
        expect(((app as any).drawingOverlay as any).selectedId).toBe(null);

        app.destroy();
        root.remove();
      } finally {
        if (prior === undefined) delete process.env.DESKTOP_GRID_MODE;
        else process.env.DESKTOP_GRID_MODE = prior;
      }
    });
  }
});
