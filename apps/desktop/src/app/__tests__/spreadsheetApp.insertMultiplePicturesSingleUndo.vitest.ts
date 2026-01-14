/**
 * @vitest-environment jsdom
 */

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

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

function createApp(root: HTMLElement, status: { activeCell: HTMLElement; selectionRange: HTMLElement; activeValue: HTMLElement }): SpreadsheetApp {
  const app = new SpreadsheetApp(root, status);
  // SpreadsheetApp seeds a demo ChartStore chart in non-collab mode. With canvas charts enabled by
  // default, that chart appears in `getDrawingObjects()` and would interfere with picture-focused
  // tests that assert object counts.
  for (const chart of app.listCharts()) {
    (app as any).chartStore.deleteChart(chart.id);
  }
  return app;
}

describe("SpreadsheetApp insertPicturesFromFiles (multi-file) undo batching", () => {
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

    // Ensure intrinsic size decode doesn't hang in jsdom.
    vi.stubGlobal(
      "createImageBitmap",
      vi.fn(async () => ({ width: 64, height: 32, close: () => {} }) as any),
    );
  });

  it("inserts 3 pictures and creates a single undo step; undo removes all pictures", async () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = createApp(root, status);
    const sheetId = app.getCurrentSheetId();
    const doc = app.getDocument();

    const file1 = new File([new Uint8Array([0x89, 0x50, 0x4e, 0x47])], "a.png", { type: "image/png" });
    const file2 = new File([new Uint8Array([0x89, 0x50, 0x4e, 0x47])], "b.png", { type: "image/png" });
    const file3 = new File([new Uint8Array([0x89, 0x50, 0x4e, 0x47])], "c.png", { type: "image/png" });

    const undoBefore = doc.getStackDepths().undo;
    await app.insertPicturesFromFiles([file1, file2, file3]);

    expect(doc.getStackDepths().undo).toBe(undoBefore + 1);
    expect(doc.undoLabel).toBe("Insert Picture");

    const drawings = doc.getSheetDrawings(sheetId) as any[];
    expect(drawings).toHaveLength(3);
    const objects = app.getDrawingObjects(sheetId).filter((obj) => obj.kind.type === "image");
    expect(objects).toHaveLength(3);
    expect(app.getSelectedDrawingId()).toBe(objects[objects.length - 1]!.id);

    expect(app.undo()).toBe(true);
    const afterUndo = doc.getSheetDrawings(sheetId) as any[];
    expect(afterUndo).toHaveLength(0);

    app.destroy();
    root.remove();
  });
});
