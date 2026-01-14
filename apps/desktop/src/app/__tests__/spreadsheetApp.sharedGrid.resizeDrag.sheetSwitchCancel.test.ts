/**
 * @vitest-environment jsdom
 */

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { SpreadsheetApp } from "../spreadsheetApp";

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

function dispatchPointerEvent(
  target: EventTarget,
  type: string,
  opts: { clientX: number; clientY: number; pointerId?: number; button?: number },
): void {
  const pointerId = opts.pointerId ?? 1;
  const button = opts.button ?? 0;
  const base = { bubbles: true, cancelable: true, clientX: opts.clientX, clientY: opts.clientY, pointerId, button };
  const event =
    typeof (globalThis as any).PointerEvent === "function"
      ? new (globalThis as any).PointerEvent(type, base)
      : (() => {
          const e = new MouseEvent(type, base);
          Object.assign(e, { pointerId });
          return e;
        })();
  target.dispatchEvent(event);
}

describe("SpreadsheetApp shared-grid axis resize sheet switching", () => {
  beforeEach(() => {
    priorGridMode = process.env.DESKTOP_GRID_MODE;
    process.env.DESKTOP_GRID_MODE = "shared";

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

  afterEach(() => {
    if (priorGridMode === undefined) delete process.env.DESKTOP_GRID_MODE;
    else process.env.DESKTOP_GRID_MODE = priorGridMode;
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
  });

  it("cancels an in-progress shared-grid column resize drag when switching sheets (prevents cross-sheet commits)", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    expect(app.getGridMode()).toBe("shared");

    const doc: any = app.getDocument();
    // Ensure Sheet2 exists and has a non-default column width so a resize pointerup would
    // trigger a sheet-view mutation if the gesture is not canceled.
    doc.setCellValue("Sheet2", { row: 0, col: 0 }, "X");
    doc.setColWidth("Sheet2", 0, 130, { label: "Seed" });

    const setColWidthSpy = vi.spyOn(doc, "setColWidth");
    setColWidthSpy.mockClear();

    const sharedGrid = (app as any).sharedGrid as any;
    expect(sharedGrid).toBeTruthy();
    const renderer = sharedGrid.renderer;

    // Grid index 1 is the first data column (doc col 0) since shared-grid mode includes 1 header col.
    const index = 1;
    const startSize = renderer.getColWidth(index);

    // Simulate an in-progress resize drag on Sheet1.
    sharedGrid.resizePointerId = 42;
    sharedGrid.resizeDrag = { kind: "col", index, startClient: 0, startSize };

    // Switch sheets before the pointerup arrives. This should cancel the in-progress resize drag so
    // the later pointerup cannot trigger an axis-size commit on Sheet2.
    app.activateSheet("Sheet2");
    expect(sharedGrid.resizePointerId).toBeNull();
    expect(sharedGrid.resizeDrag).toBeNull();

    const selectionCanvas = (app as any).selectionCanvas as HTMLCanvasElement;
    dispatchPointerEvent(selectionCanvas, "pointerup", { clientX: 0, clientY: 0, pointerId: 42 });

    expect(setColWidthSpy).not.toHaveBeenCalled();

    app.destroy();
    root.remove();
  });
});

