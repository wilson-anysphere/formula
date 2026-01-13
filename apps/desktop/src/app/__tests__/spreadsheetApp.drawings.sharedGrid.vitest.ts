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
  root.className = "grid-root";
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

function dispatchPointer(target: EventTarget, type: string, point: { x: number; y: number; pointerId?: number }): void {
  const event = new Event(type, { bubbles: true, cancelable: true }) as any;
  const pointerId = point.pointerId ?? 1;
  Object.defineProperties(event, {
    clientX: { value: point.x, configurable: true },
    clientY: { value: point.y, configurable: true },
    offsetX: { value: point.x, configurable: true },
    offsetY: { value: point.y, configurable: true },
    pointerId: { value: pointerId, configurable: true },
    pointerType: { value: "mouse", configurable: true },
    button: { value: 0, configurable: true },
    buttons: { value: 1, configurable: true },
  });
  target.dispatchEvent(event);
}

describe("SpreadsheetApp shared-grid drawings interaction", () => {
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

    // CanvasGridRenderer schedules renders via requestAnimationFrame; ensure it exists in jsdom.
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

  it("prevents shared-grid selection changes when clicking on a drawing, while allowing empty clicks to select cells", () => {
    const prior = process.env.DESKTOP_GRID_MODE;
    process.env.DESKTOP_GRID_MODE = "shared";
    try {
      const root = createRoot();
      const status = {
        activeCell: document.createElement("div"),
        selectionRange: document.createElement("div"),
        activeValue: document.createElement("div"),
      };

      const app = new SpreadsheetApp(root, status, { enableDrawingInteractions: true });
      expect(app.getGridMode()).toBe("shared");

      const selectionCanvas = (app as any).selectionCanvas as HTMLCanvasElement;
      selectionCanvas.getBoundingClientRect = root.getBoundingClientRect as any;

      // Inject a simple drawing object so the interaction controller has something to hit-test.
      const doc = (app as any).document as any;
      doc.getSheetDrawings = () => [
        {
          id: 1,
          kind: { type: "shape", label: "Box" },
          anchor: {
            type: "absolute",
            pos: { xEmu: pxToEmu(40), yEmu: pxToEmu(200) },
            size: { cx: pxToEmu(120), cy: pxToEmu(60) },
          },
          zOrder: 0,
        },
      ];

      const syncSpy = vi.spyOn(app as any, "syncSelectionFromSharedGrid");
      const initialSelection = (app as any).selection?.active ? { ...(app as any).selection.active } : null;

      // Click inside the drawing bounds (drawings are rendered under the headers).
      dispatchPointer(selectionCanvas, "pointerdown", { x: 100, y: 240, pointerId: 1 });
      dispatchPointer(selectionCanvas, "pointerup", { x: 100, y: 240, pointerId: 1 });

      expect(syncSpy).toHaveBeenCalledTimes(0);
      expect((app as any).selection.active).toEqual(initialSelection);

      // Clicking on empty cell space should still route to DesktopSharedGrid selection handlers.
      dispatchPointer(selectionCanvas, "pointerdown", { x: 100, y: 50, pointerId: 2 });
      dispatchPointer(selectionCanvas, "pointerup", { x: 100, y: 50, pointerId: 2 });

      expect(syncSpy.mock.calls.length).toBeGreaterThan(0);
      expect((app as any).selection.active).toEqual({ row: 1, col: 0 });

      app.destroy();
      root.remove();
    } finally {
      if (prior === undefined) delete process.env.DESKTOP_GRID_MODE;
      else process.env.DESKTOP_GRID_MODE = prior;
    }
  });
});
