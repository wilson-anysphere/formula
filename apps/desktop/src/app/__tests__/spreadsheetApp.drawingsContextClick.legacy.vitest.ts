/**
 * @vitest-environment jsdom
 */

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { pxToEmu } from "../../drawings/overlay";
import type { DrawingObject } from "../../drawings/types";
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
  root.getBoundingClientRect = vi.fn(
    () =>
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
      }) as any,
  );
  document.body.appendChild(root);
  return root;
}

function createPointerLikeMouseEvent(
  type: string,
  options: {
    clientX: number;
    clientY: number;
    button: number;
    ctrlKey?: boolean;
    metaKey?: boolean;
    pointerId?: number;
    pointerType?: string;
  },
): MouseEvent {
  const event = new MouseEvent(type, {
    bubbles: true,
    cancelable: true,
    button: options.button,
    clientX: options.clientX,
    clientY: options.clientY,
    ctrlKey: options.ctrlKey,
    metaKey: options.metaKey,
  });
  Object.defineProperty(event, "pointerId", { configurable: true, value: options.pointerId ?? 1 });
  Object.defineProperty(event, "pointerType", { configurable: true, value: options.pointerType ?? "mouse" });
  return event;
}

describe("SpreadsheetApp drawings context-click behavior (legacy grid)", () => {
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

    // jsdom (as used by vitest) does not provide PointerEvent in all environments.
    // SpreadsheetApp only relies on MouseEvent fields (clientX/Y, button) for drawing hit tests.
    if (typeof (globalThis as any).PointerEvent === "undefined") {
      Object.defineProperty(globalThis, "PointerEvent", { configurable: true, value: MouseEvent });
    }

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

  it("treats Ctrl+click as a context-click on macOS (does not start drawing drag/resize)", () => {
    const originalPlatform = navigator.platform;
    const restorePlatform = () => {
      try {
        Object.defineProperty(navigator, "platform", { configurable: true, value: originalPlatform });
      } catch {
        // ignore
      }
    };

    try {
      Object.defineProperty(navigator, "platform", { configurable: true, value: "MacIntel" });
    } catch {
      // If the runtime doesn't allow stubbing `navigator.platform`, skip the test.
      restorePlatform();
      return;
    }

    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };
    const app = new SpreadsheetApp(root, status);
    expect(app.getGridMode()).toBe("legacy");

    const drawing: DrawingObject = {
      id: 1,
      kind: { type: "image", imageId: "img-1" },
      anchor: {
        type: "absolute",
        pos: { xEmu: pxToEmu(0), yEmu: pxToEmu(0) },
        size: { cx: pxToEmu(100), cy: pxToEmu(100) },
      },
      zOrder: 0,
    };
    app.setDrawingObjects([drawing]);

    const bubbled = vi.fn();
    root.addEventListener("pointerdown", bubbled);

    const rowHeaderWidth = (app as any).rowHeaderWidth as number;
    const colHeaderHeight = (app as any).colHeaderHeight as number;
    const down = createPointerLikeMouseEvent("pointerdown", {
      clientX: rowHeaderWidth + 10,
      clientY: colHeaderHeight + 10,
      button: 0,
      ctrlKey: true,
      metaKey: false,
    });
    root.dispatchEvent(down);

    expect(app.getSelectedDrawingId()).toBe(1);
    expect((app as any).drawingGesture).toBeNull();
    expect(down.defaultPrevented).toBe(false);
    expect(bubbled).toHaveBeenCalledTimes(1);

    app.destroy();
    root.remove();
    restorePlatform();
  });

  it("keeps the active cell stable when right-clicking a drawing selection handle", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };
    const app = new SpreadsheetApp(root, status);
    expect(app.getGridMode()).toBe("legacy");

    // Move the active cell away so we can detect unwanted selection changes.
    app.activateCell({ row: 5, col: 5 }, { scrollIntoView: false, focus: false });
    const beforeActive = app.getActiveCell();

    const drawing: DrawingObject = {
      id: 1,
      kind: { type: "image", imageId: "img-1" },
      anchor: {
        type: "absolute",
        pos: { xEmu: pxToEmu(100), yEmu: pxToEmu(100) },
        size: { cx: pxToEmu(100), cy: pxToEmu(100) },
      },
      zOrder: 0,
    };
    app.setDrawingObjects([drawing]);
    app.selectDrawingById(1);

    const bubbled = vi.fn();
    root.addEventListener("pointerdown", bubbled);

    const rowHeaderWidth = (app as any).rowHeaderWidth as number;
    const colHeaderHeight = (app as any).colHeaderHeight as number;

    // Right-click slightly outside the drawing bounds, but within the top-left resize handle square.
    const down = createPointerLikeMouseEvent("pointerdown", {
      clientX: rowHeaderWidth + 100 - 1,
      clientY: colHeaderHeight + 100 - 1,
      button: 2,
    });
    root.dispatchEvent(down);

    expect(app.getSelectedDrawingId()).toBe(1);
    expect(app.getActiveCell()).toEqual(beforeActive);
    expect(down.defaultPrevented).toBe(false);
    expect(bubbled).toHaveBeenCalledTimes(1);

    app.destroy();
    root.remove();
  });
});

