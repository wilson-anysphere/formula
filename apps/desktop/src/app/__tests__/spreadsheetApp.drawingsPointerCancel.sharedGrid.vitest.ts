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
      drawImage: noop,
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

function createRoot(options: { width?: number; height?: number } = {}): HTMLElement {
  const root = document.createElement("div");
  root.tabIndex = 0;
  const width = options.width ?? 800;
  const height = options.height ?? 600;
  root.getBoundingClientRect = () =>
    ({
      width,
      height,
      left: 0,
      top: 0,
      right: width,
      bottom: height,
      x: 0,
      y: 0,
      toJSON: () => {},
    }) as any;
  // JSDOM doesn't always implement pointer capture APIs.
  (root as any).setPointerCapture ??= () => {};
  (root as any).releasePointerCapture ??= () => {};
  document.body.appendChild(root);
  return root;
}

function dispatchPointerEvent(
  target: EventTarget,
  type: string,
  opts: { clientX: number; clientY: number; pointerId?: number; buttons?: number; button?: number },
): void {
  const pointerId = opts.pointerId ?? 1;
  const button = opts.button ?? 0;
  const buttons = opts.buttons ?? 0;
  const base = { bubbles: true, cancelable: true, clientX: opts.clientX, clientY: opts.clientY, pointerId, button, buttons };
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

describe("SpreadsheetApp drawings pointercancel (shared grid)", () => {
  afterEach(() => {
    if (priorGridMode === undefined) delete process.env.DESKTOP_GRID_MODE;
    else process.env.DESKTOP_GRID_MODE = priorGridMode;
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
  });

  beforeEach(() => {
    priorGridMode = process.env.DESKTOP_GRID_MODE;
    process.env.DESKTOP_GRID_MODE = "shared";

    document.body.innerHTML = "";

    // Node 22 ships an experimental `localStorage` global that errors unless configured via flags.
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

    // jsdom doesn't currently ship PointerEvent; provide a minimal polyfill so
    // we can exercise pointer-driven drawing interactions.
    if (!(globalThis as any).PointerEvent) {
      (globalThis as any).PointerEvent = class PointerEvent extends MouseEvent {
        pointerId: number;
        constructor(eventType: string, init: any = {}) {
          super(eventType, init);
          this.pointerId = Number(init.pointerId ?? 0);
        }
      };
    }
  });

  it("commits a drawing drag gesture on pointercancel when drawing interactions are enabled", () => {
    const root = createRoot({ width: 800, height: 600 });
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status, { enableDrawingInteractions: true });
    expect(app.getGridMode()).toBe("shared");

    const sheetId = app.getCurrentSheetId();
    const doc: any = app.getDocument() as any;

    const startXEmu = 0;
    const startYEmu = 0;
    const drawing: DrawingObject = {
      id: 1,
      kind: { type: "shape", label: "box" },
      zOrder: 0,
      anchor: {
        type: "absolute",
        pos: { xEmu: startXEmu, yEmu: startYEmu },
        size: { cx: pxToEmu(100), cy: pxToEmu(100) },
      },
    };
    doc.setSheetDrawings(sheetId, [drawing], { label: "Insert Picture" });
    // Ensure hit testing sees the latest document state immediately.
    (app as any).drawingObjectsCache = null;
    (app as any).drawingHitTestIndex = null;
    (app as any).drawingHitTestIndexObjects = null;

    const historyBefore = doc.history.length;

    const viewport = app.getDrawingInteractionViewport();
    const headerOffsetX = Number.isFinite(viewport.headerOffsetX) ? Math.max(0, viewport.headerOffsetX!) : 0;
    const headerOffsetY = Number.isFinite(viewport.headerOffsetY) ? Math.max(0, viewport.headerOffsetY!) : 0;

    const selectionCanvas = (app as any).selectionCanvas as HTMLCanvasElement;
    expect(selectionCanvas).toBeTruthy();
    selectionCanvas.getBoundingClientRect = root.getBoundingClientRect;

    const downX = headerOffsetX + 10;
    const downY = headerOffsetY + 10;

    dispatchPointerEvent(selectionCanvas, "pointerdown", { clientX: downX, clientY: downY, pointerId: 1, button: 0, buttons: 1 });
    dispatchPointerEvent(selectionCanvas, "pointermove", {
      clientX: downX + 30,
      clientY: downY + 10,
      pointerId: 1,
      buttons: 1,
    });
    dispatchPointerEvent(selectionCanvas, "pointercancel", { clientX: downX + 30, clientY: downY + 10, pointerId: 1 });

    expect(doc.history.length).toBe(historyBefore + 1);

    const committed = doc.getSheetDrawings(sheetId)[0];
    expect(committed.anchor.type).toBe("absolute");
    expect(committed.anchor.pos.xEmu).not.toBe(startXEmu);
    expect(committed.anchor.pos.yEmu).not.toBe(startYEmu);

    if (typeof doc.undo === "function") {
      expect(doc.undo()).toBe(true);
      const reverted = doc.getSheetDrawings(sheetId)[0];
      expect(reverted.anchor.type).toBe("absolute");
      expect(reverted.anchor.pos.xEmu).toBe(startXEmu);
      expect(reverted.anchor.pos.yEmu).toBe(startYEmu);
    }

    app.destroy();
    root.remove();
  });
});

