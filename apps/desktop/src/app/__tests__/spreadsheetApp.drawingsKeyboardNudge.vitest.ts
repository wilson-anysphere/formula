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

describe("SpreadsheetApp drawings keyboard nudging", () => {
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

    // jsdom doesn't always ship PointerEvent. Provide a minimal polyfill so we can
    // exercise pointer-driven drawing interactions (Escape cancel gesture behavior).
    if (!(globalThis as any).PointerEvent) {
      (globalThis as any).PointerEvent = class PointerEvent extends MouseEvent {
        pointerId: number;
        constructor(type: string, init: any = {}) {
          super(type, init);
          this.pointerId = Number(init.pointerId ?? 0);
        }
      };
    }
  });

  it("nudges the selected drawing with arrow keys and clears selection with Escape (legacy grid)", () => {
    const prior = process.env.DESKTOP_GRID_MODE;
    process.env.DESKTOP_GRID_MODE = "legacy";
    try {
      const root = createRoot();
      const status = {
        activeCell: document.createElement("div"),
        selectionRange: document.createElement("div"),
        activeValue: document.createElement("div"),
      };

      const app = new SpreadsheetApp(root, status, { enableDrawingInteractions: true });
      const sheetId = app.getCurrentSheetId();
      const doc = app.getDocument() as any;

      doc.setSheetDrawings(sheetId, [
        {
          id: 1,
          kind: { type: "shape", label: "rect" },
          anchor: {
            type: "oneCell",
            from: { cell: { row: 0, col: 0 }, offset: { xEmu: 0, yEmu: 0 } },
            size: { cx: pxToEmu(10), cy: pxToEmu(10) },
          },
          zOrder: 0,
        },
      ]);

      app.selectDrawing(1);

      root.dispatchEvent(new KeyboardEvent("keydown", { key: "ArrowRight", bubbles: true, cancelable: true }));
      const updated = doc.getSheetDrawings(sheetId)[0];
      expect(updated.anchor.type).toBe("oneCell");
      expect(updated.anchor.from.offset.xEmu).toBe(pxToEmu(1));
      expect(updated.anchor.from.offset.yEmu).toBe(0);

      root.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape", bubbles: true, cancelable: true }));
      expect((app as any).selectedDrawingId).toBeNull();
      expect(((app as any).drawingOverlay as any).selectedId).toBe(null);
      expect(((app as any).drawingInteractionController as any).selectedId).toBe(null);

      app.destroy();
      root.remove();
    } finally {
      if (prior === undefined) delete process.env.DESKTOP_GRID_MODE;
      else process.env.DESKTOP_GRID_MODE = prior;
    }
  });

  it("accounts for zoom when nudging absolute anchors (shared grid)", () => {
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
      app.setZoom(2);
      const sheetId = app.getCurrentSheetId();
      const doc = app.getDocument() as any;

      doc.setSheetDrawings(sheetId, [
        {
          id: 1,
          kind: { type: "unknown", label: "picture" },
          anchor: { type: "absolute", pos: { xEmu: 0, yEmu: 0 }, size: { cx: pxToEmu(10), cy: pxToEmu(10) } },
          zOrder: 0,
        },
      ]);

      app.selectDrawing(1);

      root.dispatchEvent(new KeyboardEvent("keydown", { key: "ArrowRight", bubbles: true, cancelable: true }));
      const updated = doc.getSheetDrawings(sheetId)[0];
      expect(updated.anchor.type).toBe("absolute");
      // Moving by 1 screen px at 2x zoom shifts the underlying sheet position by 0.5px.
      expect(updated.anchor.pos.xEmu).toBeCloseTo(pxToEmu(0.5));
      expect(updated.anchor.pos.yEmu).toBe(0);

      root.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape", bubbles: true, cancelable: true }));
      expect((app as any).selectedDrawingId).toBeNull();
      expect(((app as any).drawingInteractionController as any).selectedId).toBe(null);

      app.destroy();
      root.remove();
    } finally {
      if (prior === undefined) delete process.env.DESKTOP_GRID_MODE;
      else process.env.DESKTOP_GRID_MODE = prior;
    }
  });

  it("does not block Escape from cancelling an active drawing drag gesture", () => {
    const prior = process.env.DESKTOP_GRID_MODE;
    process.env.DESKTOP_GRID_MODE = "legacy";
    try {
      const root = createRoot();
      const status = {
        activeCell: document.createElement("div"),
        selectionRange: document.createElement("div"),
        activeValue: document.createElement("div"),
      };

      const app = new SpreadsheetApp(root, status, { enableDrawingInteractions: true });
      const sheetId = app.getCurrentSheetId();
      const doc = app.getDocument() as any;

      doc.setSheetDrawings(sheetId, [
        {
          id: 1,
          kind: { type: "unknown", label: "picture" },
          anchor: { type: "absolute", pos: { xEmu: 0, yEmu: 0 }, size: { cx: pxToEmu(100), cy: pxToEmu(100) } },
          zOrder: 0,
        },
      ]);

      const selectionCanvas = (app as any).selectionCanvas as HTMLCanvasElement;
      const rowHeaderWidth = (app as any).rowHeaderWidth as number;
      const colHeaderHeight = (app as any).colHeaderHeight as number;

      const startClientX = rowHeaderWidth + 10;
      const startClientY = colHeaderHeight + 10;

      selectionCanvas.dispatchEvent(
        new (globalThis as any).PointerEvent("pointerdown", {
          bubbles: true,
          cancelable: true,
          clientX: startClientX,
          clientY: startClientY,
          pointerId: 1,
          button: 0,
          buttons: 1,
        }),
      );
      selectionCanvas.dispatchEvent(
        new (globalThis as any).PointerEvent("pointermove", {
          bubbles: true,
          cancelable: true,
          clientX: startClientX + 20,
          clientY: startClientY,
          pointerId: 1,
          buttons: 1,
        }),
      );

      // Drag should have moved the in-memory drawing state.
      expect(((app as any).drawingObjectsCache as any)?.objects?.[0]?.anchor?.pos?.xEmu).not.toBe(0);

      // Escape should reach the controller's window-level handler and cancel the drag.
      root.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape", bubbles: true, cancelable: true }));

      expect(((app as any).drawingObjectsCache as any)?.objects?.[0]?.anchor?.pos?.xEmu).toBe(0);

      // Releasing the pointer after cancel should not re-commit the drag.
      selectionCanvas.dispatchEvent(
        new (globalThis as any).PointerEvent("pointerup", {
          bubbles: true,
          cancelable: true,
          clientX: startClientX + 20,
          clientY: startClientY,
          pointerId: 1,
        }),
      );

      expect(doc.getSheetDrawings(sheetId)[0].anchor.pos.xEmu).toBe(0);

      app.destroy();
      root.remove();
    } finally {
      if (prior === undefined) delete process.env.DESKTOP_GRID_MODE;
      else process.env.DESKTOP_GRID_MODE = prior;
    }
  });
});
