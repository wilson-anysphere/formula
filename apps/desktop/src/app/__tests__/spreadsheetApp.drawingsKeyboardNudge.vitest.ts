/**
 * @vitest-environment jsdom
 */

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { pxToEmu } from "../../drawings/overlay";
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
          // Real pointer events bubble; match that behavior so SpreadsheetApp's root-level
          // pointer listeners see these synthetic events.
          super(type, { bubbles: true, cancelable: true, ...init });
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
      app.activateCell({ row: 5, col: 7 }, { scrollIntoView: false, focus: false });
      const activeBefore = app.getActiveCell();
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
      expect(app.getActiveCell()).toEqual(activeBefore);

      root.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape", bubbles: true, cancelable: true }));
      expect(app.getSelectedDrawingId()).toBeNull();
      expect(((app as any).drawingOverlay as any).selectedId).toBe(null);
      expect(((app as any).drawingInteractionController as any).selectedId).toBe(null);
      expect(app.getActiveCell()).toEqual(activeBefore);

      app.destroy();
      root.remove();
    } finally {
      if (prior === undefined) delete process.env.DESKTOP_GRID_MODE;
      else process.env.DESKTOP_GRID_MODE = prior;
    }
  });

  it("does not intercept grid arrow-key navigation when no drawing is selected", () => {
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
      app.activateCell({ row: 2, col: 3 }, { scrollIntoView: false, focus: false });
      const before = app.getActiveCell();

      root.dispatchEvent(new KeyboardEvent("keydown", { key: "ArrowRight", bubbles: true, cancelable: true }));

      expect(app.getActiveCell()).toEqual({ row: before.row, col: before.col + 1 });

      app.destroy();
      root.remove();
    } finally {
      if (prior === undefined) delete process.env.DESKTOP_GRID_MODE;
      else process.env.DESKTOP_GRID_MODE = prior;
    }
  });

  it("preserves non-numeric string drawing ids when nudging (avoids hashing ids into the document)", () => {
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
          id: "drawing_foo",
          kind: { type: "shape", label: "rect" },
          anchor: {
            type: "oneCell",
            from: { cell: { row: 0, col: 0 }, offset: { xEmu: 0, yEmu: 0 } },
            size: { cx: pxToEmu(10), cy: pxToEmu(10) },
          },
          zOrder: 0,
        },
      ]);

      const ui = convertDocumentSheetDrawingsToUiDrawingObjects(doc.getSheetDrawings(sheetId), { sheetId })[0]!;
      app.selectDrawing(ui.id);

      root.dispatchEvent(new KeyboardEvent("keydown", { key: "ArrowRight", bubbles: true, cancelable: true }));

      const updated = doc.getSheetDrawings(sheetId)[0];
      expect(updated.id).toBe("drawing_foo");
      expect(updated.anchor.type).toBe("oneCell");
      expect(updated.anchor.from.offset.xEmu).toBe(pxToEmu(1));

      app.destroy();
      root.remove();
    } finally {
      if (prior === undefined) delete process.env.DESKTOP_GRID_MODE;
      else process.env.DESKTOP_GRID_MODE = prior;
    }
  });

  it("nudges twoCell anchors by shifting both points", () => {
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
          kind: { type: "unknown", label: "twoCell" },
          anchor: {
            type: "twoCell",
            from: { cell: { row: 0, col: 0 }, offset: { xEmu: 0, yEmu: 0 } },
            to: { cell: { row: 1, col: 1 }, offset: { xEmu: 0, yEmu: 0 } },
          },
          zOrder: 0,
        },
      ]);

      app.selectDrawing(1);

      root.dispatchEvent(new KeyboardEvent("keydown", { key: "ArrowRight", bubbles: true, cancelable: true }));
      const updated = doc.getSheetDrawings(sheetId)[0];
      expect(updated.anchor.type).toBe("twoCell");
      expect(updated.anchor.from.offset.xEmu).toBe(pxToEmu(1));
      expect(updated.anchor.to.offset.xEmu).toBe(pxToEmu(1));
      expect(updated.anchor.from.offset.yEmu).toBe(0);
      expect(updated.anchor.to.offset.yEmu).toBe(0);

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
      app.activateCell({ row: 3, col: 4 }, { scrollIntoView: false, focus: false });
      const activeBefore = app.getActiveCell();
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
      // Anchors are persisted in integer EMUs, so expect rounding at sub-pixel steps.
      expect(updated.anchor.pos.xEmu).toBe(Math.round(pxToEmu(0.5)));
      expect(updated.anchor.pos.yEmu).toBe(0);
      expect(app.getActiveCell()).toEqual(activeBefore);

      root.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape", bubbles: true, cancelable: true }));
      expect(app.getSelectedDrawingId()).toBeNull();
      expect(((app as any).drawingInteractionController as any).selectedId).toBe(null);
      expect(app.getActiveCell()).toEqual(activeBefore);

      app.destroy();
      root.remove();
    } finally {
      if (prior === undefined) delete process.env.DESKTOP_GRID_MODE;
      else process.env.DESKTOP_GRID_MODE = prior;
    }
  });

  it("uses 10px screen steps with Shift when nudging (shared grid zoom)", () => {
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

      root.dispatchEvent(new KeyboardEvent("keydown", { key: "ArrowRight", shiftKey: true, bubbles: true, cancelable: true }));
      const updated = doc.getSheetDrawings(sheetId)[0];
      expect(updated.anchor.type).toBe("absolute");
      // 10px screen movement at 2x zoom => 5px in sheet space.
      expect(updated.anchor.pos.xEmu).toBeCloseTo(pxToEmu(5));

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
      // Ensure the app's in-memory drawing cache reflects the seeded DocumentController state.
      app.setDrawingObjects(doc.getSheetDrawings(sheetId));

      const rowHeaderWidth = (app as any).rowHeaderWidth as number;
      const colHeaderHeight = (app as any).colHeaderHeight as number;

      const startClientX = rowHeaderWidth + 10;
      const startClientY = colHeaderHeight + 10;

      // In legacy mode the interaction controller is attached to the root (not the selection canvas),
      // so dispatch events on the root to avoid relying on bubbling semantics in our PointerEvent polyfill.
      root.dispatchEvent(
        new (globalThis as any).PointerEvent("pointerdown", {
          clientX: startClientX,
          clientY: startClientY,
          pointerId: 1,
          buttons: 1,
          bubbles: true,
          cancelable: true,
        }),
      );
      root.dispatchEvent(
        new (globalThis as any).PointerEvent("pointermove", {
          clientX: startClientX + 20,
          clientY: startClientY,
          pointerId: 1,
          buttons: 1,
          bubbles: true,
          cancelable: true,
        }),
      );

      // Drag should have moved the in-memory drawing state.
      const moved = app.getDrawingObjects(sheetId).find((o) => o.id === 1);
      expect(moved?.anchor?.type).toBe("absolute");
      expect((moved?.anchor as any)?.pos?.xEmu).not.toBe(0);

      // Escape should reach the controller's window-level handler and cancel the drag.
      root.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape", bubbles: true, cancelable: true }));

      const cancelled = app.getDrawingObjects(sheetId).find((o) => o.id === 1);
      expect(cancelled?.anchor?.type).toBe("absolute");
      expect((cancelled?.anchor as any)?.pos?.xEmu).toBe(0);

      // Releasing the pointer after cancel should not re-commit the drag.
      root.dispatchEvent(
        new (globalThis as any).PointerEvent("pointerup", {
          clientX: startClientX + 20,
          clientY: startClientY,
          pointerId: 1,
          bubbles: true,
          cancelable: true,
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
