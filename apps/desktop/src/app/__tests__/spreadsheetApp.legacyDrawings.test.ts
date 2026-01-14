/**
 * @vitest-environment jsdom
 */

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { drawingObjectToViewportRect } from "../../drawings/hitTest";
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

type PointerEventInitLike = {
  type: string;
  pointerId: number;
  clientX: number;
  clientY: number;
  pointerType?: string;
  button?: number;
};

function createPointerEvent(init: PointerEventInitLike): PointerEvent {
  const event: any = {
    ...init,
    pointerType: init.pointerType ?? "mouse",
    button: init.button ?? 0,
    ctrlKey: false,
    metaKey: false,
    shiftKey: false,
    altKey: false,
    defaultPrevented: false,
    preventDefault() {
      this.defaultPrevented = true;
    },
  };
  return event as PointerEvent;
}

function dispatchPointerEvent(
  target: EventTarget,
  type: string,
  opts: { clientX: number; clientY: number; pointerId?: number; button?: number; pointerType?: string },
): void {
  const pointerId = opts.pointerId ?? 1;
  const button = opts.button ?? 0;
  const pointerType = opts.pointerType ?? "mouse";
  const base = { bubbles: true, cancelable: true, clientX: opts.clientX, clientY: opts.clientY, pointerId, button };
  const event =
    typeof (globalThis as any).PointerEvent === "function"
      ? new (globalThis as any).PointerEvent(type, { ...base, pointerType })
      : (() => {
          const e = new MouseEvent(type, base);
          Object.assign(e, { pointerId, pointerType });
          return e;
        })();
  target.dispatchEvent(event);
}

describe("SpreadsheetApp legacy drawing interactions", () => {
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

  it("click on a drawing selects it and does not move the active cell", () => {
    const prior = process.env.DESKTOP_GRID_MODE;
    process.env.DESKTOP_GRID_MODE = "legacy";
    try {
      const root = createRoot();
      const status = {
        activeCell: document.createElement("div"),
        selectionRange: document.createElement("div"),
        activeValue: document.createElement("div"),
      };

      const app = new SpreadsheetApp(root, status);
      const sheetId = app.getCurrentSheetId();

      // Put the active cell somewhere else so a missed drawing hit-test would move it.
      app.selectRange({ range: { startRow: 2, endRow: 2, startCol: 2, endCol: 2 } }, { scrollIntoView: false, focus: false });
      expect(status.activeCell.textContent).toBe("C3");

      // Insert a simple model drawing into the document (SpreadsheetApp converts it to UI objects).
      (app as any).document.setSheetDrawings(sheetId, [
        {
          id: "drawing1",
          zOrder: 0,
          kind: { type: "shape" },
          anchor: {
            type: "oneCell",
            from: { cell: { row: 0, col: 0 }, offset: { xEmu: pxToEmu(8), yEmu: pxToEmu(8) } },
            size: { cx: pxToEmu(120), cy: pxToEmu(80) },
          },
        },
      ]);
      (app as any).drawingObjectsCache = null;

      const objects = (app as any).listDrawingObjectsForSheet();
      expect(Array.isArray(objects)).toBe(true);
      expect(objects.length).toBe(1);
      const object = objects[0];
      const viewport = (app as any).getDrawingInteractionViewport();
      const geom = (app as any).drawingGeom;
      const rect = drawingObjectToViewportRect(object, viewport, geom);

      // Click near the center (not on a resize handle).
      const clientX = rect.x + rect.width / 2;
      const clientY = rect.y + rect.height / 2;
      (app as any).onPointerDown(
        createPointerEvent({ type: "pointerdown", pointerId: 1, clientX, clientY, pointerType: "mouse", button: 0 }),
      );

      expect((app as any).selectedDrawingId).toBe(object.id);
      expect(status.activeCell.textContent).toBe("C3");

      app.destroy();
      root.remove();
    } finally {
      if (prior === undefined) delete process.env.DESKTOP_GRID_MODE;
      else process.env.DESKTOP_GRID_MODE = prior;
    }
  });

  it("dragging a drawing updates anchor offsets (EMU)", () => {
    const prior = process.env.DESKTOP_GRID_MODE;
    process.env.DESKTOP_GRID_MODE = "legacy";
    try {
      const root = createRoot();
      const status = {
        activeCell: document.createElement("div"),
        selectionRange: document.createElement("div"),
        activeValue: document.createElement("div"),
      };

      const app = new SpreadsheetApp(root, status);
      const sheetId = app.getCurrentSheetId();

      (app as any).document.setSheetDrawings(sheetId, [
        {
          id: "drawing1",
          zOrder: 0,
          kind: { type: "shape" },
          anchor: {
            type: "oneCell",
            from: { cell: { row: 0, col: 0 }, offset: { xEmu: pxToEmu(8), yEmu: pxToEmu(8) } },
            size: { cx: pxToEmu(120), cy: pxToEmu(80) },
          },
        },
      ]);
      (app as any).drawingObjectsCache = null;
      const objects = (app as any).listDrawingObjectsForSheet();
      const object = objects[0];
      expect(object.anchor.type).toBe("oneCell");
      const startAnchor = object.anchor as any;
      const startFromX = startAnchor.from.offset.xEmu;
      const startFromY = startAnchor.from.offset.yEmu;

      const viewport = (app as any).getDrawingInteractionViewport();
      const geom = (app as any).drawingGeom;
      const rect = drawingObjectToViewportRect(object, viewport, geom);
      const startClientX = rect.x + rect.width / 2;
      const startClientY = rect.y + rect.height / 2;

      (app as any).onPointerDown(
        createPointerEvent({
          type: "pointerdown",
          pointerId: 1,
          clientX: startClientX,
          clientY: startClientY,
          pointerType: "mouse",
          button: 0,
        }),
      );

      const dxPx = 10;
      const dyPx = 5;
      (app as any).onPointerMove(
        createPointerEvent({
          type: "pointermove",
          pointerId: 1,
          clientX: startClientX + dxPx,
          clientY: startClientY + dyPx,
          pointerType: "mouse",
          button: 0,
        }),
      );

      (app as any).onPointerUp(
        createPointerEvent({
          type: "pointerup",
          pointerId: 1,
          clientX: startClientX + dxPx,
          clientY: startClientY + dyPx,
          pointerType: "mouse",
          button: 0,
        }),
      );

      const updatedObjects = (app as any).listDrawingObjectsForSheet();
      const updatedObject = updatedObjects.find((o: any) => o.id === object.id);
      const updatedAnchor = updatedObject.anchor as any;

      expect(updatedAnchor.from.offset.xEmu).toBe(startFromX + pxToEmu(dxPx));
      expect(updatedAnchor.from.offset.yEmu).toBe(startFromY + pxToEmu(dyPx));

      app.destroy();
      root.remove();
    } finally {
      if (prior === undefined) delete process.env.DESKTOP_GRID_MODE;
      else process.env.DESKTOP_GRID_MODE = prior;
    }
  });

  it("dragging a drawing via pointer events works in legacy mode (capture handlers do not block)", () => {
    const prior = process.env.DESKTOP_GRID_MODE;
    process.env.DESKTOP_GRID_MODE = "legacy";
    try {
      const root = createRoot();
      const status = {
        activeCell: document.createElement("div"),
        selectionRange: document.createElement("div"),
        activeValue: document.createElement("div"),
      };

      const app = new SpreadsheetApp(root, status);
      const sheetId = app.getCurrentSheetId();

      (app as any).document.setSheetDrawings(sheetId, [
        {
          id: "drawing1",
          zOrder: 0,
          kind: { type: "shape" },
          anchor: {
            type: "oneCell",
            from: { cell: { row: 0, col: 0 }, offset: { xEmu: pxToEmu(8), yEmu: pxToEmu(8) } },
            size: { cx: pxToEmu(120), cy: pxToEmu(80) },
          },
        },
      ]);
      // Ensure the selection-canvas interaction controller sees the latest drawings.
      (app as any).syncSheetDrawings();
      // Ensure listDrawingObjectsForSheet reflects the newly inserted drawings (it caches results).
      (app as any).drawingObjectsCache = null;

      const selectionCanvas = root.querySelector<HTMLCanvasElement>("canvas.grid-canvas--selection");
      expect(selectionCanvas).not.toBeNull();

      const objects = (app as any).listDrawingObjectsForSheet();
      const object = objects[0];
      expect(object).toBeTruthy();
      expect(object.anchor.type).toBe("oneCell");
      const startAnchor = object.anchor as any;
      const startFromX = startAnchor.from.offset.xEmu;
      const startFromY = startAnchor.from.offset.yEmu;

      const viewport = (app as any).getDrawingInteractionViewport();
      const geom = (app as any).drawingGeom;
      const rect = drawingObjectToViewportRect(object, viewport, geom);
      const startClientX = rect.x + rect.width / 2;
      const startClientY = rect.y + rect.height / 2;

      const dxPx = 10;
      const dyPx = 5;

      dispatchPointerEvent(selectionCanvas!, "pointerdown", {
        clientX: startClientX,
        clientY: startClientY,
        pointerId: 1,
        button: 0,
        pointerType: "mouse",
      });
      dispatchPointerEvent(selectionCanvas!, "pointermove", {
        clientX: startClientX + dxPx,
        clientY: startClientY + dyPx,
        pointerId: 1,
        button: 0,
        pointerType: "mouse",
      });
      dispatchPointerEvent(selectionCanvas!, "pointerup", {
        clientX: startClientX + dxPx,
        clientY: startClientY + dyPx,
        pointerId: 1,
        button: 0,
        pointerType: "mouse",
      });

      const updatedObjects = (app as any).listDrawingObjectsForSheet();
      const updatedObject = updatedObjects.find((o: any) => o.id === object.id);
      expect(updatedObject).toBeTruthy();
      const updatedAnchor = updatedObject.anchor as any;

      expect(updatedAnchor.from.offset.xEmu).toBe(startFromX + pxToEmu(dxPx));
      expect(updatedAnchor.from.offset.yEmu).toBe(startFromY + pxToEmu(dyPx));

      app.destroy();
      root.remove();
    } finally {
      if (prior === undefined) delete process.env.DESKTOP_GRID_MODE;
      else process.env.DESKTOP_GRID_MODE = prior;
    }
  });
});
