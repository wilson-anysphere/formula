/**
 * @vitest-environment jsdom
 */

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { SpreadsheetApp } from "../spreadsheetApp";
import { buildHitTestIndex, hitTestDrawings } from "../../drawings/hitTest";
import type { DrawingObject } from "../../drawings/types";

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

function seedTwoOverlappingDrawings(): DrawingObject[] {
  // Absolute anchors are easiest for tests (they don't depend on grid geometry).
  const anchor = { type: "absolute", pos: { xEmu: 0, yEmu: 0 }, size: { cx: 1_000_000, cy: 1_000_000 } } as const;
  return [
    { id: 1, kind: { type: "image", imageId: "img1" }, anchor, zOrder: 0 },
    { id: 2, kind: { type: "image", imageId: "img2" }, anchor, zOrder: 10 },
  ];
}

describe("SpreadsheetApp drawings arrange z-order", () => {
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

    (globalThis as any).ResizeObserver = class {
      observe() {}
      disconnect() {}
    };
  });

  it("bringSelectedDrawingForward swaps with the next higher z-order and renormalizes", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    const sheetId = app.getCurrentSheetId();

    const before = seedTwoOverlappingDrawings();
    const doc: any = app.getDocument() as any;
    doc.setSheetDrawings(sheetId, before, { label: "Seed Drawings" });
    (app as any).selectedDrawingId = 1;

    const viewport = { scrollX: 0, scrollY: 0, width: 200, height: 200, dpr: 1 };
    const geom = { cellOriginPx: () => ({ x: 0, y: 0 }), cellSizePx: () => ({ width: 100, height: 24 }) };
    expect(String(hitTestDrawings(buildHitTestIndex(before, geom), viewport, 10, 10)?.object.id)).toBe("2");

    const undoDepthBefore = doc.getStackDepths().undo;

    app.bringSelectedDrawingForward();

    const after = (doc.getSheetDrawings(sheetId) ?? []) as DrawingObject[];
    expect((app as any).selectedDrawingId).toBe(1);

    const zOrders = after.map((d) => d.zOrder).sort((a, b) => a - b);
    expect(zOrders).toEqual([0, 1]);

    const d1 = after.find((d) => String((d as any).id) === "1")!;
    const d2 = after.find((d) => String((d as any).id) === "2")!;
    expect(d1.zOrder).toBeGreaterThan(d2.zOrder);
    expect(String(hitTestDrawings(buildHitTestIndex(after, geom), viewport, 10, 10)?.object.id)).toBe("1");

    const undoDepthAfter = doc.getStackDepths().undo;
    expect(undoDepthAfter).toBe(undoDepthBefore + 1);
    expect(app.getUndoRedoState().canUndo).toBe(true);

    app.undo();
    const undone = (doc.getSheetDrawings(sheetId) ?? []) as DrawingObject[];
    const u1 = undone.find((d) => String((d as any).id) === "1")!;
    const u2 = undone.find((d) => String((d as any).id) === "2")!;
    expect(u2.zOrder).toBeGreaterThan(u1.zOrder);
    expect(String(hitTestDrawings(buildHitTestIndex(undone, geom), viewport, 10, 10)?.object.id)).toBe("2");

    app.destroy();
    root.remove();
  });

  it("sendSelectedDrawingBackward swaps with the next lower z-order and renormalizes", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    const sheetId = app.getCurrentSheetId();

    const before = seedTwoOverlappingDrawings();
    const doc: any = app.getDocument() as any;
    doc.setSheetDrawings(sheetId, before, { label: "Seed Drawings" });
    (app as any).selectedDrawingId = 2;

    const viewport = { scrollX: 0, scrollY: 0, width: 200, height: 200, dpr: 1 };
    const geom = { cellOriginPx: () => ({ x: 0, y: 0 }), cellSizePx: () => ({ width: 100, height: 24 }) };
    expect(String(hitTestDrawings(buildHitTestIndex(before, geom), viewport, 10, 10)?.object.id)).toBe("2");

    const undoDepthBefore = doc.getStackDepths().undo;

    app.sendSelectedDrawingBackward();

    const after = (doc.getSheetDrawings(sheetId) ?? []) as DrawingObject[];
    expect((app as any).selectedDrawingId).toBe(2);

    const zOrders = after.map((d) => d.zOrder).sort((a, b) => a - b);
    expect(zOrders).toEqual([0, 1]);

    const d1 = after.find((d) => String((d as any).id) === "1")!;
    const d2 = after.find((d) => String((d as any).id) === "2")!;
    expect(d2.zOrder).toBeLessThan(d1.zOrder);
    expect(String(hitTestDrawings(buildHitTestIndex(after, geom), viewport, 10, 10)?.object.id)).toBe("1");

    const undoDepthAfter = doc.getStackDepths().undo;
    expect(undoDepthAfter).toBe(undoDepthBefore + 1);
    expect(app.getUndoRedoState().canUndo).toBe(true);

    app.undo();
    const undone = (doc.getSheetDrawings(sheetId) ?? []) as DrawingObject[];
    const u1 = undone.find((d) => String((d as any).id) === "1")!;
    const u2 = undone.find((d) => String((d as any).id) === "2")!;
    expect(u2.zOrder).toBeGreaterThan(u1.zOrder);
    expect(String(hitTestDrawings(buildHitTestIndex(undone, geom), viewport, 10, 10)?.object.id)).toBe("2");

    app.destroy();
    root.remove();
  });
});
