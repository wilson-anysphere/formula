/**
 * @vitest-environment jsdom
 */

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { pxToEmu } from "../../drawings/overlay";
import { convertDocumentSheetDrawingsToUiDrawingObjects } from "../../drawings/modelAdapters";
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
  // JSDOM doesn't always implement pointer capture APIs.
  (root as any).setPointerCapture ??= () => {};
  (root as any).releasePointerCapture ??= () => {};
  document.body.appendChild(root);
  return root;
}

describe("SpreadsheetApp drawing interaction commits", () => {
  afterEach(() => {
    if (priorGridMode === undefined) delete process.env.DESKTOP_GRID_MODE;
    else process.env.DESKTOP_GRID_MODE = priorGridMode;
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
  });

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
      value: (cb: FrameRequestCallback) => {
        cb(0);
        return 0;
      },
    });
    Object.defineProperty(globalThis, "cancelAnimationFrame", { configurable: true, value: () => {} });

    Object.defineProperty(window, "devicePixelRatio", { configurable: true, value: 1 });

    Object.defineProperty(HTMLCanvasElement.prototype, "getContext", {
      configurable: true,
      value: () => createMockCanvasContext(),
    });

    (globalThis as any).ResizeObserver = class {
      observe() {}
      disconnect() {}
    };
  });

  it("persists drawing anchor/transform/preserved updates to DocumentController via onInteractionCommit (undoable)", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status, { enableDrawingInteractions: true });
    const sheetId = app.getCurrentSheetId();
    const doc = app.getDocument() as any;

    const rawDrawing = {
      id: "drawing_foo",
      zOrder: 0,
      kind: { type: "shape", label: "Box" },
      anchor: {
        type: "absolute",
        pos: { xEmu: pxToEmu(0), yEmu: pxToEmu(0) },
        size: { cx: pxToEmu(120), cy: pxToEmu(80) },
      },
      preserved: { foo: "before" },
    };
    doc.setSheetDrawings(sheetId, [rawDrawing]);

    const before = convertDocumentSheetDrawingsToUiDrawingObjects(doc.getSheetDrawings(sheetId), { sheetId })[0]!;
    const after = {
      ...before,
      anchor: {
        ...before.anchor,
        // Move it slightly and keep the same size.
        pos: { xEmu: pxToEmu(20), yEmu: pxToEmu(10) },
      },
      transform: { rotationDeg: 45, flipH: false, flipV: false },
      preserved: { foo: "after" },
    };

    const callbacks = (app as any).drawingInteractionCallbacks;
    expect(callbacks?.onInteractionCommit).toBeTypeOf("function");

    callbacks.onInteractionCommit({ kind: "rotate", id: before.id, before, after, objects: [after] });

    const updated = doc.getSheetDrawings(sheetId).find((d: any) => String(d?.id) === "drawing_foo");
    expect(updated?.id).toBe("drawing_foo");
    expect(updated?.zOrder).toBe(0);
    expect(updated?.anchor).toEqual(after.anchor);
    expect(updated?.transform).toEqual(after.transform);
    expect(updated?.preserved).toEqual(after.preserved);

    if (typeof doc.undo === "function") {
      expect(doc.undo()).toBe(true);
      const reverted = doc.getSheetDrawings(sheetId).find((d: any) => String(d?.id) === "drawing_foo");
      expect(reverted?.anchor).toEqual(rawDrawing.anchor);
      expect(reverted?.transform).toBeUndefined();
      expect(reverted?.preserved).toEqual(rawDrawing.preserved);
    }

    app.dispose();
    root.remove();
  });
});
