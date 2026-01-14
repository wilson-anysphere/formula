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

describe("SpreadsheetApp legacy drawings pointercancel", () => {
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

  it("abandons a legacy drawing gesture on pointercancel without committing to the DocumentController", () => {
    const root = createRoot({ width: 800, height: 600 });
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    // Disable DrawingInteractionController so we exercise SpreadsheetApp's legacy drawing gesture state machine.
    const app = new SpreadsheetApp(root, status, { enableDrawingInteractions: false });

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

    const historyBefore = doc.history.length;

    const selectionCanvas = (app as any).selectionCanvas as HTMLCanvasElement;
    const rowHeaderWidth = (app as any).rowHeaderWidth as number;
    const colHeaderHeight = (app as any).colHeaderHeight as number;

    // Start dragging inside the drawing bounds.
    const downX = rowHeaderWidth + 10;
    const downY = colHeaderHeight + 10;
    dispatchPointerEvent(selectionCanvas, "pointerdown", { clientX: downX, clientY: downY, pointerId: 1, button: 0, buttons: 1 });

    // Move it slightly so there is a preview delta.
    dispatchPointerEvent(selectionCanvas, "pointermove", {
      clientX: downX + 30,
      clientY: downY + 10,
      pointerId: 1,
      buttons: 1,
    });

    const previewAnchor: any = app.getDrawingObjects(sheetId)[0]!.anchor;
    expect(previewAnchor.type).toBe("absolute");
    expect(previewAnchor.pos.xEmu).not.toBe(startXEmu);

    // Cancel the gesture (e.g. OS/browser cancels pointer capture).
    dispatchPointerEvent(selectionCanvas, "pointercancel", { clientX: downX + 30, clientY: downY + 10, pointerId: 1 });

    // The document should not be mutated / no new undo step created.
    expect(doc.history.length).toBe(historyBefore);
    const committed = doc.getSheetDrawings(sheetId)[0];
    expect(committed.anchor.type).toBe("absolute");
    expect(committed.anchor.pos.xEmu).toBe(startXEmu);
    expect(committed.anchor.pos.yEmu).toBe(startYEmu);

    // The live preview should revert to the persisted document snapshot.
    const afterCancel: any = app.getDrawingObjects(sheetId)[0]!.anchor;
    expect(afterCancel.type).toBe("absolute");
    expect(afterCancel.pos.xEmu).toBe(startXEmu);
    expect(afterCancel.pos.yEmu).toBe(startYEmu);

    app.destroy();
    root.remove();
  });

  it("abandons a legacy drawing resize gesture on pointercancel without committing to the DocumentController", () => {
    const root = createRoot({ width: 800, height: 600 });
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    // Disable DrawingInteractionController so we exercise SpreadsheetApp's legacy drawing gesture state machine.
    const app = new SpreadsheetApp(root, status, { enableDrawingInteractions: false });

    const sheetId = app.getCurrentSheetId();
    const doc: any = app.getDocument() as any;

    const startCx = pxToEmu(100);
    const startCy = pxToEmu(100);
    const drawing: DrawingObject = {
      id: 1,
      kind: { type: "shape", label: "box" },
      zOrder: 0,
      anchor: {
        type: "absolute",
        pos: { xEmu: 0, yEmu: 0 },
        size: { cx: startCx, cy: startCy },
      },
    };
    doc.setSheetDrawings(sheetId, [drawing], { label: "Insert Picture" });

    const historyBefore = doc.history.length;

    const selectionCanvas = (app as any).selectionCanvas as HTMLCanvasElement;
    const rowHeaderWidth = (app as any).rowHeaderWidth as number;
    const colHeaderHeight = (app as any).colHeaderHeight as number;

    // Start resizing from the bottom-right handle (se). Handles are centered on corners, so landing
    // exactly on the corner should begin a resize gesture.
    const downX = rowHeaderWidth + 100;
    const downY = colHeaderHeight + 100;
    dispatchPointerEvent(selectionCanvas, "pointerdown", { clientX: downX, clientY: downY, pointerId: 1, button: 0, buttons: 1 });

    // Resize it slightly so there is a preview delta.
    dispatchPointerEvent(selectionCanvas, "pointermove", {
      clientX: downX + 20,
      clientY: downY + 30,
      pointerId: 1,
      buttons: 1,
    });

    const previewAnchor: any = app.getDrawingObjects(sheetId)[0]!.anchor;
    expect(previewAnchor.type).toBe("absolute");
    expect(previewAnchor.size.cx).not.toBe(startCx);
    expect(previewAnchor.size.cy).not.toBe(startCy);

    // Cancel the gesture (e.g. OS/browser cancels pointer capture).
    dispatchPointerEvent(selectionCanvas, "pointercancel", { clientX: downX + 20, clientY: downY + 30, pointerId: 1 });

    // The document should not be mutated / no new undo step created.
    expect(doc.history.length).toBe(historyBefore);
    const committed = doc.getSheetDrawings(sheetId)[0];
    expect(committed.anchor.type).toBe("absolute");
    expect(committed.anchor.size.cx).toBe(startCx);
    expect(committed.anchor.size.cy).toBe(startCy);

    // The live preview should revert to the persisted document snapshot.
    const afterCancel: any = app.getDrawingObjects(sheetId)[0]!.anchor;
    expect(afterCancel.type).toBe("absolute");
    expect(afterCancel.size.cx).toBe(startCx);
    expect(afterCancel.size.cy).toBe(startCy);

    app.destroy();
    root.remove();
  });
});
