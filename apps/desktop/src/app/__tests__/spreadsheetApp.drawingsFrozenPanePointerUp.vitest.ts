/**
 * @vitest-environment jsdom
 */

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { anchorToRectPx, effectiveScrollForAnchor, pxToEmu } from "../../drawings/overlay";
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
  let width = options.width ?? 800;
  let height = options.height ?? 600;
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
  // Allow tests to update size if needed.
  (root as any).__setSize = (w: number, h: number) => {
    width = w;
    height = h;
  };
  return root;
}

function dispatchPointerEvent(
  target: EventTarget,
  type: string,
  opts: { clientX: number; clientY: number; pointerId?: number; button?: number; buttons?: number; shiftKey?: boolean },
): void {
  const pointerId = opts.pointerId ?? 1;
  const button = opts.button ?? 0;
  const buttons = opts.buttons ?? 0;
  const base = {
    bubbles: true,
    cancelable: true,
    clientX: opts.clientX,
    clientY: opts.clientY,
    pointerId,
    button,
    buttons,
    shiftKey: Boolean(opts.shiftKey),
  };
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

describe("SpreadsheetApp drawings frozen pane pointerup commit (legacy)", () => {
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

  it("commits the same anchor shown in live preview when pointerup lands in the frozen pane", () => {
    const root = createRoot({ width: 800, height: 600 });
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    // Use legacy draw gesture state machine (no DrawingInteractionController) so this test
    // exercises SpreadsheetApp.onPointerUp's `drawingGesture` branch.
    const app = new SpreadsheetApp(root, status, { enableDrawingInteractions: false });
    expect(app.getGridMode()).toBe("legacy");

    // Freeze the first column so there is a meaningful frozen pane boundary.
    app.freezeFirstColumn();
    // Scroll horizontally to ensure the drag gesture starts in the scrollable pane.
    app.setScroll(150, 0);

    const sheetId = app.getCurrentSheetId();
    const drawing: DrawingObject = {
      id: 1,
      kind: { type: "shape", label: "box" },
      zOrder: 0,
      anchor: {
        type: "oneCell",
        from: { cell: { row: 0, col: 3 }, offset: { xEmu: 0, yEmu: 0 } }, // D1 (scrollable pane)
        size: { cx: pxToEmu(80), cy: pxToEmu(80) },
      },
    };

    const doc = app.getDocument() as any;
    doc.setSheetDrawings(sheetId, [drawing]);
    // Ensure SpreadsheetApp hit testing reads the latest document snapshot.
    (app as any).drawingObjectsCache = null;
    (app as any).drawingHitTestIndex = null;
    (app as any).drawingHitTestIndexObjects = null;

    const viewport = app.getDrawingInteractionViewport();
    const headerOffsetX = Number.isFinite(viewport.headerOffsetX) ? Math.max(0, viewport.headerOffsetX!) : 0;
    const headerOffsetY = Number.isFinite(viewport.headerOffsetY) ? Math.max(0, viewport.headerOffsetY!) : 0;
    const frozenBoundaryX = Number.isFinite(viewport.frozenWidthPx) ? Math.max(headerOffsetX, viewport.frozenWidthPx!) : headerOffsetX;

    // Compute a pointerdown coordinate inside the drawing bounds (in root client coords).
    const rect = anchorToRectPx(drawing.anchor, (app as any).drawingGeom, app.getZoom());
    const scroll = effectiveScrollForAnchor(drawing.anchor, viewport);
    const downX = rect.x - scroll.scrollX + headerOffsetX + 10;
    const downY = rect.y - scroll.scrollY + headerOffsetY + 10;

    expect(downX).toBeGreaterThanOrEqual(frozenBoundaryX);

    const selectionCanvas = (app as any).selectionCanvas as HTMLCanvasElement;

    dispatchPointerEvent(selectionCanvas, "pointerdown", { clientX: downX, clientY: downY, pointerId: 1, button: 0, buttons: 1 });

    // Move into the frozen pane (left of the frozen boundary).
    const moveX = headerOffsetX + 10;
    expect(moveX).toBeLessThan(frozenBoundaryX);
    const moveY = downY;

    dispatchPointerEvent(selectionCanvas, "pointermove", { clientX: moveX, clientY: moveY, pointerId: 1, buttons: 1 });

    const clone = <T,>(value: T): T =>
      typeof structuredClone === "function" ? (structuredClone(value) as T) : (JSON.parse(JSON.stringify(value)) as T);

    const previewAnchor = clone(app.getDrawingObjects(sheetId)[0]!.anchor);

    dispatchPointerEvent(selectionCanvas, "pointerup", { clientX: moveX, clientY: moveY, pointerId: 1 });

    const committed = doc.getSheetDrawings(sheetId)[0];
    expect(committed?.anchor).toEqual(previewAnchor);

    app.destroy();
    root.remove();
  });
});

