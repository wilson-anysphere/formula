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
        // Default all unknown properties to no-op functions so rendering code can execute.
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

function createRoot(options: { width?: number; height?: number } = {}): { root: HTMLElement; setSize: (w: number, h: number) => void } {
  const root = document.createElement("div");
  root.className = "grid-root";
  root.tabIndex = 0;
  let width = options.width ?? 200;
  let height = options.height ?? 200;
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
  return {
    root,
    setSize: (w: number, h: number) => {
      width = w;
      height = h;
    },
  };
}

function installRafQueue(): { flush: (count?: number) => void } {
  let queue: FrameRequestCallback[] = [];
  let nextId = 1;
  Object.defineProperty(globalThis, "requestAnimationFrame", {
    configurable: true,
    value: (cb: FrameRequestCallback) => {
      queue.push(cb);
      return nextId++;
    },
  });
  Object.defineProperty(globalThis, "cancelAnimationFrame", { configurable: true, value: () => {} });

  return {
    flush: (count = 1) => {
      for (let i = 0; i < count; i += 1) {
        const current = queue;
        queue = [];
        for (const cb of current) cb(0);
      }
    },
  };
}

function dispatchPointerEvent(
  target: EventTarget,
  type: string,
  opts: { clientX: number; clientY: number; pointerId?: number; button?: number; shiftKey?: boolean },
): void {
  const pointerId = opts.pointerId ?? 1;
  const button = opts.button ?? 0;
  const base = {
    bubbles: true,
    cancelable: true,
    clientX: opts.clientX,
    clientY: opts.clientY,
    pointerId,
    button,
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

describe.each(["legacy", "shared"] as const)("SpreadsheetApp drawings auto-scroll drag (%s)", (gridMode) => {
  let raf: ReturnType<typeof installRafQueue>;

  afterEach(() => {
    if (priorGridMode === undefined) delete process.env.DESKTOP_GRID_MODE;
    else process.env.DESKTOP_GRID_MODE = priorGridMode;
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
  });

  beforeEach(() => {
    priorGridMode = process.env.DESKTOP_GRID_MODE;
    process.env.DESKTOP_GRID_MODE = gridMode;
    document.body.innerHTML = "";

    // Node 22 ships an experimental `localStorage` global that errors unless configured via flags.
    const storage = createInMemoryLocalStorage();
    Object.defineProperty(globalThis, "localStorage", { configurable: true, value: storage });
    Object.defineProperty(window, "localStorage", { configurable: true, value: storage });
    storage.clear();

    raf = installRafQueue();

    // jsdom lacks a real canvas implementation; SpreadsheetApp expects a 2D context.
    Object.defineProperty(HTMLCanvasElement.prototype, "getContext", {
      configurable: true,
      value: () => createMockCanvasContext(),
    });

    // jsdom doesn't ship ResizeObserver by default.
    (globalThis as any).ResizeObserver = class {
      observe() {}
      disconnect() {}
    };
  });

  it("auto-scrolls the grid while dragging a drawing near the viewport edge", () => {
    const { root } = createRoot({ width: 200, height: 200 });
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    // In shared-grid mode, drawing drags/resizes are owned by the dedicated DrawingInteractionController.
    // Without it, SpreadsheetApp only provides capture-based drawing selection (no dragging).
    const app = new SpreadsheetApp(root, status, { enableDrawingInteractions: gridMode === "shared" });
    expect(app.getGridMode()).toBe(gridMode);

    const drawingCanvas = root.querySelector<HTMLCanvasElement>('[data-testid="drawing-layer-canvas"]');
    expect(drawingCanvas).toBeTruthy();

    const objects: DrawingObject[] = [
      {
        id: 1,
        kind: { type: "shape", label: "box" },
        zOrder: 0,
        anchor: {
          type: "absolute",
          pos: { xEmu: pxToEmu(0), yEmu: pxToEmu(0) },
          size: { cx: pxToEmu(80), cy: pxToEmu(80) },
        },
      },
    ];
    const doc = app.getDocument() as any;
    doc.setSheetDrawings(app.getCurrentSheetId(), objects);
    raf.flush(5);

    const selectionCanvas = (app as any).selectionCanvas as HTMLCanvasElement;
    expect(selectionCanvas).toBeTruthy();

    const viewport = app.getDrawingInteractionViewport();
    const headerOffsetX = Number.isFinite(viewport.headerOffsetX) ? Math.max(0, viewport.headerOffsetX!) : 0;
    const headerOffsetY = Number.isFinite(viewport.headerOffsetY) ? Math.max(0, viewport.headerOffsetY!) : 0;

    // Start dragging inside the drawing.
    dispatchPointerEvent(selectionCanvas, "pointerdown", {
      pointerId: 1,
      button: 0,
      clientX: headerOffsetX + 10,
      clientY: headerOffsetY + 10,
    });

    // Move pointer near the bottom-right edge to kick off auto-scroll.
    dispatchPointerEvent(selectionCanvas, "pointermove", {
      pointerId: 1,
      clientX: viewport.width - 1,
      clientY: viewport.height - 1,
    });

    const beforeScroll = app.getScroll();
    const beforeAnchor = (app.getDrawingObjects(app.getCurrentSheetId())[0]!.anchor as any).pos as { xEmu: number; yEmu: number };

    // Run one auto-scroll frame.
    raf.flush(1);

    const afterScroll = app.getScroll();
    const afterAnchor = (app.getDrawingObjects(app.getCurrentSheetId())[0]!.anchor as any).pos as { xEmu: number; yEmu: number };

    expect(afterScroll.x).toBeGreaterThan(beforeScroll.x);
    expect(afterScroll.y).toBeGreaterThan(beforeScroll.y);

    const deltaScrollX = afterScroll.x - beforeScroll.x;
    const deltaScrollY = afterScroll.y - beforeScroll.y;
    const zoom = app.getZoom();
    expect(afterAnchor.xEmu - beforeAnchor.xEmu).toBeCloseTo(pxToEmu(deltaScrollX / zoom), 5);
    expect(afterAnchor.yEmu - beforeAnchor.yEmu).toBeCloseTo(pxToEmu(deltaScrollY / zoom), 5);

    app.destroy();
    root.remove();
  });
});
