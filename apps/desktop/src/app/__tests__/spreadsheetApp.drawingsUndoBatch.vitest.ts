/**
 * @vitest-environment jsdom
 */

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { SpreadsheetApp } from "../spreadsheetApp";
import { pxToEmu } from "../../drawings/overlay";

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
  document.body.appendChild(root);
  // Allow tests to update size if needed.
  (root as any).__setSize = (w: number, h: number) => {
    width = w;
    height = h;
  };
  return root;
}

describe("SpreadsheetApp drawings undo batching", () => {
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
        constructor(type: string, init: any = {}) {
          super(type, init);
          this.pointerId = Number(init.pointerId ?? 0);
        }
      };
    }
  });

  it("creates exactly one undo step for a picture drag gesture", () => {
    const root = createRoot({ width: 800, height: 600 });
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    const doc = (app as any).document as any;
    const sheetId = (app as any).sheetId as string;

    // Insert a drawing (absolute anchor so the test doesn't depend on cell geometry).
    const startXEmu = 0;
    const startYEmu = 0;
    const drawing = {
      id: 1,
      kind: { type: "unknown", label: "picture" },
      anchor: { type: "absolute", pos: { xEmu: startXEmu, yEmu: startYEmu }, size: { cx: pxToEmu(100), cy: pxToEmu(100) } },
      zOrder: 0,
    };
    doc.setSheetDrawings(sheetId, [drawing], { label: "Insert Picture" });

    const historyBeforeDrag = doc.history.length;

    const selectionCanvas = (app as any).selectionCanvas as HTMLCanvasElement;
    const rowHeaderWidth = (app as any).rowHeaderWidth as number;
    const colHeaderHeight = (app as any).colHeaderHeight as number;

    const startClientX = rowHeaderWidth + 10;
    const startClientY = colHeaderHeight + 10;

    selectionCanvas.dispatchEvent(
      new PointerEvent("pointerdown", { clientX: startClientX, clientY: startClientY, pointerId: 1, buttons: 1 }),
    );

    // Multiple pointermove events should still produce a single undo entry.
    selectionCanvas.dispatchEvent(
      new PointerEvent("pointermove", { clientX: startClientX + 10, clientY: startClientY, pointerId: 1, buttons: 1 }),
    );
    selectionCanvas.dispatchEvent(
      new PointerEvent("pointermove", { clientX: startClientX + 20, clientY: startClientY + 5, pointerId: 1, buttons: 1 }),
    );
    selectionCanvas.dispatchEvent(
      new PointerEvent("pointermove", { clientX: startClientX + 30, clientY: startClientY + 10, pointerId: 1, buttons: 1 }),
    );

    selectionCanvas.dispatchEvent(
      new PointerEvent("pointerup", { clientX: startClientX + 30, clientY: startClientY + 10, pointerId: 1 }),
    );

    expect(doc.history.length).toBe(historyBeforeDrag + 1);

    const moved = doc.getSheetDrawings(sheetId)[0];
    expect(moved.anchor.type).toBe("absolute");
    expect(moved.anchor.pos.xEmu).not.toBe(startXEmu);

    doc.undo();
    const undone = doc.getSheetDrawings(sheetId)[0];
    expect(undone.anchor.type).toBe("absolute");
    expect(undone.anchor.pos.xEmu).toBe(startXEmu);
    expect(undone.anchor.pos.yEmu).toBe(startYEmu);

    app.destroy();
    root.remove();
  });
});
