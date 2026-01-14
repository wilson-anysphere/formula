/**
 * @vitest-environment jsdom
 */

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

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
  document.body.appendChild(root);
  return root;
}

function dispatchPointerEvent(
  target: EventTarget,
  type: string,
  opts: { clientX: number; clientY: number; pointerId?: number; button?: number },
): void {
  const pointerId = opts.pointerId ?? 1;
  const button = opts.button ?? 0;
  const base = { bubbles: true, cancelable: true, clientX: opts.clientX, clientY: opts.clientY, pointerId, button };
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

describe("SpreadsheetApp shared-grid selection drag sheet switching", () => {
  beforeEach(() => {
    priorGridMode = process.env.DESKTOP_GRID_MODE;
    process.env.DESKTOP_GRID_MODE = "shared";

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

  afterEach(() => {
    if (priorGridMode === undefined) delete process.env.DESKTOP_GRID_MODE;
    else process.env.DESKTOP_GRID_MODE = priorGridMode;
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
  });

  it("cancels an in-progress shared-grid range-selection drag when switching sheets (prevents pointerup focus steal)", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const formulaBar = document.createElement("div");
    document.body.appendChild(formulaBar);

    const app = new SpreadsheetApp(root, status, { formulaBar });
    expect(app.getGridMode()).toBe("shared");

    const doc: any = app.getDocument();
    doc.setCellValue("Sheet2", { row: 0, col: 0 }, "X");

    const sharedGrid = (app as any).sharedGrid as any;
    expect(sharedGrid).toBeTruthy();

    const bar = (app as any).formulaBar as any;
    expect(bar).toBeTruthy();
    const focusSpy = vi.spyOn(bar, "focus");
    const endRangeSelectionSpy = vi.spyOn(bar, "endRangeSelection");

    // Force the shared grid into range-selection mode (as if the formula bar were editing a formula).
    sharedGrid.setInteractionMode("rangeSelection");

    // Seed an in-progress range-selection drag. Without canceling this on sheet switch, a later pointerup
    // would call the range-selection end callback and focus the formula bar after the user navigated away.
    const pointerId = 42;
    const transientRange = { startRow: 1, endRow: 2, startCol: 1, endCol: 2 };
    sharedGrid.selectionPointerId = pointerId;
    sharedGrid.dragMode = "selection";
    sharedGrid.transientRange = transientRange;
    sharedGrid.renderer.setRangeSelection(transientRange);

    focusSpy.mockClear();
    endRangeSelectionSpy.mockClear();

    app.activateSheet("Sheet2");

    expect(sharedGrid.selectionPointerId).toBeNull();
    expect(sharedGrid.dragMode).toBeNull();
    expect(sharedGrid.transientRange).toBeNull();
    expect(endRangeSelectionSpy).toHaveBeenCalled();
    expect(focusSpy).not.toHaveBeenCalled();

    const selectionCanvas = (app as any).selectionCanvas as HTMLCanvasElement;
    dispatchPointerEvent(selectionCanvas, "pointerup", { clientX: 0, clientY: 0, pointerId });

    expect(focusSpy).not.toHaveBeenCalled();

    app.destroy();
    root.remove();
    formulaBar.remove();
  });
});

