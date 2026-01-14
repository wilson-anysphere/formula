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

describe("SpreadsheetApp fill-handle drag sheet switching", () => {
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

  it("cancels an in-progress legacy fill-handle drag when switching sheets (prevents cross-sheet commit)", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);

    // Remove the default demo chart so chart hit testing doesn't interfere with the fill handle.
    for (const chart of app.listCharts()) {
      (app as any).chartStore.deleteChart(chart.id);
    }

    const doc: any = app.getDocument();
    const sheet1 = app.getCurrentSheetId();

    // Ensure Sheet2 exists and has a sentinel value we can detect.
    doc.setCellValue("Sheet2", { row: 1, col: 0 }, "S2_A2");

    // Seed Sheet1 A1 with a value so a fill drag would write into A2.
    doc.setCellValue(sheet1, { row: 0, col: 0 }, 1);
    // Seed Sheet1 A2 with a sentinel value so we can detect an (incorrect) fill commit that
    // happens even after the drag is canceled.
    doc.setCellValue(sheet1, { row: 1, col: 0 }, "S1_A2");

    // Ensure a render occurs so the fill handle rect is computed.
    app.refresh();

    const handle = app.getFillHandleRect();
    expect(handle).not.toBeNull();

    const surface = document.createElement("div");
    root.appendChild(surface);

    const x = handle!.x + handle!.width / 2;
    const y = handle!.y + handle!.height / 2;

    dispatchPointerEvent(surface, "pointerdown", { clientX: x, clientY: y, pointerId: 42 });
    const dragState = (app as any).dragState;
    expect(dragState?.mode).toBe("fill");

    // Simulate the user dragging the fill handle down one row (A1 -> A2) before releasing.
    dragState.targetRange = { startRow: 0, endRow: 1, startCol: 0, endCol: 0 };
    dragState.endCell = { row: 1, col: 0 };

    // Switch sheets before the pointerup arrives. This should cancel the in-progress drag so
    // the later pointerup cannot commit into Sheet2.
    app.activateSheet("Sheet2");
    expect((app as any).dragState).toBeNull();

    dispatchPointerEvent(surface, "pointerup", { clientX: x, clientY: y, pointerId: 42 });

    const sheet2A2 = doc.getCell("Sheet2", { row: 1, col: 0 }) as any;
    expect(sheet2A2.formula).toBeNull();
    expect(sheet2A2.value).toBe("S2_A2");

    const sheet1A2 = doc.getCell(sheet1, { row: 1, col: 0 }) as any;
    expect(sheet1A2.formula).toBeNull();
    expect(sheet1A2.value).toBe("S1_A2");

    app.destroy();
    root.remove();
  });
});
