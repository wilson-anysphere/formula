/**
 * @vitest-environment jsdom
 */

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { pxToEmu } from "../../drawings/overlay";
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
  const base = { bubbles: true, clientX: opts.clientX, clientY: opts.clientY, pointerId, button };
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

describe("SpreadsheetApp chart zoom", () => {
  afterEach(() => {
    if (priorGridMode === undefined) delete process.env.DESKTOP_GRID_MODE;
    else process.env.DESKTOP_GRID_MODE = priorGridMode;
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
  });

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

  it("scales chart anchors under zoom (absolute) and converts drag deltas back to EMU correctly", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);

    const result = app.addChart({
      chart_type: "bar",
      data_range: "A2:B5",
      title: "Zoom Chart",
      position: "A1",
    });

    // Override the generated anchor so the test is independent of cell sizing details.
    (app as any).chartStore.updateChartAnchor(result.chart_id, {
      kind: "absolute",
      xEmu: pxToEmu(10),
      yEmu: pxToEmu(20),
      cxEmu: pxToEmu(50),
      cyEmu: pxToEmu(40),
    });

    app.setZoom(2);

    const chart = app.listCharts().find((c) => c.id === result.chart_id);
    expect(chart).toBeTruthy();
    expect(chart!.anchor.kind).toBe("absolute");

    const rect = (app as any).chartAnchorToViewportRect(chart!.anchor);
    expect(rect).not.toBeNull();
    expect(rect.left).toBeCloseTo(20);
    expect(rect.top).toBeCloseTo(40);
    expect(rect.width).toBeCloseTo(100);
    expect(rect.height).toBeCloseTo(80);

    const layout = (app as any).chartOverlayLayout();
    const originX = layout.originX as number;
    const originY = layout.originY as number;

    const startX = originX + rect.left + 10;
    const startY = originY + rect.top + 10;
    const endX = startX + 20; // screen delta (px)
    const endY = startY;

    dispatchPointerEvent(root, "pointerdown", { clientX: startX, clientY: startY, pointerId: 200 });
    dispatchPointerEvent(window, "pointermove", { clientX: endX, clientY: endY, pointerId: 200 });
    dispatchPointerEvent(window, "pointerup", { clientX: endX, clientY: endY, pointerId: 200 });

    const after = app.listCharts().find((c) => c.id === result.chart_id);
    expect(after).toBeTruthy();
    expect(after!.anchor.kind).toBe("absolute");
    const afterAnchor = after!.anchor as any;
    // dx=20px at zoom=2 => +10px in document space
    expect(afterAnchor.xEmu).toBeCloseTo(pxToEmu(20));
    expect(afterAnchor.yEmu).toBeCloseTo(pxToEmu(20));
    expect(afterAnchor.cxEmu).toBeCloseTo(pxToEmu(50));
    expect(afterAnchor.cyEmu).toBeCloseTo(pxToEmu(40));

    app.destroy();
    root.remove();
  });
});

