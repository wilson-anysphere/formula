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

describe("SpreadsheetApp chart selection + drag", () => {
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

  it("selects a chart on click", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    const chart = app.listCharts().find((c) => c.sheetId === app.getCurrentSheetId());
    expect(chart).toBeTruthy();

    const rect = (app as any).chartAnchorToViewportRect(chart!.anchor);
    expect(rect).not.toBeNull();

    const layout = (app as any).chartOverlayLayout();
    const originX = layout.originX as number;
    const originY = layout.originY as number;

    const clickX = originX + rect.left + 2;
    const clickY = originY + rect.top + 2;
    dispatchPointerEvent(root, "pointerdown", { clientX: clickX, clientY: clickY, pointerId: 1 });
    dispatchPointerEvent(window, "pointerup", { clientX: clickX, clientY: clickY, pointerId: 1 });

    expect(app.getSelectedChartId()).toBe(chart!.id);

    app.destroy();
    root.remove();
  });

  it("dragging a chart updates its twoCell anchor", () => {
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
      title: "Drag Chart",
      position: "A1",
    });

    const before = app.listCharts().find((c) => c.id === result.chart_id);
    expect(before).toBeTruthy();
    expect(before!.anchor.kind).toBe("twoCell");

    const beforeAnchor = before!.anchor as any;
    expect(beforeAnchor.fromCol).toBe(0);
    expect(beforeAnchor.toCol).toBeGreaterThan(0);

    const rect = (app as any).chartAnchorToViewportRect(before!.anchor);
    expect(rect).not.toBeNull();

    const layout = (app as any).chartOverlayLayout();
    const originX = layout.originX as number;
    const originY = layout.originY as number;

    const startX = originX + rect.left + 10;
    const startY = originY + rect.top + 10;
    const endX = startX + 100; // move by one column (default col width)
    const endY = startY;

    dispatchPointerEvent(root, "pointerdown", { clientX: startX, clientY: startY, pointerId: 7 });
    dispatchPointerEvent(window, "pointermove", { clientX: endX, clientY: endY, pointerId: 7 });
    dispatchPointerEvent(window, "pointerup", { clientX: endX, clientY: endY, pointerId: 7 });

    const after = app.listCharts().find((c) => c.id === result.chart_id);
    expect(after).toBeTruthy();
    expect(after!.anchor.kind).toBe("twoCell");

    const afterAnchor = after!.anchor as any;
    expect(afterAnchor.fromCol).toBe(beforeAnchor.fromCol + 1);
    expect(afterAnchor.toCol).toBe(beforeAnchor.toCol + 1);
    expect(afterAnchor.fromRow).toBe(beforeAnchor.fromRow);
    expect(afterAnchor.toRow).toBe(beforeAnchor.toRow);

    app.destroy();
    root.remove();
  });

  it("resizing a selected chart updates its twoCell anchor", () => {
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
      title: "Resize Chart",
      position: "A1",
    });

    const before = app.listCharts().find((c) => c.id === result.chart_id);
    expect(before).toBeTruthy();
    expect(before!.anchor.kind).toBe("twoCell");

    const beforeAnchor = before!.anchor as any;
    expect(beforeAnchor.fromCol).toBe(0);
    expect(beforeAnchor.fromRow).toBe(0);

    const rect = (app as any).chartAnchorToViewportRect(before!.anchor);
    expect(rect).not.toBeNull();

    const layout = (app as any).chartOverlayLayout();
    const originX = layout.originX as number;
    const originY = layout.originY as number;

    // First click selects the chart.
    const selectX = originX + rect.left + 10;
    const selectY = originY + rect.top + 10;
    dispatchPointerEvent(root, "pointerdown", { clientX: selectX, clientY: selectY, pointerId: 31 });
    dispatchPointerEvent(window, "pointerup", { clientX: selectX, clientY: selectY, pointerId: 31 });
    expect(app.getSelectedChartId()).toBe(result.chart_id);

    // Second click on the bottom-right handle starts a resize drag.
    const handleX = originX + rect.left + rect.width;
    const handleY = originY + rect.top + rect.height;
    const endX = handleX + 110; // increase width by ~1 column (default col width = 100)
    const endY = handleY;

    dispatchPointerEvent(root, "pointerdown", { clientX: handleX, clientY: handleY, pointerId: 32 });
    dispatchPointerEvent(window, "pointermove", { clientX: endX, clientY: endY, pointerId: 32 });
    dispatchPointerEvent(window, "pointerup", { clientX: endX, clientY: endY, pointerId: 32 });

    const after = app.listCharts().find((c) => c.id === result.chart_id);
    expect(after).toBeTruthy();
    expect(after!.anchor.kind).toBe("twoCell");

    const afterAnchor = after!.anchor as any;
    expect(afterAnchor.fromCol).toBe(beforeAnchor.fromCol);
    expect(afterAnchor.fromRow).toBe(beforeAnchor.fromRow);
    expect(afterAnchor.toCol).toBe(beforeAnchor.toCol + 1);
    expect(afterAnchor.toRow).toBe(beforeAnchor.toRow);

    app.destroy();
    root.remove();
  });

  it("hit testing respects shared-grid pane clipping (frozen panes)", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    const doc = app.getDocument();
    doc.setFrozen(app.getCurrentSheetId(), 1, 1, { label: "Freeze" });

    const result = app.addChart({
      chart_type: "bar",
      data_range: "A2:B5",
      title: "Frozen Chart",
      position: "A1",
    });

    const chart = app.listCharts().find((c) => c.id === result.chart_id);
    expect(chart).toBeTruthy();

    const layout = (app as any).chartOverlayLayout();
    const originX = layout.originX as number;
    const originY = layout.originY as number;

    // This point is inside the chart bounds (which extend far beyond A1), but outside the
    // top-left frozen pane (so it should not count as a hit because the chart is clipped).
    dispatchPointerEvent(root, "pointerdown", { clientX: originX + 150, clientY: originY + 10, pointerId: 99 });
    dispatchPointerEvent(window, "pointerup", { clientX: originX + 150, clientY: originY + 10, pointerId: 99 });
    expect(app.getSelectedChartId()).toBe(null);

    // This point lies in the visible (clipped) portion of the chart in the top-left pane.
    dispatchPointerEvent(root, "pointerdown", { clientX: originX + 50, clientY: originY + 10, pointerId: 100 });
    dispatchPointerEvent(window, "pointerup", { clientX: originX + 50, clientY: originY + 10, pointerId: 100 });
    expect(app.getSelectedChartId()).toBe(chart!.id);

    app.destroy();
    root.remove();
  });
});
