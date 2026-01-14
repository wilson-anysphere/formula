/**
 * @vitest-environment jsdom
 */

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { chartIdToDrawingId } from "../../charts/chartDrawingAdapter";
import { pxToEmu } from "../../drawings/overlay";
import { SpreadsheetApp } from "../spreadsheetApp";

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

function createPointerLikeMouseEvent(
  type: string,
  options: {
    clientX: number;
    clientY: number;
    button: number;
    pointerId?: number;
    pointerType?: string;
  },
): MouseEvent {
  const event = new MouseEvent(type, {
    bubbles: true,
    cancelable: true,
    clientX: options.clientX,
    clientY: options.clientY,
    button: options.button,
  });
  Object.defineProperty(event, "pointerId", { configurable: true, value: options.pointerId ?? 1 });
  Object.defineProperty(event, "pointerType", { configurable: true, value: options.pointerType ?? "mouse" });
  return event;
}

describe("SpreadsheetApp canvas chart pointer drag (read-only)", () => {
  afterEach(() => {
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
    delete process.env.DESKTOP_GRID_MODE;
    delete process.env.CANVAS_CHARTS;
    delete process.env.USE_CANVAS_CHARTS;
  });

  beforeEach(() => {
    document.body.innerHTML = "";
    process.env.DESKTOP_GRID_MODE = "legacy";
    process.env.CANVAS_CHARTS = "1";

    const storage = createInMemoryLocalStorage();
    Object.defineProperty(globalThis, "localStorage", { configurable: true, value: storage });
    Object.defineProperty(window, "localStorage", { configurable: true, value: storage });
    storage.clear();

    Object.defineProperty(globalThis, "requestAnimationFrame", {
      configurable: true,
      writable: true,
      value: (cb: FrameRequestCallback) => {
        cb(0);
        return 0;
      },
    });
    Object.defineProperty(globalThis, "cancelAnimationFrame", { configurable: true, writable: true, value: () => {} });

    Object.defineProperty(HTMLCanvasElement.prototype, "getContext", {
      configurable: true,
      value: () => createMockCanvasContext(),
    });

    (globalThis as any).ResizeObserver = class {
      observe() {}
      disconnect() {}
    };
  });

  it("moves the chart on pointer drag when editable", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };
    const app = new SpreadsheetApp(root, status);

    const { chart_id: chartId } = app.addChart({
      chart_type: "bar",
      data_range: "A2:B5",
      title: "Drag Chart",
      position: "A1",
    });

    // Use a deterministic absolute anchor for stable pointer coordinates.
    (app as any).chartStore.updateChartAnchor(chartId, {
      kind: "absolute",
      xEmu: pxToEmu(30),
      yEmu: pxToEmu(20),
      cxEmu: pxToEmu(80),
      cyEmu: pxToEmu(60),
    });

    const drawingId = chartIdToDrawingId(chartId);
    const rect = app.getDrawingRectPx(drawingId);
    expect(rect).not.toBeNull();

    const anchorBefore = JSON.parse(JSON.stringify(app.listCharts().find((c) => c.id === chartId)?.anchor ?? null));

    const startX = rect!.x + rect!.width / 2;
    const startY = rect!.y + rect!.height / 2;

    root.dispatchEvent(createPointerLikeMouseEvent("pointerdown", { clientX: startX, clientY: startY, button: 0, pointerId: 1 }));
    root.dispatchEvent(
      createPointerLikeMouseEvent("pointermove", { clientX: startX + 50, clientY: startY, button: 0, pointerId: 1 }),
    );
    root.dispatchEvent(
      createPointerLikeMouseEvent("pointerup", { clientX: startX + 50, clientY: startY, button: 0, pointerId: 1 }),
    );

    const anchorAfter = JSON.parse(JSON.stringify(app.listCharts().find((c) => c.id === chartId)?.anchor ?? null));
    expect(anchorAfter).not.toEqual(anchorBefore);

    app.destroy();
    root.remove();
  });

  it("does not move the chart on pointer drag when collab is read-only", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };
    const app = new SpreadsheetApp(root, status);
    (app as any).collabSession = { isReadOnly: () => true };

    const { chart_id: chartId } = app.addChart({
      chart_type: "bar",
      data_range: "A2:B5",
      title: "Read-only Drag Chart",
      position: "A1",
    });

    (app as any).chartStore.updateChartAnchor(chartId, {
      kind: "absolute",
      xEmu: pxToEmu(30),
      yEmu: pxToEmu(20),
      cxEmu: pxToEmu(80),
      cyEmu: pxToEmu(60),
    });

    const drawingId = chartIdToDrawingId(chartId);
    const rect = app.getDrawingRectPx(drawingId);
    expect(rect).not.toBeNull();

    const anchorBefore = JSON.parse(JSON.stringify(app.listCharts().find((c) => c.id === chartId)?.anchor ?? null));

    const startX = rect!.x + rect!.width / 2;
    const startY = rect!.y + rect!.height / 2;

    root.dispatchEvent(createPointerLikeMouseEvent("pointerdown", { clientX: startX, clientY: startY, button: 0, pointerId: 1 }));
    root.dispatchEvent(
      createPointerLikeMouseEvent("pointermove", { clientX: startX + 50, clientY: startY, button: 0, pointerId: 1 }),
    );
    root.dispatchEvent(
      createPointerLikeMouseEvent("pointerup", { clientX: startX + 50, clientY: startY, button: 0, pointerId: 1 }),
    );

    const anchorAfter = JSON.parse(JSON.stringify(app.listCharts().find((c) => c.id === chartId)?.anchor ?? null));
    expect(anchorAfter).toEqual(anchorBefore);

    app.destroy();
    root.remove();
  });

  it("shows a chart read-only toast when deleting a chart via deleteDrawingById", () => {
    const root = createRoot();
    const toastRoot = document.createElement("div");
    toastRoot.id = "toast-root";
    document.body.appendChild(toastRoot);

    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };
    const app = new SpreadsheetApp(root, status);
    (app as any).collabSession = { isReadOnly: () => true };

    const { chart_id: chartId } = app.addChart({
      chart_type: "bar",
      data_range: "A2:B5",
      title: "Read-only Delete Chart",
      position: "A1",
    });

    const drawingId = chartIdToDrawingId(chartId);
    app.deleteDrawingById(drawingId);

    expect(document.querySelector("#toast-root")?.textContent ?? "").toContain("edit charts");
    expect(app.listCharts().some((c) => c.id === chartId)).toBe(true);

    app.destroy();
    root.remove();
  });
});
