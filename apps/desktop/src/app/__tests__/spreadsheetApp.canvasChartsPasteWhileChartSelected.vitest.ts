/**
 * @vitest-environment jsdom
 */

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { chartIdToDrawingId } from "../../charts/chartDrawingAdapter";
import { SpreadsheetApp } from "../spreadsheetApp";
import * as ui from "../../extensions/ui.js";

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

describe("SpreadsheetApp paste (canvas charts)", () => {
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

  it("does not paste cell contents when a chart is selected", async () => {
    const toastSpy = vi.spyOn(ui, "showToast").mockImplementation(() => {});

    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };
    const app = new SpreadsheetApp(root, status);

    const sheetId = app.getCurrentSheetId();
    app.getDocument().setCellValue(sheetId, { row: 0, col: 0 }, "Hello");

    const provider = {
      read: vi.fn(async () => ({ text: "X" })),
      write: vi.fn(async () => {}),
    };
    (app as any).clipboardProviderPromise = Promise.resolve(provider);

    const { chart_id: chartId } = app.addChart({
      chart_type: "bar",
      data_range: "A1:B2",
      title: "Paste Chart",
      position: "C1",
    });
    app.selectDrawingById(chartIdToDrawingId(chartId));
    expect(app.getSelectedChartId()).toBe(chartId);

    await app.pasteClipboardToSelection();

    expect(app.getDocument().getCell(sheetId, { row: 0, col: 0 }).value).toBe("Hello");
    expect(toastSpy).toHaveBeenCalledWith("Paste not supported while a chart is selected.", "warning");

    app.destroy();
    root.remove();
  });

  it("pastes images as drawings even when a chart is selected", async () => {
    // 1x1 transparent PNG.
    const pngBytes = new Uint8Array(
      // eslint-disable-next-line no-undef
      Buffer.from("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO7+FeAAAAAASUVORK5CYII=", "base64"),
    );

    Object.defineProperty(globalThis, "createImageBitmap", {
      configurable: true,
      value: vi.fn(async () => ({ width: 64, height: 32 })),
    });

    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };
    const app = new SpreadsheetApp(root, status);

    const provider = {
      read: vi.fn(async () => ({ imagePng: pngBytes })),
      write: vi.fn(async () => {}),
    };
    (app as any).clipboardProviderPromise = Promise.resolve(provider);

    const { chart_id: chartId } = app.addChart({
      chart_type: "bar",
      data_range: "A1:B2",
      title: "Paste Chart",
      position: "C1",
    });
    app.selectDrawingById(chartIdToDrawingId(chartId));
    expect(app.getSelectedChartId()).toBe(chartId);

    await app.pasteClipboardToSelection();

    const objects = app.getDrawingObjects();
    const image = objects.find((obj) => obj.kind.type === "image") ?? null;
    expect(image).not.toBeNull();
    expect(app.getSelectedChartId()).toBe(null);
    expect(app.getSelectedDrawingId()).toBe(image!.id);

    app.destroy();
    root.remove();
  });
});

