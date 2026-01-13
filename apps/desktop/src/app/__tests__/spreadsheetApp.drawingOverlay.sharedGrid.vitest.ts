/**
 * @vitest-environment jsdom
 */

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { DrawingOverlay } from "../../drawings/overlay";
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

describe("SpreadsheetApp drawing overlay (shared grid)", () => {
  afterEach(() => {
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
  });

  beforeEach(() => {
    document.body.innerHTML = "";

    const storage = createInMemoryLocalStorage();
    Object.defineProperty(globalThis, "localStorage", { configurable: true, value: storage });
    Object.defineProperty(window, "localStorage", { configurable: true, value: storage });
    storage.clear();

    // CanvasGridRenderer schedules renders via requestAnimationFrame; ensure it exists in jsdom.
    Object.defineProperty(globalThis, "requestAnimationFrame", {
      configurable: true,
      value: (cb: FrameRequestCallback) => {
        cb(0);
        return 0;
      },
    });
    Object.defineProperty(globalThis, "cancelAnimationFrame", { configurable: true, value: () => {} });

    Object.defineProperty(window, "devicePixelRatio", { configurable: true, value: 2 });

    Object.defineProperty(HTMLCanvasElement.prototype, "getContext", {
      configurable: true,
      value: () => createMockCanvasContext(),
    });

    (globalThis as any).ResizeObserver = class {
      observe() {}
      disconnect() {}
    };
  });

  it("mounts the drawing canvas, resizes with DPR, and re-renders on shared-grid scroll + sheet change", () => {
    const prior = process.env.DESKTOP_GRID_MODE;
    process.env.DESKTOP_GRID_MODE = "shared";
    try {
      const resizeSpy = vi.spyOn(DrawingOverlay.prototype, "resize");
      const renderSpy = vi.spyOn(DrawingOverlay.prototype, "render");

      const root = createRoot();
      const status = {
        activeCell: document.createElement("div"),
        selectionRange: document.createElement("div"),
        activeValue: document.createElement("div"),
      };

      const app = new SpreadsheetApp(root, status);

      const canvas = root.querySelector<HTMLCanvasElement>('[data-testid="drawing-layer-canvas"]');
      expect(canvas).not.toBeNull();

      expect(resizeSpy).toHaveBeenCalledWith(
        expect.objectContaining({
          width: 800 - 48,
          height: 600 - 24,
          dpr: 2,
        }),
      );

      // `DrawingOverlay.resize` should size the backing buffer at DPR scale.
      expect(canvas!.width).toBe((800 - 48) * 2);
      expect(canvas!.height).toBe((600 - 24) * 2);

      renderSpy.mockClear();
      const sharedGrid = (app as any).sharedGrid;
      sharedGrid.scrollTo(0, 100);
      expect(renderSpy).toHaveBeenCalled();

      // Switching sheets should trigger a drawings overlay rerender so objects can refresh.
      const doc = app.getDocument();
      doc.addSheet({ sheetId: "sheet_2", name: "Sheet2" });
      doc.setCellValue("sheet_2", { row: 0, col: 0 }, "Seed2");
      renderSpy.mockClear();
      app.activateSheet("sheet_2");
      expect(renderSpy).toHaveBeenCalled();

      app.destroy();
      root.remove();
    } finally {
      if (prior === undefined) delete process.env.DESKTOP_GRID_MODE;
      else process.env.DESKTOP_GRID_MODE = prior;
    }
  });

  it("passes frozen pane metadata to the drawing overlay viewport so drawings pin + clip correctly", () => {
    const prior = process.env.DESKTOP_GRID_MODE;
    process.env.DESKTOP_GRID_MODE = "shared";
    try {
      const renderSpy = vi.spyOn(DrawingOverlay.prototype, "render");

      const root = createRoot();
      const status = {
        activeCell: document.createElement("div"),
        selectionRange: document.createElement("div"),
        activeValue: document.createElement("div"),
      };

      const app = new SpreadsheetApp(root, status);

      const doc = app.getDocument();
      doc.setFrozen(app.getCurrentSheetId(), 1, 1, { label: "Freeze" });

      renderSpy.mockClear();
      const sharedGrid = (app as any).sharedGrid;
      sharedGrid.scrollTo(50, 100);

      expect(renderSpy).toHaveBeenCalled();
      const lastCall = renderSpy.mock.calls.at(-1);
      const viewport = lastCall?.[1] as any;

      expect(viewport).toEqual(
        expect.objectContaining({
          frozenRows: 1,
          frozenCols: 1,
        }),
      );

      const gridViewport = sharedGrid.renderer.scroll.getViewportState();
      const headerWidth = sharedGrid.renderer.scroll.cols.totalSize(1);
      const headerHeight = sharedGrid.renderer.scroll.rows.totalSize(1);
      const offsetX = Math.min(headerWidth, gridViewport.width);
      const offsetY = Math.min(headerHeight, gridViewport.height);
      const cellAreaWidth = Math.max(0, gridViewport.width - offsetX);
      const cellAreaHeight = Math.max(0, gridViewport.height - offsetY);

      const expectedFrozenWidthPx = Math.min(cellAreaWidth, Math.max(0, gridViewport.frozenWidth - offsetX));
      const expectedFrozenHeightPx = Math.min(cellAreaHeight, Math.max(0, gridViewport.frozenHeight - offsetY));

      expect(viewport.frozenWidthPx).toBe(expectedFrozenWidthPx);
      expect(viewport.frozenHeightPx).toBe(expectedFrozenHeightPx);

      app.destroy();
      root.remove();
    } finally {
      if (prior === undefined) delete process.env.DESKTOP_GRID_MODE;
      else process.env.DESKTOP_GRID_MODE = prior;
    }
  });

  it("renders per-sheet drawings + images from DocumentController", () => {
    const prior = process.env.DESKTOP_GRID_MODE;
    process.env.DESKTOP_GRID_MODE = "shared";
    try {
      const renderSpy = vi.spyOn(DrawingOverlay.prototype, "render").mockResolvedValue(undefined);

      const root = createRoot();
      const status = {
        activeCell: document.createElement("div"),
        selectionRange: document.createElement("div"),
        activeValue: document.createElement("div"),
      };

      const app = new SpreadsheetApp(root, status);

      // Inject a simple workbook drawing layer: one floating image anchored to A1.
      const imageId = "image1.png";
      const bytes = new Uint8Array([1, 2, 3]);
      const doc = (app as any).document as any;
      doc.getSheetDrawings = () => ({
        drawings: [
          {
            id: 1,
            kind: { Image: { image_id: imageId } },
            anchor: {
              OneCell: {
                from: { cell: { row: 0, col: 0 }, offset: { x_emu: 0, y_emu: 0 } },
                ext: { cx: 914400, cy: 914400 },
              },
            },
            z_order: 0,
          },
        ],
      });
      doc.getImage = (id: string) => (id === imageId ? { bytes, mimeType: "image/png" } : null);

      renderSpy.mockClear();

      // Force a drawing render pass and assert that the overlay receives our object.
      (app as any).renderDrawings();
      expect(renderSpy).toHaveBeenCalled();
      const objects = renderSpy.mock.calls[0]?.[0] as any[];
      expect(objects).toHaveLength(1);
      expect(objects[0]).toMatchObject({ kind: { type: "image", imageId } });

      // Ensure the overlay image store is backed by the document's image map.
      const imageStore = (app as any).drawingImages;
      expect(imageStore.get(imageId)).toMatchObject({ id: imageId, mimeType: "image/png" });
      expect(imageStore.get(imageId)?.bytes).toEqual(bytes);

      app.destroy();
      root.remove();
    } finally {
      if (prior === undefined) delete process.env.DESKTOP_GRID_MODE;
      else process.env.DESKTOP_GRID_MODE = prior;
    }
  });
});
