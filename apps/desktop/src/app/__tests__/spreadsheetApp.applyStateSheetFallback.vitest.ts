/**
 * @vitest-environment jsdom
 */

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { DocumentController } from "../../document/documentController.js";
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

function createStatus() {
  return {
    activeCell: document.createElement("div"),
    selectionRange: document.createElement("div"),
    activeValue: document.createElement("div"),
  };
}

function seedThreeSheets(app: SpreadsheetApp): void {
  const doc = app.getDocument();
  // Ensure the default sheet is materialized (DocumentController creates sheets lazily).
  doc.getCell("Sheet1", { row: 0, col: 0 });
  doc.addSheet({ sheetId: "Sheet2", name: "Sheet2", insertAfterId: "Sheet1" });
  doc.addSheet({ sheetId: "Sheet3", name: "Sheet3", insertAfterId: "Sheet2" });
}

describe("SpreadsheetApp applyState active sheet fallback", () => {
  afterEach(() => {
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
    delete process.env.DESKTOP_GRID_MODE;
  });

  beforeEach(() => {
    document.body.innerHTML = "";
    process.env.DESKTOP_GRID_MODE = "legacy";

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

  it("activates the adjacent visible sheet when doc.applyState removes the active sheet", () => {
    const app = new SpreadsheetApp(createRoot(), createStatus());
    try {
      seedThreeSheets(app);
      const doc = app.getDocument();
      const activateSpy = vi.spyOn(app, "activateSheet");
      const outlinesBySheet = (app as any).outlinesBySheet as Map<string, unknown>;
      const workbookImageManager = (app as any).workbookImageManager as { imageRefCount: Map<string, number> };

      app.activateSheet("Sheet2");
      // Ensure per-sheet outline state is created for Sheet2 so we can verify it is cleaned up when
      // applyState removes the sheet.
      app.hideRows([0]);
      expect(outlinesBySheet.has("Sheet2")).toBe(true);

      const imageId = "test_image";
      doc.setSheetDrawings("Sheet2", [
        {
          id: 1,
          kind: { type: "image", imageId },
          // Use UI-style anchors so SpreadsheetApp can treat this as a normalized DrawingObject
          // without needing to run the model adapter (avoids test-only overlay render errors).
          anchor: {
            type: "oneCell",
            from: { cell: { row: 0, col: 0 }, offset: { xEmu: 0, yEmu: 0 } },
            size: { cx: 10_000, cy: 10_000 },
          },
          zOrder: 0,
        },
      ]);
      expect(workbookImageManager.imageRefCount.get(imageId)).toBe(1);

      const backgroundImageId = "test_background_image";
      doc.setSheetBackgroundImageId("Sheet2", backgroundImageId);
      expect(workbookImageManager.imageRefCount.get(backgroundImageId)).toBe(1);

      const cellImageId = "test_cell_image";
      doc.setCellValue("Sheet2", { row: 0, col: 0 }, { imageId: cellImageId });
      expect(workbookImageManager.imageRefCount.get(cellImageId)).toBe(1);
      activateSpy.mockClear();

      const snapshotDoc = new DocumentController();
      snapshotDoc.setCellValue("Sheet1", { row: 0, col: 0 }, "A");
      snapshotDoc.setCellValue("Sheet3", { row: 0, col: 0 }, "C");
      const snapshot = snapshotDoc.encodeState();

      doc.applyState(snapshot);

      expect(app.getCurrentSheetId()).toBe("Sheet3");
      expect(activateSpy).toHaveBeenCalledWith("Sheet3");
      // applyState deletes sheets after emitting its change event; ensure per-sheet caches do not
      // retain state for removed sheets.
      expect(outlinesBySheet.has("Sheet2")).toBe(false);
      // The workbook image ref counter should also drop image refs that were only present on the
      // removed sheet.
      expect(workbookImageManager.imageRefCount.get(imageId) ?? 0).toBe(0);
      expect(workbookImageManager.imageRefCount.get(backgroundImageId) ?? 0).toBe(0);
      expect(workbookImageManager.imageRefCount.get(cellImageId) ?? 0).toBe(0);
    } finally {
      app.destroy();
    }
  });

  it("prefers the next visible sheet to the right when applyState removes the first sheet", () => {
    const app = new SpreadsheetApp(createRoot(), createStatus());
    try {
      seedThreeSheets(app);
      const doc = app.getDocument();
      const activateSpy = vi.spyOn(app, "activateSheet");

      app.activateSheet("Sheet1");
      activateSpy.mockClear();

      // Restore a snapshot that removes Sheet1 but keeps Sheet2 + Sheet3.
      const snapshotDoc = new DocumentController();
      snapshotDoc.setCellValue("Sheet2", { row: 0, col: 0 }, "B");
      snapshotDoc.setCellValue("Sheet3", { row: 0, col: 0 }, "C");
      const snapshot = snapshotDoc.encodeState();

      doc.applyState(snapshot);

      // Excel-like: delete active sheet -> activate the next visible sheet to the right (Sheet2).
      expect(app.getCurrentSheetId()).toBe("Sheet2");
      expect(activateSpy).toHaveBeenCalledWith("Sheet2");
    } finally {
      app.destroy();
    }
  });
});
