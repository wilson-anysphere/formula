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

describe("SpreadsheetApp restoreDocumentState sheet fallback activation", () => {
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

  it("prefers the adjacent visible sheet when the active sheet is removed by restoreDocumentState (Excel-like)", async () => {
    const app = new SpreadsheetApp(createRoot(), createStatus());
    try {
      seedThreeSheets(app);
      const activateSpy = vi.spyOn(app, "activateSheet");
      app.activateSheet("Sheet2");
      activateSpy.mockClear();

      // Restore a snapshot that removes Sheet2 but keeps Sheet1 + Sheet3.
      const snapshotDoc = new DocumentController();
      snapshotDoc.setCellValue("Sheet1", { row: 0, col: 0 }, "A");
      snapshotDoc.setCellValue("Sheet3", { row: 0, col: 0 }, "C");
      const snapshot = snapshotDoc.encodeState();

      await app.restoreDocumentState(snapshot);

      // Excel-like: prefer the next visible sheet to the right (Sheet3), not the first sheet.
      expect(app.getCurrentSheetId()).toBe("Sheet3");
      // Active sheet deletion during applyState occurs *after* the change event is emitted, so the
      // restore logic updates the sheet id directly (no intermediate activateSheet calls).
      expect(activateSpy).not.toHaveBeenCalled();
    } finally {
      app.destroy();
    }
  });

  it("prefers the adjacent visible sheet when the active sheet becomes hidden by restoreDocumentState (Excel-like)", async () => {
    const app = new SpreadsheetApp(createRoot(), createStatus());
    try {
      seedThreeSheets(app);
      const activateSpy = vi.spyOn(app, "activateSheet");
      app.activateSheet("Sheet2");
      activateSpy.mockClear();

      // Restore a snapshot where Sheet2 still exists but is hidden.
      const snapshotDoc = new DocumentController();
      snapshotDoc.setCellValue("Sheet1", { row: 0, col: 0 }, "A");
      snapshotDoc.setCellValue("Sheet2", { row: 0, col: 0 }, "B");
      snapshotDoc.setCellValue("Sheet3", { row: 0, col: 0 }, "C");
      snapshotDoc.setSheetVisibility("Sheet2", "hidden");
      const snapshot = snapshotDoc.encodeState();

      await app.restoreDocumentState(snapshot);

      expect(app.getCurrentSheetId()).toBe("Sheet3");
      // restoreDocumentState updates the active sheet directly after applyState completes; it should
      // not call activateSheet during the applyState change dispatch (avoids redundant re-renders).
      expect(activateSpy).not.toHaveBeenCalled();
    } finally {
      app.destroy();
    }
  });

  it("clears outline->engine hidden column sync caches on restoreDocumentState", async () => {
    const app = new SpreadsheetApp(createRoot(), createStatus());
    try {
      // Seed non-null cache values to ensure restore clears them even when sheet ids collide.
      (app as any).lastSyncedHiddenColsEngine = { terminate: () => {} };
      (app as any).lastSyncedHiddenColsKey = "Sheet1:0";
      (app as any).lastSyncedHiddenCols = [0];

      const snapshotDoc = new DocumentController();
      snapshotDoc.setCellValue("Sheet1", { row: 0, col: 0 }, "A");
      const snapshot = snapshotDoc.encodeState();

      await app.restoreDocumentState(snapshot);

      expect((app as any).lastSyncedHiddenColsEngine).toBeNull();
      expect((app as any).lastSyncedHiddenColsKey).toBeNull();
      expect((app as any).lastSyncedHiddenCols).toBeNull();
    } finally {
      app.destroy();
    }
  });
});
