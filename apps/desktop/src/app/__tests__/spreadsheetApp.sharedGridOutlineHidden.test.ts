/**
 * @vitest-environment jsdom
 */

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { SpreadsheetApp } from "../spreadsheetApp";
import { navigateSelectionByKey } from "../../selection/navigation";

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
        // Default all unknown properties to no-op functions so rendering code can execute.
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

describe("SpreadsheetApp shared-grid outline compatibility", () => {
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

    // jsdom lacks a real canvas implementation; SpreadsheetApp expects a 2D context.
    Object.defineProperty(HTMLCanvasElement.prototype, "getContext", {
      configurable: true,
      value: () => createMockCanvasContext(),
    });

    // jsdom doesn't ship ResizeObserver by default.
    (globalThis as any).ResizeObserver = class {
      observe() {}
      disconnect() {}
    };
  });

  it("does not treat outline-hidden rows/cols as hidden for navigation in shared-grid mode", () => {
    const prior = process.env.DESKTOP_GRID_MODE;
    process.env.DESKTOP_GRID_MODE = "shared";
    try {
      const root = createRoot();
      const status = {
        activeCell: document.createElement("div"),
        selectionRange: document.createElement("div"),
        activeValue: document.createElement("div"),
      };

      const app = new SpreadsheetApp(root, status);
      expect(app.getGridMode()).toBe("shared");

      const sheetId = (app as any).sheetId as string;
      const outline = (app as any).getOutlineForSheet(sheetId) as any;
      // Shared-grid mode should not seed demo outline groups.
      expect(outline.rows.entries.size).toBe(0);
      expect(outline.cols.entries.size).toBe(0);

      // Create a collapsed outline group that would hide rows/cols in legacy mode.
      outline.groupRows(2, 4);
      outline.toggleRowGroup(5);
      outline.groupCols(2, 4);
      outline.toggleColGroup(5);

      // Sanity: the outline model considers the detail rows/cols hidden.
      // (Row/col indices are 1-based on the outline axes.)
      expect(outline.rows.entry(2).hidden.outline).toBe(true); // row 2 (0-based row 1)
      expect(outline.cols.entry(2).hidden.outline).toBe(true); // col B (0-based col 1)

      // But shared-grid navigation should ignore outline hidden state until the renderer supports it.
      const provider = (app as any).usedRangeProvider();
      expect(provider.isRowHidden(1)).toBe(false);
      expect(provider.isColHidden(1)).toBe(false);

      // Ensure Ctrl+Arrow (jump-to-edge) logic doesn't skip outline-hidden indices either.
      // Make row 2 (0-based row 1) non-empty and row 3 empty so Ctrl+ArrowDown should land on row 2.
      const documentController = (app as any).document;
      documentController.setCellValue(sheetId, { row: 1, col: 0 }, "X");
      documentController.setCellValue(sheetId, { row: 2, col: 0 }, null);

      const selection = (app as any).selection;
      const limits = (app as any).limits;
      const movedDown = navigateSelectionByKey(selection, "ArrowDown", { shift: false, primary: false }, provider, limits);
      expect(movedDown?.active.row).toBe(1);
      const movedRight = navigateSelectionByKey(selection, "ArrowRight", { shift: false, primary: false }, provider, limits);
      expect(movedRight?.active.col).toBe(1);

      const jumpedDown = navigateSelectionByKey(selection, "ArrowDown", { shift: false, primary: true }, provider, limits);
      expect(jumpedDown?.active.row).toBe(1);
      const jumpedRight = navigateSelectionByKey(selection, "ArrowRight", { shift: false, primary: true }, provider, limits);
      expect(jumpedRight?.active.col).toBe(1);

      app.destroy();
      root.remove();
    } finally {
      if (prior === undefined) delete process.env.DESKTOP_GRID_MODE;
      else process.env.DESKTOP_GRID_MODE = prior;
    }
  });

  it("skips user-hidden rows/cols for navigation in shared-grid mode", () => {
    const prior = process.env.DESKTOP_GRID_MODE;
    process.env.DESKTOP_GRID_MODE = "shared";
    try {
      const root = createRoot();
      const status = {
        activeCell: document.createElement("div"),
        selectionRange: document.createElement("div"),
        activeValue: document.createElement("div"),
      };

      const app = new SpreadsheetApp(root, status);
      expect(app.getGridMode()).toBe("shared");

      const sheetId = (app as any).sheetId as string;
      const outline = (app as any).getOutlineForSheet(sheetId) as any;
      // Hide row 2 and col B (0-based index 1) via the user-hidden flag.
      outline.rows.entryMut(2).hidden.user = true;
      outline.cols.entryMut(2).hidden.user = true;

      const provider = (app as any).usedRangeProvider();
      expect(provider.isRowHidden(1)).toBe(true);
      expect(provider.isColHidden(1)).toBe(true);

      const documentController = (app as any).document;

      // Make Ctrl+ArrowDown/Right want to stop on the hidden row/col if hidden indices weren't skipped.
      // Keep the sheet's used range large (it is seeded to D5) so jump-to-edge has room to scan.
      documentController.setCellValue(sheetId, { row: 1, col: 0 }, "X"); // hidden row
      documentController.setCellValue(sheetId, { row: 2, col: 0 }, null);
      documentController.setCellValue(sheetId, { row: 3, col: 0 }, null);
      documentController.setCellValue(sheetId, { row: 4, col: 0 }, null);

      // Ensure row 1 col 2/3 are empty so Ctrl+ArrowRight scans past the hidden column.
      documentController.setCellValue(sheetId, { row: 0, col: 2 }, null);
      documentController.setCellValue(sheetId, { row: 0, col: 3 }, null);

      const selection = (app as any).selection;
      const limits = (app as any).limits;

      const movedDown = navigateSelectionByKey(selection, "ArrowDown", { shift: false, primary: false }, provider, limits);
      expect(movedDown?.active.row).toBe(2);
      const movedRight = navigateSelectionByKey(selection, "ArrowRight", { shift: false, primary: false }, provider, limits);
      expect(movedRight?.active.col).toBe(2);

      const jumpedDown = navigateSelectionByKey(selection, "ArrowDown", { shift: false, primary: true }, provider, limits);
      expect(jumpedDown?.active.row).toBe(4);
      const jumpedRight = navigateSelectionByKey(selection, "ArrowRight", { shift: false, primary: true }, provider, limits);
      expect(jumpedRight?.active.col).toBe(3);

      app.destroy();
      root.remove();
    } finally {
      if (prior === undefined) delete process.env.DESKTOP_GRID_MODE;
      else process.env.DESKTOP_GRID_MODE = prior;
    }
  });
});
