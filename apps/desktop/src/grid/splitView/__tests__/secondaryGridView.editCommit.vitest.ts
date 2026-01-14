/**
 * @vitest-environment jsdom
 */

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { SecondaryGridView } from "../secondaryGridView";
import { DocumentController } from "../../../document/documentController.js";
import type { ImageStore } from "../../../drawings/types";

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

describe("SecondaryGridView edit commits", () => {
  afterEach(() => {
    delete (globalThis as any).__formulaSpreadsheetIsEditing;
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
  });

  beforeEach(() => {
    document.body.innerHTML = "";

    Object.defineProperty(globalThis, "requestAnimationFrame", {
      configurable: true,
      value: () => 0,
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

  const images: ImageStore = { get: () => undefined, set: () => {}, delete: () => {}, clear: () => {} };

  it("advances selection after an Enter commit", () => {
    const container = document.createElement("div");
    Object.defineProperty(container, "clientWidth", { configurable: true, value: 400 });
    Object.defineProperty(container, "clientHeight", { configurable: true, value: 300 });
    document.body.appendChild(container);

    const doc = new DocumentController();
    const editState = vi.fn();
    const gridView = new SecondaryGridView({
      container,
      document: doc,
      getSheetId: () => "Sheet1",
      rowCount: 11,
      colCount: 11,
      showFormulas: () => false,
      getComputedValue: () => null,
      getDrawingObjects: () => [],
      images,
      onEditStateChange: editState,
    });

    // Select A1 (grid coordinates include a 1x1 header at row/col 0).
    gridView.grid.setSelectionRanges([{ startRow: 1, endRow: 2, startCol: 1, endCol: 2 }], {
      activeCell: { row: 1, col: 1 },
      scrollIntoView: false,
    });

    // Start editing A1 and commit via Enter.
    (gridView as any).openEditor({ row: 1, col: 1, initialKey: "h" });
    expect(editState).toHaveBeenCalledWith(true);
    (gridView as any).editor.element.value = "hello";
    (gridView as any).editor.commit("enter", false);
    expect(editState).toHaveBeenLastCalledWith(false);

    expect(gridView.grid.renderer.getSelection()).toEqual({ row: 2, col: 1 }); // A2

    gridView.destroy();
    container.remove();
  });

  it("does not open a secondary cell editor while the spreadsheet is already editing (global edit mode)", () => {
    const container = document.createElement("div");
    Object.defineProperty(container, "clientWidth", { configurable: true, value: 400 });
    Object.defineProperty(container, "clientHeight", { configurable: true, value: 300 });
    document.body.appendChild(container);

    const doc = new DocumentController();
    const editState = vi.fn();
    const gridView = new SecondaryGridView({
      container,
      document: doc,
      getSheetId: () => "Sheet1",
      rowCount: 11,
      colCount: 11,
      showFormulas: () => false,
      getComputedValue: () => null,
      getDrawingObjects: () => [],
      images,
      onEditStateChange: editState,
    });

    // Simulate the desktop shell reporting that an edit session is already active (e.g. the primary formula bar).
    (globalThis as any).__formulaSpreadsheetIsEditing = true;

    (gridView as any).openEditor({ row: 1, col: 1, initialKey: "h" });

    expect(editState).not.toHaveBeenCalled();
    expect((gridView as any).editor.isOpen()).toBe(false);

    gridView.destroy();
    container.remove();
  });

  it("Tab commit keeps multi-cell selection and advances active cell within the range", () => {
    const container = document.createElement("div");
    Object.defineProperty(container, "clientWidth", { configurable: true, value: 400 });
    Object.defineProperty(container, "clientHeight", { configurable: true, value: 300 });
    document.body.appendChild(container);

    const doc = new DocumentController();
    const editState = vi.fn();
    const gridView = new SecondaryGridView({
      container,
      document: doc,
      getSheetId: () => "Sheet1",
      rowCount: 11,
      colCount: 11,
      showFormulas: () => false,
      getComputedValue: () => null,
      getDrawingObjects: () => [],
      images,
      onEditStateChange: editState,
    });

    // Select A1:B2 with active cell at B1.
    gridView.grid.setSelectionRanges([{ startRow: 1, endRow: 3, startCol: 1, endCol: 3 }], {
      activeCell: { row: 1, col: 2 },
      scrollIntoView: false,
    });

    (gridView as any).openEditor({ row: 1, col: 2, initialKey: "x" });
    expect(editState).toHaveBeenCalledWith(true);
    (gridView as any).editor.commit("tab", false);
    expect(editState).toHaveBeenLastCalledWith(false);

    // Tab from B1 should wrap to A2 while preserving the selection range.
    expect(gridView.grid.renderer.getSelection()).toEqual({ row: 2, col: 1 }); // A2
    expect(gridView.grid.renderer.getSelectionRanges()).toEqual([{ startRow: 1, endRow: 3, startCol: 1, endCol: 3 }]);

    gridView.destroy();
    container.remove();
  });
});
