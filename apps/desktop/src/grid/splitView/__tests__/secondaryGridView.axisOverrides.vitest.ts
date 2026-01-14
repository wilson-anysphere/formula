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
      putImageData: noop
    },
    {
      get(target, prop) {
        if (prop in target) return (target as any)[prop];
        return noop;
      },
      set(target, prop, value) {
        (target as any)[prop] = value;
        return true;
      }
    }
  );
  return context as any;
}

describe("SecondaryGridView sheet view axis overrides", () => {
  afterEach(() => {
    delete (globalThis as any).__formulaSpreadsheetIsEditing;
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
  });

  beforeEach(() => {
    document.body.innerHTML = "";

    Object.defineProperty(globalThis, "requestAnimationFrame", {
      configurable: true,
      value: () => 0
    });
    Object.defineProperty(globalThis, "cancelAnimationFrame", { configurable: true, value: () => {} });

    Object.defineProperty(HTMLCanvasElement.prototype, "getContext", {
      configurable: true,
      value: () => createMockCanvasContext()
    });

    (globalThis as any).ResizeObserver = class {
      observe() {}
      disconnect() {}
    };
  });

  const images: ImageStore = { get: () => undefined, set: () => {}, delete: () => {}, clear: () => {} };

  it("uses CanvasGridRenderer.applyAxisSizeOverrides (no per-index setters)", () => {
    const container = document.createElement("div");
    // Keep viewport 0-sized so the renderer doesn't do any expensive work in jsdom.
    Object.defineProperty(container, "clientWidth", { configurable: true, value: 0 });
    Object.defineProperty(container, "clientHeight", { configurable: true, value: 0 });
    document.body.appendChild(container);

    const doc = new DocumentController();
    const sheetId = "Sheet1";

    const gridView = new SecondaryGridView({
      container,
      document: doc,
      getSheetId: () => sheetId,
      rowCount: 10_001,
      colCount: 201,
      showFormulas: () => false,
      getComputedValue: () => null,
      getDrawingObjects: () => [],
      images,
    });

    // Seed a sheet view with many overrides (bypass DocumentController deltas to keep the test
    // focused and deterministic).
    const view = (doc.getSheetView(sheetId) ?? {}) as any;
    const rowHeights: Record<string, number> = {};
    for (let i = 0; i < 1_000; i += 1) rowHeights[String(i)] = 30;
    rowHeights["999999"] = 42;
    const colWidths: Record<string, number> = {};
    for (let i = 0; i < 150; i += 1) colWidths[String(i)] = 120;
    colWidths["999999"] = 321;
    view.rowHeights = rowHeights;
    view.colWidths = colWidths;
    (doc as any).model.setSheetView(sheetId, view);

    const renderer = gridView.grid.renderer;
    const batchSpy = vi.spyOn(renderer, "applyAxisSizeOverrides");
    const setRowSpy = vi.spyOn(renderer, "setRowHeight");
    const setColSpy = vi.spyOn(renderer, "setColWidth");
    const resetRowSpy = vi.spyOn(renderer, "resetRowHeight");
    const resetColSpy = vi.spyOn(renderer, "resetColWidth");

    gridView.syncSheetViewFromDocument();

    expect(batchSpy).toHaveBeenCalledTimes(1);
    expect(setRowSpy).not.toHaveBeenCalled();
    expect(setColSpy).not.toHaveBeenCalled();
    expect(resetRowSpy).not.toHaveBeenCalled();
    expect(resetColSpy).not.toHaveBeenCalled();

    gridView.destroy();
    container.remove();
  });

  it("does not resync axis overrides back into the same pane after a local resize edit", () => {
    const container = document.createElement("div");
    Object.defineProperty(container, "clientWidth", { configurable: true, value: 0 });
    Object.defineProperty(container, "clientHeight", { configurable: true, value: 0 });
    document.body.appendChild(container);

    const doc = new DocumentController();
    const sheetId = "Sheet1";

    const gridView = new SecondaryGridView({
      container,
      document: doc,
      getSheetId: () => sheetId,
      rowCount: 100,
      colCount: 50,
      showFormulas: () => false,
      getComputedValue: () => null,
      getDrawingObjects: () => [],
      images,
    });

    const renderer = gridView.grid.renderer;
    const batchSpy = vi.spyOn(renderer, "applyAxisSizeOverrides");
    batchSpy.mockClear();

    // Simulate the end-of-drag resize callback for this secondary pane. The renderer is already updated
    // during the drag; this should only mutate the document and should not trigger a full re-sync of
    // all overrides back into the same renderer instance.
    (gridView as any).onAxisSizeChange({
      kind: "row",
      index: 2,
      size: 40,
      previousSize: 24,
      defaultSize: 24,
      zoom: 1,
      source: "resize"
    });

    expect(batchSpy).not.toHaveBeenCalled();

    gridView.destroy();
    container.remove();
  });

  it("does not persist axis resize mutations while global spreadsheet editing is active", () => {
    const container = document.createElement("div");
    Object.defineProperty(container, "clientWidth", { configurable: true, value: 0 });
    Object.defineProperty(container, "clientHeight", { configurable: true, value: 0 });
    document.body.appendChild(container);

    const doc = new DocumentController();
    const sheetId = "Sheet1";

    const gridView = new SecondaryGridView({
      container,
      document: doc,
      getSheetId: () => sheetId,
      rowCount: 100,
      colCount: 50,
      showFormulas: () => false,
      getComputedValue: () => null,
      getDrawingObjects: () => [],
      images,
    });

    // Simulate an active edit session elsewhere (e.g. primary formula bar / editor).
    (globalThis as any).__formulaSpreadsheetIsEditing = true;

    const setColWidth = vi.spyOn(doc, "setColWidth");
    const resetColWidth = vi.spyOn(doc, "resetColWidth");

    (gridView as any).onAxisSizeChange({
      kind: "col",
      index: 2,
      size: 40,
      previousSize: 24,
      defaultSize: 24,
      zoom: 1,
      source: "resize",
    });

    expect(setColWidth).not.toHaveBeenCalled();
    expect(resetColWidth).not.toHaveBeenCalled();

    gridView.destroy();
    container.remove();
  });
});
