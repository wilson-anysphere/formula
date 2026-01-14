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

describe("SecondaryGridView fill edit rejection toasts", () => {
  afterEach(() => {
    delete (globalThis as any).__formulaSpreadsheetIsEditing;
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
  });

  beforeEach(() => {
    document.body.innerHTML = `<div id=\"toast-root\"></div>`;

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

  it("shows an encryption-aware toast and restores selection when fill applies 0 edits due to canEditCell", async () => {
    const container = document.createElement("div");
    // Keep viewport 0-sized so the renderer doesn't do any expensive work in jsdom.
    Object.defineProperty(container, "clientWidth", { configurable: true, value: 0 });
    Object.defineProperty(container, "clientHeight", { configurable: true, value: 0 });
    document.body.appendChild(container);

    const doc = new DocumentController();
    // Seed A1 so the fill would normally apply.
    doc.setCellInput("Sheet1", { row: 0, col: 0 }, 1);
    // Block all writes.
    (doc as any).canEditCell = () => false;

    const selectionChange = vi.fn();
    const gridView = new SecondaryGridView({
      container,
      document: doc,
      getSheetId: () => "Sheet1",
      rowCount: 100,
      colCount: 50,
      showFormulas: () => false,
      getComputedValue: () => null,
      getDrawingObjects: () => [],
      images,
      onSelectionChange: selectionChange,
      onSelectionRangeChange: selectionChange,
      inferCollabEditRejection: () => ({ rejectionReason: "encryption" }),
    });

    // Ensure a stable pre-fill selection exists.
    const initialRanges = [{ startRow: 1, endRow: 2, startCol: 1, endCol: 2 }];
    gridView.grid.setSelectionRanges(initialRanges, { activeIndex: 0, activeCell: { row: 1, col: 1 }, scrollIntoView: false });
    selectionChange.mockClear();

    const beginBatch = vi.spyOn(doc, "beginBatch");
    const setCellInput = vi.spyOn(doc, "setCellInput");

    // Trigger a simple fill commit (A1 -> A2). Grid ranges include a 1-row/1-col header at index 0.
    (gridView as any).onFillCommit({
      sourceRange: { startRow: 1, endRow: 2, startCol: 1, endCol: 2 },
      targetRange: { startRow: 2, endRow: 3, startCol: 1, endCol: 2 },
      mode: "formulas",
    });

    // Simulate DesktopSharedGrid expanding selection after the callback returns.
    gridView.grid.setSelectionRanges(
      [{ startRow: 1, endRow: 3, startCol: 1, endCol: 2 }],
      { activeIndex: 0, activeCell: { row: 1, col: 1 }, scrollIntoView: false },
    );

    // Flush queued microtasks (selection restore).
    await new Promise<void>((resolve) => queueMicrotask(resolve));

    expect(beginBatch).not.toHaveBeenCalled();
    expect(setCellInput).not.toHaveBeenCalled();
    expect(selectionChange).not.toHaveBeenCalled();
    expect((gridView as any).suppressSelectionCallbacks).toBe(false);

    // Selection should be restored to the pre-fill state.
    expect(gridView.grid.renderer.getSelectionRanges()).toEqual(initialRanges);

    const content = document.querySelector("#toast-root")?.textContent ?? "";
    expect(content).toContain("Missing encryption key");

    gridView.destroy();
    container.remove();
  });
});

