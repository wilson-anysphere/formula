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

describe("SecondaryGridView no-resurrection", () => {
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

  it("does not resurrect deleted sheets when syncing sheet view state", () => {
    const container = document.createElement("div");
    // Keep viewport 0-sized so the renderer doesn't do any expensive work in jsdom.
    Object.defineProperty(container, "clientWidth", { configurable: true, value: 0 });
    Object.defineProperty(container, "clientHeight", { configurable: true, value: 0 });
    document.body.appendChild(container);

    const doc = new DocumentController();
    doc.setCellValue("Sheet1", "A1", "one");
    doc.setCellValue("Sheet2", "A1", "two");
    expect(doc.getSheetIds()).toEqual(["Sheet1", "Sheet2"]);

    doc.deleteSheet("Sheet2");
    expect(doc.getSheetIds()).toEqual(["Sheet1"]);

    let activeSheetId = "Sheet1";
    const gridView = new SecondaryGridView({
      container,
      document: doc,
      getSheetId: () => activeSheetId,
      rowCount: 100,
      colCount: 50,
      showFormulas: () => false,
      getComputedValue: () => null,
      getDrawingObjects: () => [],
      images
    });

    // Simulate a stale sheet id (e.g. during sheet delete undo/redo window).
    activeSheetId = "Sheet2";
    gridView.syncSheetViewFromDocument();

    expect(doc.getSheetIds()).toEqual(["Sheet1"]);

    gridView.destroy();
    container.remove();
  });
});

