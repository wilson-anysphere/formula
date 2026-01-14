/**
 * @vitest-environment jsdom
 */

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import type { DrawingObject } from "../../drawings/types";
import { pxToEmu } from "../../drawings/overlay";
import { SecondaryGridView } from "../../grid/splitView/secondaryGridView";
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

function createRoot(rect: DOMRect): HTMLElement {
  const root = document.createElement("div");
  root.tabIndex = 0;
  root.getBoundingClientRect = vi.fn(() => rect);
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
    ctrlKey?: boolean;
    metaKey?: boolean;
  },
): MouseEvent {
  const event = new MouseEvent(type, {
    bubbles: true,
    cancelable: true,
    clientX: options.clientX,
    clientY: options.clientY,
    button: options.button,
    ctrlKey: options.ctrlKey,
    metaKey: options.metaKey,
  });
  Object.defineProperty(event, "pointerId", { configurable: true, value: options.pointerId ?? 1 });
  Object.defineProperty(event, "pointerType", { configurable: true, value: options.pointerType ?? "mouse" });
  return event;
}

describe("SpreadsheetApp drawings selection in split-view secondary pane (shared grid)", () => {
  let priorGridMode: string | undefined;

  beforeEach(() => {
    priorGridMode = process.env.DESKTOP_GRID_MODE;
    process.env.DESKTOP_GRID_MODE = "shared";

    document.body.innerHTML = "";

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

  afterEach(() => {
    if (priorGridMode === undefined) delete process.env.DESKTOP_GRID_MODE;
    else process.env.DESKTOP_GRID_MODE = priorGridMode;
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
  });

  function setup() {
    const primaryRect = {
      width: 800,
      height: 600,
      left: 0,
      top: 0,
      right: 800,
      bottom: 600,
      x: 0,
      y: 0,
      toJSON: () => {},
    } as DOMRect;
    const secondaryRect = {
      width: 800,
      height: 600,
      left: 1000,
      top: 0,
      right: 1800,
      bottom: 600,
      x: 1000,
      y: 0,
      toJSON: () => {},
    } as DOMRect;

    const root = createRoot(primaryRect);
    const secondaryContainer = createRoot(secondaryRect);

    Object.defineProperty(secondaryContainer, "clientWidth", { configurable: true, value: secondaryRect.width });
    Object.defineProperty(secondaryContainer, "clientHeight", { configurable: true, value: secondaryRect.height });

    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    // Drawing interactions are enabled by default in shared-grid mode; keep the default
    // so this test matches the desktop runtime.
    const app = new SpreadsheetApp(root, status);
    expect(app.getGridMode()).toBe("shared");

    // Wire a minimal "secondary selection → primary selection" sync (like `main.ts`).
    let suppressSync = true;
    const syncFromSecondarySelection = (selection: { row: number; col: number } | null) => {
      if (suppressSync) return;
      if (!selection) return;
      const docRow = selection.row - 1;
      const docCol = selection.col - 1;
      if (docRow < 0 || docCol < 0) return;
      app.activateCell({ row: docRow, col: docCol }, { scrollIntoView: false, focus: false });
    };

    const images = { get: () => undefined, set: () => {}, delete: () => {}, clear: () => {} };
    let secondaryView: SecondaryGridView;
    secondaryView = new SecondaryGridView({
      container: secondaryContainer,
      document: app.getDocument(),
      getSheetId: () => app.getCurrentSheetId(),
      rowCount: 30,
      colCount: 30,
      showFormulas: () => false,
      getComputedValue: () => null,
      getDrawingObjects: (sheetId) => app.getDrawingObjects(sheetId),
      images,
      getSelectedDrawingId: () => app.getSelectedDrawingId(),
      onSelectionChange: syncFromSecondarySelection,
    });

    // Ensure DesktopSharedGrid uses the same viewport origin as the container for pickCellAt.
    const selectionCanvas = secondaryContainer.querySelector<HTMLCanvasElement>("canvas.grid-canvas--selection");
    if (!selectionCanvas) {
      throw new Error("Missing secondary selection canvas");
    }
    selectionCanvas.getBoundingClientRect = secondaryContainer.getBoundingClientRect as any;

    app.setSplitViewSecondaryGridView(secondaryView);

    // Set an active cell away from A1 so selection changes are observable.
    app.activateCell({ row: 5, col: 5 }, { scrollIntoView: false, focus: false });
    const beforeActive = app.getActiveCell();

    // Mirror that selection into the secondary pane so the upcoming clicks land *outside*
    // the current selection (ensures DesktopSharedGrid would move selection without the fix).
    suppressSync = true;
    secondaryView.grid.setSelectionRanges(
      [{ startRow: 6, endRow: 7, startCol: 6, endCol: 7 }],
      { activeIndex: 0, activeCell: { row: 6, col: 6 }, scrollIntoView: false },
    );
    suppressSync = false;

    const drawing: DrawingObject = {
      id: 1,
      kind: { type: "image", imageId: "img-1" },
      anchor: {
        type: "absolute",
        pos: { xEmu: pxToEmu(0), yEmu: pxToEmu(0) },
        size: { cx: pxToEmu(100), cy: pxToEmu(100) },
      },
      zOrder: 0,
    };
    app.setDrawingObjects([drawing]);

    return { app, root, secondaryView, secondaryContainer, selectionCanvas, beforeActive };
  }

  it("selects the drawing and tags context-clicks on right click without moving the active cell", () => {
    const { app, root, secondaryView, secondaryContainer, selectionCanvas, beforeActive } = setup();

    // Secondary grid header sizes are fixed in SecondaryGridView constructor.
    const headerOffsetX = secondaryView.grid.renderer.scroll.cols.totalSize(1);
    const headerOffsetY = secondaryView.grid.renderer.scroll.rows.totalSize(1);

    const down = createPointerLikeMouseEvent("pointerdown", {
      clientX: secondaryContainer.getBoundingClientRect().left + headerOffsetX + 60,
      clientY: secondaryContainer.getBoundingClientRect().top + headerOffsetY + 30,
      button: 2,
    });
    selectionCanvas.dispatchEvent(down);

    expect(app.getSelectedDrawingId()).toBe(1);
    expect(app.getActiveCell()).toEqual(beforeActive);
    expect((down as any).__formulaDrawingContextClick).toBe(true);
    expect(down.defaultPrevented).toBe(false);

    secondaryView.destroy();
    app.destroy();
    secondaryContainer.remove();
    root.remove();
  });

  it("selects the drawing on left click without moving the active cell", () => {
    const { app, root, secondaryView, secondaryContainer, selectionCanvas, beforeActive } = setup();

    const headerOffsetX = secondaryView.grid.renderer.scroll.cols.totalSize(1);
    const headerOffsetY = secondaryView.grid.renderer.scroll.rows.totalSize(1);

    const down = createPointerLikeMouseEvent("pointerdown", {
      clientX: secondaryContainer.getBoundingClientRect().left + headerOffsetX + 60,
      clientY: secondaryContainer.getBoundingClientRect().top + headerOffsetY + 30,
      button: 0,
    });
    selectionCanvas.dispatchEvent(down);

    expect(app.getSelectedDrawingId()).toBe(1);
    expect(app.getActiveCell()).toEqual(beforeActive);
    expect((down as any).__formulaDrawingContextClick).toBeUndefined();
    expect(down.defaultPrevented).toBe(true);

    secondaryView.destroy();
    app.destroy();
    secondaryContainer.remove();
    root.remove();
  });
});
