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

    // Enable drawing interactions so the split-view secondary pane can use its dedicated
    // DrawingInteractionController for drag/resize/rotate gestures.
    const app = new SpreadsheetApp(root, status, { enableDrawingInteractions: true });
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

    return { app, root, secondaryView, secondaryContainer, selectionCanvas, beforeActive, drawing };
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

  it("preserves drawing selection on context-click misses in the secondary pane", () => {
    const { app, root, secondaryView, secondaryContainer, selectionCanvas } = setup();

    // Select the drawing first (e.g. via the selection pane).
    app.selectDrawingById(1);
    expect(app.getSelectedDrawingId()).toBe(1);

    const headerOffsetX = secondaryView.grid.renderer.scroll.cols.totalSize(1);
    const headerOffsetY = secondaryView.grid.renderer.scroll.rows.totalSize(1);

    const rect = secondaryContainer.getBoundingClientRect();
    const missClientX = rect.left + headerOffsetX + 200;
    const missClientY = rect.top + headerOffsetY + 200;
    expect(app.hitTestDrawingAtClientPoint(missClientX, missClientY)).toBeNull();

    const down = createPointerLikeMouseEvent("pointerdown", {
      clientX: missClientX,
      clientY: missClientY,
      button: 2,
    });
    selectionCanvas.dispatchEvent(down);

    expect(app.getSelectedDrawingId()).toBe(1);
    expect((down as any).__formulaDrawingContextClick).toBeUndefined();
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

  it("clears the selected drawing when clicking the secondary header area", () => {
    const { app, root, secondaryView, secondaryContainer, selectionCanvas } = setup();

    const headerOffsetX = secondaryView.grid.renderer.scroll.cols.totalSize(1);
    const headerOffsetY = secondaryView.grid.renderer.scroll.rows.totalSize(1);

    // Select the drawing first.
    selectionCanvas.dispatchEvent(
      createPointerLikeMouseEvent("pointerdown", {
        clientX: secondaryContainer.getBoundingClientRect().left + headerOffsetX + 60,
        clientY: secondaryContainer.getBoundingClientRect().top + headerOffsetY + 30,
        button: 0,
      }),
    );
    expect(app.getSelectedDrawingId()).toBe(1);

    // Click in the secondary column header area. This should behave like "click outside any drawing":
    // clear drawing selection but allow the grid header selection logic to proceed.
    selectionCanvas.dispatchEvent(
      createPointerLikeMouseEvent("pointerdown", {
        clientX: secondaryContainer.getBoundingClientRect().left + headerOffsetX + 10,
        clientY: secondaryContainer.getBoundingClientRect().top + 2,
        button: 0,
      }),
    );
    expect(app.getSelectedDrawingId()).toBeNull();

    secondaryView.destroy();
    app.destroy();
    secondaryContainer.remove();
    root.remove();
  });

  it("shows drawing hover cursors in the secondary pane", () => {
    const { app, root, secondaryView, secondaryContainer, selectionCanvas } = setup();

    const headerOffsetX = secondaryView.grid.renderer.scroll.cols.totalSize(1);
    const headerOffsetY = secondaryView.grid.renderer.scroll.rows.totalSize(1);

    // Hover inside the drawing bounds (absolute anchor at 0,0 with size 100x100).
    const insideX = headerOffsetX + 10;
    const insideY = headerOffsetY + 10;
    const inside = createPointerLikeMouseEvent("pointermove", {
      clientX: secondaryContainer.getBoundingClientRect().left + insideX,
      clientY: secondaryContainer.getBoundingClientRect().top + insideY,
      button: 0,
    });
    Object.defineProperty(inside, "offsetX", { configurable: true, value: insideX });
    Object.defineProperty(inside, "offsetY", { configurable: true, value: insideY });
    selectionCanvas.dispatchEvent(inside);

    expect(selectionCanvas.style.cursor).toBe("move");
    expect(secondaryContainer.style.cursor).toBe("move");

    // Hover over an empty area; shared-grid hover logic should restore the default cursor.
    const outsideX = headerOffsetX + 200;
    const outsideY = headerOffsetY + 200;
    const outside = createPointerLikeMouseEvent("pointermove", {
      clientX: secondaryContainer.getBoundingClientRect().left + outsideX,
      clientY: secondaryContainer.getBoundingClientRect().top + outsideY,
      button: 0,
    });
    Object.defineProperty(outside, "offsetX", { configurable: true, value: outsideX });
    Object.defineProperty(outside, "offsetY", { configurable: true, value: outsideY });
    selectionCanvas.dispatchEvent(outside);

    expect(selectionCanvas.style.cursor).toBe("default");
    expect(secondaryContainer.style.cursor).toBe("");

    secondaryView.destroy();
    app.destroy();
    secondaryContainer.remove();
    root.remove();
  });

  it("moves drawings via drag gestures in the secondary pane", () => {
    const { app, root, secondaryView, secondaryContainer, selectionCanvas, drawing } = setup();

    // Persist the drawing into the DocumentController so interaction commits update the
    // underlying model (SpreadsheetApp clears its in-memory cache at commit time).
    const sheetId = app.getCurrentSheetId();
    (app.getDocument() as any).setSheetDrawings(sheetId, [drawing], { label: "Set Drawings" });

    const headerOffsetX = secondaryView.grid.renderer.scroll.cols.totalSize(1);
    const headerOffsetY = secondaryView.grid.renderer.scroll.rows.totalSize(1);

    const rect = secondaryContainer.getBoundingClientRect();

    const startX = headerOffsetX + 40;
    const startY = headerOffsetY + 40;
    const dx = 30;
    const dy = 20;
    const endX = startX + dx;
    const endY = startY + dy;

    const down = createPointerLikeMouseEvent("pointerdown", {
      clientX: rect.left + startX,
      clientY: rect.top + startY,
      button: 0,
      pointerId: 1,
    });
    selectionCanvas.dispatchEvent(down);

    const move = createPointerLikeMouseEvent("pointermove", {
      clientX: rect.left + endX,
      clientY: rect.top + endY,
      button: 0,
      pointerId: 1,
    });
    selectionCanvas.dispatchEvent(move);

    const up = createPointerLikeMouseEvent("pointerup", {
      clientX: rect.left + endX,
      clientY: rect.top + endY,
      button: 0,
      pointerId: 1,
    });
    selectionCanvas.dispatchEvent(up);

    const persisted = (app.getDocument() as any).getSheetDrawings(sheetId) as any[];
    const next = persisted.find((d) => d?.id === 1) ?? null;
    expect(next).not.toBeNull();

    const zoom = secondaryView.grid.renderer.getZoom();
    const expectedDxEmu = pxToEmu(dx, zoom);
    const expectedDyEmu = pxToEmu(dy, zoom);

    expect(next.anchor?.type).toBe("absolute");
    expect(next.anchor?.pos?.xEmu).toBeCloseTo(expectedDxEmu, 4);
    expect(next.anchor?.pos?.yEmu).toBeCloseTo(expectedDyEmu, 4);

    secondaryView.destroy();
    app.destroy();
    secondaryContainer.remove();
    root.remove();
  });

  it("cancels an in-progress drag when switching sheets (secondary pane)", () => {
    const { app, root, secondaryView, secondaryContainer, selectionCanvas, drawing } = setup();

    const doc: any = app.getDocument();
    const sheet1 = app.getCurrentSheetId();

    // Persist the drawing into Sheet1 so a commit would normally update the document.
    doc.setSheetDrawings(sheet1, [drawing], { label: "Set Drawings" });

    // Ensure Sheet2 exists and contains a drawing with the same id. Without canceling the in-flight
    // gesture, a pointerup after `sheetId` changes could commit into the wrong sheet.
    doc.setCellValue("Sheet2", { row: 0, col: 0 }, "X");
    const sheet2Drawing = {
      ...drawing,
      kind: { type: "image", imageId: "img-2" },
      anchor: {
        type: "absolute",
        pos: { xEmu: pxToEmu(10), yEmu: pxToEmu(10) },
        size: { cx: pxToEmu(100), cy: pxToEmu(100) },
      },
      zOrder: 0,
    } satisfies DrawingObject;
    doc.setSheetDrawings("Sheet2", [sheet2Drawing], { label: "Set Drawings" });
    const sheet2AnchorBefore = JSON.parse(JSON.stringify(sheet2Drawing.anchor));

    const headerOffsetX = secondaryView.grid.renderer.scroll.cols.totalSize(1);
    const headerOffsetY = secondaryView.grid.renderer.scroll.rows.totalSize(1);
    const rect = secondaryContainer.getBoundingClientRect();

    const startX = headerOffsetX + 40;
    const startY = headerOffsetY + 40;
    const endX = startX + 30;
    const endY = startY + 20;

    selectionCanvas.dispatchEvent(
      createPointerLikeMouseEvent("pointerdown", {
        clientX: rect.left + startX,
        clientY: rect.top + startY,
        button: 0,
        pointerId: 1,
      }),
    );
    selectionCanvas.dispatchEvent(
      createPointerLikeMouseEvent("pointermove", {
        clientX: rect.left + endX,
        clientY: rect.top + endY,
        button: 0,
        pointerId: 1,
      }),
    );

    // Switch sheets while the pointer is still down (mid-gesture). This should cancel the gesture so
    // the eventual pointerup cannot commit into the newly active sheet.
    app.activateSheet("Sheet2");

    selectionCanvas.dispatchEvent(
      createPointerLikeMouseEvent("pointerup", {
        clientX: rect.left + endX,
        clientY: rect.top + endY,
        button: 0,
        pointerId: 1,
      }),
    );

    const sheet2After = Array.isArray(doc.getSheetDrawings("Sheet2")) ? doc.getSheetDrawings("Sheet2") : [];
    expect(sheet2After).toHaveLength(1);
    expect((sheet2After[0] as any).anchor).toEqual(sheet2AnchorBefore);

    const sheet1After = Array.isArray(doc.getSheetDrawings(sheet1)) ? doc.getSheetDrawings(sheet1) : [];
    expect(sheet1After).toHaveLength(1);
    expect((sheet1After[0] as any).anchor?.type).toBe("absolute");
    expect((sheet1After[0] as any).anchor?.pos?.xEmu ?? 0).toBeCloseTo(0, 4);
    expect((sheet1After[0] as any).anchor?.pos?.yEmu ?? 0).toBeCloseTo(0, 4);

    secondaryView.destroy();
    app.destroy();
    secondaryContainer.remove();
    root.remove();
  });

  it("reverts an in-progress drag when the secondary pane is torn down", () => {
    const { app, root, secondaryView, secondaryContainer, selectionCanvas, drawing } = setup();

    const sheetId = app.getCurrentSheetId();
    (app.getDocument() as any).setSheetDrawings(sheetId, [drawing], { label: "Set Drawings" });

    const headerOffsetX = secondaryView.grid.renderer.scroll.cols.totalSize(1);
    const headerOffsetY = secondaryView.grid.renderer.scroll.rows.totalSize(1);
    const rect = secondaryContainer.getBoundingClientRect();

    const startX = headerOffsetX + 40;
    const startY = headerOffsetY + 40;
    const dx = 30;
    const dy = 20;
    const endX = startX + dx;
    const endY = startY + dy;

    const down = createPointerLikeMouseEvent("pointerdown", {
      clientX: rect.left + startX,
      clientY: rect.top + startY,
      button: 0,
      pointerId: 1,
    });
    selectionCanvas.dispatchEvent(down);

    const move = createPointerLikeMouseEvent("pointermove", {
      clientX: rect.left + endX,
      clientY: rect.top + endY,
      button: 0,
      pointerId: 1,
    });
    selectionCanvas.dispatchEvent(move);

    // Simulate the split-view secondary pane being removed mid-gesture (e.g. user closes split view).
    app.setSplitViewSecondaryGridView(null);

    // The drag gesture should be canceled (no commit), and any in-memory preview state should be reverted.
    const persisted = (app.getDocument() as any).getSheetDrawings(sheetId) as any[];
    const persistedDrawing = persisted.find((d) => d?.id === 1) ?? null;
    expect(persistedDrawing).not.toBeNull();
    expect(persistedDrawing.anchor?.pos?.xEmu).toBeCloseTo(0, 4);
    expect(persistedDrawing.anchor?.pos?.yEmu).toBeCloseTo(0, 4);

    const inMemory = app.sheetDrawings.find((d) => d.id === 1) ?? null;
    expect(inMemory).not.toBeNull();
    expect(inMemory!.anchor.type).toBe("absolute");
    expect((inMemory!.anchor as any).pos?.xEmu).toBeCloseTo(0, 4);
    expect((inMemory!.anchor as any).pos?.yEmu).toBeCloseTo(0, 4);

    secondaryView.destroy();
    app.destroy();
    secondaryContainer.remove();
    root.remove();
  });
});
