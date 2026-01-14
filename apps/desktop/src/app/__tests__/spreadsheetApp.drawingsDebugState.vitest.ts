/**
 * @vitest-environment jsdom
 */

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

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

describe("SpreadsheetApp drawings debug state", () => {
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

    Object.defineProperty(HTMLCanvasElement.prototype, "getContext", {
      configurable: true,
      value: () => createMockCanvasContext(),
    });

    (globalThis as any).ResizeObserver = class {
      observe() {}
      disconnect() {}
    };
  });

  it("returns non-null picture rects in legacy grid mode", async () => {
    const prior = process.env.DESKTOP_GRID_MODE;
    process.env.DESKTOP_GRID_MODE = "legacy";
    try {
      const root = createRoot();
      const status = {
        activeCell: document.createElement("div"),
        selectionRange: document.createElement("div"),
        activeValue: document.createElement("div"),
      };
      const app = new SpreadsheetApp(root, status);

      const file = new File([new Uint8Array([1, 2, 3])], "cat.png", { type: "image/png" });
      await app.insertPicturesFromFiles([file], { placeAt: { row: 0, col: 0 } });

      const state = app.getDrawingsDebugState();
      expect(state.sheetId).toBe(app.getCurrentSheetId());
      expect(state.drawings.some((d) => d.kind === "image")).toBe(true);

      const drawing = state.drawings.find((d) => d.kind === "image")!;
      expect(drawing.rectPx).not.toBeNull();
      expect(drawing.rectPx?.width).toBeGreaterThan(0);
      expect(drawing.rectPx?.height).toBeGreaterThan(0);

      expect(app.getDrawingRectPx(drawing.id)).not.toBeNull();
      expect(app.getDrawingHandlePointsPx(drawing.id)).not.toBeNull();

      app.destroy();
      root.remove();
    } finally {
      if (prior === undefined) delete process.env.DESKTOP_GRID_MODE;
      else process.env.DESKTOP_GRID_MODE = prior;
    }
  });

  it("returns non-null picture rects in shared grid mode", async () => {
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

      const file = new File([new Uint8Array([1, 2, 3])], "cat.png", { type: "image/png" });
      await app.insertPicturesFromFiles([file], { placeAt: { row: 0, col: 0 } });

      const state = app.getDrawingsDebugState();
      expect(state.sheetId).toBe(app.getCurrentSheetId());
      expect(state.drawings.some((d) => d.kind === "image")).toBe(true);

      const drawing = state.drawings.find((d) => d.kind === "image")!;
      expect(drawing.rectPx).not.toBeNull();
      expect(drawing.rectPx?.width).toBeGreaterThan(0);
      expect(drawing.rectPx?.height).toBeGreaterThan(0);

      expect(app.getDrawingRectPx(drawing.id)).not.toBeNull();
      expect(app.getDrawingHandlePointsPx(drawing.id)).not.toBeNull();

      app.destroy();
      root.remove();
    } finally {
      if (prior === undefined) delete process.env.DESKTOP_GRID_MODE;
      else process.env.DESKTOP_GRID_MODE = prior;
    }
  });

  it("accounts for shared-grid zoom when computing drawing rects", async () => {
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

      const file = new File([new Uint8Array([1, 2, 3])], "cat.png", { type: "image/png" });
      await app.insertPicturesFromFiles([file], { placeAt: { row: 0, col: 0 } });

      const drawingId = app.getSelectedDrawingId();
      expect(drawingId).not.toBeNull();
      const rect1 = app.getDrawingRectPx(drawingId);
      expect(rect1).not.toBeNull();

      app.setZoom(2);
      const rect2 = app.getDrawingRectPx(drawingId);
      expect(rect2).not.toBeNull();

      // The on-screen size should scale proportionally with zoom.
      expect(rect2!.width).toBeGreaterThan(rect1!.width);
      expect(rect2!.height).toBeGreaterThan(rect1!.height);
      expect(rect2!.width / rect1!.width).toBeCloseTo(2, 1);
      expect(rect2!.height / rect1!.height).toBeCloseTo(2, 1);

      app.destroy();
      root.remove();
    } finally {
      if (prior === undefined) delete process.env.DESKTOP_GRID_MODE;
      else process.env.DESKTOP_GRID_MODE = prior;
    }
  });

  it("respects frozen panes when computing drawing rects (shared grid)", async () => {
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

      // Freeze first row/col (sheet space).
      const doc = app.getDocument();
      doc.setFrozen(app.getCurrentSheetId(), 1, 1, { label: "Freeze" });

      const file1 = new File([new Uint8Array([1, 2, 3])], "a.png", { type: "image/png" });
      const file2 = new File([new Uint8Array([4, 5, 6])], "b.png", { type: "image/png" });

      // Insert one picture in the frozen top-left pane and one in the scrollable pane.
      await app.insertPicturesFromFiles([file1], { placeAt: { row: 0, col: 0 } }); // A1 (frozen)
      const frozenId = app.getSelectedDrawingId();
      expect(frozenId).not.toBeNull();

      await app.insertPicturesFromFiles([file2], { placeAt: { row: 1, col: 1 } }); // B2 (scrollable)
      const scrollableId = app.getSelectedDrawingId();
      expect(scrollableId).not.toBeNull();
      expect(scrollableId).not.toBe(frozenId);

      const frozenBefore = app.getDrawingRectPx(frozenId!);
      const scrollableBefore = app.getDrawingRectPx(scrollableId!);
      expect(frozenBefore).not.toBeNull();
      expect(scrollableBefore).not.toBeNull();

      app.setScroll(100, 100);

      const frozenAfter = app.getDrawingRectPx(frozenId!);
      const scrollableAfter = app.getDrawingRectPx(scrollableId!);
      expect(frozenAfter).not.toBeNull();
      expect(scrollableAfter).not.toBeNull();

      // Frozen-pane object should not move when the sheet scrolls.
      expect(frozenAfter!.x).toBeCloseTo(frozenBefore!.x, 6);
      expect(frozenAfter!.y).toBeCloseTo(frozenBefore!.y, 6);

      // Scrollable object should move with the scroll position.
      expect(scrollableAfter!.x).toBeLessThan(scrollableBefore!.x);
      expect(scrollableAfter!.y).toBeLessThan(scrollableBefore!.y);

      app.destroy();
      root.remove();
    } finally {
      if (prior === undefined) delete process.env.DESKTOP_GRID_MODE;
      else process.env.DESKTOP_GRID_MODE = prior;
    }
  });

  it("respects frozen panes when computing drawing rects (legacy grid)", async () => {
    const prior = process.env.DESKTOP_GRID_MODE;
    process.env.DESKTOP_GRID_MODE = "legacy";
    try {
      const root = createRoot();
      const status = {
        activeCell: document.createElement("div"),
        selectionRange: document.createElement("div"),
        activeValue: document.createElement("div"),
      };
      const app = new SpreadsheetApp(root, status);

      const doc = app.getDocument();
      doc.setFrozen(app.getCurrentSheetId(), 1, 1, { label: "Freeze" });

      const file1 = new File([new Uint8Array([1, 2, 3])], "a.png", { type: "image/png" });
      const file2 = new File([new Uint8Array([4, 5, 6])], "b.png", { type: "image/png" });

      await app.insertPicturesFromFiles([file1], { placeAt: { row: 0, col: 0 } }); // A1 (frozen)
      const frozenId = app.getSelectedDrawingId();
      expect(frozenId).not.toBeNull();

      await app.insertPicturesFromFiles([file2], { placeAt: { row: 1, col: 1 } }); // B2 (scrollable)
      const scrollableId = app.getSelectedDrawingId();
      expect(scrollableId).not.toBeNull();
      expect(scrollableId).not.toBe(frozenId);

      const frozenBefore = app.getDrawingRectPx(frozenId!);
      const scrollableBefore = app.getDrawingRectPx(scrollableId!);
      expect(frozenBefore).not.toBeNull();
      expect(scrollableBefore).not.toBeNull();

      app.setScroll(100, 100);

      const frozenAfter = app.getDrawingRectPx(frozenId!);
      const scrollableAfter = app.getDrawingRectPx(scrollableId!);
      expect(frozenAfter).not.toBeNull();
      expect(scrollableAfter).not.toBeNull();

      expect(frozenAfter!.x).toBeCloseTo(frozenBefore!.x, 6);
      expect(frozenAfter!.y).toBeCloseTo(frozenBefore!.y, 6);

      expect(scrollableAfter!.x).toBeLessThan(scrollableBefore!.x);
      expect(scrollableAfter!.y).toBeLessThan(scrollableBefore!.y);

      app.destroy();
      root.remove();
    } finally {
      if (prior === undefined) delete process.env.DESKTOP_GRID_MODE;
      else process.env.DESKTOP_GRID_MODE = prior;
    }
  });
});
