// @vitest-environment jsdom

import { act } from "react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { SpreadsheetApp } from "../../../app/spreadsheetApp";
import type { DrawingObject } from "../../../drawings/types";

// React 18 relies on this flag to suppress act() warnings in test runners.
// eslint-disable-next-line @typescript-eslint/no-explicit-any
(globalThis as any).IS_REACT_ACT_ENVIRONMENT = true;

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
      width: 1200,
      height: 800,
      left: 0,
      top: 0,
      right: 1200,
      bottom: 800,
      x: 0,
      y: 0,
      toJSON: () => {},
    }) as any;
  document.body.appendChild(root);
  return root;
}

describe("Selection Pane panel", () => {
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

  afterEach(() => {
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
    document.body.innerHTML = "";
  });

  it("opens via ribbon command and lists drawings; clicking selects a drawing", async () => {
    const [{ createPanelBodyRenderer }, { PanelIds }, { mountRibbon }] = await Promise.all([
      import("../../panelBodyRenderer.js"),
      import("../../panelRegistry.js"),
      import("../../../ribbon/index.js"),
    ]);

    const sheetRoot = createRoot();
    const app = new SpreadsheetApp(sheetRoot, {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    });

    const drawings: DrawingObject[] = [
      {
        id: 1,
        kind: { type: "image", imageId: "img_1" },
        anchor: { type: "absolute", pos: { xEmu: 0, yEmu: 0 }, size: { cx: 10, cy: 10 } },
        zOrder: 0,
      },
      {
        id: 2,
        kind: { type: "image", imageId: "img_2" },
        anchor: { type: "absolute", pos: { xEmu: 0, yEmu: 0 }, size: { cx: 10, cy: 10 } },
        zOrder: 0,
      },
      {
        id: 3,
        kind: { type: "image", imageId: "img_3" },
        anchor: { type: "absolute", pos: { xEmu: 0, yEmu: 0 }, size: { cx: 10, cy: 10 } },
        zOrder: 1,
      },
    ];

    const sheetId = app.getCurrentSheetId();
    // Seed drawings via the DocumentController drawing layer so SpreadsheetApp's drawing caches update.
    (app.getDocument() as any).setSheetDrawings(sheetId, drawings);

    const panelBody = document.createElement("div");
    document.body.appendChild(panelBody);

    const panelBodyRenderer = createPanelBodyRenderer({
      getDocumentController: () => app.getDocument(),
      getSpreadsheetApp: () => app,
    });

    const ribbonRoot = document.createElement("div");
    document.body.appendChild(ribbonRoot);

    let unmountRibbon: (() => void) | null = null;
    await act(async () => {
      unmountRibbon = mountRibbon(
        ribbonRoot,
        {
          onCommand: (commandId: string) => {
            if (commandId !== "pageLayout.arrange.selectionPane") return;
            panelBodyRenderer.renderPanelBody(PanelIds.SELECTION_PANE, panelBody);
          },
        },
        { initialTabId: "pageLayout" },
      );
    });

    const commandButton = ribbonRoot.querySelector<HTMLButtonElement>('button[data-command-id="pageLayout.arrange.selectionPane"]');
    expect(commandButton).toBeInstanceOf(HTMLButtonElement);

    await act(async () => {
      commandButton!.click();
    });

    const itemEls = panelBody.querySelectorAll('[data-testid^="selection-pane-item-"]');
    expect(itemEls.length).toBe(3);
    // Topmost first (highest z-order -> id=3). Within ties, reverse render order (id=2 before id=1).
    expect(itemEls[0]?.getAttribute("data-testid")).toBe("selection-pane-item-3");
    expect(itemEls[1]?.getAttribute("data-testid")).toBe("selection-pane-item-2");
    expect(itemEls[2]?.getAttribute("data-testid")).toBe("selection-pane-item-1");

    await act(async () => {
      (itemEls[2] as HTMLElement).click();
    });

    expect(app.getSelectedDrawingId()).toBe(1);
    expect(itemEls[2]?.getAttribute("aria-selected")).toBe("true");

    // Bring Forward should update z-order and re-render the list (Picture 1 moves one step forward,
    // swapping above Picture 2 while still staying below the topmost drawing).
    const bringForwardBtn = panelBody.querySelector<HTMLButtonElement>('[data-testid="selection-pane-bring-forward-1"]');
    expect(bringForwardBtn).toBeInstanceOf(HTMLButtonElement);
    await act(async () => {
      bringForwardBtn!.click();
    });
    const reorderedItemEls = panelBody.querySelectorAll('[data-testid^="selection-pane-item-"]');
    expect(reorderedItemEls.length).toBe(3);
    expect(reorderedItemEls[0]?.getAttribute("data-testid")).toBe("selection-pane-item-3");
    expect(reorderedItemEls[1]?.getAttribute("data-testid")).toBe("selection-pane-item-1");
    expect(reorderedItemEls[2]?.getAttribute("data-testid")).toBe("selection-pane-item-2");

    // Adding a drawing should update the panel list via subscribeDrawings.
    const currentDrawings = (app.getDocument() as any).getSheetDrawings(sheetId);
    const nextDrawings: DrawingObject[] = [
      ...(Array.isArray(currentDrawings) ? currentDrawings : []),
      {
        id: 4,
        kind: { type: "shape", label: "Shape 4" },
        anchor: { type: "absolute", pos: { xEmu: 0, yEmu: 0 }, size: { cx: 10, cy: 10 } },
        zOrder: 2,
      },
    ];

    await act(async () => {
      (app.getDocument() as any).setSheetDrawings(sheetId, nextDrawings);
    });

    const updatedItemEls = panelBody.querySelectorAll('[data-testid^="selection-pane-item-"]');
    expect(updatedItemEls.length).toBe(4);
    expect(updatedItemEls[0]?.getAttribute("data-testid")).toBe("selection-pane-item-4");

    await act(async () => {
      unmountRibbon?.();
      panelBodyRenderer.cleanup([]);
    });
    app.destroy();
    sheetRoot.remove();
  });

  it("deletes drawings via per-row Delete button", async () => {
    const [{ createPanelBodyRenderer }, { PanelIds }, { mountRibbon }] = await Promise.all([
      import("../../panelBodyRenderer.js"),
      import("../../panelRegistry.js"),
      import("../../../ribbon/index.js"),
    ]);

    const sheetRoot = createRoot();
    const app = new SpreadsheetApp(sheetRoot, {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    });

    const drawings: DrawingObject[] = [
      {
        id: 1,
        kind: { type: "image", imageId: "img_1" },
        anchor: { type: "absolute", pos: { xEmu: 0, yEmu: 0 }, size: { cx: 10, cy: 10 } },
        zOrder: 0,
      },
      {
        id: 2,
        kind: { type: "image", imageId: "img_2" },
        anchor: { type: "absolute", pos: { xEmu: 0, yEmu: 0 }, size: { cx: 10, cy: 10 } },
        zOrder: 1,
      },
    ];

    const sheetId = app.getCurrentSheetId();
    (app.getDocument() as any).setSheetDrawings(sheetId, drawings);

    const panelBody = document.createElement("div");
    document.body.appendChild(panelBody);

    const panelBodyRenderer = createPanelBodyRenderer({
      getDocumentController: () => app.getDocument(),
      getSpreadsheetApp: () => app,
    });

    const ribbonRoot = document.createElement("div");
    document.body.appendChild(ribbonRoot);

    let unmountRibbon: (() => void) | null = null;
    await act(async () => {
      unmountRibbon = mountRibbon(
        ribbonRoot,
        {
          onCommand: (commandId: string) => {
            if (commandId !== "pageLayout.arrange.selectionPane") return;
            panelBodyRenderer.renderPanelBody(PanelIds.SELECTION_PANE, panelBody);
          },
        },
        { initialTabId: "pageLayout" },
      );
    });

    const commandButton = ribbonRoot.querySelector<HTMLButtonElement>('button[data-command-id="pageLayout.arrange.selectionPane"]');
    expect(commandButton).toBeInstanceOf(HTMLButtonElement);

    await act(async () => {
      commandButton!.click();
    });

    const deleteBtn = panelBody.querySelector<HTMLButtonElement>('[data-testid="selection-pane-delete-2"]');
    expect(deleteBtn).toBeInstanceOf(HTMLButtonElement);
    await act(async () => {
      deleteBtn!.click();
    });

    const remainingItemEls = panelBody.querySelectorAll('[data-testid^="selection-pane-item-"]');
    expect(remainingItemEls.length).toBe(1);
    expect(remainingItemEls[0]?.getAttribute("data-testid")).toBe("selection-pane-item-1");

    const raw = (app.getDocument() as any).getSheetDrawings(sheetId);
    expect(Array.isArray(raw)).toBe(true);
    expect(raw.length).toBe(1);
    expect(String(raw[0]?.id)).toBe("1");

    await act(async () => {
      unmountRibbon?.();
      panelBodyRenderer.cleanup([]);
    });
    app.destroy();
    sheetRoot.remove();
  });
});
