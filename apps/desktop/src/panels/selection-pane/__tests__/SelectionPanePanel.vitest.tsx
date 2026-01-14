// @vitest-environment jsdom

import { act } from "react";
import { createRoot as createReactRoot } from "react-dom/client";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { SpreadsheetApp } from "../../../app/spreadsheetApp";
import { chartIdToDrawingId } from "../../../charts/chartDrawingAdapter";
import type { DrawingObject } from "../../../drawings/types";
import { SelectionPanePanel } from "../SelectionPanePanel";

// React 18 relies on this flag to suppress act() warnings in test runners.
// eslint-disable-next-line @typescript-eslint/no-explicit-any
(globalThis as any).IS_REACT_ACT_ENVIRONMENT = true;

let priorCanvasChartsEnv: string | undefined;
let priorUseCanvasChartsEnv: string | undefined;

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
    // Some test suites use fake timers and can leak them across files when a test aborts early.
    // SelectionPanePanel relies on real timers for React scheduling, so force real timers here.
    vi.useRealTimers();

    priorCanvasChartsEnv = process.env.CANVAS_CHARTS;
    priorUseCanvasChartsEnv = process.env.USE_CANVAS_CHARTS;
    // Most Selection Pane behavior is authored/expected in legacy chart mode. In canvas charts mode,
    // ChartStore charts are rendered as drawing objects and show up in the list alongside workbook
    // drawings. These tests force legacy chart mode; suites that need canvas charts should opt in
    // explicitly (e.g. `process.env.CANVAS_CHARTS = "1"`).
    process.env.CANVAS_CHARTS = "0";
    delete process.env.USE_CANVAS_CHARTS;
    document.body.innerHTML = "";

    // Avoid leaking URL params (e.g. `?canvasCharts=0`) between tests.
    try {
      const url = new URL(window.location.href);
      url.search = "";
      url.hash = "";
      window.history.replaceState(null, "", url.toString());
    } catch {
      // ignore history errors
    }

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

    if (priorCanvasChartsEnv === undefined) delete process.env.CANVAS_CHARTS;
    else process.env.CANVAS_CHARTS = priorCanvasChartsEnv;
    if (priorUseCanvasChartsEnv === undefined) delete process.env.USE_CANVAS_CHARTS;
    else process.env.USE_CANVAS_CHARTS = priorUseCanvasChartsEnv;
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
    // Some demo/dev configurations seed ChartStore charts. Remove any charts here so this test can
    // focus on workbook drawing objects (images) without being sensitive to optional demo fixtures.
    for (const chart of app.listCharts()) {
      (app as any).chartStore.deleteChart(chart.id);
    }

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
    // Topmost object first (highest z-order -> id=3). Within ties, reverse render order (id=2 before id=1).
    expect(itemEls[0]?.getAttribute("data-testid")).toBe("selection-pane-item-3");
    expect(itemEls[1]?.getAttribute("data-testid")).toBe("selection-pane-item-2");
    expect(itemEls[2]?.getAttribute("data-testid")).toBe("selection-pane-item-1");

    // Topmost object cannot be brought forward; backmost object cannot be sent backward.
    const bringForward3 = panelBody.querySelector<HTMLButtonElement>('[data-testid="selection-pane-bring-forward-3"]');
    expect(bringForward3).toBeInstanceOf(HTMLButtonElement);
    expect(bringForward3!.disabled).toBe(true);
    const sendBackward1 = panelBody.querySelector<HTMLButtonElement>('[data-testid="selection-pane-send-backward-1"]');
    expect(sendBackward1).toBeInstanceOf(HTMLButtonElement);
    expect(sendBackward1!.disabled).toBe(true);

    const picture1Row = panelBody.querySelector<HTMLElement>('[data-testid="selection-pane-item-1"]');
    expect(picture1Row).toBeInstanceOf(HTMLElement);
    await act(async () => {
      picture1Row!.click();
    });

    expect(app.getSelectedDrawingId()).toBe(1);
    expect(picture1Row?.getAttribute("aria-selected")).toBe("true");

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
    expect(updatedItemEls[1]?.getAttribute("data-testid")).toBe("selection-pane-item-3");

    await act(async () => {
      unmountRibbon?.();
      panelBodyRenderer.cleanup([]);
    });
    app.destroy();
    sheetRoot.remove();
  });

  it("supports keyboard navigation + deletion (Arrow keys / Delete)", async () => {
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
    for (const chart of app.listCharts()) {
      (app as any).chartStore.deleteChart(chart.id);
    }

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

    const paneRoot = panelBody.querySelector<HTMLElement>('[data-testid="selection-pane"]');
    expect(paneRoot).toBeInstanceOf(HTMLElement);
    paneRoot!.focus();

    const itemEls = panelBody.querySelectorAll('[data-testid^="selection-pane-item-"]');
    expect(itemEls.length).toBe(3);

    const picture1Row = panelBody.querySelector<HTMLElement>('[data-testid="selection-pane-item-1"]');
    expect(picture1Row).toBeInstanceOf(HTMLElement);
    await act(async () => {
      picture1Row!.click();
    });
    expect(app.getSelectedDrawingId()).toBe(1);

    // ArrowUp should move selection to the previous item in the list ordering.
    await act(async () => {
      paneRoot!.dispatchEvent(new KeyboardEvent("keydown", { key: "ArrowUp", bubbles: true }));
    });
    expect(app.getSelectedDrawingId()).toBe(2);

    // Delete should remove the selected drawing (Excel-style Selection Pane behavior).
    await act(async () => {
      paneRoot!.dispatchEvent(new KeyboardEvent("keydown", { key: "Delete", bubbles: true }));
    });

    const remainingItemEls = panelBody.querySelectorAll('[data-testid^="selection-pane-item-"]');
    expect(remainingItemEls.length).toBe(2);
    expect(Array.from(remainingItemEls).map((el) => el.getAttribute("data-testid"))).not.toContain("selection-pane-item-2");

    const raw = (app.getDocument() as any).getSheetDrawings(sheetId);
    expect(Array.isArray(raw)).toBe(true);
    expect(raw.some((drawing: any) => String(drawing?.id) === "2")).toBe(false);

    // Escape should return focus to the grid surface (without clearing selection state).
    await act(async () => {
      paneRoot!.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape", bubbles: true }));
    });
    expect(document.activeElement).toBe(sheetRoot);

    await act(async () => {
      unmountRibbon?.();
      panelBodyRenderer.cleanup([]);
    });
    app.destroy();
    sheetRoot.remove();
  });

  it("Delete key deletes selected drawing even when an action button is focused", async () => {
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
    for (const chart of app.listCharts()) {
      (app as any).chartStore.deleteChart(chart.id);
    }

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

    const item2 = panelBody.querySelector<HTMLElement>('[data-testid="selection-pane-item-2"]');
    expect(item2).toBeInstanceOf(HTMLElement);
    await act(async () => {
      item2!.click();
    });
    expect(app.getSelectedDrawingId()).toBe(2);

    // Use an enabled action button; item 2 is topmost so Bring Forward is disabled.
    const sendBackward2 = panelBody.querySelector<HTMLButtonElement>('[data-testid="selection-pane-send-backward-2"]');
    expect(sendBackward2).toBeInstanceOf(HTMLButtonElement);
    sendBackward2!.focus();
    expect(document.activeElement).toBe(sendBackward2);

    await act(async () => {
      sendBackward2!.dispatchEvent(new KeyboardEvent("keydown", { key: "Delete", bubbles: true }));
    });

    const raw = (app.getDocument() as any).getSheetDrawings(sheetId);
    expect(Array.isArray(raw)).toBe(true);
    expect(raw.some((drawing: any) => String(drawing?.id) === "2")).toBe(false);

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
    for (const chart of app.listCharts()) {
      (app as any).chartStore.deleteChart(chart.id);
    }

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
    expect(Array.from(remainingItemEls).map((el) => el.getAttribute("data-testid"))).toContain("selection-pane-item-1");

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

  it("lists ChartStore charts when canvas charts are enabled; clicking selects and delete removes them", async () => {
    const [{ createPanelBodyRenderer }, { PanelIds }, { mountRibbon }] = await Promise.all([
      import("../../panelBodyRenderer.js"),
      import("../../panelRegistry.js"),
      import("../../../ribbon/index.js"),
    ]);

    const url = new URL(window.location.href);
    url.searchParams.set("canvasCharts", "1");
    window.history.replaceState(null, "", url.toString());

    const sheetRoot = createRoot();
    const app = new SpreadsheetApp(sheetRoot, {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    });

    const { chart_id: chartId } = app.addChart({ chart_type: "bar", data_range: "A1:B2", title: "Test Chart" });
    const drawingId = chartIdToDrawingId(chartId);

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

    const chartRow = panelBody.querySelector<HTMLElement>(`[data-testid="selection-pane-item-${drawingId}"]`);
    expect(chartRow).toBeInstanceOf(HTMLElement);

    const selectionChangedIds: Array<number | null> = [];
    const onSelectionChanged = (evt: Event) => {
      selectionChangedIds.push(((evt as CustomEvent)?.detail as any)?.id ?? null);
    };
    window.addEventListener("formula:drawing-selection-changed", onSelectionChanged as EventListener);

    try {
      await act(async () => {
        chartRow!.click();
      });

      expect(app.getSelectedChartId()).toBe(chartId);
      expect(app.getSelectedDrawingId()).toBe(drawingId);
      // Canvas-chart selection should also emit the window-level drawing selection event with the
      // effective selected drawing id (not null).
      expect(selectionChangedIds).toContain(drawingId);
      expect(panelBody.querySelector(`[data-testid="selection-pane-item-${drawingId}"]`)?.getAttribute("aria-selected")).toBe("true");

      const deleteBtn = panelBody.querySelector<HTMLButtonElement>(`[data-testid="selection-pane-delete-${drawingId}"]`);
      expect(deleteBtn).toBeInstanceOf(HTMLButtonElement);
      await act(async () => {
        deleteBtn!.click();
      });

      expect(app.listCharts().some((c) => c.id === chartId)).toBe(false);
      expect(panelBody.querySelector(`[data-testid="selection-pane-item-${drawingId}"]`)).toBeNull();
    } finally {
      window.removeEventListener("formula:drawing-selection-changed", onSelectionChanged as EventListener);

      await act(async () => {
        unmountRibbon?.();
        panelBodyRenderer.cleanup([]);
      });
      app.destroy();
      sheetRoot.remove();
    }
  });

  it("reorders ChartStore charts via bring forward / send backward when canvas charts are enabled", async () => {
    const url = new URL(window.location.href);
    url.searchParams.set("canvasCharts", "1");
    window.history.replaceState(null, "", url.toString());

    const sheetRoot = createRoot();
    const app = new SpreadsheetApp(sheetRoot, {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    });

    const { chart_id: chartA } = app.addChart({ chart_type: "bar", data_range: "A1:B2", title: "A" });
    const { chart_id: chartB } = app.addChart({ chart_type: "bar", data_range: "A1:B2", title: "B" });
    const drawingA = chartIdToDrawingId(chartA);
    const drawingB = chartIdToDrawingId(chartB);

    // Newer charts render above older charts.
    expect(app.listDrawingsForSheet().map((d) => d.id).slice(0, 2)).toEqual([drawingB, drawingA]);

    app.selectDrawingById(drawingA);
    expect(app.getSelectedChartId()).toBe(chartA);

    app.bringSelectedDrawingForward();
    expect(app.getSelectedChartId()).toBe(chartA);
    expect(app.listDrawingsForSheet().map((d) => d.id).slice(0, 2)).toEqual([drawingA, drawingB]);

    app.sendSelectedDrawingBackward();
    expect(app.getSelectedChartId()).toBe(chartA);
    expect(app.listDrawingsForSheet().map((d) => d.id).slice(0, 2)).toEqual([drawingB, drawingA]);

    app.destroy();
    sheetRoot.remove();
  });

  it("deletes selected ChartStore charts via Delete key when canvas charts are enabled", async () => {
    const url = new URL(window.location.href);
    url.searchParams.set("canvasCharts", "1");
    window.history.replaceState(null, "", url.toString());

    const sheetRoot = createRoot();
    const app = new SpreadsheetApp(sheetRoot, {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    });

    const { chart_id: chartId } = app.addChart({ chart_type: "bar", data_range: "A1:B2", title: "Delete Me" });
    const drawingId = chartIdToDrawingId(chartId);

    app.selectDrawingById(drawingId);
    expect(app.getSelectedChartId()).toBe(chartId);

    const evt = new KeyboardEvent("keydown", { key: "Delete", bubbles: true });
    sheetRoot.dispatchEvent(evt);

    expect(app.listCharts().some((c) => c.id === chartId)).toBe(false);
    expect(app.getSelectedChartId()).toBeNull();

    app.destroy();
    sheetRoot.remove();
  });

  it("disables arrange/delete action buttons in read-only mode", async () => {
    const drawings: DrawingObject[] = [
      {
        id: 1,
        kind: { type: "image", imageId: "img_1" },
        anchor: { type: "absolute", pos: { xEmu: 0, yEmu: 0 }, size: { cx: 10, cy: 10 } },
        zOrder: 0,
      },
    ];

    const deleteDrawingById = vi.fn();

    const app = {
      listDrawingsForSheet: () => drawings,
      subscribeDrawings: (listener: () => void) => {
        listener();
        return () => {};
      },
      getSelectedDrawingId: () => null,
      subscribeDrawingSelection: (listener: (id: number | null) => void) => {
        listener(null);
        return () => {};
      },
      selectDrawingById: vi.fn(),
      deleteDrawingById,
      bringSelectedDrawingForward: vi.fn(),
      sendSelectedDrawingBackward: vi.fn(),
      isReadOnly: () => true,
      isEditing: () => false,
    };

    const container = document.createElement("div");
    document.body.appendChild(container);
    const root = createReactRoot(container);

    await act(async () => {
      root.render(<SelectionPanePanel app={app as any} />);
    });

    const deleteBtn = container.querySelector<HTMLButtonElement>('[data-testid="selection-pane-delete-1"]');
    expect(deleteBtn).toBeInstanceOf(HTMLButtonElement);
    expect(deleteBtn!.disabled).toBe(true);

    await act(async () => {
      deleteBtn!.click();
    });
    expect(deleteDrawingById).toHaveBeenCalledTimes(0);

    act(() => root.unmount());
  });

  it("treats large-magnitude negative ids as workbook drawings (not canvas charts) for arrange gating", async () => {
    const hashedDrawingId = -0x200000000; // 2^33 (see parseDrawingObjectId in drawings/modelAdapters.ts)
    const drawings: DrawingObject[] = [
      {
        // ChartStore chart id namespace (negative 32-bit).
        id: -1,
        kind: { type: "chart", chartId: "chart_1" },
        anchor: { type: "absolute", pos: { xEmu: 0, yEmu: 0 }, size: { cx: 10, cy: 10 } },
        zOrder: 0,
      },
      {
        // Hashed drawing ids are also negative but should behave like workbook drawings.
        id: hashedDrawingId,
        kind: { type: "shape", label: "Hashed Drawing" },
        anchor: { type: "absolute", pos: { xEmu: 0, yEmu: 0 }, size: { cx: 10, cy: 10 } },
        zOrder: 0,
      },
      {
        id: 2,
        kind: { type: "shape", label: "Other Drawing" },
        anchor: { type: "absolute", pos: { xEmu: 0, yEmu: 0 }, size: { cx: 10, cy: 10 } },
        zOrder: 0,
      },
    ];

    const app = {
      listDrawingsForSheet: () => drawings,
      subscribeDrawings: (listener: () => void) => {
        listener();
        return () => {};
      },
      getSelectedDrawingId: () => null,
      subscribeDrawingSelection: (listener: (id: number | null) => void) => {
        listener(null);
        return () => {};
      },
      selectDrawingById: vi.fn(),
      deleteDrawingById: vi.fn(),
      bringSelectedDrawingForward: vi.fn(),
      sendSelectedDrawingBackward: vi.fn(),
      isReadOnly: () => false,
      isEditing: () => false,
    };

    const container = document.createElement("div");
    document.body.appendChild(container);
    const root = createReactRoot(container);

    await act(async () => {
      root.render(<SelectionPanePanel app={app as any} />);
    });

    const bringForward = container.querySelector<HTMLButtonElement>(`[data-testid="selection-pane-bring-forward-${hashedDrawingId}"]`);
    expect(bringForward).toBeInstanceOf(HTMLButtonElement);
    // Topmost drawing cannot be brought forward above the chart stack.
    expect(bringForward!.disabled).toBe(true);

    const sendBackward = container.querySelector<HTMLButtonElement>(`[data-testid="selection-pane-send-backward-${hashedDrawingId}"]`);
    expect(sendBackward).toBeInstanceOf(HTMLButtonElement);
    expect(sendBackward!.disabled).toBe(false);

    act(() => root.unmount());
  });
});
