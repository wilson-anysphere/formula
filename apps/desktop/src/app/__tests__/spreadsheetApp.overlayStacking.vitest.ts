/**
 * @vitest-environment jsdom
 */

import fs from "node:fs";
import path from "node:path";

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { SpreadsheetApp } from "../spreadsheetApp";

let priorCanvasCharts: string | undefined;
let priorUseCanvasCharts: string | undefined;

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
  root.className = "grid-root";
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

function resolveDesktopStylePath(fileName: string): string {
  const fromRepoRoot = path.resolve(process.cwd(), "apps/desktop/src/styles", fileName);
  if (fs.existsSync(fromRepoRoot)) return fromRepoRoot;

  const fromDesktopRoot = path.resolve(process.cwd(), "src/styles", fileName);
  if (fs.existsSync(fromDesktopRoot)) return fromDesktopRoot;

  throw new Error(`Unable to resolve desktop style file "${fileName}" from cwd "${process.cwd()}"`);
}

function injectCss(filePath: string): void {
  const cssText = fs.readFileSync(filePath, "utf8");
  const style = document.createElement("style");
  style.textContent = cssText;
  document.head.appendChild(style);
}

function zIndexNumber(value: string): number {
  const trimmed = String(value ?? "").trim();
  const direct = Number.parseInt(trimmed, 10);
  if (Number.isFinite(direct)) return direct;

  // jsdom may not resolve `var(...)` indirections; fall back to parsing numeric
  // fallbacks (including nested `var(..., var(..., 3))` forms).
  const fallbackMatches = Array.from(trimmed.matchAll(/,\s*([0-9]+)\s*\)/g));
  if (fallbackMatches.length > 0) {
    const parsed = Number.parseInt(fallbackMatches.at(-1)?.[1] ?? "", 10);
    if (Number.isFinite(parsed)) return parsed;
  }

  throw new Error(`Unexpected z-index value: "${trimmed}"`);
}

function expectOverlayZOrder(root: HTMLElement): void {
  const drawingLayer = root.querySelector(".drawing-layer");
  const auditingLayer = root.querySelector(".grid-canvas--auditing");
  const selectionLayer = root.querySelector(".grid-canvas--selection");
  // Legacy-only chart canvas layers. In canvas-charts mode, ChartStore charts render via the
  // drawing overlay so these canvases are not mounted.
  const chartLayer = root.querySelector(".grid-canvas--chart");
  const chartSelectionLayer = root.querySelector(".chart-selection-canvas");
  const outlineLayer = root.querySelector(".outline-layer");
  const vScrollbarTrack = root.querySelector('[data-testid="scrollbar-track-y"]');
  const hScrollbarTrack = root.querySelector('[data-testid="scrollbar-track-x"]');
  const cellEditor = root.querySelector(".cell-editor");

  expect(drawingLayer).toBeTruthy();
  expect(auditingLayer).toBeTruthy();
  expect(selectionLayer).toBeTruthy();
  expect(outlineLayer).toBeTruthy();
  expect(vScrollbarTrack).toBeTruthy();
  expect(hScrollbarTrack).toBeTruthy();
  expect(cellEditor).toBeTruthy();

  const drawingZ = zIndexNumber(getComputedStyle(drawingLayer as Element).zIndex);
  const drawingStyle = getComputedStyle(drawingLayer as Element);
  const auditingStyle = getComputedStyle(auditingLayer as Element);
  const outlineStyle = getComputedStyle(outlineLayer as Element);
  const chartStyle = chartLayer ? getComputedStyle(chartLayer as Element) : null;
  const chartSelectionStyle = chartSelectionLayer ? getComputedStyle(chartSelectionLayer as Element) : null;
  const chartZ = chartStyle ? zIndexNumber(chartStyle.zIndex) : null;
  const auditingZ = zIndexNumber(getComputedStyle(auditingLayer as Element).zIndex);
  const selectionZ = zIndexNumber(getComputedStyle(selectionLayer as Element).zIndex);
  const chartSelectionZ = chartSelectionStyle ? zIndexNumber(chartSelectionStyle.zIndex) : null;
  const outlineZ = zIndexNumber(outlineStyle.zIndex);
  const vScrollbarZ = zIndexNumber(getComputedStyle(vScrollbarTrack as Element).zIndex);
  const hScrollbarZ = zIndexNumber(getComputedStyle(hScrollbarTrack as Element).zIndex);
  const editorZ = zIndexNumber(getComputedStyle(cellEditor as Element).zIndex);

  // Non-interactive overlay canvases should never intercept pointer events.
  if (chartStyle) expect(chartStyle.pointerEvents).toBe("none");
  expect(drawingStyle.pointerEvents).toBe("none");
  expect(auditingStyle.pointerEvents).toBe("none");
  if (chartSelectionStyle) expect(chartSelectionStyle.pointerEvents).toBe("none");
  expect(outlineStyle.pointerEvents).toBe("none");

  // Overlay stacking (low → high):
  //   - auditing highlights (z=1)
  //   - chart canvas (z=2; legacy-only)
  //   - drawings/images overlay (z=3)
  //   - selection + outline (z=4; + legacy chart selection handles)
  //   - scrollbars (z=5)
  //   - cell editor (z=10)
  expect(auditingZ).toBe(1);
  expect(drawingZ).toBe(3);
  expect(selectionZ).toBe(4);
  expect(outlineZ).toBe(4);
  expect(vScrollbarZ).toBe(5);
  expect(hScrollbarZ).toBe(5);
  expect(editorZ).toBe(10);

  if (chartZ != null) {
    expect(chartZ).toBe(2);
    expect(auditingZ).toBeLessThan(chartZ);
    expect(chartZ).toBeLessThan(drawingZ);
  }
  expect(drawingZ).toBeLessThan(selectionZ);
  if (chartSelectionZ != null) {
    expect(chartSelectionZ).toBe(4);
    expect(chartSelectionZ).toBeGreaterThanOrEqual(selectionZ);
  }
  expect(outlineZ).toBeGreaterThanOrEqual(selectionZ);
  expect(vScrollbarZ).toBeGreaterThan(selectionZ);
  expect(hScrollbarZ).toBeGreaterThan(selectionZ);
  expect(editorZ).toBeGreaterThan(vScrollbarZ);

  // Selection, chart selection handles (legacy), and outline overlays share the same z-index, so ensure
  // DOM insertion order preserves the visual stacking when z-index ties.
  const children = Array.from(root.children);
  if (chartSelectionLayer && chartSelectionZ === selectionZ) {
    expect(children.indexOf(chartSelectionLayer)).toBeGreaterThan(children.indexOf(selectionLayer!));
  }
  if (outlineZ === selectionZ) {
    // Outline must be above selection chrome.
    expect(children.indexOf(outlineLayer!)).toBeGreaterThan(children.indexOf(selectionLayer!));
  }
  if (chartSelectionLayer && chartSelectionZ != null && outlineZ === chartSelectionZ) {
    expect(children.indexOf(outlineLayer!)).toBeGreaterThan(children.indexOf(chartSelectionLayer));
  }
}

function expectPresenceOverlayZOrder(root: HTMLElement): void {
  const drawingLayer = root.querySelector(".drawing-layer");
  const auditingLayer = root.querySelector(".grid-canvas--auditing");
  const presenceLayer = root.querySelector(".grid-canvas--presence");
  const selectionLayer = root.querySelector(".grid-canvas--selection");
  const chartLayer = root.querySelector(".grid-canvas--chart");

  expect(drawingLayer).toBeTruthy();
  expect(auditingLayer).toBeTruthy();
  expect(presenceLayer).toBeTruthy();
  expect(selectionLayer).toBeTruthy();

  const drawingZ = zIndexNumber(getComputedStyle(drawingLayer as Element).zIndex);
  const chartZ = chartLayer ? zIndexNumber(getComputedStyle(chartLayer as Element).zIndex) : null;
  const auditingZ = zIndexNumber(getComputedStyle(auditingLayer as Element).zIndex);
  const presenceZ = zIndexNumber(getComputedStyle(presenceLayer as Element).zIndex);
  const selectionZ = zIndexNumber(getComputedStyle(selectionLayer as Element).zIndex);

  expect(auditingZ).toBe(1);
  expect(drawingZ).toBe(3);
  expect(presenceZ).toBe(4);
  expect(selectionZ).toBe(4);

  if (chartZ != null) {
    expect(chartZ).toBe(2);
    expect(auditingZ).toBeLessThan(chartZ);
    expect(chartZ).toBeLessThan(drawingZ);
  }
  expect(drawingZ).toBeLessThan(presenceZ);
  expect(presenceZ).toBeLessThanOrEqual(selectionZ);
  expect(getComputedStyle(presenceLayer as Element).pointerEvents).toBe("none");

  // Presence and selection share the same z-index so ensure DOM insertion order preserves
  // selection above presence highlights.
  if (presenceZ === selectionZ) {
    const children = Array.from(root.children);
    expect(children.indexOf(selectionLayer!)).toBeGreaterThan(children.indexOf(presenceLayer!));
  }
}

describe("SpreadsheetApp overlay stacking", () => {
  afterEach(() => {
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
    if (priorCanvasCharts === undefined) delete process.env.CANVAS_CHARTS;
    else process.env.CANVAS_CHARTS = priorCanvasCharts;
    if (priorUseCanvasCharts === undefined) delete process.env.USE_CANVAS_CHARTS;
    else process.env.USE_CANVAS_CHARTS = priorUseCanvasCharts;
  });

  beforeEach(() => {
    priorCanvasCharts = process.env.CANVAS_CHARTS;
    priorUseCanvasCharts = process.env.USE_CANVAS_CHARTS;
    // These tests assert overlay stacking for the default (legacy charts) behavior. Ensure the
    // feature flag env vars do not enable canvas charts (which would omit the legacy chart
    // canvases entirely).
    delete process.env.CANVAS_CHARTS;
    delete process.env.USE_CANVAS_CHARTS;

    document.head.innerHTML = "";
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

    // Load the global grid/overlay CSS so we can assert stacking.
    injectCss(resolveDesktopStylePath("charts-overlay.css"));
    injectCss(resolveDesktopStylePath("scrollbars.css"));
    injectCss(resolveDesktopStylePath("shell.css"));
  });

  it("applies deterministic drawing/chart/selection z-index ordering in legacy grid mode", () => {
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
      expect(app.getGridMode()).toBe("legacy");
      // Canvas charts are enabled by default, so legacy chart canvases should not be mounted.
      expect(root.querySelector(".grid-canvas--chart")).toBeNull();
      expect(root.querySelector(".chart-selection-canvas")).toBeNull();

      expectOverlayZOrder(root);

      app.destroy();
      root.remove();
    } finally {
      if (prior === undefined) delete process.env.DESKTOP_GRID_MODE;
      else process.env.DESKTOP_GRID_MODE = prior;
    }
  });

  it("keeps drawing/chart/selection z-index ordering consistent in shared grid mode", () => {
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

       // Shared mode still mounts the modifier classes; ensure they don't break stacking.
       const drawing = root.querySelector(".drawing-layer");
       const chart = root.querySelector(".grid-canvas--chart");
       const selection = root.querySelector(".grid-canvas--selection");
        expect(drawing?.classList.contains("drawing-layer--shared")).toBe(true);
        // Canvas charts are enabled by default, so legacy chart canvases should not be mounted.
        expect(chart).toBeNull();
        expect(selection?.classList.contains("grid-canvas--shared-selection")).toBe(true);

       expectOverlayZOrder(root);

      app.destroy();
      root.remove();
    } finally {
      if (prior === undefined) delete process.env.DESKTOP_GRID_MODE;
      else process.env.DESKTOP_GRID_MODE = prior;
    }
  });

  it("keeps collab presence overlay above charts/drawings but below selection via z-index ties", () => {
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
      expect(app.getGridMode()).toBe("legacy");

      // SpreadsheetApp only mounts presenceCanvas when collaboration is enabled. For stacking
      // coverage, simulate the collab canvas with the same class so CSS + DOM order assertions
      // can execute without a full collab session.
      const selectionLayer = root.querySelector(".grid-canvas--selection");
      expect(selectionLayer).toBeTruthy();
      const presenceCanvas = document.createElement("canvas");
      presenceCanvas.className = "grid-canvas grid-canvas--presence";
      presenceCanvas.setAttribute("aria-hidden", "true");
      root.insertBefore(presenceCanvas, selectionLayer!);

      expectOverlayZOrder(root);
      expectPresenceOverlayZOrder(root);

      app.destroy();
      root.remove();
    } finally {
      if (prior === undefined) delete process.env.DESKTOP_GRID_MODE;
      else process.env.DESKTOP_GRID_MODE = prior;
    }
  });
});
