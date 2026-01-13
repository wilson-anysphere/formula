/**
 * @vitest-environment jsdom
 */

import fs from "node:fs";
import path from "node:path";

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

  // jsdom may not resolve `var(...)` indirections; fall back to parsing the default.
  const varMatch = /var\(\s*--[^,\s)]+\s*,\s*([0-9]+)\s*\)/.exec(trimmed);
  if (varMatch) {
    const parsed = Number.parseInt(varMatch[1] ?? "", 10);
    if (Number.isFinite(parsed)) return parsed;
  }

  throw new Error(`Unexpected z-index value: "${trimmed}"`);
}

function expectOverlayZOrder(root: HTMLElement): void {
  const drawingLayer = root.querySelector(".drawing-layer");
  const chartLayer = root.querySelector(".grid-canvas--chart");
  const selectionLayer = root.querySelector(".grid-canvas--selection");
  const outlineLayer = root.querySelector(".outline-layer");

  expect(drawingLayer).toBeTruthy();
  expect(chartLayer).toBeTruthy();
  expect(selectionLayer).toBeTruthy();
  expect(outlineLayer).toBeTruthy();

  const drawingZ = zIndexNumber(getComputedStyle(drawingLayer as Element).zIndex);
  const chartZ = zIndexNumber(getComputedStyle(chartLayer as Element).zIndex);
  const selectionZ = zIndexNumber(getComputedStyle(selectionLayer as Element).zIndex);
  const outlineZ = zIndexNumber(getComputedStyle(outlineLayer as Element).zIndex);

  expect(drawingZ).toBe(1);
  expect(chartZ).toBe(2);
  expect(selectionZ).toBe(3);
  expect(outlineZ).toBe(4);

  expect(drawingZ).toBeLessThan(chartZ);
  expect(chartZ).toBeLessThan(selectionZ);
  expect(selectionZ).toBeLessThan(outlineZ);
}

describe("SpreadsheetApp overlay stacking", () => {
  afterEach(() => {
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
  });

  beforeEach(() => {
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
      expect(chart?.classList.contains("grid-canvas--shared-chart")).toBe(true);
      expect(selection?.classList.contains("grid-canvas--shared-selection")).toBe(true);

      expectOverlayZOrder(root);

      app.destroy();
      root.remove();
    } finally {
      if (prior === undefined) delete process.env.DESKTOP_GRID_MODE;
      else process.env.DESKTOP_GRID_MODE = prior;
    }
  });
});
