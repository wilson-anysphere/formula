/**
 * @vitest-environment jsdom
 */

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { SpreadsheetApp } from "../spreadsheetApp";

type DrawnText = { text: string; font: string };

let priorGridMode: string | undefined;

function createMockCanvasContext(draws: DrawnText[]): CanvasRenderingContext2D {
  const noop = () => {};
  const gradient = { addColorStop: noop } as any;

  let currentFont = "";

  const target: any = {
    canvas: document.createElement("canvas"),
    measureText: (text: string) => ({ width: text.length * 8 }),
    createLinearGradient: () => gradient,
    createPattern: () => null,
    getImageData: () => ({ data: new Uint8ClampedArray(), width: 0, height: 0 }),
    putImageData: noop,
    fillText: (text: string) => {
      draws.push({ text: String(text), font: currentFont });
    }
  };

  Object.defineProperty(target, "font", {
    get: () => currentFont,
    set: (value: string) => {
      currentFont = String(value);
    }
  });

  return new Proxy(target, {
    get(obj, prop) {
      if (prop in obj) return (obj as any)[prop];
      return noop;
    },
    set(obj, prop, value) {
      (obj as any)[prop] = value;
      return true;
    }
  }) as any;
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
      toJSON: () => {}
    }) as any;
  document.body.appendChild(root);
  return root;
}

describe("SpreadsheetApp legacy grid font defaults", () => {
  beforeEach(() => {
    priorGridMode = process.env.DESKTOP_GRID_MODE;
    process.env.DESKTOP_GRID_MODE = "legacy";

    document.body.innerHTML = "";
    try {
      globalThis.localStorage?.clear?.();
    } catch {
      // ignore
    }

    // SpreadsheetApp schedules renders via requestAnimationFrame; ensure it exists in jsdom.
    Object.defineProperty(globalThis, "requestAnimationFrame", {
      configurable: true,
      value: (cb: FrameRequestCallback) => {
        cb(0);
        return 0;
      }
    });
    Object.defineProperty(globalThis, "cancelAnimationFrame", { configurable: true, value: () => {} });

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

  it("renders cell text using the monospace stack by default", () => {
    const draws: DrawnText[] = [];

    Object.defineProperty(HTMLCanvasElement.prototype, "getContext", {
      configurable: true,
      value: () => createMockCanvasContext(draws)
    });

    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div")
    };

    const app = new SpreadsheetApp(root, status);
    expect(app.getGridMode()).toBe("legacy");

    const doc = app.getDocument();
    doc.setCellValue("Sheet1", { row: 0, col: 0 }, "hello");

    // Force a render pass so the legacy renderer draws cell content.
    (app as any).renderGrid();

    const cellDraw = draws.find((d) => d.text === "hello");
    expect(cellDraw, "Expected the legacy renderer to draw a cell value").toBeTruthy();
    expect(cellDraw?.font).toContain("ui-monospace");
    expect(cellDraw?.font).toContain("monospace");
    expect(draws.some((d) => d.font.includes("system-ui"))).toBe(true);

    app.destroy();
    root.remove();
  });
});
