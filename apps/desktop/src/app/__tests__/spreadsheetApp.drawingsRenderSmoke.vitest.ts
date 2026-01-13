/**
 * @vitest-environment jsdom
 */

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { SpreadsheetApp } from "../spreadsheetApp";
import type { DrawingObject } from "../../drawings/types";
import { pxToEmu } from "../../drawings/overlay";

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

function createMockCanvasContext(): CanvasRenderingContext2D & {
  strokeRect: ReturnType<typeof vi.fn>;
  fillText: ReturnType<typeof vi.fn>;
  drawImage: ReturnType<typeof vi.fn>;
} {
  const noop = () => {};
  const gradient = { addColorStop: noop } as any;

  const base: any = {
    canvas: document.createElement("canvas"),
    measureText: (text: string) => ({ width: text.length * 8 }),
    createLinearGradient: () => gradient,
    createPattern: () => null,
    getImageData: () => ({ data: new Uint8ClampedArray(), width: 0, height: 0 }),
    putImageData: noop,

    // Methods used by the drawings overlay smoke assertion.
    strokeRect: vi.fn(),
    fillText: vi.fn(),
    drawImage: vi.fn(),
  };

  return new Proxy(base, {
    get(target, prop) {
      if (prop in target) return (target as any)[prop];
      return noop;
    },
    set(target, prop, value) {
      (target as any)[prop] = value;
      return true;
    },
  }) as any;
}

describe("SpreadsheetApp drawings overlay render smoke", () => {
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

    // DrawingOverlay uses createImageBitmap for images; stub it even though this test uses a placeholder shape.
    Object.defineProperty(globalThis, "createImageBitmap", {
      configurable: true,
      value: vi.fn(async () => ({} as any)),
    });

    // Provide a distinct context per canvas so we can assert against the drawings layer specifically.
    const contexts = new WeakMap<HTMLCanvasElement, CanvasRenderingContext2D>();
    Object.defineProperty(HTMLCanvasElement.prototype, "getContext", {
      configurable: true,
      value: function getContext(type: string) {
        if (type !== "2d") return null;
        let ctx = contexts.get(this);
        if (!ctx) {
          ctx = createMockCanvasContext();
          contexts.set(this, ctx);
        }
        return ctx;
      },
    });

    (globalThis as any).ResizeObserver = class {
      observe() {}
      disconnect() {}
    };
  });

  it("renders the drawings canvas layer in shared mode without throwing", () => {
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

      const sheetId = app.getCurrentSheetId();
      const doc = app.getDocument() as any;
      const drawings: DrawingObject[] = [
        {
          id: 1,
          kind: { type: "shape", label: "placeholder" },
          anchor: {
            type: "oneCell",
            from: { cell: { row: 0, col: 0 }, offset: { xEmu: 0, yEmu: 0 } },
            size: { cx: pxToEmu(120), cy: pxToEmu(80) },
          },
          zOrder: 0,
        },
      ];
      doc.setSheetDrawings(sheetId, drawings);
      // Ensure the next render pass re-reads the document state.
      (app as any).drawingObjectsCache = null;
      (app as any).syncSheetDrawings?.();

      const drawingsCanvas = root.querySelector<HTMLCanvasElement>("canvas.grid-canvas--drawings");
      expect(drawingsCanvas).toBeTruthy();

      const ctx = drawingsCanvas!.getContext("2d") as any;
      ctx.strokeRect.mockClear();
      ctx.fillText.mockClear();
      ctx.drawImage.mockClear();

      app.refresh();

      expect(ctx.strokeRect).toHaveBeenCalled();
      expect(ctx.fillText).toHaveBeenCalled();

      app.destroy();
      root.remove();
    } finally {
      if (prior === undefined) delete process.env.DESKTOP_GRID_MODE;
      else process.env.DESKTOP_GRID_MODE = prior;
    }
  });
});
