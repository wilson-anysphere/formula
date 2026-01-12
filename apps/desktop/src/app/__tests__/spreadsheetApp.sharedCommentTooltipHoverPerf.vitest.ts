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
  root.getBoundingClientRect = vi.fn(
    () =>
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
      }) as any,
  );
  document.body.appendChild(root);
  return root;
}

async function flushMicrotasks(): Promise<void> {
  await Promise.resolve();
}

describe("SpreadsheetApp shared-grid comment tooltip hover perf", () => {
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

  it("avoids per-move layout reads and redundant tooltip updates when hovering within the same cell", async () => {
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

      (app as any).commentManager.addComment({
        cellRef: "A1",
        kind: "threaded",
        content: "Hello",
        author: (app as any).currentUser,
      });
      (app as any).reindexCommentCells();

      const selectionCanvas = (app as any).selectionCanvas as HTMLCanvasElement;
      const selectionRectSpy = vi.spyOn(selectionCanvas, "getBoundingClientRect");

      const rootRectSpy = root.getBoundingClientRect as unknown as ReturnType<typeof vi.fn>;
      rootRectSpy.mockClear();
      selectionRectSpy.mockClear();

      const tooltip = (app as any).commentTooltip as HTMLDivElement;
      const listSpy = vi.spyOn((app as any).commentManager, "listForCell");
      const pickSpy = vi.spyOn((app as any).sharedGrid.renderer, "pickCellAt");

      const mutations: MutationRecord[] = [];
      const observer = new MutationObserver((records) => mutations.push(...records));
      observer.observe(tooltip, { attributes: true, childList: true, characterData: true, subtree: true });

      const move = (clientX: number, clientY: number) => {
        // In production, pointermove events over the grid body target the selection canvas.
        // Include `target` + `offsetX/Y` so the handler exercises the fast path that avoids
        // root-relative coordinate math.
        (app as any).onSharedPointerMove({
          clientX,
          clientY,
          offsetX: clientX,
          offsetY: clientY,
          buttons: 0,
          pointerType: "mouse",
          target: selectionCanvas,
        } as any);
      };

      move(60, 30);
      await flushMicrotasks();

      expect(tooltip.classList.contains("comment-tooltip--visible")).toBe(true);
      expect(tooltip.textContent).toBe("Hello");
      expect(listSpy).toHaveBeenCalledTimes(0);
      expect(pickSpy).toHaveBeenCalledTimes(1);
      expect(mutations.length).toBeGreaterThan(0);
      const firstMutationCount = mutations.length;
      // The initial hover may require a one-time measurement to position the tooltip.
      // Clear the spies so we can assert that *subsequent* pointer moves within the same
      // cell do not trigger additional layout reads.
      rootRectSpy.mockClear();
      selectionRectSpy.mockClear();

      for (let i = 0; i < 100; i++) {
        move(60 + (i % 3), 30 + (i % 3));
      }
      await flushMicrotasks();

      expect(listSpy).toHaveBeenCalledTimes(0);
      expect(pickSpy).toHaveBeenCalledTimes(1);
      expect(mutations.length).toBe(firstMutationCount);
      // Hover should avoid per-move layout reads; tolerate at most a very small number of
      // rect refreshes (e.g. due to throttled root position tracking).
      expect(rootRectSpy.mock.calls.length).toBeLessThanOrEqual(2);
      expect(selectionRectSpy).toHaveBeenCalledTimes(0);

      observer.disconnect();
      app.destroy();
      root.remove();
    } finally {
      if (prior === undefined) delete process.env.DESKTOP_GRID_MODE;
      else process.env.DESKTOP_GRID_MODE = prior;
    }
  });

  it("skips cell picking entirely when the active sheet has no comments", async () => {
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

      const tooltip = (app as any).commentTooltip as HTMLDivElement;
      const pickSpy = vi.spyOn((app as any).sharedGrid.renderer, "pickCellAt");

      const move = (clientX: number, clientY: number) => {
        (app as any).onSharedPointerMove({ clientX, clientY, buttons: 0 } as any);
      };

      move(60, 30);
      await flushMicrotasks();

      expect(pickSpy).toHaveBeenCalledTimes(0);
      expect(tooltip.classList.contains("comment-tooltip--visible")).toBe(false);

      app.destroy();
      root.remove();
    } finally {
      if (prior === undefined) delete process.env.DESKTOP_GRID_MODE;
      else process.env.DESKTOP_GRID_MODE = prior;
    }
  });

  it("shows a tooltip even when the first comment thread content is an empty string", async () => {
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

      // Empty comment content should still mark the cell as having a comment and show a tooltip.
      (app as any).commentManager.addComment({
        cellRef: "A1",
        kind: "threaded",
        content: "",
        author: (app as any).currentUser,
      });
      (app as any).reindexCommentCells();

      const tooltip = (app as any).commentTooltip as HTMLDivElement;
      const selectionCanvas = (app as any).selectionCanvas as HTMLCanvasElement;

      (app as any).onSharedPointerMove({
        clientX: 60,
        clientY: 30,
        offsetX: 60,
        offsetY: 30,
        buttons: 0,
        pointerType: "mouse",
        target: selectionCanvas,
      } as any);
      await flushMicrotasks();

      expect(tooltip.classList.contains("comment-tooltip--visible")).toBe(true);
      expect(tooltip.textContent).toBe("");

      app.destroy();
      root.remove();
    } finally {
      if (prior === undefined) delete process.env.DESKTOP_GRID_MODE;
      else process.env.DESKTOP_GRID_MODE = prior;
    }
  });
});
