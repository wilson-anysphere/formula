/**
 * @vitest-environment jsdom
 */

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { SpreadsheetApp } from "../spreadsheetApp";

let priorGridMode: string | undefined;

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

async function flushArgumentPreview(): Promise<void> {
  const flushRender = async (): Promise<void> => {
    await new Promise<void>((resolve) => {
      if (typeof requestAnimationFrame === "function") {
        requestAnimationFrame(() => resolve());
      } else {
        setTimeout(resolve, 0);
      }
    });
  };

  // 1) Flush the render scheduled by input/selection changes.
  await flushRender();
  // 2) Allow the preview evaluation timer (setTimeout(..., 0)) to run.
  await new Promise<void>((resolve) => setTimeout(resolve, 0));
  // 3) Flush any promise microtasks from the preview provider + Promise.race.
  await Promise.resolve();
  await Promise.resolve();
  // 4) Flush the render scheduled after the preview resolves.
  await flushRender();
}

describe("SpreadsheetApp workbook file metadata + formula-bar argument preview", () => {
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
    if (priorGridMode === undefined) delete process.env.DESKTOP_GRID_MODE;
    else process.env.DESKTOP_GRID_MODE = priorGridMode;
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
  });

  it("refreshes argument previews when workbook file metadata changes (CELL(\"filename\"))", async () => {
    const root = createRoot();
    const formulaBar = document.createElement("div");
    document.body.appendChild(formulaBar);

    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status, { formulaBar });
    const doc = app.getDocument();
    const completionUpdateSpy = vi.fn();
    // Prevent the lazy tab-completion module from loading during this test, while still
    // asserting that metadata changes request a refresh.
    (app as any).formulaBarCompletion = { update: completionUpdateSpy, destroy: vi.fn() };

    // Force multi-sheet mode so computed values are evaluated in-process (no engine cache reliance).
    doc.addSheet({ sheetId: "Sheet2", name: "Sheet2" });

    await app.whenIdle();

    // Begin editing in the formula bar.
    app.activateCell({ row: 0, col: 0 }, { focus: false, scrollIntoView: false });
    const bar = (app as any).formulaBar as any;
    bar.focus({ cursor: "end" });

    // Cursor between the inner/outer ')' so the active argument expression is `CELL("filename")`.
    const formula = '=LEN(CELL("filename"))';
    bar.textarea.value = formula;
    const cursor = formula.length - 1;
    bar.textarea.setSelectionRange(cursor, cursor);
    bar.textarea.dispatchEvent(new Event("input"));

    await flushArgumentPreview();

    const preview1 = formulaBar.querySelector<HTMLElement>('[data-testid="formula-hint-arg-preview"]');
    expect(preview1?.textContent).toContain('↳ CELL("filename")');

    // Simulate Save As (directory + filename become known). The argument preview should refresh
    // without requiring cursor movement.
    await app.setWorkbookFileMetadata("/tmp/", "Book.xlsx");
    await flushArgumentPreview();

    const preview2 = formulaBar.querySelector<HTMLElement>('[data-testid="formula-hint-arg-preview"]');
    expect(preview2?.textContent).toContain("/tmp/[Book.xlsx]Sheet1");
    expect(completionUpdateSpy).toHaveBeenCalled();

    app.destroy();
    root.remove();
    formulaBar.remove();
  });
});
