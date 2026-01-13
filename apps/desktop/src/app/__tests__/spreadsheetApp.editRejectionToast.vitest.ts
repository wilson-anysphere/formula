/**
 * @vitest-environment jsdom
 */

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import * as Y from "yjs";

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

describe("SpreadsheetApp edit rejection toasts", () => {
  let priorGridMode: string | undefined;

  afterEach(() => {
    vi.unstubAllGlobals();
    vi.restoreAllMocks();

    if (priorGridMode === undefined) delete process.env.DESKTOP_GRID_MODE;
    else process.env.DESKTOP_GRID_MODE = priorGridMode;
  });

  beforeEach(() => {
    priorGridMode = process.env.DESKTOP_GRID_MODE;
    process.env.DESKTOP_GRID_MODE = "legacy";

    document.body.innerHTML = "";

    const toastRoot = document.createElement("div");
    toastRoot.id = "toast-root";
    document.body.appendChild(toastRoot);

    const storage = createInMemoryLocalStorage();
    Object.defineProperty(globalThis, "localStorage", { configurable: true, value: storage });
    Object.defineProperty(window, "localStorage", { configurable: true, value: storage });
    storage.clear();

    // SpreadsheetApp schedules paints via requestAnimationFrame.
    Object.defineProperty(globalThis, "requestAnimationFrame", {
      configurable: true,
      value: (cb: FrameRequestCallback) => {
        cb(0);
        return 0;
      },
    });
    Object.defineProperty(globalThis, "cancelAnimationFrame", { configurable: true, value: () => {} });

    // jsdom lacks a real canvas implementation; SpreadsheetApp expects a 2D context.
    Object.defineProperty(HTMLCanvasElement.prototype, "getContext", {
      configurable: true,
      value: () => createMockCanvasContext(),
    });

    // jsdom doesn't ship ResizeObserver by default.
    (globalThis as any).ResizeObserver = class {
      observe() {}
      disconnect() {}
    };
  });

  it("shows a read-only toast when canEditCell blocks an in-cell edit", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);

    // Simulate a permissions guard installed by collab mode.
    (app as any).document.canEditCell = () => false;

    (app as any).applyEdit("Sheet1", { row: 0, col: 0 }, "hello");

    expect(document.querySelector("#toast-root")?.textContent ?? "").toContain("Read-only");

    app.destroy();
    root.remove();
  });

  it("shows a missing encryption key toast when collab encryption blocks an edit", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);

    (app as any).document.canEditCell = () => false;

    const ydoc = new Y.Doc();
    const cells = ydoc.getMap("cells");
    (app as any).collabSession = {
      cells,
      getEncryptionConfig: () => ({
        keyForCell: () => null,
        shouldEncryptCell: () => true,
      }),
    };

    (app as any).applyEdit("Sheet1", { row: 0, col: 0 }, "hello");

    expect(document.querySelector("#toast-root")?.textContent ?? "").toContain("Missing encryption key");

    app.destroy();
    root.remove();
  });

  it("shows a read-only toast when the collab session role is viewer/commenter (isReadOnly)", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);

    // Simulate a read-only collab role; SpreadsheetApp should surface a toast when the user
    // attempts to start editing (rather than silently doing nothing).
    (app as any).collabSession = { isReadOnly: () => true };

    app.openCellEditorAtActiveCell();

    expect(document.querySelector("#toast-root")?.textContent ?? "").toContain("Read-only");

    app.destroy();
    root.remove();
  });

  it("shows a read-only toast when invoking AutoSum in read-only collab mode", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    (app as any).collabSession = { isReadOnly: () => true };

    app.autoSum();

    expect(document.querySelector("#toast-root")?.textContent ?? "").toContain("Read-only");

    app.destroy();
    root.remove();
  });
});
