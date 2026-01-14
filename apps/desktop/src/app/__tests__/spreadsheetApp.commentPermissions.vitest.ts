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

describe("SpreadsheetApp comment permissions", () => {
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
  });

  it("disables the comments panel composer for viewers (canComment=false)", () => {
    const priorGridMode = process.env.DESKTOP_GRID_MODE;
    process.env.DESKTOP_GRID_MODE = "shared";
    try {
      const root = createRoot();
      const status = {
        activeCell: document.createElement("div"),
        selectionRange: document.createElement("div"),
        activeValue: document.createElement("div"),
      };

      const app = new SpreadsheetApp(root, status, { collabMode: true });
      // Simulate a collab session with viewer permissions.
      (app as any).collabSession = { canComment: () => false };

      app.toggleCommentsPanel();

      const panel = root.querySelector('[data-testid="comments-panel"]') as HTMLDivElement | null;
      if (!panel) throw new Error("Missing comments panel");

      const input = panel.querySelector('[data-testid="new-comment-input"]') as HTMLInputElement | null;
      if (!input) throw new Error("Missing new comment input");

      const submit = panel.querySelector('[data-testid="submit-comment"]') as HTMLButtonElement | null;
      if (!submit) throw new Error("Missing submit button");

      const hint = panel.querySelector('[data-testid="comments-readonly-hint"]') as HTMLDivElement | null;
      if (!hint) throw new Error("Missing read-only hint");

      expect(input.disabled).toBe(true);
      expect(submit.disabled).toBe(true);
      expect(hint.hidden).toBe(false);

      app.destroy();
      root.remove();
    } finally {
      if (priorGridMode === undefined) delete process.env.DESKTOP_GRID_MODE;
      else process.env.DESKTOP_GRID_MODE = priorGridMode;
    }
  });

  it("keeps the comments panel composer enabled for commenters even when the session is read-only", () => {
    const priorGridMode = process.env.DESKTOP_GRID_MODE;
    process.env.DESKTOP_GRID_MODE = "shared";
    try {
      const root = createRoot();
      const status = {
        activeCell: document.createElement("div"),
        selectionRange: document.createElement("div"),
        activeValue: document.createElement("div"),
      };

      const app = new SpreadsheetApp(root, status, { collabMode: true });

      // Simulate a "commenter" role: read-only for cell edits but allowed to comment.
      (app as any).collabSession = {
        canComment: () => true,
        isReadOnly: () => true,
        getPermissions: () => ({ role: "commenter" }),
      };
      // Ensure SpreadsheetApp's read-only state reflects the stubbed session.
      (app as any).syncReadOnlyState();

      app.toggleCommentsPanel();

      const panel = root.querySelector('[data-testid="comments-panel"]') as HTMLDivElement | null;
      if (!panel) throw new Error("Missing comments panel");

      const input = panel.querySelector('[data-testid="new-comment-input"]') as HTMLInputElement | null;
      if (!input) throw new Error("Missing new comment input");

      const submit = panel.querySelector('[data-testid="submit-comment"]') as HTMLButtonElement | null;
      if (!submit) throw new Error("Missing submit button");

      const hint = panel.querySelector('[data-testid="comments-readonly-hint"]') as HTMLDivElement | null;
      if (!hint) throw new Error("Missing read-only hint");

      expect(input.disabled).toBe(false);
      expect(submit.disabled).toBe(false);
      expect(hint.hidden).toBe(true);

      app.destroy();
      root.remove();
    } finally {
      if (priorGridMode === undefined) delete process.env.DESKTOP_GRID_MODE;
      else process.env.DESKTOP_GRID_MODE = priorGridMode;
    }
  });

  it("shows a read-only toast and opens the comments panel when viewers press Shift+F2", () => {
    const priorGridMode = process.env.DESKTOP_GRID_MODE;
    process.env.DESKTOP_GRID_MODE = "shared";
    try {
      const root = createRoot();
      const status = {
        activeCell: document.createElement("div"),
        selectionRange: document.createElement("div"),
        activeValue: document.createElement("div"),
      };

      const toastRoot = document.createElement("div");
      toastRoot.id = "toast-root";
      document.body.appendChild(toastRoot);

      const app = new SpreadsheetApp(root, status, { collabMode: true });
      // Simulate a collab session with viewer permissions.
      (app as any).collabSession = { canComment: () => false };

      const evt = new KeyboardEvent("keydown", { key: "F2", shiftKey: true, bubbles: true, cancelable: true });
      root.dispatchEvent(evt);

      expect(evt.defaultPrevented).toBe(true);

      const panel = root.querySelector('[data-testid="comments-panel"]') as HTMLDivElement | null;
      if (!panel) throw new Error("Missing comments panel");
      expect(panel.classList.contains("comments-panel--visible")).toBe(true);

      const input = panel.querySelector('[data-testid="new-comment-input"]') as HTMLInputElement | null;
      if (!input) throw new Error("Missing new comment input");
      expect(input.disabled).toBe(true);

      const toast = toastRoot.querySelector('[data-testid="toast"]') as HTMLDivElement | null;
      if (!toast) throw new Error("Missing toast");
      expect(toast.textContent).toContain("Read-only");

      app.destroy();
      root.remove();
    } finally {
      if (priorGridMode === undefined) delete process.env.DESKTOP_GRID_MODE;
      else process.env.DESKTOP_GRID_MODE = priorGridMode;
    }
  });
});
