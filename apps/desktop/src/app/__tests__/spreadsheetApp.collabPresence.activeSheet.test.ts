/**
 * @vitest-environment jsdom
 */

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { InMemoryAwarenessHub, PresenceManager } from "../../collab/presence/index.js";
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
        // Default all unknown properties to no-op functions so rendering code can execute.
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

describe("SpreadsheetApp collab presence", () => {
  afterEach(() => {
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
  });

  beforeEach(() => {
    document.body.innerHTML = "";

    // Node 22 ships an experimental `localStorage` global that errors unless configured via flags.
    // Provide a stable in-memory implementation for unit tests.
    const storage = createInMemoryLocalStorage();
    Object.defineProperty(globalThis, "localStorage", { configurable: true, value: storage });
    Object.defineProperty(window, "localStorage", { configurable: true, value: storage });
    storage.clear();

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

  it("updates session.presence.localPresence.activeSheet when switching sheets", () => {
    const hub = new InMemoryAwarenessHub();
    const awareness = hub.createAwareness(1);
    const presence = new PresenceManager(awareness, {
      user: { id: "u1", name: "User 1", color: "#ff0000" },
      activeSheet: "Sheet1",
      throttleMs: 0,
    });

    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    // Test shim: emulate collaboration mode by injecting a presence-enabled session.
    // SpreadsheetApp's sheet switching code should update `session.presence.activeSheet`.
    (app as any).collabSession = { presence };
    app.getDocument().setCellValue("Sheet2", { row: 0, col: 0 }, "X");

    app.activateSheet("Sheet2");

    expect(presence.localPresence.activeSheet).toBe("Sheet2");
    // Presence should also be refreshed for the new sheet.
    expect(presence.localPresence.cursor).toEqual({ row: 0, col: 0 });
    expect(presence.localPresence.selections).toEqual([{ startRow: 0, startCol: 0, endRow: 0, endCol: 0 }]);

    app.destroy();
    root.remove();
  });

  it("publishes cursor + selections when switching sheets via activateCell", () => {
    const hub = new InMemoryAwarenessHub();
    const awareness = hub.createAwareness(1);
    const presence = new PresenceManager(awareness, {
      user: { id: "u1", name: "User 1", color: "#ff0000" },
      activeSheet: "Sheet1",
      throttleMs: 0,
    });

    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    (app as any).collabSession = { presence };
    app.getDocument().setCellValue("Sheet2", { row: 0, col: 0 }, "X");

    app.activateCell({ sheetId: "Sheet2", row: 2, col: 3 }, { scrollIntoView: false, focus: false });

    expect(presence.localPresence.activeSheet).toBe("Sheet2");
    expect(presence.localPresence.cursor).toEqual({ row: 2, col: 3 });
    expect(presence.localPresence.selections).toEqual([{ startRow: 2, startCol: 3, endRow: 2, endCol: 3 }]);

    app.destroy();
    root.remove();
  });

  it("publishes cursor + selections when switching sheets via selectRange", () => {
    const hub = new InMemoryAwarenessHub();
    const awareness = hub.createAwareness(1);
    const presence = new PresenceManager(awareness, {
      user: { id: "u1", name: "User 1", color: "#ff0000" },
      activeSheet: "Sheet1",
      throttleMs: 0,
    });

    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    (app as any).collabSession = { presence };
    app.getDocument().setCellValue("Sheet2", { row: 0, col: 0 }, "X");

    app.selectRange(
      { sheetId: "Sheet2", range: { startRow: 1, startCol: 1, endRow: 2, endCol: 3 } },
      { scrollIntoView: false, focus: false },
    );

    expect(presence.localPresence.activeSheet).toBe("Sheet2");
    expect(presence.localPresence.cursor).toEqual({ row: 1, col: 1 });
    expect(presence.localPresence.selections).toEqual([{ startRow: 1, startCol: 1, endRow: 2, endCol: 3 }]);

    app.destroy();
    root.remove();
  });
});
