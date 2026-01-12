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

describe("SpreadsheetApp comments sheet qualification", () => {
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

  it("keeps A1 comments separate across sheets in collab mode (sheetId!A1)", () => {
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
      expect(app.getGridMode()).toBe("shared");

      // Ensure a second sheet exists with a stable internal id that differs from the user-facing
      // display name (mirrors real workbooks where ids are stable, e.g. `sheet_<uuid>`).
      const doc = app.getDocument();
      doc.setCellValue("Sheet1", { row: 0, col: 0 }, "Seed1");
      doc.addSheet({ sheetId: "sheet_2", name: "My Sheet" });
      doc.setCellValue("sheet_2", { row: 0, col: 0 }, "Seed2");

      // Open the comments panel and add a comment on Sheet1!A1.
      app.activateCell({ sheetId: "Sheet1", row: 0, col: 0 });
      app.toggleCommentsPanel();

      const panel = root.querySelector('[data-testid="comments-panel"]') as HTMLDivElement | null;
      if (!panel) throw new Error("Missing comments panel");

      const cellLabel = panel.querySelector('[data-testid="comments-active-cell"]') as HTMLDivElement | null;
      if (!cellLabel) throw new Error("Missing active cell label");

      const input = panel.querySelector('[data-testid="new-comment-input"]') as HTMLInputElement | null;
      if (!input) throw new Error("Missing new comment input");

      const submit = panel.querySelector('[data-testid="submit-comment"]') as HTMLButtonElement | null;
      if (!submit) throw new Error("Missing submit button");

      input.value = "Comment on Sheet1";
      submit.click();

      const manager = (app as any).commentManager as { listForCell: (ref: string) => Array<{ content: string }> };
      expect(manager.listForCell("Sheet1!A1").map((c) => c.content)).toEqual(["Comment on Sheet1"]);
      expect(manager.listForCell("sheet_2!A1").map((c) => c.content)).toEqual([]);

      expect(cellLabel.textContent).toContain("Sheet1!A1");

      // Switch to sheet_2 - panel should update to 'My Sheet'!A1 (display name) and show no comments yet.
      app.activateSheet("sheet_2");
      expect(cellLabel.textContent).toContain("'My Sheet'!A1");
      expect(cellLabel.textContent).not.toContain("sheet_2!A1");
      expect(panel.textContent).not.toContain("Comment on Sheet1");
      expect(panel.querySelectorAll('[data-testid="comment-thread"]').length).toBe(0);

      // Add a comment on sheet_2!A1.
      input.value = "Comment on My Sheet";
      submit.click();

      expect(manager.listForCell("sheet_2!A1").map((c) => c.content)).toEqual(["Comment on My Sheet"]);
      expect(panel.textContent).toContain("Comment on My Sheet");
      expect(panel.textContent).not.toContain("Comment on Sheet1");

      // Switching back should restore the Sheet1 thread, without collisions.
      app.activateSheet("Sheet1");
      expect(cellLabel.textContent).toContain("Sheet1!A1");
      expect(panel.textContent).toContain("Comment on Sheet1");
      expect(panel.textContent).not.toContain("Comment on My Sheet");

      app.destroy();
      root.remove();
    } finally {
      if (priorGridMode === undefined) delete process.env.DESKTOP_GRID_MODE;
      else process.env.DESKTOP_GRID_MODE = priorGridMode;
    }
  });

  it("shows legacy unqualified A1 comments on the default sheet in collab mode", () => {
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
      const doc = app.getDocument();
      doc.setCellValue("Sheet2", { row: 0, col: 0 }, "Seed2");

      app.activateCell({ sheetId: "Sheet1", row: 0, col: 0 });
      app.toggleCommentsPanel();

      const panel = root.querySelector('[data-testid="comments-panel"]') as HTMLDivElement | null;
      if (!panel) throw new Error("Missing comments panel");

      const cellLabel = panel.querySelector('[data-testid="comments-active-cell"]') as HTMLDivElement | null;
      if (!cellLabel) throw new Error("Missing active cell label");

      // Inject a legacy comment stored under an unqualified A1 ref.
      const manager = (app as any).commentManager as {
        addComment: (input: { cellRef: string; kind: "threaded"; content: string; author: { id: string; name: string } }) => string;
        listForCell: (ref: string) => Array<{ content: string }>;
      };
      manager.addComment({ cellRef: "A1", kind: "threaded", content: "Legacy comment", author: { id: "u1", name: "User" } });

      // Underlying storage remains unqualified.
      expect(manager.listForCell("A1").map((c) => c.content)).toEqual(["Legacy comment"]);
      expect(manager.listForCell("Sheet1!A1").map((c) => c.content)).toEqual([]);

      // UI should surface it as Sheet1!A1 (default sheet) without showing it on other sheets.
      expect(cellLabel.textContent).toContain("Sheet1!A1");
      expect(panel.textContent).toContain("Legacy comment");

      app.activateSheet("Sheet2");
      expect(cellLabel.textContent).toContain("Sheet2!A1");
      expect(panel.textContent).not.toContain("Legacy comment");
      expect(panel.querySelectorAll('[data-testid="comment-thread"]').length).toBe(0);

      app.activateSheet("Sheet1");
      expect(cellLabel.textContent).toContain("Sheet1!A1");
      expect(panel.textContent).toContain("Legacy comment");

      app.destroy();
      root.remove();
    } finally {
      if (priorGridMode === undefined) delete process.env.DESKTOP_GRID_MODE;
      else process.env.DESKTOP_GRID_MODE = priorGridMode;
    }
  });

  it("reindexes per-sheet comment indicator caches when restoreDocumentState changes the active sheet (collab mode)", async () => {
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
      const doc = app.getDocument();
      doc.setCellValue("Sheet2", { row: 0, col: 0 }, "Seed2");

      app.activateSheet("Sheet2");

      (app as any).commentManager.addComment({
        cellRef: "Sheet2!A1",
        kind: "threaded",
        content: "Comment on Sheet2",
        author: (app as any).currentUser,
      });

      // In collab mode, coord-keyed indexes are scoped to the active sheet.
      expect((app as any).commentMetaByCoord.get(0)).toEqual({ resolved: false });

      // Restore a workbook snapshot that only contains Sheet1. This should force the app
      // to fall back to Sheet1 as the active sheet, and should also reindex comment caches
      // so Sheet2's A1 comment doesn't "bleed" into Sheet1.
      const snapshot = new TextEncoder().encode(
        JSON.stringify({
          schemaVersion: 1,
          sheetOrder: ["Sheet1"],
          sheets: [{ id: "Sheet1", frozenRows: 0, frozenCols: 0, cells: [] }],
        }),
      );

      await app.restoreDocumentState(snapshot);

      expect(app.getCurrentSheetId()).toBe("Sheet1");
      expect((app as any).commentMetaByCoord.size).toBe(0);
      expect((app as any).commentPreviewByCoord.size).toBe(0);

      app.destroy();
      root.remove();
    } finally {
      if (priorGridMode === undefined) delete process.env.DESKTOP_GRID_MODE;
      else process.env.DESKTOP_GRID_MODE = priorGridMode;
    }
  });
});
