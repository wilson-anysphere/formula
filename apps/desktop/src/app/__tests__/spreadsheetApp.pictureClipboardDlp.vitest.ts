// @vitest-environment jsdom

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

vi.mock("../../ai/audit/auditStore.js", async () => {
  const { MemoryAIAuditStore } = await import("../../../../../packages/ai-audit/src/memory-store.js");
  return {
    getDesktopAIAuditStore: () => new MemoryAIAuditStore(),
  };
});

import { SpreadsheetApp } from "../spreadsheetApp";

import { createDefaultOrgPolicy } from "../../../../../packages/security/dlp/src/policy.js";
import { DLP_ACTION } from "../../../../../packages/security/dlp/src/actions.js";
import { CLASSIFICATION_SCOPE } from "../../../../../packages/security/dlp/src/selectors.js";
import { CLASSIFICATION_LEVEL } from "../../../../../packages/security/dlp/src/classification.js";

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
    }
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

describe("SpreadsheetApp picture clipboard DLP", () => {
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

    const toastRoot = document.createElement("div");
    toastRoot.id = "toast-root";
    document.body.appendChild(toastRoot);
  });

  it("does not write image clipboard data when DLP blocks picture copy", async () => {
    const workbookId = "dlp-picture-copy-test";
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status, { workbookId });
    // Canvas charts are enabled by default, so any ChartStore charts appear in drawing-layer APIs
    // (`getDrawingObjects`, Selection Pane, etc). Remove any charts so this test can focus on the
    // image drawing inserted below.
    for (const chart of app.listCharts()) {
      (app as any).chartStore.deleteChart(chart.id);
    }

    const write = vi.fn().mockResolvedValue(undefined);
    (app as any).clipboardProviderPromise = Promise.resolve({ write, read: vi.fn() });

    // Ensure active cell is *not* the restricted cell so the test verifies the picture anchor is used.
    (app as any).selection = {
      type: "range",
      ranges: [{ startRow: 5, endRow: 5, startCol: 5, endCol: 5 }],
      active: { row: 5, col: 5 },
      anchor: { row: 5, col: 5 },
      activeRangeIndex: 0,
    };

    // Seed a picture anchored at A1 (row 0, col 0).
    const drawing = {
      id: 1,
      kind: { type: "image", imageId: "img1" },
      anchor: {
        type: "oneCell",
        from: { cell: { row: 0, col: 0 }, offset: { xEmu: 0, yEmu: 0 } },
        size: { cx: 0, cy: 0 },
      },
      zOrder: 0,
    };
    (app as any).setDrawingObjectsForSheet([drawing]);
    app.getDrawingImages().set({ id: "img1", bytes: new Uint8Array([1, 2, 3]), mimeType: "image/png" });
    app.selectDrawing(1);

    // Configure DLP to block clipboard copy when the anchor cell is Restricted.
    const dlp = (app as any).dlpContext;
    const policy = createDefaultOrgPolicy();
    policy.rules[DLP_ACTION.CLIPBOARD_COPY] = {
      ...policy.rules[DLP_ACTION.CLIPBOARD_COPY],
      maxAllowed: CLASSIFICATION_LEVEL.PUBLIC,
    };
    dlp.policy = policy;
    dlp.classificationStore.upsert(
      workbookId,
      { scope: CLASSIFICATION_SCOPE.CELL, documentId: workbookId, sheetId: "Sheet1", row: 0, col: 0 },
      { level: CLASSIFICATION_LEVEL.RESTRICTED, labels: [] }
    );

    app.copy();
    await app.whenIdle();

    expect(write).not.toHaveBeenCalled();

    app.destroy();
    root.remove();
  });

  it("does not delete the picture when DLP blocks picture cut", async () => {
    const workbookId = "dlp-picture-cut-test";
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status, { workbookId });
    for (const chart of app.listCharts()) {
      (app as any).chartStore.deleteChart(chart.id);
    }

    const write = vi.fn().mockResolvedValue(undefined);
    (app as any).clipboardProviderPromise = Promise.resolve({ write, read: vi.fn() });

    (app as any).selection = {
      type: "range",
      ranges: [{ startRow: 5, endRow: 5, startCol: 5, endCol: 5 }],
      active: { row: 5, col: 5 },
      anchor: { row: 5, col: 5 },
      activeRangeIndex: 0,
    };

    const drawing = {
      id: 1,
      kind: { type: "image", imageId: "img1" },
      anchor: {
        type: "oneCell",
        from: { cell: { row: 0, col: 0 }, offset: { xEmu: 0, yEmu: 0 } },
        size: { cx: 0, cy: 0 },
      },
      zOrder: 0,
    };
    (app as any).setDrawingObjectsForSheet([drawing]);
    app.getDrawingImages().set({ id: "img1", bytes: new Uint8Array([1, 2, 3]), mimeType: "image/png" });
    app.selectDrawing(1);

    const dlp = (app as any).dlpContext;
    const policy = createDefaultOrgPolicy();
    policy.rules[DLP_ACTION.CLIPBOARD_COPY] = {
      ...policy.rules[DLP_ACTION.CLIPBOARD_COPY],
      maxAllowed: CLASSIFICATION_LEVEL.PUBLIC,
    };
    dlp.policy = policy;
    dlp.classificationStore.upsert(
      workbookId,
      { scope: CLASSIFICATION_SCOPE.CELL, documentId: workbookId, sheetId: "Sheet1", row: 0, col: 0 },
      { level: CLASSIFICATION_LEVEL.RESTRICTED, labels: [] }
    );

    app.cut();
    await app.whenIdle();

    expect(write).not.toHaveBeenCalled();
    expect(app.getDrawingObjects().filter((obj) => obj.kind.type === "image")).toHaveLength(1);

    app.destroy();
    root.remove();
  });
});
