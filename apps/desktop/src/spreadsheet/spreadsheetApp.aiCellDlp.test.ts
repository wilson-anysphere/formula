// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

vi.mock("../ai/audit/auditStore.js", async () => {
  const { MemoryAIAuditStore } = await import("../../../../packages/ai-audit/src/memory-store.js");
  return {
    getDesktopAIAuditStore: () => new MemoryAIAuditStore(),
  };
});

import { SpreadsheetApp } from "../app/spreadsheetApp";
import { evaluateFormula } from "./evaluateFormula.js";
import { AI_CELL_DLP_ERROR } from "./AiCellFunctionEngine.js";

import { createDefaultOrgPolicy } from "../../../../packages/security/dlp/src/policy.js";
import { DLP_ACTION } from "../../../../packages/security/dlp/src/actions.js";
import { LocalPolicyStore } from "../../../../packages/security/dlp/src/policyStore.js";
import { LocalClassificationStore } from "../../../../packages/security/dlp/src/classificationStore.js";
import { CLASSIFICATION_SCOPE } from "../../../../packages/security/dlp/src/selectors.js";
import { CLASSIFICATION_LEVEL } from "../../../../packages/security/dlp/src/classification.js";

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

describe("SpreadsheetApp AI cell functions (DLP wiring)", () => {
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

    Object.defineProperty(HTMLCanvasElement.prototype, "getContext", {
      configurable: true,
      value: () => createMockCanvasContext(),
    });

    (globalThis as any).ResizeObserver = class {
      observe() {}
      disconnect() {}
    };
  });

  it("constructs AiCellFunctionEngine with policy + classificationStore (no Public-default leakage)", () => {
    const workbookId = "dlp-test-workbook";
    const storage = window.localStorage as any;

    const policyStore = new LocalPolicyStore({ storage });
    const policy = createDefaultOrgPolicy();
    // Make AI cloud processing strict so restricted cell refs block immediately (instead of redacting).
    policy.rules[DLP_ACTION.AI_CLOUD_PROCESSING] = { ...policy.rules[DLP_ACTION.AI_CLOUD_PROCESSING], redactDisallowed: false };
    policyStore.setDocumentPolicy(workbookId, policy);

    const classificationStore = new LocalClassificationStore({ storage });
    classificationStore.upsert(
      workbookId,
      { scope: CLASSIFICATION_SCOPE.CELL, documentId: workbookId, sheetId: "Sheet1", row: 0, col: 0 },
      { level: CLASSIFICATION_LEVEL.RESTRICTED, labels: [] },
    );

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

    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status, { workbookId });
    const engine = (app as any).aiCellFunctions;

    const value = evaluateFormula('=AI("summarize", A1)', (ref) => (ref === "A1" ? "top secret" : null), {
      ai: engine,
      cellAddress: "Sheet1!B1",
    });
    expect(value).toBe(AI_CELL_DLP_ERROR);

    app.destroy();
    root.remove();
  });
});
