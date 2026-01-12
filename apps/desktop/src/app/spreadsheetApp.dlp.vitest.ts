// @vitest-environment jsdom

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

vi.mock("../ai/audit/auditStore.js", async () => {
  const { MemoryAIAuditStore } = await import("../../../../packages/ai-audit/src/memory-store.js");
  return {
    getDesktopAIAuditStore: () => new MemoryAIAuditStore(),
  };
});

import { SpreadsheetApp } from "./spreadsheetApp";

import { createDefaultOrgPolicy } from "../../../../packages/security/dlp/src/policy.js";
import { DLP_ACTION } from "../../../../packages/security/dlp/src/actions.js";
import { LocalPolicyStore } from "../../../../packages/security/dlp/src/policyStore.js";
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

describe("SpreadsheetApp clipboard DLP UX", () => {
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

  it("shows a toast when DLP blocks copy", async () => {
    const workbookId = "dlp-clipboard-toast-test";
    const storage = window.localStorage as any;

    const policyStore = new LocalPolicyStore({ storage });
    const policy = createDefaultOrgPolicy();
    // Make clipboard copy strict so the test is stable even if defaults change.
    policy.rules[DLP_ACTION.CLIPBOARD_COPY] = {
      ...policy.rules[DLP_ACTION.CLIPBOARD_COPY],
      maxAllowed: CLASSIFICATION_LEVEL.PUBLIC,
    };
    policyStore.setDocumentPolicy(workbookId, policy);

    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status, { workbookId });

    const write = vi.fn().mockResolvedValue(undefined);
    (app as any).clipboardProviderPromise = Promise.resolve({ write, read: vi.fn() });

    const dlp = (app as any).dlpContext;
    dlp.classificationStore.upsert(
      workbookId,
      { scope: CLASSIFICATION_SCOPE.DOCUMENT, documentId: workbookId },
      { level: CLASSIFICATION_LEVEL.RESTRICTED, labels: [] },
    );

    await (app as any).copySelectionToClipboard();

    const toast = document.querySelector('[data-testid="toast"]') as HTMLElement | null;
    expect(toast).not.toBeNull();
    expect(toast?.dataset.type).toBe("warning");
    expect(write).not.toHaveBeenCalled();

    app.destroy();
    root.remove();
  });
});
