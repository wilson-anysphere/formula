// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { DocumentController } from "../../../document/documentController.js";
import { InlineEditController } from "../inlineEditController";

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
    }
  } as Storage;
}

async function waitFor<T>(fn: () => T | null | undefined, timeoutMs = 2000): Promise<T> {
  const started = Date.now();
  while (Date.now() - started < timeoutMs) {
    const value = fn();
    if (value) return value;
    await new Promise((resolve) => setTimeout(resolve, 0));
  }
  throw new Error("Timed out waiting for condition");
}

function createBaselineDocument(): DocumentController {
  const doc = new DocumentController();
  doc.setCellValue("Sheet1", "A1", "before");
  // Clear history so assertions can reason about "no changes applied" deterministically.
  doc.applyState(doc.encodeState());
  return doc;
}

describe("InlineEditController approval gating + batching", () => {
  beforeEach(() => {
    document.body.innerHTML = "";

    // Stabilize localStorage access (Node 22 + jsdom can be flaky otherwise).
    const storage = createInMemoryLocalStorage();
    Object.defineProperty(globalThis, "localStorage", { configurable: true, value: storage });
    Object.defineProperty(window, "localStorage", { configurable: true, value: storage });
    storage.clear();
  });

  afterEach(() => {
    vi.restoreAllMocks();
    vi.unstubAllGlobals();
  });

  it("deny path: shows preview, applies no changes, cancels any batch state, and closes the overlay", async () => {
    const container = document.createElement("div");
    document.body.appendChild(container);

    const doc = createBaselineDocument();

    let callCount = 0;
    const llmClient = {
      chat: vi.fn(async () => {
        callCount += 1;
        if (callCount === 1) {
          return {
            message: {
              role: "assistant",
              content: "",
              toolCalls: [{ id: "call-1", name: "set_range", arguments: { range: "A1:A1", values: [["after"]] } }]
            }
          };
        }
        return { message: { role: "assistant", content: "done" } };
      })
    };

    const onApplied = vi.fn();
    const controller = new InlineEditController({
      container,
      document: doc,
      workbookId: "test-workbook",
      getSheetId: () => "Sheet1",
      getSelectionRange: () => ({ startRow: 0, endRow: 0, startCol: 0, endCol: 0 }),
      llmClient,
      model: "unit-test-model",
      auditStore: {
        logEntry: vi.fn(async () => {}),
        listEntries: vi.fn(async () => [])
      } as any,
      onApplied
    });

    controller.open();

    const overlay = await waitFor(() => container.querySelector<HTMLElement>('[data-testid="inline-edit-overlay"]'));
    expect(overlay.hidden).toBe(false);

    overlay.querySelector<HTMLInputElement>('[data-testid="inline-edit-prompt"]')!.value = "Set A1 to after";
    overlay.querySelector<HTMLButtonElement>('[data-testid="inline-edit-run"]')!.click();

    const preview = await waitFor(
      () => {
        const el = overlay.querySelector<HTMLElement>('[data-testid="inline-edit-preview"]');
        return el && !el.hidden ? el : null;
      },
      5000
    );
    expect(preview.hidden).toBe(false);

    // Approval gating: sheet must be unchanged until user approves.
    expect(doc.getCell("Sheet1", "A1").value).toBe("before");

    overlay.querySelector<HTMLButtonElement>('[data-testid="inline-edit-preview-cancel"]')!.click();

    await waitFor(() => (overlay.hidden ? overlay : null), 5000);

    // Denial path must not apply changes.
    expect(doc.getCell("Sheet1", "A1").value).toBe("before");
    expect((doc as any).batchDepth ?? 0).toBe(0);
    expect((doc as any).activeBatch ?? null).toBeNull();
    expect(onApplied).not.toHaveBeenCalled();

    // No "second turn" after the tool was denied.
    expect(llmClient.chat).toHaveBeenCalledTimes(1);

    // Regression guard: later edits should not be swallowed into an abandoned batch.
    expect(doc.canUndo).toBe(false);
    doc.setCellValue("Sheet1", "A1", "user-edit");
    expect(doc.getCell("Sheet1", "A1").value).toBe("user-edit");
    expect(doc.canUndo).toBe(true);
    expect(doc.undo()).toBe(true);
    expect(doc.getCell("Sheet1", "A1").value).toBe("before");
  });

  it("approve path: shows preview, applies tool changes in a batch, ends batching, and fires onApplied", async () => {
    const container = document.createElement("div");
    document.body.appendChild(container);

    const doc = createBaselineDocument();

    let callCount = 0;
    const llmClient = {
      chat: vi.fn(async () => {
        callCount += 1;
        if (callCount === 1) {
          return {
            message: {
              role: "assistant",
              content: "",
              toolCalls: [{ id: "call-1", name: "set_range", arguments: { range: "A1:A1", values: [["after"]] } }]
            }
          };
        }
        return { message: { role: "assistant", content: "done" } };
      })
    };

    const onApplied = vi.fn();
    const controller = new InlineEditController({
      container,
      document: doc,
      workbookId: "test-workbook",
      getSheetId: () => "Sheet1",
      getSelectionRange: () => ({ startRow: 0, endRow: 0, startCol: 0, endCol: 0 }),
      llmClient,
      model: "unit-test-model",
      auditStore: {
        logEntry: vi.fn(async () => {}),
        listEntries: vi.fn(async () => [])
      } as any,
      onApplied
    });

    controller.open();

    const overlay = await waitFor(() => container.querySelector<HTMLElement>('[data-testid="inline-edit-overlay"]'));
    expect(overlay.hidden).toBe(false);

    overlay.querySelector<HTMLInputElement>('[data-testid="inline-edit-prompt"]')!.value = "Set A1 to after";
    overlay.querySelector<HTMLButtonElement>('[data-testid="inline-edit-run"]')!.click();

    await waitFor(
      () => {
        const el = overlay.querySelector<HTMLElement>('[data-testid="inline-edit-preview"]');
        return el && !el.hidden ? el : null;
      },
      5000
    );

    // Approval gating: sheet must be unchanged until user approves.
    expect(doc.getCell("Sheet1", "A1").value).toBe("before");

    overlay.querySelector<HTMLButtonElement>('[data-testid="inline-edit-approve"]')!.click();

    await waitFor(() => (doc.getCell("Sheet1", "A1").value === "after" ? doc : null), 5000);
    await waitFor(() => (overlay.hidden ? overlay : null), 5000);

    expect(doc.getCell("Sheet1", "A1").value).toBe("after");
    expect((doc as any).batchDepth ?? 0).toBe(0);
    expect((doc as any).activeBatch ?? null).toBeNull();
    expect(doc.undoLabel).toBe("AI Inline Edit");

    expect(onApplied).toHaveBeenCalledTimes(1);
    expect(llmClient.chat).toHaveBeenCalledTimes(2);
  });
});

