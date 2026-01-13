// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { DocumentController } from "../../../document/documentController.js";
import { createSheetNameResolverFromIdToNameMap } from "../../../sheet/sheetNameResolver.js";
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

describe("InlineEditController sheetNameResolver regression", () => {
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

  it("resolves display sheet names via sheetNameResolver (prevents phantom sheets)", async () => {
    const container = document.createElement("div");
    document.body.appendChild(container);

    const doc = new DocumentController();
    doc.setCellValue("Sheet2", "A1", "x");
    // Clear history so assertions can reason about sheet creation deterministically.
    doc.applyState(doc.encodeState());

    const sheetIdToName = new Map<string, string>([["Sheet2", "Budget"]]);
    const sheetNameResolver = createSheetNameResolverFromIdToNameMap(sheetIdToName);

    let callCount = 0;
    const llmClient = {
      chat: vi.fn(async () => {
        callCount += 1;
        if (callCount === 1) {
          return {
            message: {
              role: "assistant",
              content: "",
              toolCalls: [
                { id: "call-1", name: "set_range", arguments: { range: "Budget!C1:C1", values: [[99]] } }
              ]
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
      getSheetId: () => "Sheet2",
      sheetNameResolver,
      getSelectionRange: () => ({ startRow: 0, endRow: 0, startCol: 2, endCol: 2 }), // C1:C1
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

    overlay.querySelector<HTMLInputElement>('[data-testid="inline-edit-prompt"]')!.value = "Set C1 to 99";
    overlay.querySelector<HTMLButtonElement>('[data-testid="inline-edit-run"]')!.click();

    await waitFor(
      () => {
        const el = overlay.querySelector<HTMLElement>('[data-testid="inline-edit-preview"]');
        return el && !el.hidden ? el : null;
      },
      5000
    );

    overlay.querySelector<HTMLButtonElement>('[data-testid="inline-edit-approve"]')!.click();

    await waitFor(() => (doc.getCell("Sheet2", "C1").value === 99 ? doc : null), 5000);
    await waitFor(() => (overlay.hidden ? overlay : null), 5000);

    expect(doc.getCell("Sheet2", "C1").value).toBe(99);
    expect(doc.getSheetIds()).toContain("Sheet2");
    expect(doc.getSheetIds()).not.toContain("Budget");

    expect((doc as any).batchDepth ?? 0).toBe(0);
    expect((doc as any).activeBatch ?? null).toBeNull();
    expect(onApplied).toHaveBeenCalledTimes(1);
    expect(llmClient.chat).toHaveBeenCalledTimes(2);
  });
});

