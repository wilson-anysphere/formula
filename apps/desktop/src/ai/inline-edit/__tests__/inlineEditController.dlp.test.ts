// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { MemoryAIAuditStore } from "@formula/ai-audit";

import { DocumentController } from "../../../document/documentController.js";
import { InlineEditController } from "../inlineEditController";

import { DLP_ACTION } from "../../../../../../packages/security/dlp/src/actions.js";
import { CLASSIFICATION_LEVEL } from "../../../../../../packages/security/dlp/src/classification.js";
import { LocalClassificationStore } from "../../../../../../packages/security/dlp/src/classificationStore.js";
import { LocalPolicyStore } from "../../../../../../packages/security/dlp/src/policyStore.js";

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

describe("InlineEditController (DLP blocked audit logging)", () => {
  beforeEach(() => {
    document.body.innerHTML = "";
    const storage = createInMemoryLocalStorage();
    Object.defineProperty(globalThis, "localStorage", { configurable: true, value: storage });
    Object.defineProperty(window, "localStorage", { configurable: true, value: storage });
    storage.clear();
  });

  afterEach(() => {
    vi.restoreAllMocks();
    vi.unstubAllGlobals();
  });

  it("logs a blocked audit entry (without cell values) when the selection is BLOCKed for AI_CLOUD_PROCESSING", async () => {
    const workbookId = "wb_inline_edit_blocked_audit";
    const restrictedValue = "TOP SECRET CELL VALUE";

    const policyStore = new LocalPolicyStore({ storage: window.localStorage as any });
    policyStore.setDocumentPolicy(workbookId, {
      version: 1,
      allowDocumentOverrides: true,
      rules: {
        [DLP_ACTION.AI_CLOUD_PROCESSING]: {
          maxAllowed: "Confidential",
          allowRestrictedContent: false,
          redactDisallowed: false
        }
      }
    });

    const classificationStore = new LocalClassificationStore({ storage: window.localStorage as any });
    classificationStore.upsert(
      workbookId,
      { scope: "cell", documentId: workbookId, sheetId: "Sheet1", row: 0, col: 0 },
      { level: CLASSIFICATION_LEVEL.RESTRICTED, labels: ["test"] }
    );

    const doc = new DocumentController();
    doc.setCellValue("Sheet1", "A1", restrictedValue);

    const container = document.createElement("div");
    document.body.appendChild(container);

    const llmClient = {
      chat: vi.fn(async () => {
        throw new Error("LLM should not be called when DLP blocks inline edit");
      })
    };
    const auditStore = new MemoryAIAuditStore();

    const controller = new InlineEditController({
      container,
      document: doc,
      workbookId,
      getSheetId: () => "Sheet1",
      getSelectionRange: () => ({ startRow: 0, endRow: 0, startCol: 0, endCol: 0 }),
      llmClient: llmClient as any,
      model: "unit-test-model",
      auditStore
    });

    controller.open();
    const overlay = await waitFor(() => container.querySelector<HTMLElement>('[data-testid="inline-edit-overlay"]'));

    const input = overlay.querySelector<HTMLInputElement>('[data-testid="inline-edit-prompt"]')!;
    input.value = "Fill something";

    overlay.querySelector<HTMLButtonElement>('[data-testid="inline-edit-run"]')!.click();

    const errorLabel = await waitFor(() => {
      const el = overlay.querySelector<HTMLElement>('[data-testid="inline-edit-error"]');
      return el && !el.hidden && el.textContent ? el : null;
    });

    expect(errorLabel.textContent).toMatch(/Sending data to cloud AI is restricted/i);
    expect(llmClient.chat).not.toHaveBeenCalled();

    const entries = await auditStore.listEntries({ workbook_id: workbookId, mode: "inline_edit" });
    expect(entries).toHaveLength(1);
    const entry = entries[0]!;
    expect((entry as any)?.input?.blocked).toBe(true);
    expect(JSON.stringify(entry.input)).not.toContain(restrictedValue);
  });
});

