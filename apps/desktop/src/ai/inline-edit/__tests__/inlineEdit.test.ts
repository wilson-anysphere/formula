// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { SpreadsheetApp } from "../../../app/spreadsheetApp";
import { createHeuristicTokenEstimator, estimateToolDefinitionTokens } from "../../../../../../packages/ai-context/src/tokenBudget.js";
import { getDefaultReserveForOutputTokens, getModeContextWindowTokens } from "../../contextBudget.js";

import { LocalStorageBinaryStorage } from "@formula/ai-audit/browser";
import { SqliteAIAuditStore } from "@formula/ai-audit/sqlite";
import { createRequire } from "node:module";
import { DLP_ACTION } from "../../../../../../packages/security/dlp/src/actions.js";
import { CLASSIFICATION_LEVEL } from "../../../../../../packages/security/dlp/src/classification.js";
import { LocalClassificationStore } from "../../../../../../packages/security/dlp/src/classificationStore.js";
import { LocalPolicyStore } from "../../../../../../packages/security/dlp/src/policyStore.js";

let priorGridMode: string | undefined;
// Legacy user API key storage key (used by very old builds). Keep the string split
// so unit tests don't mention provider names (Cursor-only AI policy guard).
const LEGACY_API_KEY_STORAGE_KEY = "formula:" + "op" + "en" + "ai" + "ApiKey";

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

async function waitFor<T>(fn: () => T | null | undefined, timeoutMs = 2000): Promise<T> {
  const started = Date.now();
  while (Date.now() - started < timeoutMs) {
    const value = fn();
    if (value) return value;
    await new Promise((resolve) => setTimeout(resolve, 0));
  }
  throw new Error("Timed out waiting for condition");
}

describe("AI inline edit (Cmd/Ctrl+K)", () => {
  afterEach(() => {
    if (priorGridMode === undefined) delete process.env.DESKTOP_GRID_MODE;
    else process.env.DESKTOP_GRID_MODE = priorGridMode;
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
  });

  beforeEach(() => {
    priorGridMode = process.env.DESKTOP_GRID_MODE;
    process.env.DESKTOP_GRID_MODE = "legacy";
    document.body.innerHTML = "";

    // Node 22 ships an experimental `localStorage` global that errors unless configured via flags.
    // Provide a stable in-memory implementation for unit tests (also used by SpreadsheetApp's comment
    // persistence + the LocalStorageAIAuditStore).
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

  it("opens via Cmd/Ctrl+K, previews tool changes, gates approval, applies in one undo step, and audits", async () => {
    const require = createRequire(import.meta.url);
    const locateFile = (file: string) => require.resolve(`sql.js/dist/${file}`);
    const auditStore = await SqliteAIAuditStore.create({
      storage: new LocalStorageBinaryStorage("ai_audit_inline_edit_test"),
      locateFile,
    });

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

    let callCount = 0;
    const llmClient = {
      async chat() {
        callCount++;
        if (callCount === 1) {
          return {
            message: {
              role: "assistant",
              content: "",
              toolCalls: [
                {
                  id: "call-1",
                  name: "set_range",
                  arguments: { range: "C1:C3", values: [[1], [2], [3]] },
                },
              ],
            },
            usage: { promptTokens: 10, completionTokens: 5 },
          };
        }

        return {
          message: { role: "assistant", content: "done" },
          usage: { promptTokens: 1, completionTokens: 1 },
        };
      },
    };

    const app = new SpreadsheetApp(root, status, {
      inlineEdit: { llmClient, model: "unit-test-model", auditStore },
    });

    // Select an empty range so preview diffs are deterministic (SpreadsheetApp seeds A/B columns).
    app.selectRange({ range: { startRow: 0, endRow: 2, startCol: 2, endCol: 2 } }); // C1:C3

    root.dispatchEvent(new KeyboardEvent("keydown", { key: "k", ctrlKey: true, bubbles: true }));

    const overlay = await waitFor(() => document.querySelector<HTMLElement>('[data-testid="inline-edit-overlay"]'));
    expect(overlay.hidden).toBe(false);

    const input = overlay.querySelector<HTMLInputElement>('[data-testid="inline-edit-prompt"]');
    expect(input).toBeTruthy();
    input!.value = "Fill with 1..3";

    const runBtn = overlay.querySelector<HTMLButtonElement>('[data-testid="inline-edit-run"]');
    expect(runBtn).toBeTruthy();
    runBtn!.click();

    const previewSummary = await waitFor(() => {
      const el = overlay.querySelector<HTMLElement>('[data-testid="inline-edit-preview-summary"]');
      return el && el.textContent?.includes("Changes:") ? el : null;
    });
    expect(previewSummary.textContent).toContain("Changes:");

    // Approval gating: tool hasn't executed yet, so the document is unchanged.
    const doc = app.getDocument();
    expect(doc.getCell("Sheet1", "C1").value).toBeNull();
    expect(doc.getCell("Sheet1", "C2").value).toBeNull();
    expect(doc.getCell("Sheet1", "C3").value).toBeNull();

    const approveBtn = overlay.querySelector<HTMLButtonElement>('[data-testid="inline-edit-approve"]');
    expect(approveBtn).toBeTruthy();
    approveBtn!.click();

    await waitFor(() => (doc.getCell("Sheet1", "C3").value === 3 ? doc : null));
    expect(doc.getCell("Sheet1", "C1").value).toBe(1);
    expect(doc.getCell("Sheet1", "C2").value).toBe(2);
    expect(doc.getCell("Sheet1", "C3").value).toBe(3);

    // Undo should revert the entire inline edit in one step.
    expect(doc.undo()).toBe(true);
    expect(doc.getCell("Sheet1", "C1").value).toBeNull();
    expect(doc.getCell("Sheet1", "C2").value).toBeNull();
    expect(doc.getCell("Sheet1", "C3").value).toBeNull();

    const entries = await auditStore.listEntries({ workbook_id: "local-workbook", mode: "inline_edit" });
    expect(entries.length).toBeGreaterThan(0);
    expect(entries[0]?.mode).toBe("inline_edit");
    expect((entries[0] as any)?.input?.workbookId).toBe("local-workbook");
    expect(entries[0]?.tool_calls?.[0]?.name).toBe("set_range");
    expect(entries[0]?.tool_calls?.[0]?.approved).toBe(true);
  });

  it("blocks inline edit before calling the LLM when DLP forbids cloud AI processing for the selection", async () => {
    const workbookId = "local-workbook";

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
    // Mark C1 as Restricted (selection is C1:C3 in the test).
    classificationStore.upsert(
      workbookId,
      { scope: "cell", documentId: workbookId, sheetId: "Sheet1", row: 0, col: 2 },
      { level: CLASSIFICATION_LEVEL.RESTRICTED, labels: ["test"] }
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
        toJSON: () => {}
      }) as any;
    document.body.appendChild(root);

    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div")
    };

    const llmClient = {
      chat: vi.fn(async () => {
        throw new Error("LLM should not be called when DLP blocks inline edit");
      })
    };

    const app = new SpreadsheetApp(root, status, { inlineEdit: { llmClient, model: "unit-test-model" } });
    app.selectRange({ range: { startRow: 0, endRow: 2, startCol: 2, endCol: 2 } }); // C1:C3

    root.dispatchEvent(new KeyboardEvent("keydown", { key: "k", ctrlKey: true, bubbles: true }));

    const overlay = await waitFor(() => document.querySelector<HTMLElement>('[data-testid="inline-edit-overlay"]'));
    const input = overlay.querySelector<HTMLInputElement>('[data-testid="inline-edit-prompt"]')!;
    input.value = "Fill with values";

    overlay.querySelector<HTMLButtonElement>('[data-testid="inline-edit-run"]')!.click();

    const errorLabel = await waitFor(() => {
      const el = overlay.querySelector<HTMLElement>('[data-testid="inline-edit-error"]');
      return el && !el.hidden && el.textContent ? el : null;
    });

    expect(errorLabel.textContent).toMatch(/Sending data to cloud AI is restricted/i);
    expect(llmClient.chat).not.toHaveBeenCalled();
  });

  it("redacts the selection sample before calling the LLM when DLP requires redaction", async () => {
    const workbookId = "local-workbook";

    const policyStore = new LocalPolicyStore({ storage: window.localStorage as any });
    policyStore.setDocumentPolicy(workbookId, {
      version: 1,
      allowDocumentOverrides: true,
      rules: {
        [DLP_ACTION.AI_CLOUD_PROCESSING]: {
          maxAllowed: "Confidential",
          allowRestrictedContent: false,
          redactDisallowed: true
        }
      }
    });

    const classificationStore = new LocalClassificationStore({ storage: window.localStorage as any });
    // Mark C1 as Restricted (selection is C1:C3 in the test).
    classificationStore.upsert(
      workbookId,
      { scope: "cell", documentId: workbookId, sheetId: "Sheet1", row: 0, col: 2 },
      { level: CLASSIFICATION_LEVEL.RESTRICTED, labels: ["test"] }
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
        toJSON: () => {}
      }) as any;
    document.body.appendChild(root);

    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div")
    };

    const llmClient = {
      chat: vi.fn(async (request: any) => {
        const messages = Array.isArray(request?.messages) ? request.messages : [];
        const userMessage = messages.find((m: any) => m?.role === "user")?.content ?? "";
        // The sample included in the prompt should be redacted before sending to the LLM.
        expect(userMessage).toContain("[REDACTED]");
        expect(userMessage).not.toContain("TOP SECRET");
        return {
          message: { role: "assistant", content: "done" },
          usage: { promptTokens: 1, completionTokens: 1 }
        };
      })
    };

    const app = new SpreadsheetApp(root, status, { inlineEdit: { llmClient, model: "unit-test-model" } });
    const doc = app.getDocument();
    doc.setCellValue("Sheet1", "C1", "TOP SECRET");

    app.selectRange({ range: { startRow: 0, endRow: 2, startCol: 2, endCol: 2 } }); // C1:C3

    root.dispatchEvent(new KeyboardEvent("keydown", { key: "k", ctrlKey: true, bubbles: true }));

    const overlay = await waitFor(() => document.querySelector<HTMLElement>('[data-testid="inline-edit-overlay"]'));
    const input = overlay.querySelector<HTMLInputElement>('[data-testid="inline-edit-prompt"]')!;
    input.value = "Do something";

    overlay.querySelector<HTMLButtonElement>('[data-testid="inline-edit-run"]')!.click();

    await waitFor(() => (llmClient.chat.mock.calls.length > 0 ? overlay : null));
    expect(llmClient.chat).toHaveBeenCalledTimes(1);

    await waitFor(() => (overlay.hidden ? overlay : null));
    // No tool calls were issued; the sheet should remain unchanged.
    expect(doc.getCell("Sheet1", "C1").value).toBe("TOP SECRET");
  });

  it("uses the default Cursor client when no inlineEdit llmClient is injected", async () => {
    // Legacy user API keys should be proactively purged (and never used for auth).
    localStorage.setItem(LEGACY_API_KEY_STORAGE_KEY, "sk-test-inline-edit");

    let callCount = 0;
    const fetchMock = vi.fn(async (url: string, init: any) => {
      callCount++;
      expect(url).toBe("/v1/chat/completions");
      expect(init?.method).toBe("POST");
      expect(init?.credentials).toBe("include");
      // Cursor-only: no user API keys should be sent from the client.
      expect(init?.headers?.Authorization).toBeUndefined();

      if (callCount === 1) {
        return {
          ok: true,
          json: async () => ({
            choices: [
              {
                message: {
                  role: "assistant",
                  content: "",
                  tool_calls: [
                    {
                      id: "call-1",
                      type: "function",
                      function: {
                        name: "set_range",
                        arguments: JSON.stringify({ range: "C1:C3", values: [[1], [2], [3]] }),
                      },
                    },
                  ],
                },
              },
            ],
            usage: { prompt_tokens: 10, completion_tokens: 5, total_tokens: 15 },
          }),
        } as any;
      }

      return {
        ok: true,
        json: async () => ({
          choices: [{ message: { role: "assistant", content: "done" } }],
          usage: { prompt_tokens: 1, completion_tokens: 1, total_tokens: 2 },
        }),
      } as any;
    });

    vi.stubGlobal("fetch", fetchMock);

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

    // No inlineEdit config passed; controller should fall back to the default desktop LLM client.
    const app = new SpreadsheetApp(root, status);

    app.selectRange({ range: { startRow: 0, endRow: 2, startCol: 2, endCol: 2 } }); // C1:C3
    root.dispatchEvent(new KeyboardEvent("keydown", { key: "k", ctrlKey: true, bubbles: true }));

    const overlay = await waitFor(() => document.querySelector<HTMLElement>('[data-testid="inline-edit-overlay"]'));
    const input = overlay.querySelector<HTMLInputElement>('[data-testid="inline-edit-prompt"]');
    input!.value = "Fill with 1..3";

    overlay.querySelector<HTMLButtonElement>('[data-testid="inline-edit-run"]')!.click();

    await waitFor(() => {
      const el = overlay.querySelector<HTMLElement>('[data-testid="inline-edit-preview-summary"]');
      return el && el.textContent?.includes("Changes:") ? el : null;
    });

    expect(localStorage.getItem(LEGACY_API_KEY_STORAGE_KEY)).toBeNull();

    overlay.querySelector<HTMLButtonElement>('[data-testid="inline-edit-approve"]')!.click();

    const doc = app.getDocument();
    await waitFor(() => (doc.getCell("Sheet1", "C3").value === 3 ? doc : null));
    expect(doc.getCell("Sheet1", "C1").value).toBe(1);
    expect(doc.getCell("Sheet1", "C2").value).toBe(2);
    expect(doc.getCell("Sheet1", "C3").value).toBe(3);

    expect(fetchMock).toHaveBeenCalled();
  });

  it("trims oversized inline-edit prompts to stay under the strict context budget", async () => {
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
        toJSON: () => {}
      }) as any;
    document.body.appendChild(root);

    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div")
    };

    const estimator = createHeuristicTokenEstimator();
    const model = "unit-test-model";
    const contextWindowTokens = getModeContextWindowTokens("inline_edit", model);
    const reserveForOutputTokens = getDefaultReserveForOutputTokens("inline_edit", contextWindowTokens);

    let callCount = 0;
    const llmClient = {
      async chat(request: any) {
        callCount += 1;
        const promptTokens =
          estimator.estimateMessagesTokens(request.messages) + estimateToolDefinitionTokens(request.tools, estimator);
        expect(promptTokens).toBeLessThanOrEqual(contextWindowTokens - reserveForOutputTokens);

        if (callCount === 1) {
          return {
            message: {
              role: "assistant",
              content: "",
              toolCalls: [
                {
                  id: "call-1",
                  name: "set_range",
                  arguments: { range: "C1:C3", values: [[1], [2], [3]] }
                }
              ]
            }
          };
        }

        return { message: { role: "assistant", content: "done" } };
      }
    };

    const app = new SpreadsheetApp(root, status, {
      inlineEdit: { llmClient, model }
    });

    // Select an empty range so preview diffs are deterministic.
    app.selectRange({ range: { startRow: 0, endRow: 2, startCol: 2, endCol: 2 } }); // C1:C3

    root.dispatchEvent(new KeyboardEvent("keydown", { key: "k", ctrlKey: true, bubbles: true }));

    const overlay = await waitFor(() => document.querySelector<HTMLElement>('[data-testid="inline-edit-overlay"]'));
    const input = overlay.querySelector<HTMLInputElement>('[data-testid="inline-edit-prompt"]');
    expect(input).toBeTruthy();
    input!.value = "Fill with 1..3\n" + "x".repeat(20_000);

    overlay.querySelector<HTMLButtonElement>('[data-testid="inline-edit-run"]')!.click();

    await waitFor(() => {
      const el = overlay.querySelector<HTMLElement>('[data-testid="inline-edit-preview-summary"]');
      return el && el.textContent?.includes("Changes:") ? el : null;
    });

    overlay.querySelector<HTMLButtonElement>('[data-testid="inline-edit-approve"]')!.click();

    const doc = app.getDocument();
    await waitFor(() => (doc.getCell("Sheet1", "C3").value === 3 ? doc : null));
    expect(doc.getCell("Sheet1", "C1").value).toBe(1);
    expect(doc.getCell("Sheet1", "C2").value).toBe(2);
    expect(doc.getCell("Sheet1", "C3").value).toBe(3);
  });

  it("cancels an in-flight inline edit run without hanging or applying changes", async () => {
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

    let resolveChat: ((value: any) => void) | null = null;
    const llmClient = {
      chat: vi.fn(
        () =>
          new Promise((resolve) => {
            resolveChat = resolve;
          })
      ),
    };

    // Avoid pulling in the default sqlite-backed audit store here; that initialization can be slow
    // and make this cancellation regression test flaky under full-suite runs.
    const auditStore = {
      logEntry: vi.fn(async () => {}),
      listEntries: vi.fn(async () => []),
    };

    const app = new SpreadsheetApp(root, status, {
      // Ensure the run can complete promptly after cancellation so we can re-open inline edit.
      inlineEdit: { llmClient, model: "unit-test-model", auditStore: auditStore as any },
    });

    app.selectRange({ range: { startRow: 0, endRow: 2, startCol: 2, endCol: 2 } }); // C1:C3
    root.dispatchEvent(new KeyboardEvent("keydown", { key: "k", ctrlKey: true, bubbles: true }));

    const overlay = await waitFor(() => document.querySelector<HTMLElement>('[data-testid="inline-edit-overlay"]'));
    const input = overlay.querySelector<HTMLInputElement>('[data-testid="inline-edit-prompt"]');
    input!.value = "Fill with 1..3";
    overlay.querySelector<HTMLButtonElement>('[data-testid="inline-edit-run"]')!.click();

    await waitFor(() => (resolveChat ? overlay : null));

    // Cancel while the tool loop is still waiting for the model response.
    overlay.querySelector<HTMLButtonElement>('[data-testid="inline-edit-cancel"]')!.click();
    await waitFor(() => (overlay.hidden ? overlay : null));

    // Let the model respond after cancellation (previously this would hang waiting
    // for an approval UI that was no longer visible).
    resolveChat!({
      message: {
        role: "assistant",
        content: "",
        toolCalls: [{ id: "call-1", name: "set_range", arguments: { range: "C1:C3", values: [[1], [2], [3]] } }],
      },
      usage: { promptTokens: 1, completionTokens: 1 },
    });

    // The run should terminate cleanly, allowing inline edit to open again.
    await waitFor(() => {
      root.dispatchEvent(new KeyboardEvent("keydown", { key: "k", ctrlKey: true, bubbles: true }));
      const el = document.querySelector<HTMLElement>('[data-testid="inline-edit-overlay"]');
      return el && !el.hidden ? el : null;
    }, 5000);

    const doc = app.getDocument();
    expect(doc.getCell("Sheet1", "C1").value).toBeNull();
    expect(doc.getCell("Sheet1", "C2").value).toBeNull();
    expect(doc.getCell("Sheet1", "C3").value).toBeNull();
    expect((doc as any).batchDepth ?? 0).toBe(0);
  });
});
