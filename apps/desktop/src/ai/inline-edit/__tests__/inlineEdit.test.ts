// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { SpreadsheetApp } from "../../../app/spreadsheetApp";

import { LocalStorageBinaryStorage } from "@formula/ai-audit/browser";
import { SqliteAIAuditStore } from "@formula/ai-audit/sqlite";
import { createRequire } from "node:module";

const OPENAI_API_KEY_STORAGE_KEY = "formula:openaiApiKey";

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
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
  });

  beforeEach(() => {
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
    expect(overlay.style.display).not.toBe("none");

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

  it("uses the OpenAIClient fallback when no inlineEdit llmClient is injected (localStorage key)", async () => {
    const apiKey = "sk-test-inline-edit";
    localStorage.setItem(OPENAI_API_KEY_STORAGE_KEY, apiKey);

    let callCount = 0;
    const fetchMock = vi.fn(async (_url: string, init: any) => {
      callCount++;
      expect(init?.headers?.Authorization).toBe(`Bearer ${apiKey}`);

      if (callCount === 1) {
        return {
          ok: true,
          json: async () => ({
            choices: [
              {
                message: {
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
            usage: { prompt_tokens: 10, completion_tokens: 5 },
          }),
        } as any;
      }

      return {
        ok: true,
        json: async () => ({
          choices: [{ message: { content: "done" } }],
          usage: { prompt_tokens: 1, completion_tokens: 1 },
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

    // No inlineEdit config passed; controller should pick up localStorage key.
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

    overlay.querySelector<HTMLButtonElement>('[data-testid="inline-edit-approve"]')!.click();

    const doc = app.getDocument();
    await waitFor(() => (doc.getCell("Sheet1", "C3").value === 3 ? doc : null));
    expect(doc.getCell("Sheet1", "C1").value).toBe(1);
    expect(doc.getCell("Sheet1", "C2").value).toBe(2);
    expect(doc.getCell("Sheet1", "C3").value).toBe(3);

    expect(fetchMock).toHaveBeenCalled();
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

    const app = new SpreadsheetApp(root, status, {
      inlineEdit: { llmClient, model: "unit-test-model" },
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
    await waitFor(() => (overlay.style.display === "none" ? overlay : null));

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
      return el && el.style.display !== "none" ? el : null;
    });

    const doc = app.getDocument();
    expect(doc.getCell("Sheet1", "C1").value).toBeNull();
    expect(doc.getCell("Sheet1", "C2").value).toBeNull();
    expect(doc.getCell("Sheet1", "C3").value).toBeNull();
    expect((doc as any).batchDepth ?? 0).toBe(0);
  });
});
