// @vitest-environment jsdom

import { act } from "react";
import { afterEach, describe, expect, it, vi } from "vitest";

import { DocumentController } from "../../document/documentController.js";

const TEST_TIMEOUT_MS = 30_000;

const mocks = vi.hoisted(() => {
  const sentinelClient = { __sentinel: "desktop-llm-client" };

  return {
    sentinelClient,
    createAiChatOrchestrator: vi.fn((_options: any) => ({
      sendMessage: vi.fn(),
      sessionId: "test-session",
      dispose: vi.fn(async () => {}),
    })),
    getDesktopAIAuditStore: vi.fn(() => ({
      logEntry: vi.fn(async () => {}),
      listEntries: vi.fn(async () => []),
    })),
    getDesktopLLMClient: vi.fn(() => sentinelClient),
    getDesktopModel: vi.fn(() => "test-model"),
    purgeLegacyDesktopLLMSettings: vi.fn(),
  };
});

vi.mock("../../ai/chat/orchestrator.js", () => ({
  createAiChatOrchestrator: mocks.createAiChatOrchestrator,
}));

vi.mock("../../ai/audit/auditStore.js", () => ({
  getDesktopAIAuditStore: mocks.getDesktopAIAuditStore,
}));

vi.mock("../../ai/llm/desktopLLMClient.js", () => ({
  getDesktopLLMClient: mocks.getDesktopLLMClient,
  getDesktopModel: mocks.getDesktopModel,
  purgeLegacyDesktopLLMSettings: mocks.purgeLegacyDesktopLLMSettings,
}));

// React 18 relies on this flag to suppress act() warnings in test runners.
// eslint-disable-next-line @typescript-eslint/no-explicit-any
(globalThis as any).IS_REACT_ACT_ENVIRONMENT = true;

// Import the panel modules at file-evaluation time so Vite transform work does not
// count toward the per-test timeout.
//
// Keep these as dynamic imports (not static `import` statements) so the `vi.mock(...)`
// declarations above apply before the modules are evaluated.
const [{ createPanelBodyRenderer }, { PanelIds }] = await Promise.all([
  import("../panelBodyRenderer.js"),
  import("../panelRegistry.js"),
]);
// Preload the AI chat panel container module that `PanelBodyRenderer` loads via `React.lazy()`.
//
// Under full-suite Vitest runs, Vite may need to transform a large dependency graph the first
// time this module is imported. If that work happens inside an individual test, the DOM can take
// long enough to mount that we hit the smaller `waitFor()` timeouts and flake.
//
// Warming the module cache here keeps the tests focused on `renderPanelBody` behavior rather than
// transform timing.
await import("./AIChatPanelContainer.js");

async function waitFor(assertion: () => void, timeoutMs = 2_000) {
  const started = Date.now();
  // eslint-disable-next-line no-constant-condition
  while (true) {
    try {
      assertion();
      return;
    } catch (err) {
      if (Date.now() - started > timeoutMs) throw err;
    }
    await act(async () => {
      await new Promise<void>((resolve) => setTimeout(resolve, 10));
    });
  }
}

function clearAiStorage() {
  // Node 25 ships an experimental `globalThis.localStorage` accessor that throws
  // unless Node is started with `--localstorage-file`. Guard all access so our
  // jsdom-based UI tests don't crash on newer Node versions.
  try {
    globalThis.localStorage?.clear();
  } catch {
    // ignore
  }
}

describe("AI chat panel", () => {
  afterEach(() => {
    document.body.innerHTML = "";
    clearAiStorage();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    delete (globalThis as any).__formulaSpreadsheetIsEditing;
    mocks.createAiChatOrchestrator.mockClear();
    mocks.getDesktopAIAuditStore.mockClear();
    mocks.getDesktopLLMClient.mockClear();
    mocks.getDesktopModel.mockClear();
    mocks.purgeLegacyDesktopLLMSettings.mockClear();
  });

  it(
    "mounts via renderPanelBody and shows the chat runtime UI",
    async () => {
      const doc = new DocumentController();
      const getDocumentController = vi.fn(() => doc);
      const app = { getCellComputedValueForSheet: vi.fn(() => 123) };
      const getSpreadsheetApp = vi.fn(() => app);

      const renderer = createPanelBodyRenderer({ getDocumentController, getSpreadsheetApp });

      const body = document.createElement("div");
      document.body.appendChild(body);

      await act(async () => {
        renderer.renderPanelBody(PanelIds.AI_CHAT, body);
      });

      await waitFor(() => {
        expect(body.querySelector('[data-testid="ai-tab-chat"]')).toBeInstanceOf(HTMLButtonElement);
      }, 10_000);

      expect(getDocumentController).toHaveBeenCalled();
      expect(mocks.getDesktopLLMClient).toHaveBeenCalled();
      expect(mocks.getDesktopAIAuditStore).toHaveBeenCalled();
      expect(mocks.purgeLegacyDesktopLLMSettings).toHaveBeenCalled();
      // When the toggle is off, we should not eagerly call into SpreadsheetApp for computed values.
      expect(getSpreadsheetApp).not.toHaveBeenCalled();

      expect(body.querySelector('[data-testid="ai-tab-chat"]')).toBeInstanceOf(HTMLButtonElement);
      expect(body.querySelector('[data-testid="ai-tab-agent"]')).toBeInstanceOf(HTMLButtonElement);
      const includeFormulaValues = body.querySelector('[data-testid="ai-include-formula-values"]');
      expect(includeFormulaValues).toBeInstanceOf(HTMLInputElement);
      expect((includeFormulaValues as HTMLInputElement).checked).toBe(false);

      // Cursor-only backend: the Settings button should be gone.
      expect(body.querySelector('[data-testid="ai-open-settings"]')).toBeNull();

      const lastCall = mocks.createAiChatOrchestrator.mock.calls.at(-1)?.[0] as any;
      expect(lastCall?.llmClient).toBe(mocks.sentinelClient);
      expect(lastCall?.model).toBe("test-model");
      expect(lastCall?.toolExecutorOptions?.include_formula_values).toBe(false);
      expect(lastCall?.getCellComputedValueForSheet).toBeUndefined();
      const orchestrator = mocks.createAiChatOrchestrator.mock.results.at(-1)?.value as any;
      act(() => {
        renderer.cleanup([]);
      });

      expect(orchestrator.dispose).toHaveBeenCalled();
    },
    TEST_TIMEOUT_MS
  );

  it(
    "hydrates include_formula_values from localStorage",
    async () => {
      const doc = new DocumentController();
      const getDocumentController = vi.fn(() => doc);

      window.localStorage.setItem("formula.ai.includeFormulaValues", "true");

      const renderer = createPanelBodyRenderer({ getDocumentController });

      const body = document.createElement("div");
      document.body.appendChild(body);

      await act(async () => {
        renderer.renderPanelBody(PanelIds.AI_CHAT, body);
      });

      await waitFor(() => {
        expect(body.querySelector('[data-testid="ai-include-formula-values"]')).toBeInstanceOf(HTMLInputElement);
      }, 10_000);

      const includeFormulaValues = body.querySelector('[data-testid="ai-include-formula-values"]');
      expect(includeFormulaValues).toBeInstanceOf(HTMLInputElement);
      expect((includeFormulaValues as HTMLInputElement).checked).toBe(true);

      const lastCall = mocks.createAiChatOrchestrator.mock.calls.at(-1)?.[0] as any;
      expect(lastCall?.toolExecutorOptions?.include_formula_values).toBe(true);

      const orchestrator = mocks.createAiChatOrchestrator.mock.results.at(-1)?.value as any;
      act(() => {
        renderer.cleanup([]);
      });
      expect(orchestrator.dispose).toHaveBeenCalled();
    },
    TEST_TIMEOUT_MS
  );

  it(
    "persists include_formula_values toggle to localStorage and recreates orchestrator",
    async () => {
      const doc = new DocumentController();
      const getDocumentController = vi.fn(() => doc);
      const app = { getCellComputedValueForSheet: vi.fn(() => 123) };
      const getSpreadsheetApp = vi.fn(() => app);

      const renderer = createPanelBodyRenderer({ getDocumentController, getSpreadsheetApp });

      const body = document.createElement("div");
      document.body.appendChild(body);

      await act(async () => {
        renderer.renderPanelBody(PanelIds.AI_CHAT, body);
      });

      await waitFor(() => {
        expect(body.querySelector('[data-testid="ai-include-formula-values"]')).toBeInstanceOf(HTMLInputElement);
      }, 10_000);

      const checkbox = body.querySelector('[data-testid="ai-include-formula-values"]') as HTMLInputElement | null;
      expect(checkbox).toBeInstanceOf(HTMLInputElement);
      expect(checkbox?.checked).toBe(false);
      expect(window.localStorage.getItem("formula.ai.includeFormulaValues")).toBeNull();

      const priorCalls = mocks.createAiChatOrchestrator.mock.calls.length;

      await act(async () => {
        checkbox!.click();
      });

      await waitFor(() => {
        expect(window.localStorage.getItem("formula.ai.includeFormulaValues")).toBe("true");
      });
      await waitFor(() => {
        // When enabled, the chat panel should consult SpreadsheetApp for computed values.
        expect(getSpreadsheetApp).toHaveBeenCalled();
        expect(mocks.createAiChatOrchestrator.mock.calls.length).toBeGreaterThan(priorCalls);
      });

      const lastCall = mocks.createAiChatOrchestrator.mock.calls.at(-1)?.[0] as any;
      expect(lastCall?.toolExecutorOptions?.include_formula_values).toBe(true);
      expect(lastCall?.getCellComputedValueForSheet).toBeInstanceOf(Function);
      expect(lastCall?.getCellComputedValueForSheet("Sheet1", { row: 0, col: 0 })).toBe(123);
      expect(app.getCellComputedValueForSheet).toHaveBeenCalledWith("Sheet1", { row: 0, col: 0 });

      const orchestrator = mocks.createAiChatOrchestrator.mock.results.at(-1)?.value as any;
      act(() => {
        renderer.cleanup([]);
      });
      expect(orchestrator.dispose).toHaveBeenCalled();
    },
    TEST_TIMEOUT_MS
  );

  it(
    "disables the agent Run button while spreadsheet edit mode is active",
    async () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      (globalThis as any).__formulaSpreadsheetIsEditing = true;

      const doc = new DocumentController();
      const getDocumentController = vi.fn(() => doc);

      const renderer = createPanelBodyRenderer({ getDocumentController });

      const body = document.createElement("div");
      document.body.appendChild(body);

      await act(async () => {
        renderer.renderPanelBody(PanelIds.AI_CHAT, body);
      });

      await waitFor(() => {
        expect(body.querySelector('[data-testid="ai-tab-agent"]')).toBeInstanceOf(HTMLButtonElement);
      }, 10_000);

      const agentTab = body.querySelector('[data-testid="ai-tab-agent"]') as HTMLButtonElement | null;
      expect(agentTab).toBeInstanceOf(HTMLButtonElement);
      await act(async () => {
        agentTab?.click();
      });

      await waitFor(() => {
        expect(body.querySelector('[data-testid="agent-goal"]')).toBeInstanceOf(HTMLTextAreaElement);
      });

      const goal = body.querySelector('[data-testid="agent-goal"]') as HTMLTextAreaElement | null;
      expect(goal).toBeInstanceOf(HTMLTextAreaElement);
      await act(async () => {
        if (!goal) return;
        goal.value = "Summarize this workbook";
        goal.dispatchEvent(new Event("input", { bubbles: true }));
      });

      await waitFor(() => {
        expect(body.querySelector('[data-testid="agent-run"]')).toBeInstanceOf(HTMLButtonElement);
      });

      const runButton = body.querySelector('[data-testid="agent-run"]') as HTMLButtonElement | null;
      expect(runButton).toBeInstanceOf(HTMLButtonElement);
      expect(runButton?.disabled).toBe(true);

      act(() => {
        renderer.cleanup([]);
      });
    },
    TEST_TIMEOUT_MS,
  );
});
