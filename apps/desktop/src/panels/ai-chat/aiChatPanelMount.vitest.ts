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

      const renderer = createPanelBodyRenderer({ getDocumentController });

      const body = document.createElement("div");
      document.body.appendChild(body);

      await act(async () => {
        renderer.renderPanelBody(PanelIds.AI_CHAT, body);
      });

      expect(getDocumentController).toHaveBeenCalled();
      expect(mocks.getDesktopLLMClient).toHaveBeenCalled();
      expect(mocks.getDesktopAIAuditStore).toHaveBeenCalled();
      expect(mocks.purgeLegacyDesktopLLMSettings).toHaveBeenCalled();

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
});
