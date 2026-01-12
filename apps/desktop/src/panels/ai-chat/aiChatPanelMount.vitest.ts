// @vitest-environment jsdom

import { act } from "react";
import { afterEach, describe, expect, it, vi } from "vitest";

import { DocumentController } from "../../document/documentController.js";

const TEST_TIMEOUT_MS = 30_000;

const mocks = vi.hoisted(() => {
  const sentinelClient = { __sentinel: "desktop-llm-client" };

  return {
    sentinelClient,
    createAiChatOrchestrator: vi.fn((_options: any) => ({ sendMessage: vi.fn(), sessionId: "test-session" })),
    getDesktopAIAuditStore: vi.fn(() => ({
      logEntry: vi.fn(async () => {}),
      listEntries: vi.fn(async () => []),
    })),
    getDesktopLLMClient: vi.fn(() => sentinelClient),
    getDesktopModel: vi.fn(() => "test-model"),
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
}));

// React 18 relies on this flag to suppress act() warnings in test runners.
// eslint-disable-next-line @typescript-eslint/no-explicit-any
(globalThis as any).IS_REACT_ACT_ENVIRONMENT = true;

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
  });

  it(
    "mounts via renderPanelBody and shows the chat runtime UI",
    async () => {
      const doc = new DocumentController();
      const getDocumentController = vi.fn(() => doc);

      const { createPanelBodyRenderer } = await import("../panelBodyRenderer.js");
      const { PanelIds } = await import("../panelRegistry.js");
      const renderer = createPanelBodyRenderer({
        getDocumentController,
      });

      const body = document.createElement("div");
      document.body.appendChild(body);

      await act(async () => {
        renderer.renderPanelBody(PanelIds.AI_CHAT, body);
      });

      expect(getDocumentController).toHaveBeenCalled();
      expect(mocks.getDesktopLLMClient).toHaveBeenCalled();
      expect(mocks.getDesktopAIAuditStore).toHaveBeenCalled();

      expect(body.querySelector('[data-testid="ai-tab-chat"]')).toBeInstanceOf(HTMLButtonElement);
      expect(body.querySelector('[data-testid="ai-tab-agent"]')).toBeInstanceOf(HTMLButtonElement);

      // Cursor-only backend: the Settings button should be gone.
      expect(body.querySelector('[data-testid="ai-open-settings"]')).toBeNull();

      const lastCall = mocks.createAiChatOrchestrator.mock.calls.at(-1)?.[0] as any;
      expect(lastCall?.llmClient).toBe(mocks.sentinelClient);
      expect(lastCall?.model).toBe("test-model");

      act(() => {
        renderer.cleanup([]);
      });
    },
    TEST_TIMEOUT_MS,
  );
});

