// @vitest-environment jsdom

import { act } from "react";
import { afterEach, describe, expect, it, vi } from "vitest";

import { AnthropicClient } from "../../../../../packages/llm/src/anthropic.js";
import { OllamaChatClient } from "../../../../../packages/llm/src/ollama.js";

import { clearDesktopLLMConfig } from "../../ai/llm/settings.js";
import { DocumentController } from "../../document/documentController.js";

const TEST_TIMEOUT_MS = 30_000;

const mocks = vi.hoisted(() => {
  return {
    createAiChatOrchestrator: vi.fn(() => ({ sendMessage: vi.fn(), sessionId: "test-session" })),
    getDesktopAIAuditStore: vi.fn(() => ({
      logEntry: vi.fn(async () => {}),
      listEntries: vi.fn(async () => []),
    })),
  };
});

vi.mock("../../ai/chat/orchestrator.js", () => ({
  createAiChatOrchestrator: mocks.createAiChatOrchestrator,
}));

vi.mock("../../ai/audit/auditStore.js", () => ({
  getDesktopAIAuditStore: mocks.getDesktopAIAuditStore,
}));

// React 18 relies on this flag to suppress act() warnings in test runners.
// eslint-disable-next-line @typescript-eslint/no-explicit-any
(globalThis as any).IS_REACT_ACT_ENVIRONMENT = true;

function setNativeInputValue(input: HTMLInputElement, value: string) {
  const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, "value")?.set;
  if (!setter) throw new Error("Missing HTMLInputElement.value setter");
  setter.call(input, value);
}

function setNativeSelectValue(select: HTMLSelectElement, value: string) {
  const setter = Object.getOwnPropertyDescriptor(HTMLSelectElement.prototype, "value")?.set;
  if (!setter) throw new Error("Missing HTMLSelectElement.value setter");
  setter.call(select, value);
}

function clearAiStorage() {
  // Node 25 ships an experimental `globalThis.localStorage` accessor that throws
  // unless Node is started with `--localstorage-file`. Guard all access so our
  // jsdom-based UI tests don't crash on newer Node versions.
  try {
    clearDesktopLLMConfig();
  } catch {
    // ignore
  }
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
  });

  it(
    "mounts via renderPanelBody and shows setup state when no API key is set",
    async () => {
      clearAiStorage();

      const { createPanelBodyRenderer } = await import("../panelBodyRenderer.js");
      const { PanelIds } = await import("../panelRegistry.js");
      const renderer = createPanelBodyRenderer({
      getDocumentController: () => {
        throw new Error("document controller should not be requested when API key is missing");
      },
    });

    const body = document.createElement("div");
    document.body.appendChild(body);

    await act(async () => {
      renderer.renderPanelBody(PanelIds.AI_CHAT, body);
    });

    expect(body.textContent).toContain("AI chat setup");

      act(() => {
        renderer.cleanup([]);
      });
    },
    TEST_TIMEOUT_MS
  );

  it(
    "can save an OpenAI API key and transition into the chat UI",
    async () => {
      clearAiStorage();

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

    const keyInput = body.querySelector('[data-testid="ai-openai-api-key"]');
    expect(keyInput).toBeInstanceOf(HTMLInputElement);

    await act(async () => {
      setNativeInputValue(keyInput as HTMLInputElement, "sk-test-key");
      keyInput?.dispatchEvent(new Event("input", { bubbles: true }));
      keyInput?.dispatchEvent(new Event("change", { bubbles: true }));
    });

    const saveBtn = body.querySelector('[data-testid="ai-save-settings"]');
    expect(saveBtn).toBeInstanceOf(HTMLButtonElement);

    await act(async () => {
      (saveBtn as HTMLButtonElement).click();
    });

    expect(getDocumentController).toHaveBeenCalled();
    expect(mocks.getDesktopAIAuditStore).toHaveBeenCalled();
    const chatTab = body.querySelector('[data-testid="ai-tab-chat"]');
    expect(chatTab).toBeInstanceOf(HTMLButtonElement);

    const agentTab = body.querySelector('[data-testid="ai-tab-agent"]');
    expect(agentTab).toBeInstanceOf(HTMLButtonElement);

    await act(async () => {
      (agentTab as HTMLButtonElement).click();
    });

    expect(body.querySelector('[data-testid="agent-goal"]')).toBeTruthy();

      act(() => {
        renderer.cleanup([]);
      });
    },
    TEST_TIMEOUT_MS
  );

  it(
    "selecting Anthropic provider wires up AnthropicClient",
    async () => {
      clearAiStorage();
      mocks.createAiChatOrchestrator.mockClear();

      const doc = new DocumentController();
    const getDocumentController = vi.fn(() => doc);

    const { createPanelBodyRenderer } = await import("../panelBodyRenderer.js");
    const { PanelIds } = await import("../panelRegistry.js");

    const renderer = createPanelBodyRenderer({ getDocumentController });
    const body = document.createElement("div");
    document.body.appendChild(body);

    await act(async () => {
      renderer.renderPanelBody(PanelIds.AI_CHAT, body);
    });

    const providerSelect = body.querySelector('[data-testid="ai-provider-select"]');
    expect(providerSelect).toBeInstanceOf(HTMLSelectElement);

    await act(async () => {
      setNativeSelectValue(providerSelect as HTMLSelectElement, "anthropic");
      providerSelect?.dispatchEvent(new Event("input", { bubbles: true }));
      providerSelect?.dispatchEvent(new Event("change", { bubbles: true }));
    });

    const keyInput = body.querySelector('[data-testid="ai-anthropic-api-key"]');
    expect(keyInput).toBeInstanceOf(HTMLInputElement);

    await act(async () => {
      setNativeInputValue(keyInput as HTMLInputElement, "sk-ant-test");
      keyInput?.dispatchEvent(new Event("input", { bubbles: true }));
      keyInput?.dispatchEvent(new Event("change", { bubbles: true }));
    });

    const saveBtn = body.querySelector('[data-testid="ai-save-settings"]') as HTMLButtonElement | null;
    expect(saveBtn).toBeInstanceOf(HTMLButtonElement);

    await act(async () => {
      saveBtn?.click();
    });

    const lastCall = mocks.createAiChatOrchestrator.mock.calls.at(-1)?.[0] as any;
    expect(lastCall?.llmClient).toBeInstanceOf(AnthropicClient);

      act(() => {
        renderer.cleanup([]);
      });
    },
    TEST_TIMEOUT_MS
  );

  it(
    "selecting Ollama provider wires up OllamaChatClient",
    async () => {
      clearAiStorage();
      mocks.createAiChatOrchestrator.mockClear();

      const doc = new DocumentController();
    const getDocumentController = vi.fn(() => doc);

    const { createPanelBodyRenderer } = await import("../panelBodyRenderer.js");
    const { PanelIds } = await import("../panelRegistry.js");

    const renderer = createPanelBodyRenderer({ getDocumentController });
    const body = document.createElement("div");
    document.body.appendChild(body);

    await act(async () => {
      renderer.renderPanelBody(PanelIds.AI_CHAT, body);
    });

    const providerSelect = body.querySelector('[data-testid="ai-provider-select"]');
    expect(providerSelect).toBeInstanceOf(HTMLSelectElement);

    await act(async () => {
      setNativeSelectValue(providerSelect as HTMLSelectElement, "ollama");
      providerSelect?.dispatchEvent(new Event("input", { bubbles: true }));
      providerSelect?.dispatchEvent(new Event("change", { bubbles: true }));
    });

    const baseUrlInput = body.querySelector('[data-testid="ai-ollama-base-url"]');
    expect(baseUrlInput).toBeInstanceOf(HTMLInputElement);

    await act(async () => {
      setNativeInputValue(baseUrlInput as HTMLInputElement, "http://127.0.0.1:11434");
      baseUrlInput?.dispatchEvent(new Event("input", { bubbles: true }));
      baseUrlInput?.dispatchEvent(new Event("change", { bubbles: true }));
    });

    const modelInput = body.querySelector('[data-testid="ai-ollama-model"]');
    expect(modelInput).toBeInstanceOf(HTMLInputElement);

    await act(async () => {
      setNativeInputValue(modelInput as HTMLInputElement, "llama3.1");
      modelInput?.dispatchEvent(new Event("input", { bubbles: true }));
      modelInput?.dispatchEvent(new Event("change", { bubbles: true }));
    });

    const saveBtn = body.querySelector('[data-testid="ai-save-settings"]') as HTMLButtonElement | null;
    expect(saveBtn).toBeInstanceOf(HTMLButtonElement);

    await act(async () => {
      saveBtn?.click();
    });

    const lastCall = mocks.createAiChatOrchestrator.mock.calls.at(-1)?.[0] as any;
    expect(lastCall?.llmClient).toBeInstanceOf(OllamaChatClient);

      act(() => {
        renderer.cleanup([]);
      });
    },
    TEST_TIMEOUT_MS
  );
});
