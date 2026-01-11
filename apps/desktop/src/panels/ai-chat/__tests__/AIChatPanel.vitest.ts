// @vitest-environment jsdom
import React, { act } from "react";
import { createRoot } from "react-dom/client";
import { afterEach, describe, expect, it, vi } from "vitest";

import { AIChatPanel } from "../AIChatPanel";
import type { LLMClient, ToolExecutor } from "../../../../../../packages/llm/src/types.js";

function flushPromises() {
  return new Promise<void>((resolve) => setTimeout(resolve, 0));
}

function setNativeInputValue(input: HTMLInputElement, value: string) {
  const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, "value")?.set;
  if (!setter) throw new Error("Missing HTMLInputElement.value setter");
  setter.call(input, value);
}

async function waitFor(condition: () => boolean, timeoutMs = 1_000) {
  const start = Date.now();
  while (Date.now() - start < timeoutMs) {
    if (condition()) return;
    // Give React + microtasks a chance to run.
    // eslint-disable-next-line no-await-in-loop
    await flushPromises();
  }
  throw new Error("Timed out waiting for condition");
}

afterEach(() => {
  document.body.innerHTML = "";
  vi.restoreAllMocks();
});

describe("AIChatPanel tool-calling history", () => {
  it("does not send UI-only tool messages (missing toolCallId) on subsequent turns", async () => {
    // React 18+ requires this flag for `act` to behave correctly in non-Jest runners.
    // https://react.dev/reference/react/act#configuring-your-test-environment
    (globalThis as any).IS_REACT_ACT_ENVIRONMENT = true;
    // jsdom may not implement `crypto.randomUUID`; the panel uses it for message IDs.
    let uuid = 0;
    vi.stubGlobal("crypto", { randomUUID: () => `uuid-${uuid++}` } as any);

    let sawInvalidToolMessage = false;

    let callIndex = 0;
    const chat = vi.fn(async (request) => {
      for (const m of request.messages) {
        if (m.role === "tool" && !(m as any).toolCallId) {
          sawInvalidToolMessage = true;
        }
      }

      callIndex += 1;
      // 1st call (turn 1): request tool call
      if (callIndex === 1) {
        return {
          message: {
            role: "assistant",
            content: "",
            toolCalls: [{ id: "call-1", name: "getData", arguments: { range: "A1:B2" } }],
          },
        };
      }

      // 2nd call (turn 1): final response after tool execution
      if (callIndex === 2) {
        return {
          message: {
            role: "assistant",
            content: "Here are the results.",
          },
        };
      }

      // 3rd call (turn 2): should include tool result message with toolCallId
      if (callIndex === 3) {
        return {
          message: {
            role: "assistant",
            content: "Second turn response.",
          },
        };
      }

      throw new Error(`Unexpected chat call index: ${callIndex}`);
    });

    const client: LLMClient = { chat };
    const toolExecutor: ToolExecutor = {
      tools: [
        {
          name: "getData",
          description: "Mock tool",
          parameters: { type: "object", properties: {} },
        },
      ],
      execute: vi.fn(async () => ({ value: 42 })),
    };

    const container = document.createElement("div");
    document.body.appendChild(container);
    const root = createRoot(container);

    await act(async () => {
      root.render(React.createElement(AIChatPanel, { client, toolExecutor, systemPrompt: "system" }));
    });

    const input = container.querySelector("input");
    const sendButton = Array.from(container.querySelectorAll("button")).find((b) => b.textContent === "Send");
    expect(input).toBeInstanceOf(HTMLInputElement);
    expect(sendButton).toBeInstanceOf(HTMLButtonElement);

    const inputEl = input as HTMLInputElement;
    const buttonEl = sendButton as HTMLButtonElement;

    // Turn 1
    await act(async () => {
      setNativeInputValue(inputEl, "First question");
      inputEl.dispatchEvent(new Event("input", { bubbles: true }));
      inputEl.dispatchEvent(new Event("change", { bubbles: true }));
    });
    await act(async () => {
      buttonEl.click();
      await waitFor(() => callIndex === 2);
    });
    await act(async () => {
      await waitFor(() => container.textContent?.includes("Here are the results.") ?? false);
    });

    // Turn 2
    await act(async () => {
      setNativeInputValue(inputEl, "Follow-up question");
      inputEl.dispatchEvent(new Event("input", { bubbles: true }));
      inputEl.dispatchEvent(new Event("change", { bubbles: true }));
    });
    await act(async () => {
      buttonEl.click();
      await waitFor(() => callIndex === 3);
    });

    const thirdRequest = chat.mock.calls[2]?.[0] as any;
    const toolMessages = (thirdRequest.messages ?? []).filter((m: any) => m.role === "tool");

    expect(sawInvalidToolMessage).toBe(false);
    expect(toolMessages).toHaveLength(1);
    expect(toolMessages[0].toolCallId).toBe("call-1");
    expect(toolMessages[0].content).toContain("\"value\": 42");
    // The UI-only tool "call display" message should never be part of the LLM history.
    expect(toolMessages[0].content).not.toContain("getData(");

    await act(async () => {
      root.unmount();
    });
  });

  it("clears the pending assistant placeholder when sendMessage throws", async () => {
    (globalThis as any).IS_REACT_ACT_ENVIRONMENT = true;
    vi.stubGlobal("crypto", { randomUUID: () => "uuid-1" } as any);

    const sendMessage = vi.fn(async () => {
      throw new Error("Boom");
    });

    const container = document.createElement("div");
    document.body.appendChild(container);
    const root = createRoot(container);

    await act(async () => {
      root.render(
        React.createElement(AIChatPanel, {
          systemPrompt: "system",
          sendMessage,
        }),
      );
    });

    const input = container.querySelector("input");
    const sendButton = Array.from(container.querySelectorAll("button")).find((b) => b.textContent === "Send");
    expect(input).toBeInstanceOf(HTMLInputElement);
    expect(sendButton).toBeInstanceOf(HTMLButtonElement);

    const inputEl = input as HTMLInputElement;
    const buttonEl = sendButton as HTMLButtonElement;

    await act(async () => {
      setNativeInputValue(inputEl, "Hi");
      inputEl.dispatchEvent(new Event("input", { bubbles: true }));
      inputEl.dispatchEvent(new Event("change", { bubbles: true }));
    });

    await act(async () => {
      buttonEl.click();
      await waitFor(() => sendMessage.mock.calls.length === 1);
    });

    await act(async () => {
      await waitFor(() => container.textContent?.includes("Error: Boom") ?? false);
    });

    // Regression: previously the panel would append an error message but leave the
    // pending assistant placeholder stuck in a permanent "thinking…" state.
    expect(container.textContent).not.toContain("thinking");

    await act(async () => {
      root.unmount();
    });
  });

  it("shows a warning banner when verification is low confidence", async () => {
    (globalThis as any).IS_REACT_ACT_ENVIRONMENT = true;
    vi.stubGlobal("crypto", { randomUUID: () => "uuid-1" } as any);

    const sendMessage = vi.fn(async () => {
      return {
        messages: [],
        final: "Answer.",
        verification: { needs_tools: true, used_tools: true, verified: true, confidence: 0.5, warnings: [] }
      };
    });

    const container = document.createElement("div");
    document.body.appendChild(container);
    const root = createRoot(container);

    await act(async () => {
      root.render(
        React.createElement(AIChatPanel, {
          systemPrompt: "system",
          sendMessage
        })
      );
    });

    const input = container.querySelector("input");
    const sendButton = Array.from(container.querySelectorAll("button")).find((b) => b.textContent === "Send");
    expect(input).toBeInstanceOf(HTMLInputElement);
    expect(sendButton).toBeInstanceOf(HTMLButtonElement);

    const inputEl = input as HTMLInputElement;
    const buttonEl = sendButton as HTMLButtonElement;

    await act(async () => {
      setNativeInputValue(inputEl, "Hi");
      inputEl.dispatchEvent(new Event("input", { bubbles: true }));
      inputEl.dispatchEvent(new Event("change", { bubbles: true }));
    });

    await act(async () => {
      buttonEl.click();
      await waitFor(() => sendMessage.mock.calls.length === 1);
    });

    await act(async () => {
      await waitFor(() => container.textContent?.includes("Unverified answer.") ?? false);
    });

    await act(async () => {
      root.unmount();
    });
  });
});
