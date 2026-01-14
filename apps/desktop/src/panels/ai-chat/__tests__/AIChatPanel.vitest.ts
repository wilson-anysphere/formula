// @vitest-environment jsdom
import React, { act } from "react";
import { createRoot } from "react-dom/client";
import { afterEach, describe, expect, it, vi } from "vitest";

import { AIChatPanel } from "../AIChatPanel";
import type { LLMClient, ToolExecutor } from "../../../../../../packages/llm/src/index.js";

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
    const chat = vi.fn<LLMClient["chat"]>(async (request) => {
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
    expect(toolMessages[0].content).toContain("\"value\":42");
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

  it("suppresses streamed text once tool calling begins", async () => {
    (globalThis as any).IS_REACT_ACT_ENVIRONMENT = true;
    vi.stubGlobal("crypto", { randomUUID: () => "uuid-1" } as any);

    let resumeStream: (() => void) | null = null;
    const gate = new Promise<void>((resolve) => {
      resumeStream = resolve;
    });

    const sendMessage = vi.fn(async (args: any) => {
      args.onStreamEvent?.({ type: "text", delta: "Hello" });
      await gate;

      args.onStreamEvent?.({ type: "tool_call_start", id: "call-1", name: "getData" });
      // This text should be ignored by the UI (we only stream the final answer).
      args.onStreamEvent?.({ type: "text", delta: "SHOULD_NOT_RENDER" });
      args.onStreamEvent?.({ type: "done" });

      return { messages: [], final: "Final" };
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
    const sendEl = sendButton as HTMLButtonElement;

    await act(async () => {
      setNativeInputValue(inputEl, "Hi");
      inputEl.dispatchEvent(new Event("input", { bubbles: true }));
      inputEl.dispatchEvent(new Event("change", { bubbles: true }));
    });

    await act(async () => {
      sendEl.click();
      await waitFor(() => sendMessage.mock.calls.length === 1);
    });

    await act(async () => {
      await waitFor(() => container.textContent?.includes("Hello") ?? false);
    });

    // Resume the mocked stream.
    await act(async () => {
      resumeStream?.();
    });

    await waitFor(() => container.textContent?.includes("Final") ?? false);

    expect(container.textContent).not.toContain("SHOULD_NOT_RENDER");

    await act(async () => {
      root.unmount();
    });
  });

  it("can cancel an in-flight streaming response", async () => {
    (globalThis as any).IS_REACT_ACT_ENVIRONMENT = true;
    vi.stubGlobal("crypto", { randomUUID: () => "uuid-1" } as any);

    const client: LLMClient = {
      chat: vi.fn(async () => {
        throw new Error("chat() should not be called when streamChat is available");
      }),
      streamChat: (async function* (request: any) {
        yield { type: "text", delta: "Hello" };
        await new Promise<void>((resolve) => {
          request.signal?.addEventListener("abort", () => resolve(), { once: true });
        });
        const err = new Error("Aborted");
        (err as any).name = "AbortError";
        throw err;
      }) as any,
    };

    const toolExecutor: ToolExecutor = {
      tools: [],
      execute: vi.fn(async () => ({})),
    };

    const container = document.createElement("div");
    document.body.appendChild(container);
    const root = createRoot(container);

    await act(async () => {
      root.render(React.createElement(AIChatPanel, { client, toolExecutor, systemPrompt: "system" }));
    });

    const input = container.querySelector("input");
    const sendButton = Array.from(container.querySelectorAll("button")).find((b) => b.textContent === "Send");
    const cancelButton = Array.from(container.querySelectorAll("button")).find((b) => b.textContent === "Cancel");
    expect(input).toBeInstanceOf(HTMLInputElement);
    expect(sendButton).toBeInstanceOf(HTMLButtonElement);
    expect(cancelButton).toBeInstanceOf(HTMLButtonElement);

    const inputEl = input as HTMLInputElement;
    const sendEl = sendButton as HTMLButtonElement;
    const cancelEl = cancelButton as HTMLButtonElement;

    await act(async () => {
      setNativeInputValue(inputEl, "Hi");
      inputEl.dispatchEvent(new Event("input", { bubbles: true }));
      inputEl.dispatchEvent(new Event("change", { bubbles: true }));
    });

    await act(async () => {
      sendEl.click();
    });

    await act(async () => {
      await waitFor(() => container.textContent?.includes("Hello") ?? false);
    });

    // Cancel while the stream is still pending.
    await act(async () => {
      cancelEl.click();
    });

    await act(async () => {
      await waitFor(() => container.textContent?.includes("Cancelled.") ?? false);
    });

    expect(container.textContent).not.toContain("thinking");

    await act(async () => {
      root.unmount();
    });
  });
});

describe("AIChatPanel attachments UI", () => {
  it("never inlines raw table/range attachment data into the prompt (demo mode)", async () => {
    (globalThis as any).IS_REACT_ACT_ENVIRONMENT = true;
    vi.stubGlobal("crypto", { randomUUID: () => "uuid-1" } as any);

    const secret = "TOP SECRET";
    const chat = vi.fn<LLMClient["chat"]>(async () => {
      return {
        message: {
          role: "assistant",
          content: "Ok",
        },
      };
    });

    const client: LLMClient = { chat };
    const toolExecutor: ToolExecutor = {
      tools: [],
      execute: vi.fn(async () => ({})),
    };

    const getSelectionAttachment = vi.fn(() => ({
      type: "range" as const,
      reference: "Sheet1!A1:B2",
      data: { snapshot: secret },
    }));

    const container = document.createElement("div");
    document.body.appendChild(container);
    const root = createRoot(container);

    await act(async () => {
      root.render(
        React.createElement(AIChatPanel, {
          client,
          toolExecutor,
          systemPrompt: "system",
          getSelectionAttachment,
        }),
      );
    });

    const attachSelectionBtn = container.querySelector('[data-testid="ai-chat-attach-selection"]') as HTMLButtonElement | null;
    expect(attachSelectionBtn).toBeInstanceOf(HTMLButtonElement);
    expect(attachSelectionBtn?.disabled).toBe(false);

    await act(async () => {
      attachSelectionBtn!.click();
    });

    const input = container.querySelector("input") as HTMLInputElement | null;
    const sendButton = Array.from(container.querySelectorAll("button")).find((b) => b.textContent === "Send") as
      | HTMLButtonElement
      | undefined;
    expect(input).toBeInstanceOf(HTMLInputElement);
    expect(sendButton).toBeInstanceOf(HTMLButtonElement);

    await act(async () => {
      setNativeInputValue(input!, "Hello");
      input!.dispatchEvent(new Event("input", { bubbles: true }));
      input!.dispatchEvent(new Event("change", { bubbles: true }));
    });

    await act(async () => {
      sendButton!.click();
      await waitFor(() => chat.mock.calls.length === 1);
    });

    const request = chat.mock.calls[0]?.[0] as any;
    const combined = JSON.stringify(request?.messages ?? []);
    expect(combined).not.toContain(secret);

    await act(async () => {
      root.unmount();
    });
  });

  it("disables attach selection when no selection is available and exposes tooltip text", async () => {
    (globalThis as any).IS_REACT_ACT_ENVIRONMENT = true;
    vi.stubGlobal("crypto", { randomUUID: () => "uuid-1" } as any);

    const sendMessage = vi.fn(async () => {
      return { messages: [], final: "Ok" };
    });

    const getSelectionAttachment = vi.fn(() => null);

    const container = document.createElement("div");
    document.body.appendChild(container);
    const root = createRoot(container);

    await act(async () => {
      root.render(
        React.createElement(AIChatPanel, {
          systemPrompt: "system",
          sendMessage,
          getSelectionAttachment,
        }),
      );
    });

    const attachSelectionBtn = container.querySelector('[data-testid="ai-chat-attach-selection"]') as HTMLButtonElement | null;
    expect(attachSelectionBtn).toBeInstanceOf(HTMLButtonElement);
    expect(attachSelectionBtn?.disabled).toBe(true);

    // Disabled buttons don't reliably show native title tooltips, so we wrap them.
    const wrap = attachSelectionBtn?.parentElement;
    expect(wrap?.classList.contains("ai-chat-panel__attachment-button-wrap")).toBe(true);
    expect(wrap?.getAttribute("title")).toBe("No selection available");

    await act(async () => {
      root.unmount();
    });
  });

  it("shows a toast warning when attaching a clamped selection", async () => {
    (globalThis as any).IS_REACT_ACT_ENVIRONMENT = true;
    vi.stubGlobal("crypto", { randomUUID: () => "uuid-1" } as any);

    const toastRoot = document.createElement("div");
    toastRoot.id = "toast-root";
    document.body.appendChild(toastRoot);

    const sendMessage = vi.fn(async () => {
      return { messages: [], final: "Ok" };
    });

    const getSelectionAttachment = vi.fn(() => ({
      type: "range" as const,
      reference: "Sheet1!A1:D10",
      data: {
        clamped: { originalCellCount: 1_000_000, attachedCellCount: 200_000, maxCells: 200_000 },
      },
    }));

    const container = document.createElement("div");
    document.body.appendChild(container);
    const root = createRoot(container);

    await act(async () => {
      root.render(
        React.createElement(AIChatPanel, {
          systemPrompt: "system",
          sendMessage,
          getSelectionAttachment,
        }),
      );
    });

    const attachSelectionBtn = container.querySelector('[data-testid="ai-chat-attach-selection"]') as HTMLButtonElement | null;
    expect(attachSelectionBtn).toBeInstanceOf(HTMLButtonElement);

    await act(async () => {
      attachSelectionBtn!.click();
    });

    const toast = toastRoot.querySelector('[data-testid="toast"]');
    expect(toast?.textContent).toContain("Selection is too large");
    expect(toast?.textContent).toContain("Attached 200000");

    await act(async () => {
      root.unmount();
    });
  });

  it("disables attach table when no tables are available and exposes tooltip text", async () => {
    (globalThis as any).IS_REACT_ACT_ENVIRONMENT = true;
    vi.stubGlobal("crypto", { randomUUID: () => "uuid-1" } as any);

    const sendMessage = vi.fn(async () => {
      return { messages: [], final: "Ok" };
    });

    const getTableOptions = vi.fn(() => []);

    const container = document.createElement("div");
    document.body.appendChild(container);
    const root = createRoot(container);

    await act(async () => {
      root.render(
        React.createElement(AIChatPanel, {
          systemPrompt: "system",
          sendMessage,
          getTableOptions,
        }),
      );
    });

    const attachTableBtn = container.querySelector('[data-testid="ai-chat-attach-table"]') as HTMLButtonElement | null;
    expect(attachTableBtn).toBeInstanceOf(HTMLButtonElement);
    expect(attachTableBtn?.disabled).toBe(true);

    const wrap = attachTableBtn?.parentElement;
    expect(wrap?.classList.contains("ai-chat-panel__attachment-button-wrap")).toBe(true);
    expect(wrap?.getAttribute("title")).toBe("No tables available");

    await act(async () => {
      root.unmount();
    });
  });

  it("shows a toast when attach table is clicked but tables are unavailable", async () => {
    (globalThis as any).IS_REACT_ACT_ENVIRONMENT = true;
    vi.stubGlobal("crypto", { randomUUID: () => "uuid-1" } as any);

    const toastRoot = document.createElement("div");
    toastRoot.id = "toast-root";
    document.body.appendChild(toastRoot);

    const sendMessage = vi.fn(async () => {
      return { messages: [], final: "Ok" };
    });

    // Simulate a stale render where the button appears enabled, but by the time the user clicks
    // there are no tables. Use mutable state so the test remains robust even if the panel
    // re-renders (e.g. due to attachment-toolbar refreshes).
    let currentTables: Array<{ name: string }> = [{ name: "SalesTable" }];
    const getTableOptions = vi.fn<NonNullable<React.ComponentProps<typeof AIChatPanel>["getTableOptions"]>>(
      () => currentTables as any,
    );

    const container = document.createElement("div");
    document.body.appendChild(container);
    const root = createRoot(container);

    await act(async () => {
      root.render(
        React.createElement(AIChatPanel, {
          systemPrompt: "system",
          sendMessage,
          getTableOptions,
        }),
      );
    });

    const attachTableBtn = container.querySelector('[data-testid="ai-chat-attach-table"]') as HTMLButtonElement | null;
    expect(attachTableBtn).toBeInstanceOf(HTMLButtonElement);
    expect(attachTableBtn?.disabled).toBe(false);

    // Tables disappear before the click handler reads them.
    currentTables = [];

    await act(async () => {
      attachTableBtn!.click();
    });

    const toast = toastRoot.querySelector('[data-testid="toast"]');
    expect(toast?.textContent).toContain("No tables available");

    await act(async () => {
      root.unmount();
    });
  });

  it("shows a toast when attach formula is clicked but no formula is available", async () => {
    (globalThis as any).IS_REACT_ACT_ENVIRONMENT = true;
    vi.stubGlobal("crypto", { randomUUID: () => "uuid-1" } as any);

    const toastRoot = document.createElement("div");
    toastRoot.id = "toast-root";
    document.body.appendChild(toastRoot);

    const sendMessage = vi.fn(async () => {
      return { messages: [], final: "Ok" };
    });

    const getFormulaAttachment = vi.fn(() => null);

    const container = document.createElement("div");
    document.body.appendChild(container);
    const root = createRoot(container);

    await act(async () => {
      root.render(
        React.createElement(AIChatPanel, {
          systemPrompt: "system",
          sendMessage,
          getFormulaAttachment,
        }),
      );
    });

    const attachFormulaBtn = container.querySelector('[data-testid="ai-chat-attach-formula"]') as HTMLButtonElement | null;
    expect(attachFormulaBtn).toBeInstanceOf(HTMLButtonElement);

    await act(async () => {
      attachFormulaBtn!.click();
    });

    const toast = toastRoot.querySelector('[data-testid="toast"]');
    expect(toast?.textContent).toContain("No active cell formula");

    await act(async () => {
      root.unmount();
    });
  });

  it("disables attach chart with a no-charts tooltip when there are no charts", async () => {
    (globalThis as any).IS_REACT_ACT_ENVIRONMENT = true;
    vi.stubGlobal("crypto", { randomUUID: () => "uuid-1" } as any);

    const sendMessage = vi.fn(async () => {
      return { messages: [], final: "Ok" };
    });

    const getChartOptions = vi.fn(() => []);

    const container = document.createElement("div");
    document.body.appendChild(container);
    const root = createRoot(container);

    await act(async () => {
      root.render(
        React.createElement(AIChatPanel, {
          systemPrompt: "system",
          sendMessage,
          getChartOptions,
        }),
      );
    });

    const attachChartBtn = container.querySelector('[data-testid="ai-chat-attach-chart"]') as HTMLButtonElement | null;
    expect(attachChartBtn).toBeInstanceOf(HTMLButtonElement);
    expect(attachChartBtn?.disabled).toBe(true);

    const wrap = attachChartBtn?.parentElement;
    expect(wrap?.classList.contains("ai-chat-panel__attachment-button-wrap")).toBe(true);
    expect(wrap?.getAttribute("title")).toBe("No charts available");

    await act(async () => {
      root.unmount();
    });
  });

  it("disables attach chart with a no-selection tooltip when using a selected-chart provider and none is selected", async () => {
    (globalThis as any).IS_REACT_ACT_ENVIRONMENT = true;
    vi.stubGlobal("crypto", { randomUUID: () => "uuid-1" } as any);

    const sendMessage = vi.fn(async () => {
      return { messages: [], final: "Ok" };
    });

    const getChartAttachment = vi.fn(() => null);

    const container = document.createElement("div");
    document.body.appendChild(container);
    const root = createRoot(container);

    await act(async () => {
      root.render(
        React.createElement(AIChatPanel, {
          systemPrompt: "system",
          sendMessage,
          getChartAttachment,
        }),
      );
    });

    const attachChartBtn = container.querySelector('[data-testid="ai-chat-attach-chart"]') as HTMLButtonElement | null;
    expect(attachChartBtn).toBeInstanceOf(HTMLButtonElement);
    expect(attachChartBtn?.disabled).toBe(true);

    const wrap = attachChartBtn?.parentElement;
    expect(wrap?.classList.contains("ai-chat-panel__attachment-button-wrap")).toBe(true);
    expect(wrap?.getAttribute("title")).toBe("No chart selected");

    await act(async () => {
      root.unmount();
    });
  });

  it("shows a toast when attach chart is clicked but charts are unavailable", async () => {
    (globalThis as any).IS_REACT_ACT_ENVIRONMENT = true;
    vi.stubGlobal("crypto", { randomUUID: () => "uuid-1" } as any);

    const toastRoot = document.createElement("div");
    toastRoot.id = "toast-root";
    document.body.appendChild(toastRoot);

    const sendMessage = vi.fn(async () => {
      return { messages: [], final: "Ok" };
    });

    // Simulate a stale render where charts existed, but are gone by click time. Use
    // mutable state so the test remains robust even if the panel re-renders.
    let currentCharts: Array<{ id: string; label: string }> = [{ id: "chart_1", label: "Chart 1" }];
    const getChartOptions = vi.fn<NonNullable<React.ComponentProps<typeof AIChatPanel>["getChartOptions"]>>(
      () => currentCharts as any,
    );

    const container = document.createElement("div");
    document.body.appendChild(container);
    const root = createRoot(container);

    await act(async () => {
      root.render(
        React.createElement(AIChatPanel, {
          systemPrompt: "system",
          sendMessage,
          getChartOptions,
        }),
      );
    });

    const attachChartBtn = container.querySelector('[data-testid="ai-chat-attach-chart"]') as HTMLButtonElement | null;
    expect(attachChartBtn).toBeInstanceOf(HTMLButtonElement);
    expect(attachChartBtn?.disabled).toBe(false);

    // Charts disappear before click time.
    currentCharts = [];

    await act(async () => {
      attachChartBtn!.click();
    });

    const toast = toastRoot.querySelector('[data-testid="toast"]');
    expect(toast?.textContent).toContain("No charts available");

    await act(async () => {
      root.unmount();
    });
  });

  it("can attach and remove a selection before sending (and includes attachments on the user message)", async () => {
    (globalThis as any).IS_REACT_ACT_ENVIRONMENT = true;
    vi.stubGlobal("crypto", { randomUUID: () => "uuid-1" } as any);

    const sendMessage = vi.fn(async () => {
      return { messages: [], final: "Ok" };
    });

    const getSelectionAttachment = vi.fn(() => ({ type: "range" as const, reference: "Sheet1!A1:D10" }));

    const container = document.createElement("div");
    document.body.appendChild(container);
    const root = createRoot(container);

    await act(async () => {
      root.render(
        React.createElement(AIChatPanel, {
          systemPrompt: "system",
          sendMessage,
          getSelectionAttachment,
        }),
      );
    });

    // Demo-only placeholder UI should be gone.
    expect(container.textContent).not.toContain("Attachments API placeholder");
    expect(container.textContent).not.toContain("Range (demo)");

    const attachSelectionBtn = container.querySelector('[data-testid="ai-chat-attach-selection"]') as HTMLButtonElement | null;
    expect(attachSelectionBtn).toBeInstanceOf(HTMLButtonElement);

    await act(async () => {
      attachSelectionBtn!.click();
    });

    const chip = container.querySelector('[data-testid="ai-chat-attachment-chip-0"]');
    expect(chip?.textContent).toContain("range: Sheet1!A1:D10");

    const removeBtn = container.querySelector('[data-testid="ai-chat-attachment-remove-0"]') as HTMLButtonElement | null;
    expect(removeBtn).toBeInstanceOf(HTMLButtonElement);

    await act(async () => {
      removeBtn!.click();
    });

    expect(container.querySelector('[data-testid="ai-chat-attachment-chip-0"]')).toBeNull();

    // Re-add and send.
    await act(async () => {
      attachSelectionBtn!.click();
    });

    const input = container.querySelector("input") as HTMLInputElement | null;
    const sendButton = Array.from(container.querySelectorAll("button")).find((b) => b.textContent === "Send") as HTMLButtonElement | undefined;
    expect(input).toBeInstanceOf(HTMLInputElement);
    expect(sendButton).toBeInstanceOf(HTMLButtonElement);

    await act(async () => {
      setNativeInputValue(input!, "Hello");
      input!.dispatchEvent(new Event("input", { bubbles: true }));
      input!.dispatchEvent(new Event("change", { bubbles: true }));
    });

    await act(async () => {
      sendButton!.click();
      await waitFor(() => sendMessage.mock.calls.length === 1);
    });

    const callArgs = sendMessage.mock.calls[0]?.[0] as any;
    expect(callArgs.attachments).toEqual([{ type: "range", reference: "Sheet1!A1:D10" }]);

    // The user message should render its attachments (separate from the pending chips UI).
    expect(container.textContent).toContain("Attachments:");
    expect(container.textContent).toContain("range: Sheet1!A1:D10");

    // Pending attachment chips should clear after sending.
    expect(container.querySelector('[data-testid="ai-chat-pending-attachments"]')).toBeNull();

    await act(async () => {
      root.unmount();
    });
  });

  it("can attach a formula and includes it on the user message", async () => {
    (globalThis as any).IS_REACT_ACT_ENVIRONMENT = true;
    vi.stubGlobal("crypto", { randomUUID: () => "uuid-1" } as any);

    const sendMessage = vi.fn(async () => {
      return { messages: [], final: "Ok" };
    });

    const formulaAttachment = {
      type: "formula" as const,
      reference: "Sheet1!A1",
      data: { formula: "=SUM(A1:A3)" },
    };
    const getFormulaAttachment = vi.fn(() => formulaAttachment);

    const container = document.createElement("div");
    document.body.appendChild(container);
    const root = createRoot(container);

    await act(async () => {
      root.render(
        React.createElement(AIChatPanel, {
          systemPrompt: "system",
          sendMessage,
          getFormulaAttachment,
        }),
      );
    });

    const attachFormulaBtn = container.querySelector('[data-testid="ai-chat-attach-formula"]') as HTMLButtonElement | null;
    expect(attachFormulaBtn).toBeInstanceOf(HTMLButtonElement);

    await act(async () => {
      attachFormulaBtn!.click();
    });

    const chip = container.querySelector('[data-testid="ai-chat-attachment-chip-0"]');
    expect(chip?.textContent).toContain("formula: Sheet1!A1");

    const input = container.querySelector("input") as HTMLInputElement | null;
    const sendButton = Array.from(container.querySelectorAll("button")).find((b) => b.textContent === "Send") as HTMLButtonElement | undefined;
    expect(input).toBeInstanceOf(HTMLInputElement);
    expect(sendButton).toBeInstanceOf(HTMLButtonElement);

    await act(async () => {
      setNativeInputValue(input!, "Hello");
      input!.dispatchEvent(new Event("input", { bubbles: true }));
      input!.dispatchEvent(new Event("change", { bubbles: true }));
    });

    await act(async () => {
      sendButton!.click();
      await waitFor(() => sendMessage.mock.calls.length === 1);
    });

    const callArgs = sendMessage.mock.calls[0]?.[0] as any;
    expect(callArgs.attachments).toEqual([formulaAttachment]);

    await act(async () => {
      root.unmount();
    });
  });

  it("can attach a table via quick pick and includes it on the user message", async () => {
    (globalThis as any).IS_REACT_ACT_ENVIRONMENT = true;
    vi.stubGlobal("crypto", { randomUUID: () => "uuid-1" } as any);

    const sendMessage = vi.fn(async () => {
      return { messages: [], final: "Ok" };
    });

    const getTableOptions = vi.fn(() => [{ name: "SalesTable", description: "Sheet1!A1:C10" }]);

    const container = document.createElement("div");
    document.body.appendChild(container);
    const root = createRoot(container);

    await act(async () => {
      root.render(
        React.createElement(AIChatPanel, {
          systemPrompt: "system",
          sendMessage,
          getTableOptions,
        }),
      );
    });

    const attachTableBtn = container.querySelector('[data-testid="ai-chat-attach-table"]') as HTMLButtonElement | null;
    expect(attachTableBtn).toBeInstanceOf(HTMLButtonElement);

    await act(async () => {
      attachTableBtn!.click();
    });

    const quickPickItem = document.body.querySelector('[data-testid="quick-pick-item-0"]') as HTMLButtonElement | null;
    expect(quickPickItem).toBeInstanceOf(HTMLButtonElement);

    await act(async () => {
      quickPickItem!.click();
    });

    await act(async () => {
      await waitFor(() => (container.textContent?.includes("table: SalesTable") ?? false));
    });

    const input = container.querySelector("input") as HTMLInputElement | null;
    const sendButton = Array.from(container.querySelectorAll("button")).find((b) => b.textContent === "Send") as HTMLButtonElement | undefined;
    expect(input).toBeInstanceOf(HTMLInputElement);
    expect(sendButton).toBeInstanceOf(HTMLButtonElement);

    await act(async () => {
      setNativeInputValue(input!, "Hello");
      input!.dispatchEvent(new Event("input", { bubbles: true }));
      input!.dispatchEvent(new Event("change", { bubbles: true }));
    });

    await act(async () => {
      sendButton!.click();
      await waitFor(() => sendMessage.mock.calls.length === 1);
    });

    const callArgs = sendMessage.mock.calls[0]?.[0] as any;
    expect(callArgs.attachments).toEqual([{ type: "table", reference: "SalesTable" }]);

    await act(async () => {
      root.unmount();
    });
  });

  it("can attach a chart via quick pick and includes it on the user message", async () => {
    (globalThis as any).IS_REACT_ACT_ENVIRONMENT = true;
    vi.stubGlobal("crypto", { randomUUID: () => "uuid-1" } as any);

    const sendMessage = vi.fn(async () => {
      return { messages: [], final: "Ok" };
    });

    const getChartOptions = vi.fn(() => [{ id: "chart_1", label: "Revenue Chart", description: "Sheet1 • bar" }]);

    const container = document.createElement("div");
    document.body.appendChild(container);
    const root = createRoot(container);

    await act(async () => {
      root.render(
        React.createElement(AIChatPanel, {
          systemPrompt: "system",
          sendMessage,
          getChartOptions,
        }),
      );
    });

    const attachChartBtn = container.querySelector('[data-testid="ai-chat-attach-chart"]') as HTMLButtonElement | null;
    expect(attachChartBtn).toBeInstanceOf(HTMLButtonElement);

    await act(async () => {
      attachChartBtn!.click();
    });

    const quickPickItem = document.body.querySelector('[data-testid="quick-pick-item-0"]') as HTMLButtonElement | null;
    expect(quickPickItem).toBeInstanceOf(HTMLButtonElement);

    await act(async () => {
      quickPickItem!.click();
    });

    await act(async () => {
      await waitFor(() => (container.textContent?.includes("chart: chart_1") ?? false));
    });

    const input = container.querySelector("input") as HTMLInputElement | null;
    const sendButton = Array.from(container.querySelectorAll("button")).find((b) => b.textContent === "Send") as HTMLButtonElement | undefined;
    expect(input).toBeInstanceOf(HTMLInputElement);
    expect(sendButton).toBeInstanceOf(HTMLButtonElement);

    await act(async () => {
      setNativeInputValue(input!, "Hello");
      input!.dispatchEvent(new Event("input", { bubbles: true }));
      input!.dispatchEvent(new Event("change", { bubbles: true }));
    });

    await act(async () => {
      sendButton!.click();
      await waitFor(() => sendMessage.mock.calls.length === 1);
    });

    const callArgs = sendMessage.mock.calls[0]?.[0] as any;
    expect(callArgs.attachments).toEqual([{ type: "chart", reference: "chart_1" }]);

    await act(async () => {
      root.unmount();
    });
  });

  it("prefers attaching the currently selected chart when available (no picker)", async () => {
    (globalThis as any).IS_REACT_ACT_ENVIRONMENT = true;
    vi.stubGlobal("crypto", { randomUUID: () => "uuid-1" } as any);

    const sendMessage = vi.fn(async () => {
      return { messages: [], final: "Ok" };
    });

    const getChartAttachment = vi.fn(() => ({ type: "chart" as const, reference: "chart_selected" }));
    const getChartOptions = vi.fn(() => [
      { id: "chart_selected", label: "Selected chart" },
      { id: "chart_other", label: "Other chart" },
    ]);

    const container = document.createElement("div");
    document.body.appendChild(container);
    const root = createRoot(container);

    await act(async () => {
      root.render(
        React.createElement(AIChatPanel, {
          systemPrompt: "system",
          sendMessage,
          getChartAttachment,
          getChartOptions,
        }),
      );
    });

    const attachChartBtn = container.querySelector('[data-testid="ai-chat-attach-chart"]') as HTMLButtonElement | null;
    expect(attachChartBtn).toBeInstanceOf(HTMLButtonElement);

    await act(async () => {
      attachChartBtn!.click();
    });

    // Should attach immediately without opening a quick pick dialog.
    expect(container.textContent).toContain("chart: chart_selected");
    expect(document.body.querySelector('[data-testid="quick-pick"]')).toBeNull();

    const input = container.querySelector("input") as HTMLInputElement | null;
    const sendButton = Array.from(container.querySelectorAll("button")).find((b) => b.textContent === "Send") as HTMLButtonElement | undefined;
    expect(input).toBeInstanceOf(HTMLInputElement);
    expect(sendButton).toBeInstanceOf(HTMLButtonElement);

    await act(async () => {
      setNativeInputValue(input!, "Hello");
      input!.dispatchEvent(new Event("input", { bubbles: true }));
      input!.dispatchEvent(new Event("change", { bubbles: true }));
    });

    await act(async () => {
      sendButton!.click();
      await waitFor(() => sendMessage.mock.calls.length === 1);
    });

    const callArgs = sendMessage.mock.calls[0]?.[0] as any;
    expect(callArgs.attachments).toEqual([{ type: "chart", reference: "chart_selected" }]);

    await act(async () => {
      root.unmount();
    });
  });
});
