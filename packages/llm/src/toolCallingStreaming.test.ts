import { describe, expect, it, vi } from "vitest";

import { runChatWithTools } from "./toolCalling.js";
import { runChatWithToolsStreaming } from "./toolCallingStreaming.js";

import type { ChatStreamEvent, ToolExecutor } from "./types.js";

describe("runChatWithToolsStreaming", () => {
  it("reconstructs streamed tool calls + emits text deltas", async () => {
    const toolExecutor: ToolExecutor = {
      tools: [
        {
          name: "read_range",
          description: "read range",
          parameters: { type: "object", properties: { range: { type: "string" } }, required: ["range"] },
        },
      ],
      execute: vi.fn(async (call: any) => {
        expect(call.name).toBe("read_range");
        expect(call.arguments).toEqual({ range: "Sheet1!A1:A1" });
        return { ok: true, data: { values: [[42]] } };
      }),
    };

    let streamCalls = 0;
    const client = {
      async chat() {
        throw new Error("chat() should not be called when streamChat is available");
      },
      async *streamChat() {
        streamCalls += 1;
        if (streamCalls === 1) {
          yield { type: "tool_call_start", id: "call-1", name: "read_range" } satisfies ChatStreamEvent;
          yield { type: "tool_call_delta", id: "call-1", delta: '{"range":"' } satisfies ChatStreamEvent;
          yield { type: "tool_call_delta", id: "call-1", delta: 'Sheet1!A1:A1"}' } satisfies ChatStreamEvent;
          yield { type: "tool_call_end", id: "call-1" } satisfies ChatStreamEvent;
          yield { type: "done" } satisfies ChatStreamEvent;
          return;
        }

        yield { type: "text", delta: "A1 is " } satisfies ChatStreamEvent;
        yield { type: "text", delta: "42." } satisfies ChatStreamEvent;
        yield { type: "done" } satisfies ChatStreamEvent;
      },
    };

    const receivedEvents: ChatStreamEvent[] = [];

    const streamingResult = await runChatWithToolsStreaming({
      client: client as any,
      toolExecutor,
      messages: [{ role: "user", content: "What's in A1?" }],
      onStreamEvent: (event) => receivedEvents.push(event),
    });

    expect(streamingResult.final).toBe("A1 is 42.");
    expect(toolExecutor.execute).toHaveBeenCalledTimes(1);
    expect(receivedEvents.filter((e) => e.type === "text").map((e) => (e as any).delta).join("")).toBe("A1 is 42.");

    // Compare to the non-streaming loop (should reconstruct the same assistant messages).
    let chatCalls = 0;
    const nonStreamingClient = {
      async chat() {
        chatCalls += 1;
        if (chatCalls === 1) {
          return {
            message: {
              role: "assistant",
              content: "",
              toolCalls: [{ id: "call-1", name: "read_range", arguments: { range: "Sheet1!A1:A1" } }],
            },
          };
        }
        return { message: { role: "assistant", content: "A1 is 42." } };
      },
    };

    const nonStreamingToolExecutor: ToolExecutor = {
      ...toolExecutor,
      execute: vi.fn(async (call: any) => {
        expect(call.arguments).toEqual({ range: "Sheet1!A1:A1" });
        return { ok: true, data: { values: [[42]] } };
      }),
    };

    const nonStreamingResult = await runChatWithTools({
      client: nonStreamingClient as any,
      toolExecutor: nonStreamingToolExecutor,
      messages: [{ role: "user", content: "What's in A1?" }],
    });

    expect(nonStreamingResult.final).toBe(streamingResult.final);
    expect(nonStreamingResult.messages).toEqual(streamingResult.messages);
  });

  it("trims streamed tool call names before executing tools", async () => {
    const toolExecutor: ToolExecutor = {
      tools: [
        {
          name: "read_range",
          description: "read range",
          parameters: { type: "object", properties: { range: { type: "string" } }, required: ["range"] },
        },
      ],
      execute: vi.fn(async (call: any) => {
        expect(call.name).toBe("read_range");
        expect(call.arguments).toEqual({ range: "Sheet1!A1:A1" });
        return { ok: true, data: { values: [[42]] } };
      }),
    };

    let streamCalls = 0;
    const client = {
      async chat() {
        throw new Error("chat() should not be called when streamChat is available");
      },
      async *streamChat() {
        streamCalls += 1;
        if (streamCalls === 1) {
          yield { type: "tool_call_start", id: "call-1", name: "  read_range  " } satisfies ChatStreamEvent;
          yield { type: "tool_call_delta", id: "call-1", delta: '{"range":"Sheet1!A1:A1"}' } satisfies ChatStreamEvent;
          yield { type: "done" } satisfies ChatStreamEvent;
          return;
        }

        yield { type: "text", delta: "done" } satisfies ChatStreamEvent;
        yield { type: "done" } satisfies ChatStreamEvent;
      },
    };

    const result = await runChatWithToolsStreaming({
      client: client as any,
      toolExecutor,
      messages: [{ role: "user", content: "What's in A1?" }],
    });

    expect(result.final).toBe("done");
    expect(toolExecutor.execute).toHaveBeenCalledTimes(1);
  });

  it("summarizes large tool results before appending them to the next streamed request", async () => {
    const bigValues = Array.from({ length: 100 }, (_, r) => Array.from({ length: 100 }, (_, c) => r * 100 + c));
    const maxToolResultChars = 1_000;

    const toolExecutor: ToolExecutor = {
      tools: [
        {
          name: "read_range",
          description: "read range",
          parameters: { type: "object", properties: { range: { type: "string" } }, required: ["range"] },
        },
      ],
      execute: vi.fn(async (call: any) => {
        return {
          tool: "read_range",
          ok: true,
          timing: { started_at_ms: 0, duration_ms: 0 },
          data: { range: call.arguments.range, values: bigValues },
        };
      }),
    };

    let streamCalls = 0;
    const client = {
      async chat() {
        throw new Error("chat() should not be called when streamChat is available");
      },
      async *streamChat(request: any) {
        streamCalls += 1;
        if (streamCalls === 1) {
          yield { type: "tool_call_start", id: "call-1", name: "read_range" } satisfies ChatStreamEvent;
          yield { type: "tool_call_delta", id: "call-1", delta: '{"range":"Sheet1!A1:CV100"}' } satisfies ChatStreamEvent;
          yield { type: "done" } satisfies ChatStreamEvent;
          return;
        }

        const last = request.messages.at(-1);
        expect(last.role).toBe("tool");
        expect(last.toolCallId).toBe("call-1");
        expect(typeof last.content).toBe("string");
        expect(last.content.length).toBeLessThanOrEqual(maxToolResultChars);

        const payload = JSON.parse(last.content);
        expect(payload.tool).toBe("read_range");
        expect(payload.ok).toBe(true);
        expect(payload.data?.truncated).toBe(true);
        expect(payload.data?.shape).toEqual({ rows: 100, cols: 100 });

        yield { type: "text", delta: "done" } satisfies ChatStreamEvent;
        yield { type: "done" } satisfies ChatStreamEvent;
      },
    };

    const result = await runChatWithToolsStreaming({
      client: client as any,
      toolExecutor,
      messages: [{ role: "user", content: "Read a big range" }],
      maxToolResultChars,
    });

    expect(result.final).toBe("done");
    expect(toolExecutor.execute).toHaveBeenCalledTimes(1);
  });

  it("surfaces unknown tool calls as tool results and continues the loop (streaming)", async () => {
    const toolExecutor: ToolExecutor = {
      tools: [
        {
          name: "read_range",
          description: "read",
          parameters: { type: "object", properties: {} },
        },
      ],
      execute: vi.fn(async () => {
        throw new Error("should not execute unknown tool");
      }),
    };

    let streamCalls = 0;
    const client = {
      async chat() {
        throw new Error("chat() should not be called when streamChat is available");
      },
      async *streamChat(request: any) {
        streamCalls += 1;
        if (streamCalls === 1) {
          yield { type: "tool_call_start", id: "call-1", name: "nonexistent_tool" } satisfies ChatStreamEvent;
          yield { type: "tool_call_delta", id: "call-1", delta: "{}" } satisfies ChatStreamEvent;
          yield { type: "done" } satisfies ChatStreamEvent;
          return;
        }

        const last = request.messages.at(-1);
        expect(last.role).toBe("tool");
        expect(last.toolCallId).toBe("call-1");
        const payload = JSON.parse(last.content);
        expect(payload.ok).toBe(false);
        expect(payload.error?.code).toBe("unknown_tool");

        yield { type: "text", delta: "Recovered." } satisfies ChatStreamEvent;
        yield { type: "done" } satisfies ChatStreamEvent;
      },
    };

    const result = await runChatWithToolsStreaming({
      client: client as any,
      toolExecutor,
      messages: [{ role: "user", content: "call an unknown tool" }],
    });

    expect(result.final).toBe("Recovered.");
    expect(streamCalls).toBe(2);
    expect(toolExecutor.execute).not.toHaveBeenCalled();
  });

  it("wraps tool execution errors when continueOnToolError=true (streaming)", async () => {
    const toolExecutor: ToolExecutor = {
      tools: [
        {
          name: "read_range",
          description: "read",
          parameters: { type: "object", properties: {} },
        },
      ],
      execute: vi.fn(async () => {
        throw new Error("boom");
      }),
    };

    let streamCalls = 0;
    const client = {
      async chat() {
        throw new Error("chat() should not be called when streamChat is available");
      },
      async *streamChat(request: any) {
        streamCalls += 1;
        if (streamCalls === 1) {
          yield { type: "tool_call_start", id: "call-1", name: "read_range" } satisfies ChatStreamEvent;
          yield { type: "tool_call_delta", id: "call-1", delta: '{"range":"Sheet1!A1:A1"}' } satisfies ChatStreamEvent;
          yield { type: "done" } satisfies ChatStreamEvent;
          return;
        }

        const last = request.messages.at(-1);
        expect(last.role).toBe("tool");
        expect(last.toolCallId).toBe("call-1");
        const payload = JSON.parse(last.content);
        expect(payload.ok).toBe(false);
        expect(payload.error?.code).toBe("tool_execution_error");
        expect(String(payload.error?.message ?? "")).toMatch(/boom/);

        yield { type: "text", delta: "Recovered." } satisfies ChatStreamEvent;
        yield { type: "done" } satisfies ChatStreamEvent;
      },
    };

    const result = await runChatWithToolsStreaming({
      client: client as any,
      toolExecutor,
      messages: [{ role: "user", content: "trigger tool error" }],
      continueOnToolError: true,
    });

    expect(result.final).toBe("Recovered.");
    expect(streamCalls).toBe(2);
    expect(toolExecutor.execute).toHaveBeenCalledTimes(1);
  });

  it("aborts while waiting for approval (does not surface as approval denied)", async () => {
    const abortController = new AbortController();

    const toolExecutor: ToolExecutor = {
      tools: [
        {
          name: "write_cell",
          description: "write",
          parameters: { type: "object", properties: {} },
          requiresApproval: true,
        },
      ],
      execute: vi.fn(async () => ({ ok: true })),
    };

    const requireApproval = vi.fn(async () => {
      queueMicrotask(() => abortController.abort());
      return new Promise<boolean>(() => {});
    });

    const client = {
      async chat() {
        throw new Error("chat() should not be called when streamChat is available");
      },
      async *streamChat() {
        yield { type: "tool_call_start", id: "call-1", name: "write_cell" } satisfies ChatStreamEvent;
        yield { type: "tool_call_delta", id: "call-1", delta: '{"cell":"A1","value":1}' } satisfies ChatStreamEvent;
        yield { type: "done" } satisfies ChatStreamEvent;
      },
    };

    await expect(
      runChatWithToolsStreaming({
        client: client as any,
        toolExecutor,
        messages: [{ role: "user", content: "Set A1 to 1" }],
        requireApproval,
        signal: abortController.signal,
      }),
    ).rejects.toMatchObject({ name: "AbortError" });

    expect(requireApproval).toHaveBeenCalledTimes(1);
    expect(toolExecutor.execute).not.toHaveBeenCalled();
  });
});
