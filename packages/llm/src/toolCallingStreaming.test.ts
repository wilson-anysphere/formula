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
});

