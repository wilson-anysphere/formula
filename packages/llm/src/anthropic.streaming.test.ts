import { afterEach, describe, expect, it, vi } from "vitest";

import { AnthropicClient } from "./anthropic.js";
import type { ChatStreamEvent } from "./types.js";

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

function readableStreamFromChunks(chunks: string[]): ReadableStream<Uint8Array> {
  const encoder = new TextEncoder();
  return new ReadableStream<Uint8Array>({
    start(controller) {
      for (const chunk of chunks) controller.enqueue(encoder.encode(chunk));
      controller.close();
    },
  });
}

describe("AnthropicClient.streamChat", () => {
  it("emits text + tool call deltas + done usage from SSE chunks", async () => {
    const chunks = [
      'data: {"type":"message_start","message":{"id":"msg_1","usage":{"input_tokens":10,"output_tokens":0}}}\n\n',
      'data: {"type":"content_block_start","index":0,"content_block":{"type":"tool_use","id":"toolu_1","name":"getData","input":{}}}\n\n',
      'data: {"type":"content_block_delta","index":0,"delta":{"type":"input_json_delta","partial_json":"{\\"range\\":\\""}}\n\n',
      'data: {"type":"content_block_delta","index":0,"delta":{"type":"input_json_delta","partial_json":"A1\\"}"}}\n\n',
      'data: {"type":"content_block_stop","index":0}\n\n',
      'data: {"type":"content_block_delta","index":1,"delta":{"type":"text_delta","text":"Hello"}}\n\n',
      'data: {"type":"message_delta","delta":{},"usage":{"output_tokens":5}}\n\n',
      'data: {"type":"message_stop"}\n\n',
    ];

    vi.stubGlobal(
      "fetch",
      vi.fn(async () => {
        return new Response(readableStreamFromChunks(chunks), { status: 200 });
      }) as any,
    );

    const client = new AnthropicClient({
      apiKey: "test",
      baseUrl: "https://example.com",
      timeoutMs: 1_000,
      model: "claude-test",
    });

    const events: ChatStreamEvent[] = [];
    for await (const event of client.streamChat({ messages: [{ role: "user", content: "hi" }] as any })) {
      events.push(event);
    }

    expect(events).toEqual([
      { type: "tool_call_start", id: "toolu_1", name: "getData" },
      { type: "tool_call_delta", id: "toolu_1", delta: '{"range":"' },
      { type: "tool_call_delta", id: "toolu_1", delta: 'A1"}' },
      { type: "tool_call_end", id: "toolu_1" },
      { type: "text", delta: "Hello" },
      { type: "done", usage: { promptTokens: 10, completionTokens: 5, totalTokens: 15 } },
    ]);
  });

  it("diffs tool call JSON deltas when a backend repeats the full partial_json string", async () => {
    const chunks = [
      'data: {"type":"message_start","message":{"id":"msg_1","usage":{"input_tokens":10,"output_tokens":0}}}\n\n',
      'data: {"type":"content_block_start","index":0,"content_block":{"type":"tool_use","id":"toolu_1","name":"getData","input":{}}}\n\n',
      'data: {"type":"content_block_delta","index":0,"delta":{"type":"input_json_delta","partial_json":"{\\"range\\":\\""}}\n\n',
      // Full (not delta) string repeating the already-emitted prefix.
      'data: {"type":"content_block_delta","index":0,"delta":{"type":"input_json_delta","partial_json":"{\\"range\\":\\"A1\\"}"}}\n\n',
      'data: {"type":"content_block_stop","index":0}\n\n',
      'data: {"type":"message_stop"}\n\n',
    ];

    vi.stubGlobal(
      "fetch",
      vi.fn(async () => {
        return new Response(readableStreamFromChunks(chunks), { status: 200 });
      }) as any,
    );

    const client = new AnthropicClient({
      apiKey: "test",
      baseUrl: "https://example.com",
      timeoutMs: 1_000,
      model: "claude-test",
    });

    const events: ChatStreamEvent[] = [];
    for await (const event of client.streamChat({ messages: [{ role: "user", content: "hi" }] as any })) {
      events.push(event);
    }

    expect(events).toEqual([
      { type: "tool_call_start", id: "toolu_1", name: "getData" },
      { type: "tool_call_delta", id: "toolu_1", delta: '{"range":"' },
      { type: "tool_call_delta", id: "toolu_1", delta: 'A1"}' },
      { type: "tool_call_end", id: "toolu_1" },
      { type: "done", usage: { promptTokens: 10, completionTokens: 0, totalTokens: 10 } },
    ]);
  });
});
