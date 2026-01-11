import { afterEach, describe, expect, it, vi } from "vitest";

import { OllamaChatClient } from "./ollama.js";
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

describe("OllamaChatClient.streamChat", () => {
  it("emits tool call deltas + done usage from NDJSON chunks", async () => {
    const chunks = [
      `${JSON.stringify({
        message: {
          role: "assistant",
          content: "",
          tool_calls: [{ id: "call-1", function: { name: "getData", arguments: '{"range":"' } }],
        },
      })}\n`,
      `${JSON.stringify({
        done: true,
        prompt_eval_count: 7,
        eval_count: 4,
        message: {
          role: "assistant",
          content: "",
          tool_calls: [{ id: "call-1", function: { name: "getData", arguments: '{"range":"A1"}' } }],
        },
      })}\n`,
    ];

    vi.stubGlobal(
      "fetch",
      vi.fn(async () => {
        return new Response(readableStreamFromChunks(chunks), { status: 200 });
      }) as any,
    );

    const client = new OllamaChatClient({
      baseUrl: "https://example.com",
      model: "llama-test",
      timeoutMs: 1_000,
    });

    const events: ChatStreamEvent[] = [];
    for await (const event of client.streamChat({ messages: [{ role: "user", content: "hi" }] as any })) {
      events.push(event);
    }

    expect(events).toEqual([
      { type: "tool_call_start", id: "call-1", name: "getData" },
      { type: "tool_call_delta", id: "call-1", delta: '{"range":"' },
      { type: "tool_call_delta", id: "call-1", delta: 'A1"}' },
      { type: "tool_call_end", id: "call-1" },
      { type: "done", usage: { promptTokens: 7, completionTokens: 4, totalTokens: 11 } },
    ]);
  });

  it("buffers tool call argument deltas until the call name is known", async () => {
    const chunks = [
      `${JSON.stringify({
        message: {
          role: "assistant",
          content: "",
          tool_calls: [{ id: "call-1", function: { arguments: '{"range":"' } }],
        },
      })}\n`,
      `${JSON.stringify({
        done: true,
        prompt_eval_count: 7,
        eval_count: 4,
        message: {
          role: "assistant",
          content: "",
          tool_calls: [{ id: "call-1", function: { name: "getData", arguments: '{"range":"A1"}' } }],
        },
      })}\n`,
    ];

    vi.stubGlobal(
      "fetch",
      vi.fn(async () => {
        return new Response(readableStreamFromChunks(chunks), { status: 200 });
      }) as any,
    );

    const client = new OllamaChatClient({
      baseUrl: "https://example.com",
      model: "llama-test",
      timeoutMs: 1_000,
    });

    const events: ChatStreamEvent[] = [];
    for await (const event of client.streamChat({ messages: [{ role: "user", content: "hi" }] as any })) {
      events.push(event);
    }

    expect(events).toEqual([
      { type: "tool_call_start", id: "call-1", name: "getData" },
      { type: "tool_call_delta", id: "call-1", delta: '{"range":"' },
      { type: "tool_call_delta", id: "call-1", delta: 'A1"}' },
      { type: "tool_call_end", id: "call-1" },
      { type: "done", usage: { promptTokens: 7, completionTokens: 4, totalTokens: 11 } },
    ]);
  });

  it("buffers tool call events until a stable id is available", async () => {
    const chunks = [
      `${JSON.stringify({
        message: {
          role: "assistant",
          content: "",
          tool_calls: [{ function: { name: "getData", arguments: '{"range":"' } }],
        },
      })}\n`,
      `${JSON.stringify({
        done: true,
        prompt_eval_count: 2,
        eval_count: 1,
        message: {
          role: "assistant",
          content: "",
          tool_calls: [{ id: "call-1", function: { name: "getData", arguments: '{"range":"A1"}' } }],
        },
      })}\n`,
    ];

    vi.stubGlobal(
      "fetch",
      vi.fn(async () => {
        return new Response(readableStreamFromChunks(chunks), { status: 200 });
      }) as any,
    );

    const client = new OllamaChatClient({
      baseUrl: "https://example.com",
      model: "llama-test",
      timeoutMs: 1_000,
    });

    const events: ChatStreamEvent[] = [];
    for await (const event of client.streamChat({ messages: [{ role: "user", content: "hi" }] as any })) {
      events.push(event);
    }

    expect(events).toEqual([
      { type: "tool_call_start", id: "call-1", name: "getData" },
      { type: "tool_call_delta", id: "call-1", delta: '{"range":"' },
      { type: "tool_call_delta", id: "call-1", delta: 'A1"}' },
      { type: "tool_call_end", id: "call-1" },
      { type: "done", usage: { promptTokens: 2, completionTokens: 1, totalTokens: 3 } },
    ]);
  });

  it("falls back to chat() when response has no stream reader and preserves usage", async () => {
    let callCount = 0;
    vi.stubGlobal(
      "fetch",
      vi.fn(async () => {
        callCount += 1;
        if (callCount === 1) {
          return new Response(null, { status: 200 });
        }

        return new Response(
          JSON.stringify({
            message: {
              role: "assistant",
              content: "Hello",
              tool_calls: [{ id: "call-1", function: { name: "getData", arguments: '{"range":"A1"}' } }],
            },
            prompt_eval_count: 2,
            eval_count: 3,
          }),
          { status: 200, headers: { "content-type": "application/json" } },
        );
      }) as any,
    );

    const client = new OllamaChatClient({
      baseUrl: "https://example.com",
      model: "llama-test",
      timeoutMs: 1_000,
    });

    const events: ChatStreamEvent[] = [];
    for await (const event of client.streamChat({ messages: [{ role: "user", content: "hi" }] as any })) {
      events.push(event);
    }

    expect(callCount).toBe(2);
    expect(events).toEqual([
      { type: "text", delta: "Hello" },
      { type: "tool_call_start", id: "call-1", name: "getData" },
      { type: "tool_call_delta", id: "call-1", delta: '{"range":"A1"}' },
      { type: "tool_call_end", id: "call-1" },
      { type: "done", usage: { promptTokens: 2, completionTokens: 3, totalTokens: 5 } },
    ]);
  });
});
