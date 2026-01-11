import { afterEach, describe, expect, it, vi } from "vitest";

import { OpenAIClient } from "./openai.js";
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

describe("OpenAIClient.streamChat", () => {
  it("synthesizes missing tool call ids in non-streaming chat responses", async () => {
    vi.stubGlobal(
      "fetch",
      vi.fn(async () => {
        return new Response(
          JSON.stringify({
            choices: [
              {
                message: {
                  role: "assistant",
                  content: "",
                  tool_calls: [
                    {
                      type: "function",
                      function: { name: "getData", arguments: '{"range":"A1"}' },
                    },
                  ],
                },
              },
            ],
          }),
          { status: 200, headers: { "content-type": "application/json" } },
        );
      }) as any,
    );

    const client = new OpenAIClient({
      apiKey: "test",
      baseUrl: "https://example.com",
      timeoutMs: 1_000,
      model: "gpt-test",
    });

    const response = await client.chat({ messages: [{ role: "user", content: "hi" }] as any });
    expect(response.message.toolCalls).toEqual([{ id: "toolcall-0", name: "getData", arguments: { range: "A1" } }]);
  });

  it("emits text + tool call deltas from SSE chunks", async () => {
    const chunks = [
      // Split the first SSE frame across chunks to ensure buffering works.
      'data: {"choices":[{"delta":{"content":"Hel',
      'lo"},"finish_reason":null}]}\n\n',
      'data: {"choices":[{"delta":{"tool_calls":[{"index":0,"id":"call_1","type":"function","function":{"name":"getData","arguments":"{\\"range\\":\\""}}]},"finish_reason":null}]}\n\n',
      'data: {"choices":[{"delta":{"tool_calls":[{"index":0,"function":{"arguments":"A1\\"}"}}]},"finish_reason":null}]}\n\n',
      'data: {"choices":[{"delta":{},"finish_reason":"tool_calls"}]}\n\n',
      'data: {"choices":[{"delta":{},"finish_reason":null}],"usage":{"prompt_tokens":10,"completion_tokens":5,"total_tokens":15}}\n\n',
      "data: [DONE]\n\n",
    ];

    vi.stubGlobal(
      "fetch",
      vi.fn(async () => {
        return new Response(readableStreamFromChunks(chunks), { status: 200 });
      }) as any,
    );

    const client = new OpenAIClient({
      apiKey: "test",
      baseUrl: "https://example.com",
      timeoutMs: 1_000,
      model: "gpt-test",
    });

    const events: ChatStreamEvent[] = [];
    for await (const event of client.streamChat({ messages: [{ role: "user", content: "hi" }] as any })) {
      events.push(event);
    }

    expect(events).toEqual([
      { type: "text", delta: "Hello" },
      { type: "tool_call_start", id: "call_1", name: "getData" },
      { type: "tool_call_delta", id: "call_1", delta: '{"range":"' },
      { type: "tool_call_delta", id: "call_1", delta: 'A1"}' },
      { type: "tool_call_end", id: "call_1" },
      { type: "done", usage: { promptTokens: 10, completionTokens: 5, totalTokens: 15 } },
    ]);
  });

  it("buffers tool call deltas until the tool call id is available", async () => {
    const chunks = [
      'data: {"choices":[{"delta":{"tool_calls":[{"index":0,"type":"function","function":{"name":"getData","arguments":"{\\"range\\":\\""}}]},"finish_reason":null}]}\n\n',
      'data: {"choices":[{"delta":{"tool_calls":[{"index":0,"id":"call_1","type":"function"}]},"finish_reason":null}]}\n\n',
      'data: {"choices":[{"delta":{"tool_calls":[{"index":0,"function":{"arguments":"A1\\"}"}}]},"finish_reason":null}]}\n\n',
      'data: {"choices":[{"delta":{},"finish_reason":"tool_calls"}]}\n\n',
      "data: [DONE]\n\n",
    ];

    vi.stubGlobal(
      "fetch",
      vi.fn(async () => {
        return new Response(readableStreamFromChunks(chunks), { status: 200 });
      }) as any,
    );

    const client = new OpenAIClient({
      apiKey: "test",
      baseUrl: "https://example.com",
      timeoutMs: 1_000,
      model: "gpt-test",
    });

    const events: ChatStreamEvent[] = [];
    for await (const event of client.streamChat({ messages: [{ role: "user", content: "hi" }] as any })) {
      events.push(event);
    }

    expect(events).toEqual([
      { type: "tool_call_start", id: "call_1", name: "getData" },
      { type: "tool_call_delta", id: "call_1", delta: '{"range":"' },
      { type: "tool_call_delta", id: "call_1", delta: 'A1"}' },
      { type: "tool_call_end", id: "call_1" },
      { type: "done" },
    ]);
  });

  it("retries without stream_options when a backend rejects it", async () => {
    const chunks = ['data: {"choices":[{"delta":{"content":"Hi"},"finish_reason":null}]}\n\n', "data: [DONE]\n\n"];

    const fetchMock = vi
      .fn()
      .mockImplementationOnce(async () => new Response("Unrecognized request argument supplied: stream_options", { status: 400 }))
      .mockImplementationOnce(async (_url: string, init: any) => {
        const body = JSON.parse(init.body as string);
        expect(body.stream_options).toBeUndefined();
        return new Response(readableStreamFromChunks(chunks), { status: 200 });
      });

    vi.stubGlobal("fetch", fetchMock as any);

    const client = new OpenAIClient({
      apiKey: "test",
      baseUrl: "https://example.com",
      timeoutMs: 1_000,
      model: "gpt-test",
    });

    const events: ChatStreamEvent[] = [];
    for await (const event of client.streamChat({ messages: [{ role: "user", content: "hi" }] as any })) {
      events.push(event);
    }

    expect(fetchMock).toHaveBeenCalledTimes(2);

    const firstBody = JSON.parse((fetchMock.mock.calls[0]?.[1] as any).body as string);
    expect(firstBody.stream_options).toEqual({ include_usage: true });

    expect(events).toEqual([
      { type: "text", delta: "Hi" },
      { type: "done" },
    ]);
  });

  it("buffers tool call argument fragments until the call is started", async () => {
    const chunks = [
      // Arguments arrive before the backend provides an id/name (some proxies do this).
      'data: {"choices":[{"delta":{"tool_calls":[{"index":0,"function":{"arguments":"{\\"range\\":\\""}}]},"finish_reason":null}]}\n\n',
      'data: {"choices":[{"delta":{"tool_calls":[{"index":0,"id":"call_1","type":"function","function":{"name":"getData","arguments":"A1\\"}"}}]},"finish_reason":null}]}\n\n',
      'data: {"choices":[{"delta":{},"finish_reason":"tool_calls"}]}\n\n',
      "data: [DONE]\n\n",
    ];

    vi.stubGlobal(
      "fetch",
      vi.fn(async () => {
        return new Response(readableStreamFromChunks(chunks), { status: 200 });
      }) as any,
    );

    const client = new OpenAIClient({
      apiKey: "test",
      baseUrl: "https://example.com",
      timeoutMs: 1_000,
      model: "gpt-test",
    });

    const events: ChatStreamEvent[] = [];
    for await (const event of client.streamChat({ messages: [{ role: "user", content: "hi" }] as any })) {
      events.push(event);
    }

    expect(events).toEqual([
      { type: "tool_call_start", id: "call_1", name: "getData" },
      { type: "tool_call_delta", id: "call_1", delta: '{"range":"' },
      { type: "tool_call_delta", id: "call_1", delta: 'A1"}' },
      { type: "tool_call_end", id: "call_1" },
      { type: "done" },
    ]);
  });

  it("diffs tool call arguments when a backend repeatedly streams the full string", async () => {
    const chunks = [
      `data: ${JSON.stringify({
        choices: [
          {
            delta: {
              tool_calls: [
                {
                  index: 0,
                  id: "call_1",
                  type: "function",
                  function: { name: "getData", arguments: '{"range":"' },
                },
              ],
            },
            finish_reason: null,
          },
        ],
      })}\n\n`,
      `data: ${JSON.stringify({
        choices: [
          {
            delta: {
              tool_calls: [
                {
                  index: 0,
                  // id may be omitted on subsequent chunks
                  function: { arguments: '{"range":"A1"}' },
                },
              ],
            },
            finish_reason: null,
          },
        ],
      })}\n\n`,
      'data: {"choices":[{"delta":{},"finish_reason":"tool_calls"}]}\n\n',
      "data: [DONE]\n\n",
    ];

    vi.stubGlobal(
      "fetch",
      vi.fn(async () => {
        return new Response(readableStreamFromChunks(chunks), { status: 200 });
      }) as any,
    );

    const client = new OpenAIClient({
      apiKey: "test",
      baseUrl: "https://example.com",
      timeoutMs: 1_000,
      model: "gpt-test",
    });

    const events: ChatStreamEvent[] = [];
    for await (const event of client.streamChat({ messages: [{ role: "user", content: "hi" }] as any })) {
      events.push(event);
    }

    expect(events).toEqual([
      { type: "tool_call_start", id: "call_1", name: "getData" },
      { type: "tool_call_delta", id: "call_1", delta: '{"range":"' },
      { type: "tool_call_delta", id: "call_1", delta: 'A1"}' },
      { type: "tool_call_end", id: "call_1" },
      { type: "done" },
    ]);
  });
});
