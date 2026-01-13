import { afterEach, describe, expect, it, vi } from "vitest";

import { CursorLLMClient } from "./cursor.js";
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

describe("CursorLLMClient.streamChat", () => {
  it("aborts the underlying request when the consumer stops iterating early", async () => {
    let capturedSignal: AbortSignal | undefined;
    let cancelCalled = false;
    const encoder = new TextEncoder();

    const stream = new ReadableStream<Uint8Array>({
      start(controller) {
        controller.enqueue(
          encoder.encode('data: {"choices":[{"delta":{"content":"Hi"},"finish_reason":null}]}\n\n'),
        );
        // Keep the stream open so early-cancel cleanup must trigger cancellation.
      },
      cancel() {
        cancelCalled = true;
      },
    });

    vi.stubGlobal(
      "fetch",
      vi.fn(async (_url: string, init: any) => {
        capturedSignal = init.signal;
        return new Response(stream, { status: 200 });
      }) as any,
    );

    const client = new CursorLLMClient({ baseUrl: "https://example.com", timeoutMs: 1_000, model: "gpt-test" });

    const events: ChatStreamEvent[] = [];
    for await (const event of client.streamChat({ messages: [{ role: "user", content: "hi" }] as any })) {
      events.push(event);
      break;
    }

    expect(events).toEqual([{ type: "text", delta: "Hi" }]);
    expect(capturedSignal?.aborted).toBe(true);
    expect(cancelCalled).toBe(true);
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

    vi.stubGlobal("fetch", vi.fn(async () => new Response(readableStreamFromChunks(chunks), { status: 200 })) as any);

    const client = new CursorLLMClient({ baseUrl: "https://example.com", timeoutMs: 1_000, model: "gpt-test" });

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

  it("ignores non-JSON heartbeat frames like `data: ping`", async () => {
    const chunks = [
      'data: {"choices":[{"delta":{"content":"Hi"},"finish_reason":null}]}\n\n',
      "data: ping\n\n",
      'data: {"choices":[{"delta":{"content":" there"},"finish_reason":null}]}\n\n',
      "data: [DONE]\n\n",
    ];

    vi.stubGlobal("fetch", vi.fn(async () => new Response(readableStreamFromChunks(chunks), { status: 200 })) as any);

    const client = new CursorLLMClient({ baseUrl: "https://example.com", timeoutMs: 1_000, model: "gpt-test" });

    const events: ChatStreamEvent[] = [];
    for await (const event of client.streamChat({ messages: [{ role: "user", content: "hi" }] as any })) {
      events.push(event);
    }

    expect(events).toEqual([
      { type: "text", delta: "Hi" },
      { type: "text", delta: " there" },
      { type: "done" },
    ]);
  });

  it("processes a trailing SSE frame when the stream ends without `\\n\\n`", async () => {
    const chunks = [
      'data: {"choices":[{"delta":{"content":"Hi"},"finish_reason":null}]}\n\n',
      // Missing the terminating blank line and no [DONE] frame.
      'data: {"choices":[{"delta":{"content":"!"},"finish_reason":null}]}',
    ];

    vi.stubGlobal("fetch", vi.fn(async () => new Response(readableStreamFromChunks(chunks), { status: 200 })) as any);

    const client = new CursorLLMClient({ baseUrl: "https://example.com", timeoutMs: 1_000, model: "gpt-test" });

    const events: ChatStreamEvent[] = [];
    for await (const event of client.streamChat({ messages: [{ role: "user", content: "hi" }] as any })) {
      events.push(event);
    }

    expect(events).toEqual([
      { type: "text", delta: "Hi" },
      { type: "text", delta: "!" },
      { type: "done" },
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

    vi.stubGlobal("fetch", vi.fn(async () => new Response(readableStreamFromChunks(chunks), { status: 200 })) as any);

    const client = new CursorLLMClient({ baseUrl: "https://example.com", timeoutMs: 1_000, model: "gpt-test" });

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
      .mockImplementationOnce(
        async () => new Response("Unrecognized request argument supplied: stream_options", { status: 400 }),
      )
      .mockImplementationOnce(async (_url: string, init: any) => {
        const body = JSON.parse(init.body as string);
        expect(body.stream_options).toBeUndefined();
        return new Response(readableStreamFromChunks(chunks), { status: 200 });
      });

    vi.stubGlobal("fetch", fetchMock as any);

    const client = new CursorLLMClient({ baseUrl: "https://example.com", timeoutMs: 1_000, model: "gpt-test" });

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

    vi.stubGlobal("fetch", vi.fn(async () => new Response(readableStreamFromChunks(chunks), { status: 200 })) as any);

    const client = new CursorLLMClient({ baseUrl: "https://example.com", timeoutMs: 1_000, model: "gpt-test" });

    const events: ChatStreamEvent[] = [];
    for await (const event of client.streamChat({ messages: [{ role: "user", content: "hi" }] as any })) {
      events.push(event);
    }

    expect(events[0]).toEqual({ type: "tool_call_start", id: "call_1", name: "getData" });
    const args = events
      .filter((e) => e.type === "tool_call_delta")
      .map((e) => e.delta)
      .join("");
    expect(args).toBe('{"range":"A1"}');
    expect(events.at(-2)).toEqual({ type: "tool_call_end", id: "call_1" });
    expect(events.at(-1)).toEqual({ type: "done" });
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

    vi.stubGlobal("fetch", vi.fn(async () => new Response(readableStreamFromChunks(chunks), { status: 200 })) as any);

    const client = new CursorLLMClient({ baseUrl: "https://example.com", timeoutMs: 1_000, model: "gpt-test" });

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

  it("synthesizes missing tool call ids in streaming responses", async () => {
    const chunks = [
      'data: {"choices":[{"delta":{"tool_calls":[{"index":0,"type":"function","function":{"name":"getData","arguments":"{\\"range\\":\\"A1\\"}"}}]},"finish_reason":null}]}\n\n',
      'data: {"choices":[{"delta":{},"finish_reason":"tool_calls"}]}\n\n',
      "data: [DONE]\n\n",
    ];

    vi.stubGlobal("fetch", vi.fn(async () => new Response(readableStreamFromChunks(chunks), { status: 200 })) as any);

    const client = new CursorLLMClient({ baseUrl: "https://example.com", timeoutMs: 1_000, model: "gpt-test" });

    const events: ChatStreamEvent[] = [];
    for await (const event of client.streamChat({ messages: [{ role: "user", content: "hi" }] as any })) {
      events.push(event);
    }

    expect(events).toEqual([
      { type: "tool_call_start", id: "toolcall-0", name: "getData" },
      { type: "tool_call_delta", id: "toolcall-0", delta: '{"range":"A1"}' },
      { type: "tool_call_end", id: "toolcall-0" },
      { type: "done" },
    ]);
  });

  it("preserves tool call order when synthesizing ids", async () => {
    const chunks = [
      // Intentionally emit index=1 before index=0.
      'data: {"choices":[{"delta":{"tool_calls":[{"index":1,"type":"function","function":{"name":"toolB","arguments":"{\\"b\\":1}"}}]},"finish_reason":null}]}\n\n',
      'data: {"choices":[{"delta":{"tool_calls":[{"index":0,"type":"function","function":{"name":"toolA","arguments":"{\\"a\\":1}"}}]},"finish_reason":null}]}\n\n',
      'data: {"choices":[{"delta":{},"finish_reason":"tool_calls"}]}\n\n',
      "data: [DONE]\n\n",
    ];

    vi.stubGlobal("fetch", vi.fn(async () => new Response(readableStreamFromChunks(chunks), { status: 200 })) as any);

    const client = new CursorLLMClient({ baseUrl: "https://example.com", timeoutMs: 1_000, model: "gpt-test" });

    const events: ChatStreamEvent[] = [];
    for await (const event of client.streamChat({ messages: [{ role: "user", content: "hi" }] as any })) {
      events.push(event);
    }

    expect(events).toEqual([
      { type: "tool_call_start", id: "toolcall-0", name: "toolA" },
      { type: "tool_call_delta", id: "toolcall-0", delta: '{"a":1}' },
      { type: "tool_call_start", id: "toolcall-1", name: "toolB" },
      { type: "tool_call_delta", id: "toolcall-1", delta: '{"b":1}' },
      { type: "tool_call_end", id: "toolcall-0" },
      { type: "tool_call_end", id: "toolcall-1" },
      { type: "done" },
    ]);
  });

  it("preserves tool call order when chunks arrive out of order but ids are present", async () => {
    const chunks = [
      'data: {"choices":[{"delta":{"tool_calls":[{"index":1,"id":"call_1","type":"function","function":{"name":"toolB","arguments":"{\\"b\\":1}"}}]},"finish_reason":null}]}\n\n',
      'data: {"choices":[{"delta":{"tool_calls":[{"index":0,"id":"call_0","type":"function","function":{"name":"toolA","arguments":"{\\"a\\":1}"}}]},"finish_reason":null}]}\n\n',
      'data: {"choices":[{"delta":{},"finish_reason":"tool_calls"}]}\n\n',
      "data: [DONE]\n\n",
    ];

    vi.stubGlobal("fetch", vi.fn(async () => new Response(readableStreamFromChunks(chunks), { status: 200 })) as any);

    const client = new CursorLLMClient({ baseUrl: "https://example.com", timeoutMs: 1_000, model: "gpt-test" });

    const events: ChatStreamEvent[] = [];
    for await (const event of client.streamChat({ messages: [{ role: "user", content: "hi" }] as any })) {
      events.push(event);
    }

    expect(events).toEqual([
      { type: "tool_call_start", id: "call_0", name: "toolA" },
      { type: "tool_call_delta", id: "call_0", delta: '{"a":1}' },
      { type: "tool_call_start", id: "call_1", name: "toolB" },
      { type: "tool_call_delta", id: "call_1", delta: '{"b":1}' },
      { type: "tool_call_end", id: "call_0" },
      { type: "tool_call_end", id: "call_1" },
      { type: "done" },
    ]);
  });

  it("falls back to chat() when the streaming body is unavailable", async () => {
    const fetchMock = vi
      .fn()
      // streamChat attempt (no body => triggers fallback)
      .mockImplementationOnce(async () => new Response(null, { status: 200 }))
      // chat() fallback
      .mockImplementationOnce(async () => {
        return new Response(
          JSON.stringify({
            choices: [
              {
                message: {
                  role: "assistant",
                  content: "Hello",
                  tool_calls: [{ id: "call_1", type: "function", function: { name: "getData", arguments: '{"range":"A1"}' } }],
                },
              },
            ],
          }),
          { status: 200, headers: { "content-type": "application/json" } },
        );
      });

    vi.stubGlobal("fetch", fetchMock as any);

    const client = new CursorLLMClient({ baseUrl: "https://example.com", timeoutMs: 1_000, model: "gpt-test" });
    const events: ChatStreamEvent[] = [];
    for await (const event of client.streamChat({ messages: [{ role: "user", content: "hi" }] as any })) {
      events.push(event);
    }

    expect(fetchMock).toHaveBeenCalledTimes(2);
    expect(events).toEqual([
      { type: "text", delta: "Hello" },
      { type: "tool_call_start", id: "call_1", name: "getData" },
      { type: "tool_call_delta", id: "call_1", delta: '{"range":"A1"}' },
      { type: "tool_call_end", id: "call_1" },
      { type: "done" },
    ]);
  });
});
