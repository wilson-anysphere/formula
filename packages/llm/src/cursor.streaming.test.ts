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
  it("yields ChatStreamEvents from an SSE stream", async () => {
    const originalBaseUrl = process.env.CURSOR_AI_BASE_URL;
    process.env.CURSOR_AI_BASE_URL = "https://cursor.test";

    try {
      const chunks = [
        'data: {"type":"text","delta":"Hel',
        'lo"}\n\n',
        'data: {"type":"tool_call_start","id":"call-1","name":"read_range"}\n\n',
        'data: {"type":"tool_call_delta","id":"call-1","delta":"{\\"range\\":\\"A1:A1\\"}"}\n\n',
        'data: {"type":"tool_call_end","id":"call-1"}\n\n',
        'data: {"type":"done","usage":{"promptTokens":1,"completionTokens":2,"totalTokens":3}}\n\n',
      ];

      const fetchMock = vi.fn(async (_url: string, init: any) => {
        const body = JSON.parse(init.body as string);
        expect(body.messages).toEqual([{ role: "user", content: "hi" }]);

        return new Response(readableStreamFromChunks(chunks), {
          status: 200,
          headers: { "content-type": "text/event-stream" },
        });
      });

      vi.stubGlobal("fetch", fetchMock as any);

      const client = new CursorLLMClient();
      const events: ChatStreamEvent[] = [];
      for await (const event of client.streamChat({ messages: [{ role: "user", content: "hi" }] as any })) {
        events.push(event);
      }

      expect(events).toEqual([
        { type: "text", delta: "Hello" },
        { type: "tool_call_start", id: "call-1", name: "read_range" },
        { type: "tool_call_delta", id: "call-1", delta: '{"range":"A1:A1"}' },
        { type: "tool_call_end", id: "call-1" },
        { type: "done", usage: { promptTokens: 1, completionTokens: 2, totalTokens: 3 } },
      ]);
    } finally {
      if (originalBaseUrl === undefined) delete process.env.CURSOR_AI_BASE_URL;
      else process.env.CURSOR_AI_BASE_URL = originalBaseUrl;
    }
  });

  it("yields ChatStreamEvents from an NDJSON stream", async () => {
    const originalBaseUrl = process.env.CURSOR_AI_BASE_URL;
    process.env.CURSOR_AI_BASE_URL = "https://cursor.test";

    try {
      const chunks = [
        '{"type":"text","delta":"Hel',
        'lo"}\n',
        '{"type":"done","usage":{"promptTokens":1,"completionTokens":2,"totalTokens":3}}\n',
      ];

      const fetchMock = vi.fn(async () => {
        return new Response(readableStreamFromChunks(chunks), {
          status: 200,
          headers: { "content-type": "application/x-ndjson" },
        });
      });

      vi.stubGlobal("fetch", fetchMock as any);

      const client = new CursorLLMClient();
      const events: ChatStreamEvent[] = [];
      for await (const event of client.streamChat({ messages: [{ role: "user", content: "hi" }] as any })) {
        events.push(event);
      }

      expect(events).toEqual([
        { type: "text", delta: "Hello" },
        { type: "done", usage: { promptTokens: 1, completionTokens: 2, totalTokens: 3 } },
      ]);
    } finally {
      if (originalBaseUrl === undefined) delete process.env.CURSOR_AI_BASE_URL;
      else process.env.CURSOR_AI_BASE_URL = originalBaseUrl;
    }
  });

  it("falls back to chat() when the streaming body is unavailable", async () => {
    const originalBaseUrl = process.env.CURSOR_AI_BASE_URL;
    process.env.CURSOR_AI_BASE_URL = "https://cursor.test";

    try {
      const fetchMock = vi
        .fn()
        // streamChat attempt (no body => triggers fallback)
        .mockImplementationOnce(async () => new Response(null, { status: 200 }))
        // chat() fallback
        .mockImplementationOnce(async (_url: string, init: any) => {
          const body = JSON.parse(init.body as string);
          expect(body.messages).toEqual([{ role: "user", content: "hi" }]);
          return new Response(
            JSON.stringify({
              message: {
                role: "assistant",
                content: "Hello",
                toolCalls: [{ id: "call-1", name: "read_range", arguments: { range: "A1:A1" } }],
              },
              usage: { promptTokens: 1, completionTokens: 2, totalTokens: 3 },
            }),
            { status: 200, headers: { "content-type": "application/json" } },
          );
        });

      vi.stubGlobal("fetch", fetchMock as any);

      const client = new CursorLLMClient();
      const events: ChatStreamEvent[] = [];
      for await (const event of client.streamChat({ messages: [{ role: "user", content: "hi" }] as any })) {
        events.push(event);
      }

      expect(fetchMock).toHaveBeenCalledTimes(2);
      expect(events).toEqual([
        { type: "text", delta: "Hello" },
        { type: "tool_call_start", id: "call-1", name: "read_range" },
        { type: "tool_call_delta", id: "call-1", delta: '{"range":"A1:A1"}' },
        { type: "tool_call_end", id: "call-1" },
        { type: "done", usage: { promptTokens: 1, completionTokens: 2, totalTokens: 3 } },
      ]);
    } finally {
      if (originalBaseUrl === undefined) delete process.env.CURSOR_AI_BASE_URL;
      else process.env.CURSOR_AI_BASE_URL = originalBaseUrl;
    }
  });
});
