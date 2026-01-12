import { afterEach, describe, expect, it, vi } from "vitest";

import { CursorLLMClient } from "./cursor.js";

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

describe("CursorLLMClient.chat (OpenAI-compatible tool calling)", () => {
  it("serializes `role: tool` messages with `tool_call_id`", async () => {
    const fetchMock = vi.fn(async (_url: string, init: any) => {
      const body = JSON.parse(init.body as string);
      expect(body.messages).toEqual([
        { role: "user", content: "hi" },
        { role: "tool", tool_call_id: "call_123", content: "tool result" },
      ]);

      return new Response(JSON.stringify({ choices: [{ message: { role: "assistant", content: "ok" } }] }), {
        status: 200,
        headers: { "content-type": "application/json" },
      });
    });

    vi.stubGlobal("fetch", fetchMock as any);

    const client = new CursorLLMClient({ baseUrl: "https://example.com", model: "gpt-test", timeoutMs: 1_000 });
    await client.chat({
      messages: [
        { role: "user", content: "hi" },
        { role: "tool", toolCallId: "call_123", content: "tool result" },
      ] as any,
    });
  });

  it("serializes assistant `toolCalls` as OpenAI `tool_calls`", async () => {
    const fetchMock = vi.fn(async (_url: string, init: any) => {
      const body = JSON.parse(init.body as string);
      expect(body.messages).toEqual([
        {
          role: "assistant",
          content: "calling tool",
          tool_calls: [
            {
              id: "call_1",
              type: "function",
              function: {
                name: "getData",
                arguments: '{"range":"A1"}',
              },
            },
          ],
        },
      ]);

      return new Response(JSON.stringify({ choices: [{ message: { role: "assistant", content: "ok" } }] }), {
        status: 200,
        headers: { "content-type": "application/json" },
      });
    });

    vi.stubGlobal("fetch", fetchMock as any);

    const client = new CursorLLMClient({ baseUrl: "https://example.com", model: "gpt-test", timeoutMs: 1_000 });
    await client.chat({
      messages: [
        {
          role: "assistant",
          content: "calling tool",
          toolCalls: [{ id: "call_1", name: "getData", arguments: { range: "A1" } }],
        },
      ] as any,
    });
  });

  it("parses `choices[0].message.tool_calls` into internal `toolCalls` (JSON parsing args)", async () => {
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
                  tool_calls: [{ id: "call_1", type: "function", function: { name: "getData", arguments: '{"a":1}' } }],
                },
              },
            ],
          }),
          { status: 200, headers: { "content-type": "application/json" } },
        );
      }) as any,
    );

    const client = new CursorLLMClient({ baseUrl: "https://example.com", model: "gpt-test", timeoutMs: 1_000 });
    const response = await client.chat({ messages: [{ role: "user", content: "hi" }] as any });
    expect(response.message.toolCalls).toEqual([{ id: "call_1", name: "getData", arguments: { a: 1 } }]);
  });

  it("synthesizes missing tool call ids (`toolcall-0`, ...)", async () => {
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
                    { type: "function", function: { name: "getData", arguments: '{"range":"A1"}' } },
                    { type: "function", function: { name: "getOther", arguments: '{"x":2}' } },
                  ],
                },
              },
            ],
          }),
          { status: 200, headers: { "content-type": "application/json" } },
        );
      }) as any,
    );

    const client = new CursorLLMClient({ baseUrl: "https://example.com", model: "gpt-test", timeoutMs: 1_000 });
    const response = await client.chat({ messages: [{ role: "user", content: "hi" }] as any });
    expect(response.message.toolCalls).toEqual([
      { id: "toolcall-0", name: "getData", arguments: { range: "A1" } },
      { id: "toolcall-1", name: "getOther", arguments: { x: 2 } },
    ]);
  });
});
