import { afterEach, describe, expect, it, vi } from "vitest";

import { CursorLLMClient } from "./cursor.js";

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

describe("CursorLLMClient.chat (chat completions tool calling)", () => {
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

  it("serializes assistant `toolCalls` as `tool_calls`", async () => {
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

  it("trims tool call names parsed from chat completions responses", async () => {
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
                    { id: "call_1", type: "function", function: { name: "  getData  ", arguments: '{"a":1}' } },
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

  it("posts to /chat/completions and forwards messages/tools/toolChoice + parses tool calls", async () => {
    const fetchMock = vi.fn(async (url: string, init: any) => {
      expect(url).toBe("https://cursor.test/v1/chat/completions");
      expect(init.method).toBe("POST");
      expect(init.credentials).toBe("include");
      expect(init.headers).toEqual({
        Authorization: "Bearer test-token",
        "Content-Type": "application/json",
      });

      const body = JSON.parse(init.body as string);
      expect(body).toEqual({
        model: "cursor-test-model",
        messages: [
          { role: "user", content: "hi" },
          { role: "tool", tool_call_id: "call-1", content: '{"ok":true}' },
        ],
        tools: [
          {
            type: "function",
            function: { name: "read_range", description: "read", parameters: { type: "object", properties: {} } },
          },
        ],
        tool_choice: "auto",
        temperature: 0.25,
        max_tokens: 123,
        stream: false,
      });

      return new Response(
        JSON.stringify({
          choices: [
            {
              message: {
                role: "assistant",
                content: "ok",
                tool_calls: [
                  {
                    id: "call-1",
                    type: "function",
                    function: { name: "read_range", arguments: '{"range":"A1:A1"}' },
                  },
                ],
              },
            },
          ],
          usage: { prompt_tokens: 1, completion_tokens: 2, total_tokens: 3 },
        }),
        { status: 200, headers: { "content-type": "application/json" } },
      );
    });

    vi.stubGlobal("fetch", fetchMock as any);

    const client = new CursorLLMClient({ baseUrl: "https://cursor.test", authToken: "test-token" });
    const response = await client.chat({
      messages: [
        { role: "user", content: "hi" },
        { role: "tool", toolCallId: "call-1", content: '{"ok":true}' },
      ] as any,
      tools: [{ name: "read_range", description: "read", parameters: { type: "object", properties: {} } }] as any,
      toolChoice: "auto",
      model: "cursor-test-model",
      temperature: 0.25,
      maxTokens: 123,
    });

    expect(response.message.role).toBe("assistant");
    expect(response.message.toolCalls).toEqual([{ id: "call-1", name: "read_range", arguments: { range: "A1:A1" } }]);
    expect(response.usage).toEqual({ promptTokens: 1, completionTokens: 2, totalTokens: 3 });
  });
});
