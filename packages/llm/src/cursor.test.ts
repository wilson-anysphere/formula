import { afterEach, describe, expect, it, vi } from "vitest";

import { CursorLLMClient } from "./cursor.js";

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

describe("CursorLLMClient.chat", () => {
  it("posts to the configured base URL and forwards messages/tools/toolChoice", async () => {
    const fetchMock = vi.fn(async (url: string, init: any) => {
      expect(url).toBe("https://cursor.test/chat/completions");
      expect(init.method).toBe("POST");
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
