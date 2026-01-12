import { afterEach, describe, expect, it, vi } from "vitest";

import { CursorLLMClient } from "./cursor.js";

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

describe("CursorLLMClient.chat", () => {
  it("posts to the configured base URL and forwards messages/tools/toolChoice", async () => {
    const originalBaseUrl = process.env.CURSOR_AI_BASE_URL;
    process.env.CURSOR_AI_BASE_URL = "https://cursor.test";

    try {
      const fetchMock = vi.fn(async (url: string, init: any) => {
        expect(url).toBe("https://cursor.test/v1/chat");
        expect(init.method).toBe("POST");
        expect(init.credentials).toBe("include");

        const body = JSON.parse(init.body as string);
        expect(body).toEqual({
          messages: [
            { role: "user", content: "hi" },
            { role: "tool", toolCallId: "call-1", content: '{"ok":true}' },
          ],
          tools: [{ name: "read_range", description: "read", parameters: { type: "object", properties: {} } }],
          toolChoice: "auto",
          model: "cursor-test-model",
          temperature: 0.25,
          maxTokens: 123,
        });

        return new Response(
          JSON.stringify({
            message: {
              role: "assistant",
              content: "ok",
              toolCalls: [{ id: "call-1", name: "read_range", arguments: { range: "A1:A1" } }],
            },
            usage: { promptTokens: 1, completionTokens: 2, totalTokens: 3 },
          }),
          { status: 200, headers: { "content-type": "application/json" } },
        );
      });

      vi.stubGlobal("fetch", fetchMock as any);

      const client = new CursorLLMClient();
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
    } finally {
      if (originalBaseUrl === undefined) delete process.env.CURSOR_AI_BASE_URL;
      else process.env.CURSOR_AI_BASE_URL = originalBaseUrl;
    }
  });
});
