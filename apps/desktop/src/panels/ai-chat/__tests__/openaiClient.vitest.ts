import { afterEach, describe, expect, it, vi } from "vitest";

import { OpenAIClient } from "../../../../../../packages/llm/src/openai.js";
import type { LLMMessage } from "../../../../../../packages/llm/src/types.js";

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

describe("OpenAIClient message transformation", () => {
  it("serializes tool messages with tool_call_id", async () => {
    const fetchMock = vi.fn(async (_url: string, init: any) => {
      return {
        ok: true,
        json: async () => ({
          choices: [{ message: { content: "ok" } }],
        }),
      } as any;
    });

    vi.stubGlobal("fetch", fetchMock);

    const client = new OpenAIClient({
      apiKey: "test",
      baseUrl: "https://example.com",
      timeoutMs: 1_000,
      model: "gpt-test",
    });

    const messages: LLMMessage[] = [
      { role: "system", content: "system" },
      {
        role: "assistant",
        content: "",
        toolCalls: [{ id: "call-1", name: "getData", arguments: { range: "A1" } }],
      },
      { role: "tool", toolCallId: "call-1", content: "{\"value\":42}" },
    ];

    await client.chat({ messages });

    expect(fetchMock).toHaveBeenCalledTimes(1);
    const init = fetchMock.mock.calls[0]?.[1] as any;
    const body = JSON.parse(init.body as string);

    const toolMsgs = body.messages.filter((m: any) => m.role === "tool");
    expect(toolMsgs).toHaveLength(1);
    expect(toolMsgs[0].tool_call_id).toBe("call-1");
    expect(toolMsgs[0].content).toBe("{\"value\":42}");
  });
});
