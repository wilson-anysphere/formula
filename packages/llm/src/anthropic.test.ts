import { afterEach, describe, expect, it, vi } from "vitest";

import { AnthropicClient } from "./anthropic.js";

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

describe("AnthropicClient", () => {
  it("formats messages + tools for the Messages API", async () => {
    const fetchMock = vi.fn(async (url: string, init: any) => {
      expect(url).toBe("https://api.anthropic.com/v1/messages");
      expect(init?.headers?.["x-api-key"]).toBe("test-key");
      expect(init?.headers?.["anthropic-version"]).toBe("2023-06-01");
      const body = JSON.parse(init.body);
      expect(body).toMatchObject({
        model: "claude-test",
        system: "sys",
        max_tokens: 42,
        temperature: 0.1,
        tool_choice: { type: "auto" },
      });
      expect(body.messages).toEqual([{ role: "user", content: "hi" }]);
      expect(body.tools).toEqual([
        {
          name: "read_range",
          description: "Read a range",
          input_schema: {
            type: "object",
            properties: { range: { type: "string" } },
            required: ["range"],
          },
        },
      ]);
      return {
        ok: true,
        json: async () => ({
          content: [{ type: "text", text: "ok" }],
          usage: { input_tokens: 1, output_tokens: 2 },
        }),
      } as any;
    });

    vi.stubGlobal("fetch", fetchMock);

    const client = new AnthropicClient({ apiKey: "test-key", model: "claude-test" });
    await client.chat({
      messages: [
        { role: "system", content: "sys" },
        { role: "user", content: "hi" },
      ],
      tools: [
        {
          name: "read_range",
          description: "Read a range",
          parameters: {
            type: "object",
            properties: { range: { type: "string" } },
            required: ["range"],
          },
        },
      ],
      toolChoice: "auto",
      temperature: 0.1,
      maxTokens: 42,
    });

    expect(fetchMock).toHaveBeenCalledTimes(1);
  });

  it("parses tool_use blocks into toolCalls", async () => {
    const fetchMock = vi.fn(async () => {
      return {
        ok: true,
        json: async () => ({
          content: [
            { type: "text", text: "Checking…" },
            { type: "tool_use", id: "toolu_1", name: "read_range", input: { range: "Sheet1!A1:A1" } },
          ],
          usage: { input_tokens: 10, output_tokens: 5 },
        }),
      } as any;
    });
    vi.stubGlobal("fetch", fetchMock);

    const client = new AnthropicClient({ apiKey: "test-key" });
    const result = await client.chat({ messages: [{ role: "user", content: "What's in A1?" }] });

    expect(result.message.role).toBe("assistant");
    expect(result.message.content).toBe("Checking…");
    expect(result.message.toolCalls).toEqual([
      { id: "toolu_1", name: "read_range", arguments: { range: "Sheet1!A1:A1" } },
    ]);
    expect(result.usage).toMatchObject({ promptTokens: 10, completionTokens: 5 });
  });

  it("converts role:tool messages into tool_result blocks", async () => {
    const fetchMock = vi.fn(async (_url: string, init: any) => {
      const body = JSON.parse(init.body);
      expect(body.system).toBeUndefined();
      expect(body.messages).toEqual([
        { role: "user", content: "hi" },
        {
          role: "assistant",
          content: [{ type: "tool_use", id: "call-1", name: "read_range", input: { range: "Sheet1!A1:A1" } }],
        },
        {
          role: "user",
          content: [{ type: "tool_result", tool_use_id: "call-1", content: '{"ok":true}' }],
        },
        { role: "user", content: "done" },
      ]);
      return {
        ok: true,
        json: async () => ({ content: [{ type: "text", text: "ok" }] }),
      } as any;
    });

    vi.stubGlobal("fetch", fetchMock);

    const client = new AnthropicClient({ apiKey: "test-key", model: "claude-test" });
    await client.chat({
      messages: [
        { role: "user", content: "hi" },
        {
          role: "assistant",
          content: "",
          toolCalls: [{ id: "call-1", name: "read_range", arguments: { range: "Sheet1!A1:A1" } }],
        },
        { role: "tool", toolCallId: "call-1", content: '{"ok":true}' },
        { role: "user", content: "done" },
      ],
      maxTokens: 42,
    });
  });

  it("propagates abort errors when the request signal is cancelled", async () => {
    vi.stubGlobal(
      "fetch",
      vi.fn(async (_url: string, init: any) => {
        const signal = init?.signal as AbortSignal | undefined;
        return new Promise((_resolve, reject) => {
          const error = new Error("Aborted");
          (error as any).name = "AbortError";
          if (!signal) return reject(error);
          if (signal.aborted) return reject(error);
          signal.addEventListener("abort", () => reject(error), { once: true });
        });
      }) as any,
    );

    const client = new AnthropicClient({ apiKey: "test-key", timeoutMs: 5_000 });
    const controller = new AbortController();
    const promise = client.chat({ messages: [{ role: "user", content: "hi" }] as any, signal: controller.signal });
    controller.abort();
    await expect(promise).rejects.toMatchObject({ name: "AbortError" });
  });
});
