import { afterEach, describe, expect, it, vi } from "vitest";

import { OllamaChatClient } from "./ollama.js";

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

describe("OllamaChatClient.chat", () => {
  it("formats messages + tools for /api/chat", async () => {
    const fetchMock = vi.fn(async (url: string, init: any) => {
      expect(url).toBe("https://example.com/api/chat");
      const body = JSON.parse(init.body);
      expect(body).toMatchObject({
        model: "llama-test",
        stream: false,
      });
      expect(body.messages).toEqual([{ role: "user", content: "hi" }]);
      expect(body.tools).toEqual([
        {
          type: "function",
          function: {
            name: "read_range",
            description: "Read a range",
            parameters: {
              type: "object",
              properties: { range: { type: "string" } },
              required: ["range"],
            },
          },
        },
      ]);

      return {
        ok: true,
        json: async () => ({ message: { role: "assistant", content: "ok" } }),
      } as any;
    });

    vi.stubGlobal("fetch", fetchMock as any);

    const client = new OllamaChatClient({ baseUrl: "https://example.com", model: "llama-test" });
    const result = await client.chat({
      messages: [{ role: "user", content: "hi" }] as any,
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
    });

    expect(result.message.content).toBe("ok");
    expect(fetchMock).toHaveBeenCalledTimes(1);
  });

  it("degrades gracefully when tool calling is unsupported (no tool_calls)", async () => {
    vi.stubGlobal(
      "fetch",
      vi.fn(async () => {
        return {
          ok: true,
          json: async () => ({ message: { role: "assistant", content: "hello" } }),
        } as any;
      }),
    );

    const client = new OllamaChatClient({ baseUrl: "https://example.com", model: "llama-test" });
    const result = await client.chat({
      messages: [{ role: "user", content: "hi" }] as any,
      tools: [
        {
          name: "read_range",
          description: "Read a range",
          parameters: { type: "object", properties: {}, required: [] },
        },
      ],
      toolChoice: "auto",
    });

    expect(result.message.content).toBe("hello");
    expect(result.message.toolCalls).toBeUndefined();
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

    const client = new OllamaChatClient({ baseUrl: "https://example.com", model: "llama-test", timeoutMs: 5_000 });
    const controller = new AbortController();
    const promise = client.chat({ messages: [{ role: "user", content: "hi" }] as any, signal: controller.signal });
    controller.abort();
    await expect(promise).rejects.toMatchObject({ name: "AbortError" });
  });
});

