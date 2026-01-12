import { afterEach, describe, expect, it, vi } from "vitest";

import { CursorLLMClient } from "./cursor.js";

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

describe("CursorLLMClient auth integration", () => {
  it("merges `getAuthHeaders()` into the fetch request headers", async () => {
    const fetchMock = vi.fn(async (_url: string, init: any) => {
      expect(init.headers).toMatchObject({
        "Content-Type": "application/json",
        "x-cursor-session": "session-123",
      });
      return new Response(JSON.stringify({ choices: [{ message: { role: "assistant", content: "ok" } }] }), {
        status: 200,
        headers: { "content-type": "application/json" },
      });
    });

    vi.stubGlobal("fetch", fetchMock as any);

    const client = new CursorLLMClient({
      baseUrl: "https://example.com",
      model: "gpt-test",
      timeoutMs: 1_000,
      getAuthHeaders: async () => ({ "x-cursor-session": "session-123" }),
    });
    await client.chat({ messages: [{ role: "user", content: "hi" }] as any });
  });

  it("adds `Authorization: Bearer <token>` when `authToken` is provided", async () => {
    const fetchMock = vi.fn(async (_url: string, init: any) => {
      expect(init.headers).toMatchObject({
        Authorization: "Bearer test-token",
      });
      return new Response(JSON.stringify({ choices: [{ message: { role: "assistant", content: "ok" } }] }), {
        status: 200,
        headers: { "content-type": "application/json" },
      });
    });

    vi.stubGlobal("fetch", fetchMock as any);

    const client = new CursorLLMClient({
      baseUrl: "https://example.com",
      model: "gpt-test",
      timeoutMs: 1_000,
      authToken: "test-token",
    });
    await client.chat({ messages: [{ role: "user", content: "hi" }] as any });
  });

  it("does not read provider env vars (OPENAI_API_KEY / ANTHROPIC_API_KEY / OLLAMA_HOST)", async () => {
    const originals = {
      OPENAI_API_KEY: process.env.OPENAI_API_KEY,
      ANTHROPIC_API_KEY: process.env.ANTHROPIC_API_KEY,
      OLLAMA_HOST: process.env.OLLAMA_HOST,
    };

    process.env.OPENAI_API_KEY = "should-not-be-used";
    process.env.ANTHROPIC_API_KEY = "should-not-be-used";
    process.env.OLLAMA_HOST = "http://ollama-env.invalid:11434";

    try {
      const fetchMock = vi.fn(async (url: string, init: any) => {
        expect(url).not.toContain("ollama-env.invalid");
        expect(init.headers?.Authorization).toBeUndefined();
        return new Response(JSON.stringify({ choices: [{ message: { role: "assistant", content: "ok" } }] }), {
          status: 200,
          headers: { "content-type": "application/json" },
        });
      });

      vi.stubGlobal("fetch", fetchMock as any);

      const client = new CursorLLMClient({ model: "gpt-test", timeoutMs: 1_000 });
      await client.chat({ messages: [{ role: "user", content: "hi" }] as any });
    } finally {
      if (originals.OPENAI_API_KEY === undefined) delete process.env.OPENAI_API_KEY;
      else process.env.OPENAI_API_KEY = originals.OPENAI_API_KEY;

      if (originals.ANTHROPIC_API_KEY === undefined) delete process.env.ANTHROPIC_API_KEY;
      else process.env.ANTHROPIC_API_KEY = originals.ANTHROPIC_API_KEY;

      if (originals.OLLAMA_HOST === undefined) delete process.env.OLLAMA_HOST;
      else process.env.OLLAMA_HOST = originals.OLLAMA_HOST;
    }
  });
});

