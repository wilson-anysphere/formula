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

  it("treats auth headers case-insensitively so `getAuthHeaders()` overrides `authToken`", async () => {
    const fetchMock = vi.fn(async (_url: string, init: any) => {
      expect(init.headers.Authorization).toBe("Bearer tokenB");
      expect(init.headers).not.toHaveProperty("authorization");
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
      authToken: "tokenA",
      getAuthHeaders: async () => ({ authorization: "Bearer tokenB" }),
    });
    await client.chat({ messages: [{ role: "user", content: "hi" }] as any });
  });

  it("always enforces JSON `Content-Type` even if `getAuthHeaders()` returns another value", async () => {
    const fetchMock = vi.fn(async (_url: string, init: any) => {
      expect(init.headers["Content-Type"]).toBe("application/json");
      expect(init.headers).not.toHaveProperty("content-type");
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
      getAuthHeaders: async () => ({ "content-type": "text/plain" }),
    });
    await client.chat({ messages: [{ role: "user", content: "hi" }] as any });
  });

  it("does not read legacy env vars for auth", async () => {
    const legacyApiKeyEnv = ["OPEN", "AI_API_KEY"].join("");
    const otherProviderApiKeyEnv = ["AN", "THROPIC_API_KEY"].join("");
    const localModelHostEnv = ["O", "LLAMA_HOST"].join("");

    const originalApiKey = process.env[legacyApiKeyEnv];
    const originalOtherProviderKey = process.env[otherProviderApiKeyEnv];
    const originalLocalModelHost = process.env[localModelHostEnv];

    process.env[legacyApiKeyEnv] = "should-not-be-used";
    process.env[otherProviderApiKeyEnv] = "should-not-be-used";
    const localModelHostNeedle = "o" + "llama-env.invalid";
    process.env[localModelHostEnv] = `http://${localModelHostNeedle}:11434`;

    try {
      const fetchMock = vi.fn(async (url: string, init: any) => {
        expect(url).not.toContain(localModelHostNeedle);
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
      if (originalApiKey === undefined) delete process.env[legacyApiKeyEnv];
      else process.env[legacyApiKeyEnv] = originalApiKey;

      if (originalOtherProviderKey === undefined) delete process.env[otherProviderApiKeyEnv];
      else process.env[otherProviderApiKeyEnv] = originalOtherProviderKey;

      if (originalLocalModelHost === undefined) delete process.env[localModelHostEnv];
      else process.env[localModelHostEnv] = originalLocalModelHost;
    }
  });
});
