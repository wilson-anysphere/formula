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
});
