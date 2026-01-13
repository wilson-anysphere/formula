import { afterEach, describe, expect, it, vi } from "vitest";

import { CursorLLMClient } from "./cursor.js";

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

function setEnv(key: string, value: string | undefined) {
  const previous = process.env[key];
  if (value === undefined) {
    delete process.env[key];
  } else {
    process.env[key] = value;
  }

  return () => {
    if (previous === undefined) {
      delete process.env[key];
    } else {
      process.env[key] = previous;
    }
  };
}

function okResponse() {
  return new Response(JSON.stringify({ choices: [{ message: { role: "assistant", content: "ok" } }] }), {
    status: 200,
    headers: { "content-type": "application/json" },
  });
}

describe("CursorLLMClient config (baseUrl normalization)", () => {
  it.each([
    ["https://cursor.test", "https://cursor.test/v1/chat/completions"],
    ["https://cursor.test/v1", "https://cursor.test/v1/chat/completions"],
    ["https://cursor.test/v1/chat", "https://cursor.test/v1/chat/completions"],
    ["https://cursor.test/v1/chat/completions", "https://cursor.test/v1/chat/completions"],
    ["", "/v1/chat/completions"],
  ])("posts to %s => %s", async (baseUrl, expectedEndpoint) => {
    const fetchMock = vi.fn(async () => okResponse());
    vi.stubGlobal("fetch", fetchMock as any);

    const client = new CursorLLMClient({ baseUrl, timeoutMs: 1_000 });
    await client.chat({ messages: [{ role: "user", content: "hi" }] as any });

    expect(fetchMock).toHaveBeenCalledTimes(1);
    expect(fetchMock.mock.calls[0]?.[0]).toBe(expectedEndpoint);
  });

  it("defaults to same-origin /v1/chat/completions when baseUrl is omitted", async () => {
    const restoreBaseUrl = setEnv("CURSOR_AI_BASE_URL", undefined);
    try {
      const fetchMock = vi.fn(async () => okResponse());
      vi.stubGlobal("fetch", fetchMock as any);

      const client = new CursorLLMClient({ timeoutMs: 1_000 });
      await client.chat({ messages: [{ role: "user", content: "hi" }] as any });

      expect(fetchMock).toHaveBeenCalledTimes(1);
      expect(fetchMock.mock.calls[0]?.[0]).toBe("/v1/chat/completions");
    } finally {
      restoreBaseUrl();
    }
  });
});

describe("CursorLLMClient config (Node env var overrides)", () => {
  it("uses CURSOR_AI_BASE_URL when baseUrl is not provided", async () => {
    const restoreBaseUrl = setEnv("CURSOR_AI_BASE_URL", "https://cursor.env");
    try {
      const fetchMock = vi.fn(async () => okResponse());
      vi.stubGlobal("fetch", fetchMock as any);

      const client = new CursorLLMClient({ timeoutMs: 1_000 });
      await client.chat({ messages: [{ role: "user", content: "hi" }] as any });

      expect(fetchMock).toHaveBeenCalledTimes(1);
      expect(fetchMock.mock.calls[0]?.[0]).toBe("https://cursor.env/v1/chat/completions");
    } finally {
      restoreBaseUrl();
    }
  });

  it("uses CURSOR_AI_TIMEOUT_MS when timeoutMs is not provided", () => {
    const restoreTimeout = setEnv("CURSOR_AI_TIMEOUT_MS", "1234");
    try {
      const client = new CursorLLMClient({ baseUrl: "https://cursor.test" });
      expect((client as any).timeoutMs).toBe(1234);
    } finally {
      restoreTimeout();
    }
  });
});

