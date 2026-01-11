import { describe, expect, it } from "vitest";

import { createLLMClient } from "./createLLMClient.js";
import { OpenAIClient } from "./openai.js";
import { AnthropicClient } from "./anthropic.js";
import { OllamaChatClient } from "./ollama.js";

describe("createLLMClient", () => {
  it("creates an OpenAI client", () => {
    const client = createLLMClient({ provider: "openai", apiKey: "test-key" });
    expect(client).toBeInstanceOf(OpenAIClient);
  });

  it("creates an Anthropic client", () => {
    const client = createLLMClient({ provider: "anthropic", apiKey: "test-key" });
    expect(client).toBeInstanceOf(AnthropicClient);
  });

  it("creates an Ollama client", () => {
    const client = createLLMClient({ provider: "ollama", baseUrl: "http://127.0.0.1:11434", model: "llama3.1" });
    expect(client).toBeInstanceOf(OllamaChatClient);
  });

  it("honors env fallbacks for OpenAI API key", () => {
    const originalKey = process.env.OPENAI_API_KEY;
    try {
      process.env.OPENAI_API_KEY = "env-test";
      const client = createLLMClient({ provider: "openai" });
      expect(client).toBeInstanceOf(OpenAIClient);
      expect((client as any).apiKey).toBe("env-test");
    } finally {
      if (originalKey === undefined) delete process.env.OPENAI_API_KEY;
      else process.env.OPENAI_API_KEY = originalKey;
    }
  });

  it("honors env fallbacks for Ollama host", () => {
    const originalHost = process.env.OLLAMA_HOST;
    try {
      process.env.OLLAMA_HOST = "http://example.com:11434/";
      const client = createLLMClient({ provider: "ollama" });
      expect(client).toBeInstanceOf(OllamaChatClient);
      expect((client as any).baseUrl).toBe("http://example.com:11434");
    } finally {
      if (originalHost === undefined) delete process.env.OLLAMA_HOST;
      else process.env.OLLAMA_HOST = originalHost;
    }
  });
});

