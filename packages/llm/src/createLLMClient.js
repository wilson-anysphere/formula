import { OpenAIClient } from "./openai.js";
import { AnthropicClient } from "./anthropic.js";
import { OllamaChatClient } from "./ollama.js";

/**
 * @typedef {{
 *   provider: "openai",
 *   apiKey?: string,
 *   model?: string,
 *   baseUrl?: string,
 *   timeoutMs?: number
 * }} OpenAIClientConfig
 *
 * @typedef {{
 *   provider: "anthropic",
 *   apiKey?: string,
 *   model?: string,
 *   baseUrl?: string,
 *   timeoutMs?: number,
 *   maxTokens?: number
 * }} AnthropicClientConfig
 *
 * @typedef {{
 *   provider: "ollama",
 *   baseUrl?: string,
 *   model?: string,
 *   timeoutMs?: number
 * }} OllamaClientConfig
 *
 * @typedef {OpenAIClientConfig | AnthropicClientConfig | OllamaClientConfig} LLMClientConfig
 */

/**
 * Create an LLM client from a provider configuration.
 *
 * This helper is intentionally dependency-free and works in both Node and
 * browser runtimes (relies on global `fetch`).
 *
 * @param {LLMClientConfig} config
 * @returns {import("./types.js").LLMClient}
 */
export function createLLMClient(config) {
  if (!config || typeof config !== "object") {
    throw new Error("createLLMClient requires a provider configuration object.");
  }

  if (config.provider === "openai") {
    return new OpenAIClient({
      apiKey: config.apiKey,
      model: config.model,
      baseUrl: config.baseUrl,
      timeoutMs: config.timeoutMs,
    });
  }

  if (config.provider === "anthropic") {
    return new AnthropicClient({
      apiKey: config.apiKey,
      model: config.model,
      baseUrl: config.baseUrl,
      timeoutMs: config.timeoutMs,
      maxTokens: config.maxTokens,
    });
  }

  if (config.provider === "ollama") {
    return new OllamaChatClient({
      baseUrl: config.baseUrl,
      model: config.model,
      timeoutMs: config.timeoutMs,
    });
  }

  // Exhaustive check for future providers.
  // eslint-disable-next-line @typescript-eslint/restrict-template-expressions
  throw new Error(`Unknown LLM provider: ${(config).provider}`);
}

