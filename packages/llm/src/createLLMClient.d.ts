import type { LLMClient } from "./types.js";

export interface OpenAIClientConfig {
  provider: "openai";
  apiKey?: string;
  model?: string;
  baseUrl?: string;
  timeoutMs?: number;
}

export interface AnthropicClientConfig {
  provider: "anthropic";
  apiKey?: string;
  model?: string;
  baseUrl?: string;
  timeoutMs?: number;
  maxTokens?: number;
}

export interface OllamaClientConfig {
  provider: "ollama";
  baseUrl?: string;
  model?: string;
  timeoutMs?: number;
}

export type LLMClientConfig = OpenAIClientConfig | AnthropicClientConfig | OllamaClientConfig;

export function createLLMClient(config: LLMClientConfig): LLMClient;

