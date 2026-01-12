import { createLLMClient } from "../../../../../packages/llm/src/createLLMClient.js";
import type { LLMClient } from "../../../../../packages/llm/src/types.js";

let cachedClient: LLMClient | null = null;

/**
 * Desktop LLM client for Formula.
 *
 * All inference is routed through the Cursor desktop backend (no provider selection
 * or API keys).
 */
export function getDesktopLLMClient(): LLMClient {
  if (!cachedClient) cachedClient = createLLMClient();
  return cachedClient;
}

/**
 * Default model identifier used for prompt budgeting and as the `model` field on
 * LLM requests.
 *
 * Cursor's backend may choose a different underlying model, but downstream code
 * still needs a stable identifier for context-window heuristics.
 */
export function getDesktopModel(): string {
  return "gpt-4o-mini";
}

