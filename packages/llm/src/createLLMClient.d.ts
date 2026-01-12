import type { LLMClient } from "./types.js";
import type { CursorClientOptions } from "./cursor.js";

/**
 * Create a Cursor-backed LLM client.
 *
 * Note: provider selection (OpenAI / Anthropic / Ollama) is no longer supported.
 */
export function createLLMClient(config?: CursorClientOptions): LLMClient;
