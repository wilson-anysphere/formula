import type { LLMClient } from "./types.js";
import type { CursorClientOptions } from "./cursor.js";

/**
 * Create a Cursor-backed LLM client.
 *
 * Note: provider selection is no longer supported; all AI uses Cursor backend.
 */
export function createLLMClient(config?: CursorClientOptions): LLMClient;
