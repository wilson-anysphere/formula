import { CursorLLMClient } from "./cursor.js";

/**
 * Create an LLM client backed by the Cursor inference backend.
 *
 * @returns {import("./types.js").LLMClient}
 */
export function createLLMClient() {
  if (arguments.length > 0) {
    throw new Error(
      "createLLMClient no longer accepts provider configuration. All inference is routed through the Cursor backend.",
    );
  }
  return new CursorLLMClient();
}
