import { CursorLLMClient } from "./cursor.js";

/**
 * @typedef {import("./cursor.js").CursorLLMClient} CursorLLMClient
 * @typedef {import("./cursor.js").CursorClientOptions} CursorClientOptions
 */

/**
 * Create an LLM client that talks to the Cursor backend.
 *
 * This helper is intentionally dependency-free and works in both Node and
 * browser runtimes (relies on global `fetch`).
 *
 * @param {CursorClientOptions & { provider?: unknown }} [config]
 * @returns {import("./types.js").LLMClient}
 */
export function createLLMClient(config) {
  if (config == null) {
    return new CursorLLMClient();
  }

  if (!config || typeof config !== "object") {
    throw new Error("createLLMClient expects an optional configuration object.");
  }

  if ("provider" in config) {
    throw new Error("Provider selection is no longer supported; all AI uses Cursor backend.");
  }

  if ("apiKey" in config) {
    throw new Error(
      "User API keys are not supported; all AI uses Cursor backend auth via request headers (getAuthHeaders/authToken).",
    );
  }

  return new CursorLLMClient(config);
}
