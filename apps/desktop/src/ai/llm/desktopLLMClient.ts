import { createLLMClient } from "../../../../../packages/llm/src/createLLMClient.js";
import type { LLMClient } from "../../../../../packages/llm/src/types.js";

let cachedClient: LLMClient | null = null;

const LEGACY_OPENAI_API_KEY_STORAGE_KEY = "formula:openaiApiKey";
const LLM_PROVIDER_STORAGE_KEY = "formula:llm:provider";
const LLM_SETTINGS_PREFIX = "formula:llm:";

function getLocalStorageOrNull(): Storage | null {
  // Prefer `window.localStorage` when available (jsdom + browser runtimes).
  if (typeof window !== "undefined") {
    try {
      const storage = window.localStorage;
      if (!storage) return null;
      if (typeof storage.getItem !== "function" || typeof storage.removeItem !== "function") return null;
      return storage;
    } catch {
      // ignore
    }
  }

  // Node 22+/25 can expose `globalThis.localStorage` as a throwing accessor.
  try {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const storage = (globalThis as any).localStorage as Storage | undefined;
    if (!storage) return null;
    if (typeof storage.getItem !== "function" || typeof storage.removeItem !== "function") return null;
    return storage;
  } catch {
    return null;
  }
}

/**
 * Best-effort removal of legacy desktop LLM settings (provider selection + API keys).
 *
 * Cursor desktop no longer supports user-provided API keys/provider selection, but
 * old builds persisted secrets in localStorage. We proactively delete those keys
 * so stale secrets are not left behind on disk.
 */
export function purgeLegacyDesktopLLMSettings(): void {
  const storage = getLocalStorageOrNull();
  if (!storage) return;

  // Remove known legacy keys even if enumeration fails.
  try {
    storage.removeItem(LEGACY_OPENAI_API_KEY_STORAGE_KEY);
  } catch {
    // ignore
  }
  try {
    storage.removeItem(LLM_PROVIDER_STORAGE_KEY);
  } catch {
    // ignore
  }

  // Best-effort: remove every `formula:llm:*` key, including provider + per-provider settings.
  try {
    const keysToRemove: string[] = [];
    const length = storage.length;
    for (let i = 0; i < length; i += 1) {
      let key: string | null = null;
      try {
        key = storage.key(i);
      } catch {
        continue;
      }
      if (typeof key === "string" && key.startsWith(LLM_SETTINGS_PREFIX)) keysToRemove.push(key);
    }

    for (const key of keysToRemove) {
      try {
        storage.removeItem(key);
      } catch {
        // ignore
      }
    }
  } catch {
    // ignore
  }
}

/**
 * Desktop LLM client for Formula.
 *
 * All inference is routed through the Cursor backend (no provider selection or API keys).
 */
export function getDesktopLLMClient(): LLMClient {
  // One-time best-effort cleanup: older desktop builds persisted provider/API keys
  // in localStorage. Those settings are no longer used; purge them on first AI use
  // so stale secrets aren't left behind.
  try {
    purgeLegacyDesktopLLMSettings();
  } catch {
    // ignore
  }

  if (!cachedClient) cachedClient = createLLMClient();
  return cachedClient;
}

/**
 * Default model identifier used for prompt budgeting and as the `model` field on LLM requests.
 *
 * Cursor's backend may choose a different underlying model, but downstream code still needs a
 * stable identifier for context-window heuristics.
 */
export function getDesktopModel(): string {
  return "gpt-4o-mini";
}
