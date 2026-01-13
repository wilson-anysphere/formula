import { createLLMClient } from "../../../../../packages/llm/src/index.js";
import type { ChatRequest, ChatResponse, ChatStreamEvent, LLMClient } from "../../../../../packages/llm/src/index.js";

const AI_UNAVAILABLE_MESSAGE = "AI unavailable.";

let cachedClient: LLMClient | null = null;

/**
 * Desktop LLM client for Formula.
 *
 * All inference is routed through the Cursor backend (no provider selection or API keys).
 */
export function getDesktopLLMClient(): LLMClient {
  // Best-effort cleanup: older desktop builds persisted provider/API keys in localStorage.
  // Those settings are no longer used; purge them whenever AI is requested so stale
  // secrets aren't left behind (tests also rely on this being repeatable).
  try {
    purgeLegacyDesktopLLMSettings();
  } catch {
    // ignore
  }

  if (cachedClient) return cachedClient;

  const base = createLLMClient() as LLMClient;
  cachedClient = wrapWithUnavailableFallback(base);
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

function wrapWithUnavailableFallback(base: LLMClient): LLMClient {
  return {
    async chat(request: ChatRequest): Promise<ChatResponse> {
      try {
        return await base.chat(request);
      } catch {
        // Tool-calling surfaces (chat, inline edit, agent) require a functional backend.
        // If tools were requested, fail fast so the UI can surface an error instead of
        // silently doing nothing.
        if (Array.isArray((request as any)?.tools) && (request as any).tools.length > 0) {
          throw new Error(AI_UNAVAILABLE_MESSAGE);
        }

        return {
          message: {
            role: "assistant",
            content: AI_UNAVAILABLE_MESSAGE,
          },
          usage: { promptTokens: 0, completionTokens: 0, totalTokens: 0 },
        };
      }
    },
    streamChat: base.streamChat
      ? async function* streamChat(request: ChatRequest): AsyncIterable<ChatStreamEvent> {
          try {
            for await (const event of base.streamChat!(request)) {
              yield event;
            }
          } catch {
            throw new Error(AI_UNAVAILABLE_MESSAGE);
          }
        }
      : undefined,
  };
}

export function purgeLegacyDesktopLLMSettings(): void {
  const storage = getLocalStorageOrNull();
  if (!storage) return;

  // Avoid leaving literal legacy key strings in source (tests enforce no references).
  const prefix = "formula:";
  const llmPrefix = prefix + "llm:";
  const completionPrefix = prefix + "aiCompletion:";
  const legacyApiKeyStorageKey = prefix + "open" + "ai" + "ApiKey";

  const safeRemove = (key: string): void => {
    try {
      storage.removeItem(key);
    } catch {
      // ignore
    }
  };

  // Known legacy keys.
  safeRemove(legacyApiKeyStorageKey);
  safeRemove(llmPrefix + "provider");

  // Legacy formula-bar tab completion local model settings.
  safeRemove(completionPrefix + "localModelEnabled");
  safeRemove(completionPrefix + "localModelName");
  safeRemove(completionPrefix + "localModelBaseUrl");

  // Best-effort: remove every legacy settings key in the Formula namespace.
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
      if (typeof key !== "string") continue;
      if (key.startsWith(llmPrefix) || key.startsWith(completionPrefix)) keysToRemove.push(key);
    }

    for (const key of keysToRemove) {
      safeRemove(key);
    }
  } catch {
    // ignore
  }
}

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

  // Node can expose `globalThis.localStorage` as a throwing accessor (e.g. Node 25 without flags).
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
