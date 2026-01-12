export type LLMProvider = "openai" | "anthropic" | "ollama";

export const LEGACY_OPENAI_API_KEY_STORAGE_KEY = "formula:openaiApiKey";

export const LLM_PROVIDER_STORAGE_KEY = "formula:llm:provider";

export const OPENAI_API_KEY_STORAGE_KEY = "formula:llm:openai:apiKey";
export const OPENAI_BASE_URL_STORAGE_KEY = "formula:llm:openai:baseUrl";
export const OPENAI_MODEL_STORAGE_KEY = "formula:llm:openai:model";
export const ANTHROPIC_API_KEY_STORAGE_KEY = "formula:llm:anthropic:apiKey";
export const ANTHROPIC_MODEL_STORAGE_KEY = "formula:llm:anthropic:model";

export const OLLAMA_BASE_URL_STORAGE_KEY = "formula:llm:ollama:baseUrl";
export const OLLAMA_MODEL_STORAGE_KEY = "formula:llm:ollama:model";

export const DEFAULT_OLLAMA_BASE_URL = "http://127.0.0.1:11434";
export const DEFAULT_OLLAMA_MODEL = "llama3.1";

export type DesktopLLMConfig =
  | { provider: "openai"; apiKey: string; model?: string; baseUrl?: string }
  | { provider: "anthropic"; apiKey: string; model?: string }
  | { provider: "ollama"; baseUrl: string; model: string };

function getLocalStorageOrNull(): Storage | null {
  // Prefer `window.localStorage` when available.
  if (typeof window !== "undefined") {
    try {
      const storage = window.localStorage;
      if (!storage) return null;
      if (typeof storage.getItem !== "function" || typeof storage.setItem !== "function") return null;
      return storage;
    } catch {
      // ignore
    }
  }

  try {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const storage = (globalThis as any).localStorage as Storage | undefined;
    if (!storage) return null;
    if (typeof storage.getItem !== "function" || typeof storage.setItem !== "function") return null;
    return storage;
  } catch {
    return null;
  }
}

function purgeDeprecatedDesktopAiLocalStorageKeys(storage: Storage | null): void {
  // Legacy formula-bar tab completion local model settings.
  //
  // NOTE: We intentionally avoid embedding the full key prefix as a single string
  // literal so grep-based checks can ensure no codepaths accidentally *read* these
  // keys to re-enable disallowed local models.
  const completionPrefix = "formula:" + "aiCompletion:";
  safeRemove(storage, completionPrefix + "localModelEnabled");
  safeRemove(storage, completionPrefix + "localModelName");
  safeRemove(storage, completionPrefix + "localModelBaseUrl");
}

function readViteEnv(key: string): string | null {
  try {
    const env = (import.meta as any)?.env;
    const value = env?.[key];
    if (typeof value === "string" && value.length > 0) return value;
  } catch {
    // ignore (not running under Vite)
  }
  return null;
}

export function migrateLegacyOpenAIKey(): void {
  migrateLegacyOpenAIKeyWithStorage(getLocalStorageOrNull());
}

function migrateLegacyOpenAIKeyWithStorage(storage: Storage | null): void {
  if (!storage) return;
  try {
    const legacy = storage.getItem(LEGACY_OPENAI_API_KEY_STORAGE_KEY);
    if (legacy && !storage.getItem(OPENAI_API_KEY_STORAGE_KEY)) {
      storage.setItem(OPENAI_API_KEY_STORAGE_KEY, legacy);
    }
    if (legacy && !storage.getItem(LLM_PROVIDER_STORAGE_KEY)) {
      storage.setItem(LLM_PROVIDER_STORAGE_KEY, "openai");
    }
  } catch {
    // ignore
  }
}

export function loadDesktopLLMConfig(): DesktopLLMConfig | null {
  const storage = getLocalStorageOrNull();
  purgeDeprecatedDesktopAiLocalStorageKeys(storage);
  migrateLegacyOpenAIKeyWithStorage(storage);

  /** @type {LLMProvider} */
  let provider: LLMProvider = "openai";
  try {
    const raw = storage?.getItem(LLM_PROVIDER_STORAGE_KEY);
    if (raw === "openai" || raw === "anthropic" || raw === "ollama") provider = raw;
  } catch {
    // ignore
  }

  if (provider === "openai") {
    const apiKey =
      safeGet(storage, OPENAI_API_KEY_STORAGE_KEY) ??
      safeGet(storage, LEGACY_OPENAI_API_KEY_STORAGE_KEY) ??
      readViteEnv("VITE_OPENAI_API_KEY");
    if (!apiKey) return null;
    const model = safeGet(storage, OPENAI_MODEL_STORAGE_KEY) ?? undefined;
    const baseUrl = safeGet(storage, OPENAI_BASE_URL_STORAGE_KEY) ?? readViteEnv("VITE_OPENAI_BASE_URL") ?? undefined;
    return { provider: "openai", apiKey, model, baseUrl };
  }

  if (provider === "anthropic") {
    const apiKey = safeGet(storage, ANTHROPIC_API_KEY_STORAGE_KEY) ?? readViteEnv("VITE_ANTHROPIC_API_KEY");
    if (!apiKey) return null;
    const model = safeGet(storage, ANTHROPIC_MODEL_STORAGE_KEY) ?? undefined;
    return { provider: "anthropic", apiKey, model };
  }

  const baseUrl = safeGet(storage, OLLAMA_BASE_URL_STORAGE_KEY) ?? DEFAULT_OLLAMA_BASE_URL;
  const model = safeGet(storage, OLLAMA_MODEL_STORAGE_KEY) ?? DEFAULT_OLLAMA_MODEL;
  return { provider: "ollama", baseUrl, model };
}

function safeGet(storage: Storage | null, key: string): string | null {
  if (!storage) return null;
  try {
    return storage.getItem(key);
  } catch {
    return null;
  }
}

function safeSet(storage: Storage | null, key: string, value: string): void {
  if (!storage) return;
  try {
    storage.setItem(key, value);
  } catch {
    // ignore
  }
}

function safeRemove(storage: Storage | null, key: string): void {
  if (!storage) return;
  try {
    storage.removeItem(key);
  } catch {
    // ignore
  }
}

export function saveDesktopLLMConfig(config: DesktopLLMConfig): void {
  const storage = getLocalStorageOrNull();
  safeSet(storage, LLM_PROVIDER_STORAGE_KEY, config.provider);

  if (config.provider === "openai") {
    safeSet(storage, OPENAI_API_KEY_STORAGE_KEY, config.apiKey);
    // Maintain backward compatibility for older builds/tests that only read the legacy key.
    safeSet(storage, LEGACY_OPENAI_API_KEY_STORAGE_KEY, config.apiKey);
    if (config.model) safeSet(storage, OPENAI_MODEL_STORAGE_KEY, config.model);
    else safeRemove(storage, OPENAI_MODEL_STORAGE_KEY);
    if (config.baseUrl) safeSet(storage, OPENAI_BASE_URL_STORAGE_KEY, config.baseUrl);
    else safeRemove(storage, OPENAI_BASE_URL_STORAGE_KEY);
    return;
  }

  if (config.provider === "anthropic") {
    safeSet(storage, ANTHROPIC_API_KEY_STORAGE_KEY, config.apiKey);
    if (config.model) safeSet(storage, ANTHROPIC_MODEL_STORAGE_KEY, config.model);
    else safeRemove(storage, ANTHROPIC_MODEL_STORAGE_KEY);
    return;
  }

  safeSet(storage, OLLAMA_BASE_URL_STORAGE_KEY, config.baseUrl);
  safeSet(storage, OLLAMA_MODEL_STORAGE_KEY, config.model);
}

export function clearDesktopLLMConfig(): void {
  const storage = getLocalStorageOrNull();
  safeRemove(storage, LLM_PROVIDER_STORAGE_KEY);

  safeRemove(storage, OPENAI_API_KEY_STORAGE_KEY);
  safeRemove(storage, OPENAI_MODEL_STORAGE_KEY);
  safeRemove(storage, OPENAI_BASE_URL_STORAGE_KEY);
  safeRemove(storage, ANTHROPIC_API_KEY_STORAGE_KEY);
  safeRemove(storage, ANTHROPIC_MODEL_STORAGE_KEY);

  safeRemove(storage, OLLAMA_BASE_URL_STORAGE_KEY);
  safeRemove(storage, OLLAMA_MODEL_STORAGE_KEY);

  // Keep legacy key clear to match user intent.
  safeRemove(storage, LEGACY_OPENAI_API_KEY_STORAGE_KEY);
}
