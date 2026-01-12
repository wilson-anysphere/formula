type TokenKey = string;

const MEMORY_TOKENS = new Map<TokenKey, string>();

function makeTokenKey(wsUrl: string, docId: string): TokenKey {
  const ws = String(wsUrl ?? "").trim();
  const doc = String(docId ?? "").trim();
  // Tokens are never included in the key.
  return `formula:collab:token:${ws}|${doc}`;
}

function getSessionStorage(): Storage | null {
  try {
    if (typeof window === "undefined") return null;
    return window.sessionStorage ?? null;
  } catch {
    return null;
  }
}

/**
 * Store the sync-server token for the current browser session.
 *
 * IMPORTANT:
 * - Tokens must never be logged.
 * - We intentionally avoid persisting tokens in localStorage; sessionStorage is
 *   ephemeral and cleared when the tab/app session ends.
 */
export function storeCollabToken(opts: { wsUrl: string; docId: string; token: string }): void {
  const wsUrl = String(opts.wsUrl ?? "").trim();
  const docId = String(opts.docId ?? "").trim();
  const token = String(opts.token ?? "");
  if (!wsUrl || !docId || !token) return;

  const key = makeTokenKey(wsUrl, docId);
  const storage = getSessionStorage();
  if (storage) {
    try {
      storage.setItem(key, token);
      return;
    } catch {
      // Fall back to in-memory store.
    }
  }

  MEMORY_TOKENS.set(key, token);
}

export function loadCollabToken(opts: { wsUrl: string; docId: string }): string | null {
  const wsUrl = String(opts.wsUrl ?? "").trim();
  const docId = String(opts.docId ?? "").trim();
  if (!wsUrl || !docId) return null;

  const key = makeTokenKey(wsUrl, docId);
  const storage = getSessionStorage();
  if (storage) {
    try {
      const value = storage.getItem(key);
      if (typeof value === "string" && value.length > 0) return value;
    } catch {
      // ignore
    }
  }

  return MEMORY_TOKENS.get(key) ?? null;
}

export function deleteCollabToken(opts: { wsUrl: string; docId: string }): void {
  const wsUrl = String(opts.wsUrl ?? "").trim();
  const docId = String(opts.docId ?? "").trim();
  if (!wsUrl || !docId) return;

  const key = makeTokenKey(wsUrl, docId);
  const storage = getSessionStorage();
  if (storage) {
    try {
      storage.removeItem(key);
    } catch {
      // ignore
    }
  }

  MEMORY_TOKENS.delete(key);
}

