import { CollabTokenKeychainStore, hasTauriInvoke, type CollabTokenKeychainEntry } from "./collabTokenKeychainStore.js";

type TokenKey = string;

type CollabTokenRecord = {
  token: string;
  expiresAtMs: number | null;
};

const MEMORY_TOKENS = new Map<TokenKey, CollabTokenRecord>();
// Cache of values loaded from the OS keychain-backed store (desktop).
const KEYCHAIN_TOKENS = new Map<TokenKey, CollabTokenRecord>();

const DEFAULT_OPAQUE_TOKEN_TTL_MS = 7 * 24 * 60 * 60 * 1000; // 7 days

let cachedKeychainStore: CollabTokenKeychainStore | null = null;
function getKeychainStore(): CollabTokenKeychainStore | null {
  if (cachedKeychainStore) return cachedKeychainStore;
  if (!hasTauriInvoke()) return null;
  try {
    cachedKeychainStore = new CollabTokenKeychainStore();
    return cachedKeychainStore;
  } catch {
    return null;
  }
}

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

function decodeBase64UrlToString(value: string): string | null {
  const input = String(value ?? "");
  if (!input) return null;
  const padded = input.replace(/-/g, "+").replace(/_/g, "/").padEnd(Math.ceil(input.length / 4) * 4, "=");
  try {
    if (typeof atob === "function") return atob(padded);
  } catch {
    // fall through
  }
  try {
    // Node (vitest) fallback.
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const buf = (globalThis as any).Buffer?.from?.(padded, "base64");
    if (buf) return buf.toString("utf8");
  } catch {
    // ignore
  }
  return null;
}

function jwtExpMs(token: string): number | null {
  const raw = String(token ?? "");
  const parts = raw.split(".");
  if (parts.length < 2) return null;
  const payload = parts[1] ?? "";
  const decoded = decodeBase64UrlToString(payload);
  if (!decoded) return null;
  try {
    const parsed = JSON.parse(decoded) as any;
    const exp = parsed?.exp;
    if (typeof exp === "number" && Number.isFinite(exp)) {
      return Math.trunc(exp * 1000);
    }
    if (typeof exp === "string" && exp.trim() !== "") {
      const expNum = Number(exp);
      if (Number.isFinite(expNum)) return Math.trunc(expNum * 1000);
    }
  } catch {
    // ignore parse failures
  }
  return null;
}

function opaqueTokenTtlMs(): number {
  // Allow overriding the conservative TTL for opaque tokens (e.g. dev builds / tests).
  const configured = (globalThis as any).__FORMULA_COLLAB_OPAQUE_TOKEN_TTL_MS;
  if (typeof configured === "number" && Number.isFinite(configured)) return configured;
  // `import.meta.env` is present in Vite builds; keep the access defensive for tests.
  const envRaw = (import.meta as any)?.env?.VITE_COLLAB_OPAQUE_TOKEN_TTL_MS;
  if (typeof envRaw === "string" && envRaw.trim() !== "") {
    const parsed = Number(envRaw);
    if (Number.isFinite(parsed)) return parsed;
  }
  return DEFAULT_OPAQUE_TOKEN_TTL_MS;
}

function computeExpiresAtMs(token: string, nowMs: number): number | null {
  const expFromJwt = jwtExpMs(token);
  if (expFromJwt != null) return expFromJwt;

  const ttl = Math.trunc(opaqueTokenTtlMs());
  if (!Number.isFinite(ttl) || ttl <= 0) return null;
  return nowMs + ttl;
}

function isExpired(record: CollabTokenRecord, nowMs: number): boolean {
  const exp = record.expiresAtMs;
  if (exp == null) return false;
  if (typeof exp !== "number" || !Number.isFinite(exp)) return true;
  return exp <= nowMs;
}

function recordFromUnknown(value: unknown): CollabTokenRecord | null {
  if (value == null) return null;
  if (typeof value === "string") {
    const token = value;
    if (!token) return null;
    // Legacy payloads stored token strings directly. Only JWT tokens have a reliable expiry.
    const exp = jwtExpMs(token);
    return { token, expiresAtMs: exp };
  }

  if (typeof value !== "object") return null;
  const token = typeof (value as any).token === "string" ? (value as any).token : "";
  if (!token) return null;
  const expiresAtMsRaw = (value as any).expiresAtMs;
  const expiresAtMs =
    expiresAtMsRaw == null ? null : typeof expiresAtMsRaw === "number" ? expiresAtMsRaw : Number(expiresAtMsRaw);
  return { token, expiresAtMs: Number.isFinite(expiresAtMs as number) ? (expiresAtMs as number) : null };
}

function parseStorageRecord(raw: string | null): CollabTokenRecord | null {
  if (!raw) return null;
  const trimmed = raw.trim();
  if (!trimmed) return null;
  if (trimmed.startsWith("{")) {
    try {
      return recordFromUnknown(JSON.parse(trimmed));
    } catch {
      // fall through to treating as raw token.
    }
  }
  return recordFromUnknown(trimmed);
}

function persistRecordToSessionStorage(key: TokenKey, record: CollabTokenRecord): void {
  const storage = getSessionStorage();
  if (!storage) return;
  try {
    storage.setItem(key, JSON.stringify(record));
  } catch {
    // ignore; fall back to in-memory only
  }
}

function persistRecordLocally(key: TokenKey, record: CollabTokenRecord): void {
  persistRecordToSessionStorage(key, record);
  MEMORY_TOKENS.set(key, record);
}

/**
 * Load (and cache) a collab token from the secure desktop store into session/in-memory caches.
 *
 * This should be invoked early in desktop startup so synchronous collab option resolution
 * can auto-reconnect before the user interacts with the app.
 */
export async function preloadCollabTokenFromKeychain(opts: { wsUrl: string; docId: string }): Promise<void> {
  const wsUrl = String(opts.wsUrl ?? "").trim();
  const docId = String(opts.docId ?? "").trim();
  if (!wsUrl || !docId) return;

  const store = getKeychainStore();
  if (!store) return;

  const key = makeTokenKey(wsUrl, docId);
  let entry: CollabTokenKeychainEntry | null = null;
  try {
    entry = await store.get(key);
  } catch {
    return;
  }
  if (!entry) return;

  const record = recordFromUnknown(entry);
  if (!record) return;

  const now = Date.now();
  if (isExpired(record, now)) {
    KEYCHAIN_TOKENS.delete(key);
    MEMORY_TOKENS.delete(key);
    const storage = getSessionStorage();
    if (storage) {
      try {
        storage.removeItem(key);
      } catch {
        // ignore
      }
    }
    try {
      await store.delete(key);
    } catch {
      // ignore
    }
    return;
  }

  KEYCHAIN_TOKENS.set(key, record);
  persistRecordLocally(key, record);
}

/**
 * Store the sync-server token for the current browser session (and persist to the desktop
 * secure store when available).
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
  const now = Date.now();
  const expiresAtMs = computeExpiresAtMs(token, now);
  const record: CollabTokenRecord = { token, expiresAtMs };

  if (isExpired(record, now)) {
    deleteCollabToken({ wsUrl, docId });
    return;
  }

  persistRecordLocally(key, record);
  KEYCHAIN_TOKENS.set(key, record);

  const store = getKeychainStore();
  if (store) {
    // Fire-and-forget: do not block the UI thread on keychain I/O.
    void store
      .set(key, { token, expiresAtMs })
      .catch(() => {
        // Best-effort.
      });
  }
}

export function loadCollabToken(opts: { wsUrl: string; docId: string }): string | null {
  const wsUrl = String(opts.wsUrl ?? "").trim();
  const docId = String(opts.docId ?? "").trim();
  if (!wsUrl || !docId) return null;

  const key = makeTokenKey(wsUrl, docId);
  const now = Date.now();

  // 1) Secure store cache (desktop).
  const fromKeychain = KEYCHAIN_TOKENS.get(key);
  if (fromKeychain) {
    if (isExpired(fromKeychain, now)) {
      deleteCollabToken({ wsUrl, docId });
      return null;
    }
    return fromKeychain.token;
  }

  // 2) Session-scoped storage.
  const storage = getSessionStorage();
  if (storage) {
    try {
      const value = parseStorageRecord(storage.getItem(key));
      if (value) {
        if (isExpired(value, now)) {
          deleteCollabToken({ wsUrl, docId });
          return null;
        }
        return value.token;
      }
    } catch {
      // ignore
    }
  }

  // 3) In-memory fallback.
  const fromMemory = MEMORY_TOKENS.get(key);
  if (fromMemory) {
    if (isExpired(fromMemory, now)) {
      deleteCollabToken({ wsUrl, docId });
      return null;
    }
    return fromMemory.token;
  }

  return null;
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
  KEYCHAIN_TOKENS.delete(key);

  const store = getKeychainStore();
  if (store) {
    void store.delete(key).catch(() => {
      // Best-effort.
    });
  }
}
