type TauriInvoke = (cmd: string, args?: Record<string, unknown>) => Promise<unknown>;

function getTauriInvokeOrNull(): TauriInvoke | null {
  const invoke = (globalThis as any).__TAURI__?.core?.invoke as TauriInvoke | undefined;
  return typeof invoke === "function" ? invoke : null;
}

// Collaborative cell encryption currently uses 32-byte AES-256-GCM keys.
const CELL_ENCRYPTION_KEY_BYTES = 32;

type CellEncryptionKey = {
  keyId: string;
  keyBytes: Uint8Array;
};

function assertKeyBytes(keyBytes: Uint8Array): void {
  if (!(keyBytes instanceof Uint8Array)) {
    throw new TypeError("keyBytes must be a Uint8Array");
  }
  if (keyBytes.byteLength !== CELL_ENCRYPTION_KEY_BYTES) {
    throw new RangeError(`keyBytes must be ${CELL_ENCRYPTION_KEY_BYTES} bytes (got ${keyBytes.byteLength})`);
  }
}

function bytesToBase64(bytes: Uint8Array): string {
  if (typeof Buffer !== "undefined") {
    return Buffer.from(bytes).toString("base64");
  }
  if (typeof btoa !== "function") {
    throw new Error("base64 encoding is unavailable in this environment");
  }
  let bin = "";
  for (const b of bytes) bin += String.fromCharCode(b);
  return btoa(bin);
}

function base64ToBytes(value: string): Uint8Array {
  const str = String(value ?? "");
  if (typeof Buffer !== "undefined") {
    return new Uint8Array(Buffer.from(str, "base64"));
  }
  if (typeof atob !== "function") {
    throw new Error("base64 decoding is unavailable in this environment");
  }
  const bin = atob(str);
  const out = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i += 1) out[i] = bin.charCodeAt(i);
  return out;
}

export type CollabEncryptionKeyEntry = {
  keyId: string;
  keyBytesBase64: string;
};

export type CollabEncryptionKeyIdEntry = {
  keyId: string;
};

type CachedKey = {
  keyBytesBase64: string;
  keyBytes?: Uint8Array;
};

function normalizeKeyBytesBase64(value: string): string {
  const bytes = base64ToBytes(String(value ?? "").trim());
  assertKeyBytes(bytes);
  return bytesToBase64(bytes);
}

/**
 * Durable (Tauri) or in-memory (fallback) key store for collaborative cell encryption keys.
 *
 * Keys are stored by `(docId, keyId)` and persisted via Tauri when available.
 * The store also maintains an in-memory cache so synchronous `keyForCell` hooks
 * can consult previously-loaded keys.
 */
export class CollabEncryptionKeyStore {
  private readonly invoke: TauriInvoke | null;
  private readonly cache = new Map<string, Map<string, CachedKey>>();

  constructor(opts?: { invoke?: TauriInvoke | null }) {
    this.invoke = opts?.invoke ?? getTauriInvokeOrNull();
  }

  private docCache(docId: string): Map<string, CachedKey> {
    let existing = this.cache.get(docId);
    if (!existing) {
      existing = new Map();
      this.cache.set(docId, existing);
    }
    return existing;
  }

  private getCachedEntry(docId: string, keyId: string): CachedKey | null {
    const map = this.cache.get(docId);
    if (!map) return null;
    return map.get(keyId) ?? null;
  }

  /**
   * Synchronous cached lookup for use in CollabSession `encryption.keyForCell`.
   */
  getCachedKey(docId: string, keyId: string): CellEncryptionKey | null {
    const entry = this.getCachedEntry(docId, keyId);
    if (!entry) return null;

    if (!entry.keyBytes) {
      try {
        const bytes = base64ToBytes(entry.keyBytesBase64);
        assertKeyBytes(bytes);
        entry.keyBytes = bytes;
      } catch {
        return null;
      }
    }

    return { keyId, keyBytes: entry.keyBytes };
  }

  async get(docId: string, keyId: string): Promise<CollabEncryptionKeyEntry | null> {
    const docIdStr = String(docId ?? "").trim();
    const keyIdStr = String(keyId ?? "").trim();
    if (!docIdStr || !keyIdStr) return null;

    if (this.invoke) {
      try {
        const payload = await this.invoke("collab_encryption_key_get", { doc_id: docIdStr, key_id: keyIdStr });
        if (payload == null) return null;
        if (!payload || typeof payload !== "object") return null;
        const entry = payload as any;
        if (typeof entry.keyId !== "string" || entry.keyId.length === 0) return null;
        if (typeof entry.keyBytesBase64 !== "string" || entry.keyBytesBase64.length === 0) return null;
        const keyBytesBase64 = normalizeKeyBytesBase64(entry.keyBytesBase64);
        this.docCache(docIdStr).set(entry.keyId, { keyBytesBase64 });
        return { keyId: entry.keyId, keyBytesBase64 };
      } catch {
        // Graceful degradation (older backends / invoke failures): fall back to memory.
      }
    }

    const cached = this.getCachedEntry(docIdStr, keyIdStr);
    if (!cached) return null;
    return { keyId: keyIdStr, keyBytesBase64: cached.keyBytesBase64 };
  }

  async set(docId: string, keyId: string, keyBytesBase64: string): Promise<CollabEncryptionKeyIdEntry> {
    const docIdStr = String(docId ?? "").trim();
    const keyIdStr = String(keyId ?? "").trim();
    if (!docIdStr || !keyIdStr) {
      throw new Error("docId and keyId are required");
    }
    const normalized = normalizeKeyBytesBase64(keyBytesBase64);

    if (this.invoke) {
      try {
        const payload = await this.invoke("collab_encryption_key_set", {
          doc_id: docIdStr,
          key_id: keyIdStr,
          key_bytes_base64: normalized,
        });
        // Backend returns `{ keyId }`; tolerate missing payload for forward-compat.
        const returnedKeyId = (payload as any)?.keyId;
        const finalKeyId = typeof returnedKeyId === "string" && returnedKeyId ? returnedKeyId : keyIdStr;
        this.docCache(docIdStr).set(finalKeyId, { keyBytesBase64: normalized });
        return { keyId: finalKeyId };
      } catch {
        // Fall back to memory.
      }
    }

    this.docCache(docIdStr).set(keyIdStr, { keyBytesBase64: normalized });
    return { keyId: keyIdStr };
  }

  async delete(docId: string, keyId: string): Promise<void> {
    const docIdStr = String(docId ?? "").trim();
    const keyIdStr = String(keyId ?? "").trim();
    if (!docIdStr || !keyIdStr) return;

    if (this.invoke) {
      try {
        await this.invoke("collab_encryption_key_delete", { doc_id: docIdStr, key_id: keyIdStr });
      } catch {
        // Ignore and still clear local cache.
      }
    }

    const map = this.cache.get(docIdStr);
    map?.delete(keyIdStr);
  }

  async list(docId: string): Promise<CollabEncryptionKeyIdEntry[]> {
    const docIdStr = String(docId ?? "").trim();
    if (!docIdStr) return [];

    if (this.invoke) {
      try {
        const payload = await this.invoke("collab_encryption_key_list", { doc_id: docIdStr });
        if (!Array.isArray(payload)) return [];
        return payload
          .filter((e) => e && typeof e === "object")
          .map((e: any) => ({ keyId: String(e.keyId ?? "") }))
          .filter((e) => e.keyId.length > 0);
      } catch {
        // Fall through to in-memory cache.
      }
    }

    const map = this.cache.get(docIdStr);
    if (!map) return [];
    return Array.from(map.keys()).map((keyId) => ({ keyId }));
  }

  /**
   * Best-effort: load all known keys for `docId` from the persistent store into
   * the in-memory cache.
   */
  async hydrateDoc(docId: string): Promise<void> {
    const docIdStr = String(docId ?? "").trim();
    if (!docIdStr) return;

    const entries = await this.list(docIdStr);
    await Promise.all(entries.map((e) => this.get(docIdStr, e.keyId)));
  }
}
