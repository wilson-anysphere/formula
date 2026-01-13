import { afterEach, describe, expect, it, vi } from "vitest";

import { CollabEncryptionKeyStore } from "../encryptionKeyStore";

describe("CollabEncryptionKeyStore", () => {
  const originalTauri = (globalThis as any).__TAURI__;

  afterEach(() => {
    (globalThis as any).__TAURI__ = originalTauri;
    vi.restoreAllMocks();
  });

  it("falls back to in-memory storage when __TAURI__.core.invoke is unavailable", async () => {
    (globalThis as any).__TAURI__ = undefined;

    const store = new CollabEncryptionKeyStore();
    const docId = "doc-1";
    const keyId = "key-1";
    const keyBytesBase64 = Buffer.from(new Uint8Array(32).fill(7)).toString("base64");

    await expect(store.set(docId, keyId, keyBytesBase64)).resolves.toEqual({ keyId });
    await expect(store.list(docId)).resolves.toEqual([{ keyId }]);
    await expect(store.get(docId, keyId)).resolves.toEqual({ keyId, keyBytesBase64 });

    const cached = store.getCachedKey(docId, keyId);
    expect(cached?.keyId).toBe(keyId);
    expect(cached?.keyBytes).toBeInstanceOf(Uint8Array);
    expect(cached?.keyBytes.byteLength).toBe(32);

    await store.delete(docId, keyId);
    await expect(store.get(docId, keyId)).resolves.toBeNull();
  });

  it("roundtrips through Tauri invoke when available", async () => {
    const backend = new Map<string, string>();
    const invoke = vi.fn(async (cmd: string, args?: Record<string, unknown>) => {
      const docId = String((args as any)?.doc_id ?? "");
      const keyId = String((args as any)?.key_id ?? "");
      const composite = `${docId}:${keyId}`;

      switch (cmd) {
        case "collab_encryption_key_get": {
          const keyBytesBase64 = backend.get(composite);
          return keyBytesBase64 ? { keyId, keyBytesBase64 } : null;
        }
        case "collab_encryption_key_set": {
          const keyBytesBase64 = String((args as any)?.key_bytes_base64 ?? "");
          backend.set(composite, keyBytesBase64);
          return { keyId };
        }
        case "collab_encryption_key_delete": {
          backend.delete(composite);
          return null;
        }
        case "collab_encryption_key_list": {
          const docPrefix = `${docId}:`;
          return Array.from(backend.keys())
            .filter((k) => k.startsWith(docPrefix))
            .map((k) => ({ keyId: k.slice(docPrefix.length) }));
        }
        default:
          throw new Error(`Unexpected command: ${cmd}`);
      }
    });

    (globalThis as any).__TAURI__ = { core: { invoke } };

    const store = new CollabEncryptionKeyStore();
    const docId = "doc-1";
    const keyId = "key-1";
    const keyBytesBase64 = Buffer.from(new Uint8Array(32).fill(1)).toString("base64");

    await store.set(docId, keyId, keyBytesBase64);

    expect(invoke).toHaveBeenCalledWith("collab_encryption_key_set", {
      doc_id: docId,
      key_id: keyId,
      key_bytes_base64: keyBytesBase64,
    });

    await expect(store.list(docId)).resolves.toEqual([{ keyId }]);
    await expect(store.get(docId, keyId)).resolves.toEqual({ keyId, keyBytesBase64 });

    const cached = store.getCachedKey(docId, keyId);
    expect(cached?.keyId).toBe(keyId);
    expect(cached?.keyBytes.byteLength).toBe(32);

    await store.delete(docId, keyId);
    await expect(store.get(docId, keyId)).resolves.toBeNull();
  });
});
