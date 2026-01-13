export interface SqliteBinaryStorage {
  load(): Promise<Uint8Array | null>;
  save(data: Uint8Array): Promise<void>;
}

export class InMemoryBinaryStorage implements SqliteBinaryStorage {
  private data: Uint8Array | null = null;

  async load(): Promise<Uint8Array | null> {
    return this.data ? new Uint8Array(this.data) : null;
  }

  async save(data: Uint8Array): Promise<void> {
    this.data = new Uint8Array(data);
  }
}

export class LocalStorageBinaryStorage implements SqliteBinaryStorage {
  readonly key: string;

  constructor(key: string = "ai_audit_db") {
    this.key = key;
  }

  async load(): Promise<Uint8Array | null> {
    const storage = safeLocalStorage();
    if (!storage) return null;
    try {
      const encoded = storage.getItem(this.key);
      if (!encoded) return null;
      return fromBase64(encoded);
    } catch {
      return null;
    }
  }

  async save(data: Uint8Array): Promise<void> {
    const storage = safeLocalStorage();
    if (!storage) return;
    try {
      storage.setItem(this.key, toBase64(data));
    } catch {
      // Ignore persistence failures (e.g. quota exceeded / private mode).
    }
  }
}

function safeLocalStorage(): Storage | undefined {
  try {
    const storage = globalThis.localStorage;
    if (storage) return storage;
  } catch {}

  try {
    return (globalThis as any).window?.localStorage;
  } catch {
    return undefined;
  }
}

function toBase64(data: Uint8Array): string {
  if (typeof Buffer !== "undefined") {
    return Buffer.from(data).toString("base64");
  }

  // Browser path: avoid building a giant binary string via `binary += ...` which can
  // lead to O(n^2) behavior and large intermediate allocations for multi‑MB blobs.
  //
  // We encode in chunks and concatenate the base64 output. For correctness, all
  // chunks except the last must be a multiple of 3 bytes (base64's 24‑bit groups)
  // so that we don't introduce padding in the middle of the output.
  const MAX_FROM_CHAR_CODE_ARGS = 0x8000; // 32KB-ish; safe for `String.fromCharCode(...chunk)`.
  const CHUNK_BYTES = MAX_FROM_CHAR_CODE_ARGS - (MAX_FROM_CHAR_CODE_ARGS % 3); // keep base64 alignment

  const parts: string[] = [];
  for (let i = 0; i < data.length; i += CHUNK_BYTES) {
    const chunk = data.subarray(i, i + CHUNK_BYTES);
    // eslint-disable-next-line no-undef
    parts.push(btoa(String.fromCharCode(...chunk)));
  }
  return parts.join("");
}

function fromBase64(encoded: string): Uint8Array {
  if (typeof Buffer !== "undefined") {
    return new Uint8Array(Buffer.from(encoded, "base64"));
  }

  // Browser path: decode base64 in chunks so we don't create a single multi‑MB
  // binary string via `atob(encoded)` (peak memory) for large payloads.
  //
  // Chunk boundaries must be on 4-character boundaries (base64's 6‑bit groups).
  const padding = encoded.endsWith("==") ? 2 : encoded.endsWith("=") ? 1 : 0;
  const outLen = Math.floor((encoded.length * 3) / 4) - padding;
  const bytes = new Uint8Array(outLen);

  const CHUNK_CHARS = 64 * 1024; // divisible by 4
  const chunkSize = CHUNK_CHARS - (CHUNK_CHARS % 4);

  let offset = 0;
  for (let i = 0; i < encoded.length; i += chunkSize) {
    const chunk = encoded.slice(i, i + chunkSize);
    // eslint-disable-next-line no-undef
    const binary = atob(chunk);
    for (let j = 0; j < binary.length; j++) {
      bytes[offset++] = binary.charCodeAt(j);
    }
  }

  return bytes;
}
