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

  let binary = "";
  for (const byte of data) binary += String.fromCharCode(byte);
  // eslint-disable-next-line no-undef
  return btoa(binary);
}

function fromBase64(encoded: string): Uint8Array {
  if (typeof Buffer !== "undefined") {
    return new Uint8Array(Buffer.from(encoded, "base64"));
  }
  // eslint-disable-next-line no-undef
  const binary = atob(encoded);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
  return bytes;
}
