import { promises as fs } from "node:fs";

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

export class NodeFileBinaryStorage implements SqliteBinaryStorage {
  readonly filePath: string;

  constructor(filePath: string) {
    this.filePath = filePath;
  }

  async load(): Promise<Uint8Array | null> {
    try {
      const buffer = await fs.readFile(this.filePath);
      return new Uint8Array(buffer);
    } catch (error) {
      if ((error as NodeJS.ErrnoException).code === "ENOENT") return null;
      throw error;
    }
  }

  async save(data: Uint8Array): Promise<void> {
    await fs.writeFile(this.filePath, data);
  }
}

export class LocalStorageBinaryStorage implements SqliteBinaryStorage {
  readonly key: string;

  constructor(key: string = "ai_audit_db") {
    this.key = key;
  }

  async load(): Promise<Uint8Array | null> {
    if (typeof localStorage === "undefined") return null;
    const encoded = localStorage.getItem(this.key);
    if (!encoded) return null;
    return fromBase64(encoded);
  }

  async save(data: Uint8Array): Promise<void> {
    if (typeof localStorage === "undefined") return;
    localStorage.setItem(this.key, toBase64(data));
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

