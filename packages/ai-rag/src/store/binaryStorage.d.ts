export type BinaryStorage = {
  load(): Promise<Uint8Array | null>;
  save(data: Uint8Array): Promise<void>;
  remove?(): Promise<void>;
};

export function toBase64(data: Uint8Array): string;

export function fromBase64(encoded: string): Uint8Array;

export class InMemoryBinaryStorage implements BinaryStorage {
  load(): Promise<Uint8Array | null>;
  save(data: Uint8Array): Promise<void>;
  remove(): Promise<void>;
}

export class LocalStorageBinaryStorage implements BinaryStorage {
  constructor(opts: { workbookId: string; namespace?: string });
  readonly key: string;
  load(): Promise<Uint8Array | null>;
  save(data: Uint8Array): Promise<void>;
  remove(): Promise<void>;
}
export class ChunkedLocalStorageBinaryStorage implements BinaryStorage {
  constructor(opts: { workbookId: string; namespace?: string; chunkSizeChars?: number });
  readonly key: string;
  load(): Promise<Uint8Array | null>;
  save(data: Uint8Array): Promise<void>;
  remove(): Promise<void>;
}

export class IndexedDBBinaryStorage implements BinaryStorage {
  constructor(opts: { workbookId: string; namespace?: string; dbName?: string });
  readonly dbName: string;
  readonly key: string;
  readonly namespace: string;
  load(): Promise<Uint8Array | null>;
  save(data: Uint8Array): Promise<void>;
  remove(): Promise<void>;
}
