export type BinaryStorage = {
  load(): Promise<Uint8Array | null>;
  save(data: Uint8Array): Promise<void>;
};

export class InMemoryBinaryStorage implements BinaryStorage {
  load(): Promise<Uint8Array | null>;
  save(data: Uint8Array): Promise<void>;
}

export class LocalStorageBinaryStorage implements BinaryStorage {
  constructor(opts: { workbookId: string; namespace?: string });
  readonly key: string;
  load(): Promise<Uint8Array | null>;
  save(data: Uint8Array): Promise<void>;
}

