import type { BinaryStorage } from "./binaryStorage.js";

/**
 * Node-only BinaryStorage implementation backed by a single file on disk.
 */
export class NodeFileBinaryStorage implements BinaryStorage {
  constructor(filePath: string);

  readonly filePath: string;

  load(): Promise<Uint8Array | null>;
  save(data: Uint8Array): Promise<void>;
  remove(): Promise<void>;
}

