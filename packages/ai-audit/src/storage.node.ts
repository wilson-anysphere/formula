import { promises as fs } from "node:fs";

import type { SqliteBinaryStorage } from "./storage.ts";

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
