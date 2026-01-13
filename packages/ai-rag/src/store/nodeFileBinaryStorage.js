import { mkdir, readFile, rename, unlink, writeFile } from "node:fs/promises";
import path from "node:path";

export class NodeFileBinaryStorage {
  /**
   * @param {string} filePath
   */
  constructor(filePath) {
    if (!filePath) throw new Error("NodeFileBinaryStorage requires filePath");
    this.filePath = filePath;
  }

  async load() {
    try {
      const buffer = await readFile(this.filePath);
      return new Uint8Array(buffer);
    } catch (err) {
      if (err && err.code === "ENOENT") return null;
      throw err;
    }
  }

  /**
   * @param {Uint8Array} data
   */
  async save(data) {
    const dir = path.dirname(this.filePath);
    await mkdir(dir, { recursive: true });
    const tmp = `${this.filePath}.tmp`;
    await writeFile(tmp, data);
    await rename(tmp, this.filePath);
  }

  async remove() {
    try {
      await unlink(this.filePath);
    } catch (err) {
      if (err && err.code === "ENOENT") return;
      throw err;
    }
  }
}
