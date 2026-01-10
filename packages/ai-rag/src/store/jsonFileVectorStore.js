import { mkdir, readFile, rename, writeFile } from "node:fs/promises";
import path from "node:path";

import { InMemoryVectorStore } from "./inMemoryVectorStore.js";

/**
 * A small, dependency-free persistent store.
 *
 * This intentionally keeps a full in-memory copy for queries, then snapshots
 * to disk on mutation. It is not intended for very large corpora, but it
 * provides the persistence and interface needed for workbook-scale RAG.
 */
export class JsonFileVectorStore extends InMemoryVectorStore {
  /**
   * @param {{ filePath: string, dimension: number }} opts
   */
  constructor(opts) {
    super({ dimension: opts.dimension });
    if (!opts?.filePath) throw new Error("JsonFileVectorStore requires filePath");
    this._filePath = opts.filePath;
    this._loaded = false;
  }

  /**
   * Load records from disk (idempotent).
   */
  async load() {
    if (this._loaded) return;
    this._loaded = true;
    try {
      const raw = await readFile(this._filePath, "utf8");
      const parsed = JSON.parse(raw);
      if (
        !parsed ||
        typeof parsed !== "object" ||
        parsed.version !== 1 ||
        parsed.dimension !== this.dimension ||
        !Array.isArray(parsed.records)
      ) {
        return;
      }
      await super.upsert(
        parsed.records.map((r) => ({ id: r.id, vector: r.vector, metadata: r.metadata }))
      );
    } catch (err) {
      if (err && err.code === "ENOENT") return;
      throw err;
    }
  }

  async _persist() {
    const dir = path.dirname(this._filePath);
    await mkdir(dir, { recursive: true });

    const records = await super.list();
    const payload = JSON.stringify(
      {
        version: 1,
        dimension: this.dimension,
        records: records.map((r) => ({
          id: r.id,
          vector: Array.from(r.vector),
          metadata: r.metadata,
        })),
      },
      null,
      2
    );

    const tmpPath = `${this._filePath}.tmp`;
    await writeFile(tmpPath, payload, "utf8");
    await rename(tmpPath, this._filePath);
  }

  async upsert(records) {
    await this.load();
    await super.upsert(records);
    await this._persist();
  }

  async delete(ids) {
    await this.load();
    await super.delete(ids);
    await this._persist();
  }

  async get(id) {
    await this.load();
    return super.get(id);
  }

  async list(opts) {
    await this.load();
    return super.list(opts);
  }

  async query(vector, topK, opts) {
    await this.load();
    return super.query(vector, topK, opts);
  }

  async close() {}
}
