import { InMemoryBinaryStorage } from "./binaryStorage.js";
import { normalizeL2, toFloat32Array } from "./vectorMath.js";

function createAbortError(message = "Aborted") {
  const err = new Error(message);
  err.name = "AbortError";
  return err;
}

function throwIfAborted(signal) {
  if (signal?.aborted) throw createAbortError();
}

function locateSqlJsFile(file, prefix = "") {
  try {
    if (typeof import.meta.resolve === "function") {
      const resolved = import.meta.resolve(`sql.js/dist/${file}`);
      if (resolved) {
        if (resolved.startsWith("file://")) {
          let pathname = decodeURIComponent(new URL(resolved).pathname);
          if (/^\/[A-Za-z]:\//.test(pathname)) pathname = pathname.slice(1);
          return pathname;
        }
        return resolved;
      }
    }
  } catch {
    // ignore
  }
  // Emscripten calls locateFile(path, prefix). When we can't fully resolve,
  // preserve the default behaviour (prefix + path).
  return prefix ? `${prefix}${file}` : file;
}

const sqlJsPromises = new Map();
async function getSqlJs(locateFile) {
  const locator = locateFile ?? locateSqlJsFile;
  if (!sqlJsPromises.has(locator)) {
    sqlJsPromises.set(
      locator,
      import("sql.js").then((mod) => {
        const initSqlJs = mod.default ?? mod;
        return initSqlJs({ locateFile: locator });
      })
    );
  }
  return sqlJsPromises.get(locator);
}

function blobToFloat32(blob) {
  if (!(blob instanceof Uint8Array)) {
    throw new Error("Expected SQLite BLOB to be Uint8Array");
  }
  if (blob.byteLength % 4 !== 0) {
    throw new Error(`Invalid vector blob length: ${blob.byteLength}`);
  }
  return new Float32Array(blob.buffer, blob.byteOffset, blob.byteLength / 4);
}

function float32ToBlob(vec) {
  const v = vec instanceof Float32Array ? vec : Float32Array.from(vec);
  return new Uint8Array(v.buffer, v.byteOffset, v.byteLength);
}

/**
 * A persistent vector store backed by sql.js (SQLite compiled to WASM).
 *
 * Notes:
 * - Vectors are stored as Float32Array BLOBs (L2-normalized at write time).
 * - Query uses a custom SQLite scalar function `dot(vectorBlob, queryBlob)` and
 *   `ORDER BY dot(...) DESC LIMIT topK`.
 * - This is still an O(n) scan inside SQLite; it provides persistence and
 *   reasonable workbook-scale performance. If we later add a true ANN extension
 *   (e.g. sqlite-vss) the interface stays the same.
 */
export class SqliteVectorStore {
  /**
   * @param {any} db
   * @param {{ storage: any, dimension: number, autoSave: boolean }} opts
   */
  constructor(db, opts) {
    this._db = db;
    this._storage = opts.storage;
    this._dimension = opts.dimension;
    this._autoSave = opts.autoSave;

    this._ensureSchema();
    this._registerFunctions();
  }

  static async create(opts) {
    if (!opts || !Number.isFinite(opts.dimension) || opts.dimension <= 0) {
      throw new Error("SqliteVectorStore requires a positive dimension");
    }

    if (opts.filePath) {
      throw new Error(
        "SqliteVectorStore.create no longer accepts filePath. Pass { storage } instead (e.g. LocalStorageBinaryStorage / NodeFileBinaryStorage)."
      );
    }

    const storage = opts.storage ?? new InMemoryBinaryStorage();
    const SQL = await getSqlJs(opts.locateFile);
    const existing = await storage.load();

    const db = existing ? new SQL.Database(existing) : new SQL.Database();
    const store = new SqliteVectorStore(db, {
      storage,
      dimension: opts.dimension,
      autoSave: opts.autoSave ?? true,
    });

    store._ensureMeta(opts.dimension);
    return store;
  }

  get dimension() {
    return this._dimension;
  }

  _ensureSchema() {
    this._db.run(`
      CREATE TABLE IF NOT EXISTS vector_store_meta (
        key TEXT PRIMARY KEY,
        value TEXT NOT NULL
      );

      CREATE TABLE IF NOT EXISTS vectors (
        id TEXT PRIMARY KEY,
        workbook_id TEXT,
        vector BLOB NOT NULL,
        metadata_json TEXT NOT NULL
      );

      CREATE INDEX IF NOT EXISTS idx_vectors_workbook_id ON vectors(workbook_id);
    `);
  }

  /**
   * @param {number} dimension
   */
  _ensureMeta(dimension) {
    const stmt = this._db.prepare("SELECT value FROM vector_store_meta WHERE key = 'dimension' LIMIT 1;");
    const hasRow = stmt.step();
    const existing = hasRow ? stmt.get()[0] : null;
    stmt.free();

    if (existing == null) {
      const insert = this._db.prepare("INSERT INTO vector_store_meta (key, value) VALUES ('dimension', ?);");
      insert.run([String(dimension)]);
      insert.free();
      return;
    }

    const existingDim = Number(existing);
    if (existingDim !== dimension) {
      throw new Error(
        `SqliteVectorStore dimension mismatch: db has ${existingDim}, requested ${dimension}`
      );
    }
  }

  _registerFunctions() {
    // sql.js exposes create_function(name, fn)
    // eslint-disable-next-line no-underscore-dangle
    this._db.create_function("dot", (a, b) => {
      const va = blobToFloat32(a);
      const vb = blobToFloat32(b);
      const len = Math.min(va.length, vb.length);
      let dot = 0;
      for (let i = 0; i < len; i += 1) dot += va[i] * vb[i];
      return dot;
    });
  }

  async _persist() {
    const data = this._db.export();
    // sql.js drops custom scalar functions (like our `dot`) after `export()`. Re-register
    // them so subsequent queries can still prepare statements that reference `dot(...)`.
    this._registerFunctions();
    await this._storage.save(data);
  }

  /**
   * @param {{ id: string, vector: ArrayLike<number>, metadata: any }[]} records
   */
  async upsert(records) {
    if (!records.length) return;

    const stmt = this._db.prepare(`
      INSERT INTO vectors (id, workbook_id, vector, metadata_json)
      VALUES (?, ?, ?, ?)
      ON CONFLICT(id) DO UPDATE SET
        workbook_id = excluded.workbook_id,
        vector = excluded.vector,
        metadata_json = excluded.metadata_json;
    `);

    this._db.run("BEGIN;");
    try {
      for (const r of records) {
        const vec = toFloat32Array(r.vector);
        if (vec.length !== this._dimension) {
          throw new Error(
            `Vector dimension mismatch for id=${r.id}: expected ${this._dimension}, got ${vec.length}`
          );
        }
        const norm = normalizeL2(vec);
        const blob = float32ToBlob(norm);
        const meta = r.metadata ?? {};
        const workbookId = meta.workbookId ?? null;
        stmt.run([r.id, workbookId, blob, JSON.stringify(meta)]);
      }
      this._db.run("COMMIT;");
    } catch (err) {
      this._db.run("ROLLBACK;");
      throw err;
    } finally {
      stmt.free();
    }

    if (this._autoSave) await this._persist();
  }

  /**
   * @param {string[]} ids
   */
  async delete(ids) {
    if (!ids.length) return;
    const stmt = this._db.prepare("DELETE FROM vectors WHERE id = ?;");
    this._db.run("BEGIN;");
    try {
      for (const id of ids) stmt.run([id]);
      this._db.run("COMMIT;");
    } catch (err) {
      this._db.run("ROLLBACK;");
      throw err;
    } finally {
      stmt.free();
    }
    if (this._autoSave) await this._persist();
  }

  /**
   * @param {string} id
   */
  async get(id) {
    const stmt = this._db.prepare("SELECT vector, metadata_json FROM vectors WHERE id = ? LIMIT 1;");
    stmt.bind([id]);
    if (!stmt.step()) {
      stmt.free();
      return null;
    }
    const row = stmt.get();
    stmt.free();

    const vec = blobToFloat32(row[0]);
    const metadata = JSON.parse(row[1]);
    return { id, vector: new Float32Array(vec), metadata };
  }

  /**
   * @param {{
   *   filter?: (metadata: any, id: string) => boolean,
   *   workbookId?: string,
   *   includeVector?: boolean
   * }} [opts]
   */
  async list(opts) {
    const signal = opts?.signal;
    const filter = opts?.filter;
    const workbookId = opts?.workbookId;
    const includeVector = opts?.includeVector ?? true;

    throwIfAborted(signal);
    const sql = workbookId
      ? `SELECT id, ${includeVector ? "vector," : ""} metadata_json FROM vectors WHERE workbook_id = ?`
      : `SELECT id, ${includeVector ? "vector," : ""} metadata_json FROM vectors`;
    const stmt = this._db.prepare(sql);
    try {
      if (workbookId) stmt.bind([workbookId]);

      const out = [];
      while (true) {
        throwIfAborted(signal);
        if (!stmt.step()) break;
        const row = stmt.get();
        const id = row[0];
        const vecBlob = includeVector ? row[1] : null;
        const metaJson = includeVector ? row[2] : row[1];
        const metadata = JSON.parse(metaJson);
        if (filter && !filter(metadata, id)) continue;
        out.push({
          id,
          vector: includeVector ? new Float32Array(blobToFloat32(vecBlob)) : undefined,
          metadata,
        });
      }
      throwIfAborted(signal);
      return out;
    } finally {
      stmt.free();
    }
  }

  /**
   * @param {ArrayLike<number>} vector
   * @param {number} topK
   * @param {{ filter?: (metadata: any, id: string) => boolean, workbookId?: string }} [opts]
   */
  async query(vector, topK, opts) {
    const signal = opts?.signal;
    const filter = opts?.filter;
    const workbookId = opts?.workbookId;
    throwIfAborted(signal);
    if (topK <= 0) return [];
    const q = normalizeL2(vector);
    const qBlob = float32ToBlob(q);

    /**
     * When a JS-level filter is provided, we can't apply it inside SQLite. To match
     * InMemoryVectorStore semantics ("return up to topK *matching* records"), we
     * progressively over-fetch from SQLite and apply the filter as we iterate:
     *
     *   1) Start with LIMIT = max(topK * oversampleFactor, minLimit)
     *   2) If we still have < topK matches, double the LIMIT and retry
     *   3) Stop once we have topK matches or SQLite returns fewer rows than
     *      requested (no more candidates)
     *
     * This keeps the common case fast while still guaranteeing correctness.
     */
    const oversampleFactor = filter ? 4 : 1;
    const minLimit = filter ? 64 : topK;
    let limit = Math.max(topK * oversampleFactor, minLimit);

    const sql = workbookId
      ? `
        SELECT id, metadata_json, dot(vector, ?) AS score
        FROM vectors
        WHERE workbook_id = ?
        ORDER BY score DESC
        LIMIT ?;
      `
      : `
        SELECT id, metadata_json, dot(vector, ?) AS score
        FROM vectors
        ORDER BY score DESC
        LIMIT ?;
      `;

    while (true) {
      throwIfAborted(signal);
      const stmt = this._db.prepare(sql);
      /** @type {{ id: string, score: number, metadata: any }[]} */
      const out = [];
      let rows = 0;
      try {
        if (workbookId) stmt.bind([qBlob, workbookId, limit]);
        else stmt.bind([qBlob, limit]);

        while (true) {
          throwIfAborted(signal);
          if (!stmt.step()) break;
          rows += 1;
          const row = stmt.get();
          const id = row[0];
          const metadata = JSON.parse(row[1]);
          if (filter && !filter(metadata, id)) continue;
          const score = Number(row[2]);
          out.push({ id, score, metadata });
          if (out.length >= topK) break;
        }
        throwIfAborted(signal);
      } finally {
        stmt.free();
      }

      if (out.length >= topK) return out;
      // If SQLite returned fewer rows than we requested, we've exhausted the
      // candidate set, so return whatever matches we found.
      if (rows < limit) return out;

      // Still not enough matches; widen the window and retry.
      limit *= 2;
    }
  }

  async close() {
    await this._persist();
    this._db.close();
  }
}
