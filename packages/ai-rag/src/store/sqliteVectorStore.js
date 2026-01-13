import { InMemoryBinaryStorage } from "./binaryStorage.js";
import { normalizeL2, toFloat32Array } from "./vectorMath.js";
import { throwIfAborted } from "../utils/abort.js";

const SCHEMA_VERSION = 2;

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
 * Split incoming metadata into structured fields stored in columns vs extra keys
 * kept in metadata_json.
 *
 * @param {any} metadata
 */
function splitMetadata(metadata) {
  const meta = metadata ?? {};
  const rect = meta?.rect ?? null;

  /** @type {any} */
  const extra = { ...(meta ?? {}) };
  delete extra.workbookId;
  delete extra.sheetName;
  delete extra.kind;
  delete extra.title;
  delete extra.rect;
  delete extra.contentHash;
  delete extra.tokenCount;
  delete extra.text;

  const workbookId = typeof meta.workbookId === "string" ? meta.workbookId : null;
  const sheetName = typeof meta.sheetName === "string" ? meta.sheetName : null;
  const kind = typeof meta.kind === "string" ? meta.kind : null;
  const title = typeof meta.title === "string" ? meta.title : null;

  const r0 = Number.isFinite(rect?.r0) ? rect.r0 : null;
  const c0 = Number.isFinite(rect?.c0) ? rect.c0 : null;
  const r1 = Number.isFinite(rect?.r1) ? rect.r1 : null;
  const c1 = Number.isFinite(rect?.c1) ? rect.c1 : null;

  const contentHash = typeof meta.contentHash === "string" ? meta.contentHash : null;
  const tokenCount = Number.isFinite(meta.tokenCount) ? meta.tokenCount : null;
  const text = typeof meta.text === "string" ? meta.text : null;

  return {
    workbookId,
    sheetName,
    kind,
    title,
    r0,
    c0,
    r1,
    c1,
    contentHash,
    tokenCount,
    text,
    metadataJson: JSON.stringify(extra),
  };
}

/**
 * Parse extra metadata_json, treating "{}" as empty to avoid JSON.parse work in
 * the common case where there are no extra keys.
 *
 * @param {string} json
 */
function parseExtraMetadata(json) {
  if (!json || json === "{}") return {};
  const parsed = JSON.parse(json);
  return parsed && typeof parsed === "object" ? parsed : {};
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
    this._dirty = false;

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
    store._ensureSchemaVersion();
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
        sheet_name TEXT,
        kind TEXT,
        title TEXT,
        r0 INTEGER,
        c0 INTEGER,
        r1 INTEGER,
        c1 INTEGER,
        content_hash TEXT,
        token_count INTEGER,
        text TEXT,
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

  _getMetaValue(key) {
    const stmt = this._db.prepare("SELECT value FROM vector_store_meta WHERE key = ? LIMIT 1;");
    stmt.bind([key]);
    const hasRow = stmt.step();
    const value = hasRow ? stmt.get()[0] : null;
    stmt.free();
    return value;
  }

  _setMetaValue(key, value) {
    const stmt = this._db.prepare(`
      INSERT INTO vector_store_meta (key, value) VALUES (?, ?)
      ON CONFLICT(key) DO UPDATE SET value = excluded.value;
    `);
    stmt.run([key, value]);
    stmt.free();
  }

  _getTableColumns(table) {
    const stmt = this._db.prepare(`PRAGMA table_info(${table});`);
    /** @type {Set<string>} */
    const cols = new Set();
    while (stmt.step()) {
      const row = stmt.get();
      cols.add(String(row[1]));
    }
    stmt.free();
    return cols;
  }

  _ensureVectorsColumnsV2() {
    const cols = this._getTableColumns("vectors");
    /** @type {[string, string][]} */
    const desired = [
      ["sheet_name", "TEXT"],
      ["kind", "TEXT"],
      ["title", "TEXT"],
      ["r0", "INTEGER"],
      ["c0", "INTEGER"],
      ["r1", "INTEGER"],
      ["c1", "INTEGER"],
      ["content_hash", "TEXT"],
      ["token_count", "INTEGER"],
      ["text", "TEXT"],
    ];

    for (const [name, type] of desired) {
      if (cols.has(name)) continue;
      // SQLite has no "ADD COLUMN IF NOT EXISTS" in older versions; probe via
      // PRAGMA table_info to keep migrations idempotent.
      this._db.run(`ALTER TABLE vectors ADD COLUMN ${name} ${type};`);
    }
  }

  _migrateVectorsToV2() {
    this._ensureVectorsColumnsV2();

    const select = this._db.prepare(`
      SELECT
        id,
        workbook_id,
        sheet_name,
        kind,
        title,
        r0,
        c0,
        r1,
        c1,
        content_hash,
        token_count,
        text,
        metadata_json
      FROM vectors;
    `);

    const update = this._db.prepare(`
      UPDATE vectors SET
        workbook_id = ?,
        sheet_name = ?,
        kind = ?,
        title = ?,
        r0 = ?,
        c0 = ?,
        r1 = ?,
        c1 = ?,
        content_hash = ?,
        token_count = ?,
        text = ?,
        metadata_json = ?
      WHERE id = ?;
    `);

    try {
      while (select.step()) {
        const row = select.get();
        const id = row[0];
        const workbookIdCol = row[1];
        const sheetNameCol = row[2];
        const kindCol = row[3];
        const titleCol = row[4];
        const r0Col = row[5];
        const c0Col = row[6];
        const r1Col = row[7];
        const c1Col = row[8];
        const contentHashCol = row[9];
        const tokenCountCol = row[10];
        const textCol = row[11];
        const metaJson = row[12];

        const parsedMeta = metaJson ? JSON.parse(metaJson) : {};
        const meta = parsedMeta && typeof parsedMeta === "object" ? parsedMeta : {};

        const rect = meta?.rect ?? null;

        const workbookId =
          typeof workbookIdCol === "string"
            ? workbookIdCol
            : typeof meta.workbookId === "string"
              ? meta.workbookId
              : null;
        const sheetName =
          typeof meta.sheetName === "string"
            ? meta.sheetName
            : typeof sheetNameCol === "string"
              ? sheetNameCol
              : null;
        const kind =
          typeof meta.kind === "string" ? meta.kind : typeof kindCol === "string" ? kindCol : null;
        const title =
          typeof meta.title === "string"
            ? meta.title
            : typeof titleCol === "string"
              ? titleCol
              : null;

        const r0 = Number.isFinite(rect?.r0) ? rect.r0 : Number.isFinite(r0Col) ? r0Col : null;
        const c0 = Number.isFinite(rect?.c0) ? rect.c0 : Number.isFinite(c0Col) ? c0Col : null;
        const r1 = Number.isFinite(rect?.r1) ? rect.r1 : Number.isFinite(r1Col) ? r1Col : null;
        const c1 = Number.isFinite(rect?.c1) ? rect.c1 : Number.isFinite(c1Col) ? c1Col : null;

        const contentHash =
          typeof meta.contentHash === "string"
            ? meta.contentHash
            : typeof contentHashCol === "string"
              ? contentHashCol
              : null;
        const tokenCount =
          Number.isFinite(meta.tokenCount) ? meta.tokenCount : Number.isFinite(tokenCountCol) ? tokenCountCol : null;
        const text = typeof meta.text === "string" ? meta.text : typeof textCol === "string" ? textCol : null;

        // Rewrite metadata_json to include only non-standard keys.
        /** @type {any} */
        const extra = { ...(meta ?? {}) };
        delete extra.workbookId;
        delete extra.sheetName;
        delete extra.kind;
        delete extra.title;
        delete extra.rect;
        delete extra.contentHash;
        delete extra.tokenCount;
        delete extra.text;

        update.run([
          workbookId,
          sheetName,
          kind,
          title,
          r0,
          c0,
          r1,
          c1,
          contentHash,
          tokenCount,
          text,
          JSON.stringify(extra),
          id,
        ]);
      }
    } finally {
      select.free();
      update.free();
    }
  }

  _ensureSchemaVersion() {
    const raw = this._getMetaValue("schema_version");
    const current = raw == null ? 1 : Number(raw);
    if (Number.isFinite(current) && current >= SCHEMA_VERSION) return;

    this._db.run("BEGIN;");
    try {
      // v1 -> v2: add structured metadata columns and shrink metadata_json.
      this._migrateVectorsToV2();

      this._setMetaValue("schema_version", String(SCHEMA_VERSION));
      this._db.run("COMMIT;");
    } catch (err) {
      this._db.run("ROLLBACK;");
      throw err;
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
    if (!this._dirty) return;
    // Mark clean optimistically so concurrent writes during export/save can
    // re-dirty the store. If persistence fails, flip back to dirty.
    this._dirty = false;
    try {
      const data = this._db.export();
      // sql.js drops custom scalar functions (like our `dot`) after `export()`. Re-register
      // them so subsequent queries can still prepare statements that reference `dot(...)`.
      this._registerFunctions();
      await this._storage.save(data);
    } catch (err) {
      this._dirty = true;
      // In case `export()` dropped custom functions before throwing or before `save()` failed.
      try {
        this._registerFunctions();
      } catch {
        // ignore
      }
      throw err;
    }
  }

  /**
   * @param {{ id: string, vector: ArrayLike<number>, metadata: any }[]} records
   */
  async upsert(records) {
    if (!records.length) return;

    const stmt = this._db.prepare(`
      INSERT INTO vectors (
        id,
        workbook_id,
        vector,
        sheet_name,
        kind,
        title,
        r0,
        c0,
        r1,
        c1,
        content_hash,
        token_count,
        text,
        metadata_json
      )
      VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
      ON CONFLICT(id) DO UPDATE SET
        workbook_id = excluded.workbook_id,
        vector = excluded.vector,
        sheet_name = excluded.sheet_name,
        kind = excluded.kind,
        title = excluded.title,
        r0 = excluded.r0,
        c0 = excluded.c0,
        r1 = excluded.r1,
        c1 = excluded.c1,
        content_hash = excluded.content_hash,
        token_count = excluded.token_count,
        text = excluded.text,
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
        const meta = splitMetadata(r.metadata);
        stmt.run([
          r.id,
          meta.workbookId,
          blob,
          meta.sheetName,
          meta.kind,
          meta.title,
          meta.r0,
          meta.c0,
          meta.r1,
          meta.c1,
          meta.contentHash,
          meta.tokenCount,
          meta.text,
          meta.metadataJson,
        ]);
      }
      this._db.run("COMMIT;");
      this._dirty = true;
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
      this._dirty = true;
    } catch (err) {
      this._db.run("ROLLBACK;");
      throw err;
    } finally {
      stmt.free();
    }
    if (this._autoSave) await this._persist();
  }

  /**
   * Compact the underlying SQLite database to reclaim space after large deletions.
   *
   * SQLite does not release free pages back to the database file until `VACUUM`
   * is run (unless `auto_vacuum` is enabled). Since sql.js stores the database in
   * memory and persists via `db.export()`, calling `VACUUM` can substantially
   * reduce the exported/persisted byte size after deleting many rows.
   *
   * Note:
   * - When `autoSave` is enabled (default), this will persist the compacted DB
   *   via the configured BinaryStorage.
   * - When `autoSave` is disabled, the compaction will only be persisted on the
   *   next `close()` (consistent with `upsert()` / `delete()`).
   */
  async compact() {
    // `VACUUM` cannot run inside a transaction. All of our public methods free
    // prepared statements and close transactions before returning, so we can
    // safely run it directly here.
    this._db.run("VACUUM;");
    // Re-register custom SQL functions just in case SQLite/sql.js resets them.
    // (sql.js definitely drops them after `export()`; VACUUM may or may not.)
    this._registerFunctions();
    this._dirty = true;
    if (this._autoSave) await this._persist();
  }

  /**
   * Alias for {@link compact}. Kept for API discoverability for users familiar
   * with SQLite's `VACUUM` terminology.
   */
  async vacuum() {
    await this.compact();
  }

  /**
   * @param {string} id
   */
  async get(id) {
    const stmt = this._db.prepare(`
      SELECT
        vector,
        workbook_id,
        sheet_name,
        kind,
        title,
        r0,
        c0,
        r1,
        c1,
        content_hash,
        token_count,
        text,
        metadata_json
      FROM vectors
      WHERE id = ?
      LIMIT 1;
    `);
    stmt.bind([id]);
    if (!stmt.step()) {
      stmt.free();
      return null;
    }
    const row = stmt.get();
    stmt.free();

    const vec = blobToFloat32(row[0]);

    const base = {};
    if (row[1] != null) base.workbookId = row[1];
    if (row[2] != null) base.sheetName = row[2];
    if (row[3] != null) base.kind = row[3];
    if (row[4] != null) base.title = row[4];
    if (row[5] != null && row[6] != null && row[7] != null && row[8] != null) {
      base.rect = { r0: row[5], c0: row[6], r1: row[7], c1: row[8] };
    }
    if (row[9] != null) base.contentHash = row[9];
    if (row[10] != null) base.tokenCount = row[10];
    if (row[11] != null) base.text = row[11];

    const extra = parseExtraMetadata(row[12]);
    const metadata = { ...extra, ...base };
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
      ? `
        SELECT
          id,
          ${includeVector ? "vector," : ""}
          workbook_id,
          sheet_name,
          kind,
          title,
          r0,
          c0,
          r1,
          c1,
          content_hash,
          token_count,
          text,
          metadata_json
        FROM vectors
        WHERE workbook_id = ?
      `
      : `
        SELECT
          id,
          ${includeVector ? "vector," : ""}
          workbook_id,
          sheet_name,
          kind,
          title,
          r0,
          c0,
          r1,
          c1,
          content_hash,
          token_count,
          text,
          metadata_json
        FROM vectors
      `;

    const stmt = this._db.prepare(sql);
    try {
      if (workbookId) stmt.bind([workbookId]);

      const out = [];
      while (true) {
        throwIfAborted(signal);
        if (!stmt.step()) break;
        const row = stmt.get();
        const id = row[0];
        const offset = includeVector ? 2 : 1;
        const vecBlob = includeVector ? row[1] : null;

        const base = {};
        if (row[offset] != null) base.workbookId = row[offset];
        if (row[offset + 1] != null) base.sheetName = row[offset + 1];
        if (row[offset + 2] != null) base.kind = row[offset + 2];
        if (row[offset + 3] != null) base.title = row[offset + 3];
        if (
          row[offset + 4] != null &&
          row[offset + 5] != null &&
          row[offset + 6] != null &&
          row[offset + 7] != null
        ) {
          base.rect = {
            r0: row[offset + 4],
            c0: row[offset + 5],
            r1: row[offset + 6],
            c1: row[offset + 7],
          };
        }
        if (row[offset + 8] != null) base.contentHash = row[offset + 8];
        if (row[offset + 9] != null) base.tokenCount = row[offset + 9];
        if (row[offset + 10] != null) base.text = row[offset + 10];

        const extra = parseExtraMetadata(row[offset + 11]);
        const metadata = { ...extra, ...base };
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
        SELECT
          id,
          workbook_id,
          sheet_name,
          kind,
          title,
          r0,
          c0,
          r1,
          c1,
          content_hash,
          token_count,
          text,
          metadata_json,
          dot(vector, ?) AS score
        FROM vectors
        WHERE workbook_id = ?
        ORDER BY score DESC, id ASC
        LIMIT ?;
      `
      : `
        SELECT
          id,
          workbook_id,
          sheet_name,
          kind,
          title,
          r0,
          c0,
          r1,
          c1,
          content_hash,
          token_count,
          text,
          metadata_json,
          dot(vector, ?) AS score
        FROM vectors
        ORDER BY score DESC, id ASC
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

          const base = {};
          if (row[1] != null) base.workbookId = row[1];
          if (row[2] != null) base.sheetName = row[2];
          if (row[3] != null) base.kind = row[3];
          if (row[4] != null) base.title = row[4];
          if (row[5] != null && row[6] != null && row[7] != null && row[8] != null) {
            base.rect = { r0: row[5], c0: row[6], r1: row[7], c1: row[8] };
          }
          if (row[9] != null) base.contentHash = row[9];
          if (row[10] != null) base.tokenCount = row[10];
          if (row[11] != null) base.text = row[11];

          const extra = parseExtraMetadata(row[12]);
          const metadata = { ...extra, ...base };
          if (filter && !filter(metadata, id)) continue;

          const score = Number(row[13]);
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
    if (this._dirty) await this._persist();
    this._db.close();
  }
}
