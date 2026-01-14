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
  // Typed array views in JS require byteOffset alignment for the element size.
  // If we receive an unaligned Uint8Array (possible when crossing WASM/JS boundaries),
  // copy into a fresh buffer so we can safely reinterpret as Float32Array.
  const bytes = blob.byteOffset % 4 === 0 ? blob : new Uint8Array(blob);
  return new Float32Array(bytes.buffer, bytes.byteOffset, bytes.byteLength / 4);
}

/**
 * Wrap blob decode errors with a more descriptive context so callers don't need to
 * guess which record/operation triggered the failure.
 *
 * @param {Uint8Array} blob
 * @param {string} context
 */
function blobToFloat32WithContext(blob, context) {
  try {
    return blobToFloat32(blob);
  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err);
    throw new Error(`${context}: ${msg}`);
  }
}

function createDimensionMismatchError(existingDim, requestedDim) {
  const err = new Error(
    `SqliteVectorStore dimension mismatch: db has ${existingDim}, requested ${requestedDim}`
  );
  err.name = "SqliteVectorStoreDimensionMismatchError";
  // Non-standard fields for programmatic handling/debugging.
  err.dbDimension = existingDim;
  err.requestedDimension = requestedDim;
  return err;
}

function createInvalidDimensionMetadataError(rawValue) {
  const err = new Error(
    `SqliteVectorStore invalid dimension metadata: expected positive number, got ${JSON.stringify(rawValue)}`
  );
  err.name = "SqliteVectorStoreInvalidMetadataError";
  // Non-standard field for programmatic handling/debugging.
  err.rawDimension = rawValue;
  return err;
}

/**
 * @param {Float32Array} vec
 * @param {number} expectedDim
 * @param {string} context
 */
function assertVectorDim(vec, expectedDim, context) {
  if (vec.length !== expectedDim) {
    throw new Error(`${context}: expected ${expectedDim}, got ${vec.length}`);
  }
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
  delete extra.metadataHash;
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
  const metadataHash = typeof meta.metadataHash === "string" ? meta.metadataHash : null;
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
    metadataHash,
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
    // Serialize `storage.save()` calls to prevent out-of-order async persists from
    // overwriting newer snapshots.
    /** @type {Promise<void>} */
    this._persistQueue = Promise.resolve();
    this._batchDepth = 0;

    this._ensureSchema();
    this._registerFunctions();
  }

  /**
   * @param {{
   *   storage?: any,
   *   dimension: number,
   *   autoSave?: boolean,
   *   resetOnCorrupt?: boolean,
   *   resetOnDimensionMismatch?: boolean,
   *   locateFile?: (file: string, prefix?: string) => string
   * }} opts
   * @param {boolean} [opts.resetOnCorrupt]
   *   When true (default), any failure to load/validate an existing persisted
   *   database causes the underlying storage to be cleared (via `storage.remove()`
   *   when available, otherwise by overwriting with an empty payload) and a fresh
   *   empty database to be created. When false, the initialization error is
   *   rethrown.
   * @param {boolean} [opts.resetOnDimensionMismatch]
   *   When true (default), if the persisted database's `dimension` metadata does
   *   not match the requested `opts.dimension`, the persisted bytes are wiped and
   *   an empty database is created so callers can re-index. When false, the
   *   mismatch is treated as fatal and the initialization error is rethrown.
   */
  static async create(opts) {
    if (!opts || !Number.isFinite(opts.dimension) || opts.dimension <= 0) {
      throw new Error("SqliteVectorStore requires a positive dimension");
    }

    if (opts.filePath) {
      throw new Error(
        "SqliteVectorStore.create no longer accepts filePath. Pass { storage } instead (e.g. IndexedDBBinaryStorage / ChunkedLocalStorageBinaryStorage / LocalStorageBinaryStorage / NodeFileBinaryStorage)."
      );
    }

    const storage = opts.storage ?? new InMemoryBinaryStorage();
    // The vector store is a derived cache (it can always be re-indexed). By
    // default we prefer a resilient startup experience, automatically clearing
    // invalid/incompatible persisted bytes and allowing callers to re-index.
    const resetOnCorrupt = opts.resetOnCorrupt ?? true;
    const resetOnDimensionMismatch = opts.resetOnDimensionMismatch ?? true;
    const SQL = await getSqlJs(opts.locateFile);

    async function clearPersistedBytes() {
      try {
        const remove = storage?.remove;
        if (typeof remove === "function") {
          await remove.call(storage);
          return;
        }
      } catch {
        // ignore
      }
      // Older BinaryStorage implementations might not provide `remove()`. Overwrite
      // with an empty payload to avoid retaining stale bytes.
      try {
        await storage.save(new Uint8Array());
      } catch {
        // ignore
      }
    }

    let existing = null;
    let loadFailed = false;
    try {
      existing = await storage.load();
    } catch (err) {
      if (!resetOnCorrupt) throw err;
      loadFailed = true;
      await clearPersistedBytes();
      existing = null;
    }
    const hadExisting = Boolean(existing) || loadFailed;

    /**
     * @param {any} db
     */
    function initStore(db) {
      const store = new SqliteVectorStore(db, {
        storage,
        dimension: opts.dimension,
        autoSave: opts.autoSave ?? true,
      });
      store._ensureMeta(opts.dimension);
      store._ensureSchemaVersion();
      store._ensureIndexes();
      return store;
    }

    /** @type {SqliteVectorStore} */
    let store;
    if (!existing) {
      store = initStore(new SQL.Database());
      if (!hadExisting) {
        // Initializing a brand new in-memory DB will insert meta rows (dimension,
        // schema version) and mark the store dirty. However, when there is no
        // persisted payload yet, we don't want `close()` to immediately write an
        // empty SQLite file to storage — persist only after the first real
        // mutation (upsert/delete/etc) triggers `_dirty`.
        store._dirty = false;
      } else {
        // Loading the persisted payload threw (e.g. a corrupted base64 decode in a
        // BinaryStorage implementation). Best-effort persist an empty DB so future
        // opens don't repeatedly hit the load error even when `remove()` isn't
        // supported.
        try {
          store._dirty = true;
          await store._persist();
        } catch {
          // ignore
        }
      }
      return store;
    }

    let db = null;
    try {
      db = new SQL.Database(existing);
      store = initStore(db);
    } catch (err) {
      try {
        db?.close?.();
      } catch {
        // ignore
      }

      const message = err instanceof Error ? err.message : String(err);
      const isDimMismatch =
        (err && typeof err === "object" && err.name === "SqliteVectorStoreDimensionMismatchError") ||
        message.includes("SqliteVectorStore dimension mismatch");
      if (isDimMismatch) {
        if (!resetOnDimensionMismatch) throw err;
      } else {
        if (!resetOnCorrupt) throw err;
      }

      await clearPersistedBytes();
      store = initStore(new SQL.Database());

      // Best-effort persist the empty DB so future opens don't repeatedly hit the
      // mismatch/corruption path even if `remove()` isn't supported.
      try {
        store._dirty = true;
        await store._persist();
      } catch {
        // ignore
      }
    }

    // If we migrated an existing database and autoSave is enabled, persist the
    // migrated schema immediately so subsequent opens don't re-run the migration
    // work (and so incremental indexing benefits right away).
    if (store._autoSave && hadExisting) await store._persist();
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
        metadata_hash TEXT,
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
      this._dirty = true;
      return;
    }

    const existingDim = Number(existing);
    if (!Number.isFinite(existingDim) || existingDim <= 0) {
      throw createInvalidDimensionMetadataError(existing);
    }
    if (existingDim !== dimension) {
      throw createDimensionMismatchError(existingDim, dimension);
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
    this._dirty = true;
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

  _getTableIndexes(table) {
    const stmt = this._db.prepare(`PRAGMA index_list(${table});`);
    /** @type {Set<string>} */
    const indexes = new Set();
    while (stmt.step()) {
      const row = stmt.get();
      // PRAGMA index_list: [seq, name, unique, origin, partial]
      indexes.add(String(row[1]));
    }
    stmt.free();
    return indexes;
  }

  _getIndexSql(name) {
    const stmt = this._db.prepare(
      "SELECT sql FROM sqlite_master WHERE type = 'index' AND name = ? LIMIT 1;"
    );
    stmt.bind([name]);
    const hasRow = stmt.step();
    const sql = hasRow ? stmt.get()[0] : null;
    stmt.free();
    return sql;
  }

  _ensureIndexes() {
    const indexes = this._getTableIndexes("vectors");
    // Covering index for incremental indexing state scans (id + hashes) without
    // touching the main table payload (vector/text/metadata_json).
    const indexName = "idx_vectors_workbook_hashes";
    // Include `length(vector)` as an index expression so `listContentHashes()` can
    // validate vector byte lengths without needing to read the full table payload.
    const desiredSql =
      "CREATE INDEX IF NOT EXISTS idx_vectors_workbook_hashes ON vectors(workbook_id, id, content_hash, metadata_hash, length(vector));";

    if (indexes.has(indexName)) {
      const existingSql = this._getIndexSql(indexName);
      const normalized = typeof existingSql === "string" ? existingSql.toLowerCase() : "";
      if (normalized.includes("length(vector)")) return;

      // Index exists but with an older definition; replace it.
      this._db.run(`DROP INDEX IF EXISTS ${indexName};`);
      this._db.run(desiredSql);
      this._dirty = true;
      return;
    }

    this._db.run(desiredSql);
    this._dirty = true;
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
      ["metadata_hash", "TEXT"],
      ["token_count", "INTEGER"],
      ["text", "TEXT"],
    ];

    let altered = false;
    for (const [name, type] of desired) {
      if (cols.has(name)) continue;
      // SQLite has no "ADD COLUMN IF NOT EXISTS" in older versions; probe via
      // PRAGMA table_info to keep migrations idempotent.
      this._db.run(`ALTER TABLE vectors ADD COLUMN ${name} ${type};`);
      altered = true;
    }
    if (altered) this._dirty = true;
  }

  /**
   * @param {{ preferExistingColumns?: boolean }} [opts]
   */
  _migrateVectorsToV2(opts) {
    const preferExistingColumns = Boolean(opts?.preferExistingColumns);
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
        metadata_hash,
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
        metadata_hash = ?,
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
        const metadataHashCol = row[10];
        const tokenCountCol = row[11];
        const textCol = row[12];
        const metaJson = row[13];

        const parsedMeta = metaJson ? JSON.parse(metaJson) : {};
        const meta = parsedMeta && typeof parsedMeta === "object" ? parsedMeta : {};

        const rect = meta?.rect ?? null;

        // When migrating v1 DBs, `metadata_json` is the only source of truth. For
        // repairing schema_version=2 DBs that are merely missing newer columns
        // (e.g. `metadata_hash`), prefer existing structured columns so stale
        // metadata_json (if present) does not overwrite newer column values.
        const stringFrom = (col, metaVal) => {
          if (preferExistingColumns) {
            return typeof col === "string" ? col : typeof metaVal === "string" ? metaVal : null;
          }
          return typeof metaVal === "string" ? metaVal : typeof col === "string" ? col : null;
        };
        const numberFrom = (col, metaVal) => {
          if (preferExistingColumns) {
            return Number.isFinite(col) ? col : Number.isFinite(metaVal) ? metaVal : null;
          }
          return Number.isFinite(metaVal) ? metaVal : Number.isFinite(col) ? col : null;
        };

        const workbookId =
          typeof workbookIdCol === "string"
            ? workbookIdCol
            : typeof meta.workbookId === "string"
              ? meta.workbookId
              : null;
        const sheetName = stringFrom(sheetNameCol, meta.sheetName);
        const kind = stringFrom(kindCol, meta.kind);
        const title = stringFrom(titleCol, meta.title);

        const r0 = numberFrom(r0Col, rect?.r0);
        const c0 = numberFrom(c0Col, rect?.c0);
        const r1 = numberFrom(r1Col, rect?.r1);
        const c1 = numberFrom(c1Col, rect?.c1);

        const contentHash = stringFrom(contentHashCol, meta.contentHash);
        const metadataHash = stringFrom(metadataHashCol, meta.metadataHash);
        const tokenCount = numberFrom(tokenCountCol, meta.tokenCount);
        const text = stringFrom(textCol, meta.text);

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
        delete extra.metadataHash;
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
          metadataHash,
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
    // Even for schema_version >= 2, older DBs may be missing newly-added v2 columns.
    // Check for required columns so migrations remain forward-compatible within the
    // same schema_version (and so SQL statements selecting those columns don't fail).
    const cols = this._getTableColumns("vectors");
    const required = [
      "sheet_name",
      "kind",
      "title",
      "r0",
      "c0",
      "r1",
      "c1",
      "content_hash",
      "metadata_hash",
      "token_count",
      "text",
    ];
    const missingRequiredColumn = required.some((c) => !cols.has(c));
    if (Number.isFinite(current) && current >= SCHEMA_VERSION && !missingRequiredColumn) return;

    this._db.run("BEGIN;");
    try {
      // v1 -> v2: add structured metadata columns and shrink metadata_json.
      this._migrateVectorsToV2({ preferExistingColumns: Number.isFinite(current) && current >= SCHEMA_VERSION });

      this._setMetaValue("schema_version", String(SCHEMA_VERSION));
      this._db.run("COMMIT;");
      // Migration changes must be persisted even if the caller performs no other
      // writes (e.g. incremental indexing run with 0 changes).
      this._dirty = true;
    } catch (err) {
      this._db.run("ROLLBACK;");
      throw err;
    }
  }

  _registerFunctions() {
    // sql.js exposes create_function(name, fn)
    // eslint-disable-next-line no-underscore-dangle
    this._db.create_function("dot", (a, b) => {
      // NOTE: When a user-defined function throws inside sql.js, the error message
      // is only preserved if we throw a *string* (sql.js forwards it to
      // `sqlite3_result_error`). Throwing an Error object results in a generic
      // "Error" with an empty message.
      //
      // This is important because we rely on descriptive dot() errors to surface
      // invalid/corrupt vector blobs in both tests and production debugging.
      const decode = (blob, context) => {
        try {
          return blobToFloat32(blob);
        } catch (err) {
          const msg = err instanceof Error ? err.message : String(err);
          throw `${context}: ${msg}`;
        }
      };

      const va = decode(a, "SqliteVectorStore dot() failed to decode arg0 vector blob");
      const vb = decode(b, "SqliteVectorStore dot() failed to decode arg1 vector blob");
      if (va.length !== vb.length) {
        throw `SqliteVectorStore dot() dimension mismatch: expected ${this._dimension}, got ${va.length} vs ${vb.length}`;
      }
      if (va.length !== this._dimension) {
        throw `SqliteVectorStore dot() dimension mismatch: expected ${this._dimension}, got ${va.length}`;
      }
      let dot = 0;
      for (let i = 0; i < va.length; i += 1) dot += va[i] * vb[i];
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

  async _enqueuePersist() {
    const task = () => this._persist();
    // Ensure the queue keeps flowing even if a previous persist failed.
    const next = this._persistQueue.then(task, task);
    this._persistQueue = next;
    return next;
  }

  /**
   * Batch multiple mutations into a single persistence snapshot.
   *
   * When autoSave is enabled, `upsert()`/`delete()` normally persist after each call.
   * `batch()` temporarily suppresses those intermediate saves, then persists once at
   * the end if anything changed.
   *
   * @template T
   * @param {() => Promise<T> | T} fn
   * @returns {Promise<T>}
   */
  async batch(fn) {
    const isOutermost = this._batchDepth === 0;
    const prevAutoSave = isOutermost ? this._autoSave : null;
    if (isOutermost) this._autoSave = false;
    this._batchDepth += 1;
    /** @type {any} */
    let result;
    try {
      result = await fn();
    } finally {
      this._batchDepth -= 1;
      if (isOutermost) this._autoSave = prevAutoSave;
    }

    // Only persist for successful (non-throwing) batches, and only when autoSave was
    // enabled at the start of the batch.
    if (isOutermost && prevAutoSave && this._dirty) await this._enqueuePersist();
    return result;
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
        metadata_hash,
        token_count,
        text,
        metadata_json
      )
      VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
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
        metadata_hash = excluded.metadata_hash,
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
          meta.metadataHash,
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

    if (this._autoSave) await this._enqueuePersist();
  }

  /**
   * Update stored metadata without touching vectors.
   *
   * @param {{ id: string, metadata: any }[]} records
   */
  async updateMetadata(records) {
    if (!records.length) return;

    const stmt = this._db.prepare(`
      UPDATE vectors
      SET
        workbook_id = ?,
        sheet_name = ?,
        kind = ?,
        title = ?,
        r0 = ?,
        c0 = ?,
        r1 = ?,
        c1 = ?,
        content_hash = ?,
        metadata_hash = ?,
        token_count = ?,
        text = ?,
        metadata_json = ?
      WHERE id = ?;
    `);

    this._db.run("BEGIN;");
    try {
      for (const r of records) {
        const meta = splitMetadata(r.metadata);
        stmt.run([
          meta.workbookId,
          meta.sheetName,
          meta.kind,
          meta.title,
          meta.r0,
          meta.c0,
          meta.r1,
          meta.c1,
          meta.contentHash,
          meta.metadataHash,
          meta.tokenCount,
          meta.text,
          meta.metadataJson,
          r.id,
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

    if (this._autoSave) await this._enqueuePersist();
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
    if (this._autoSave) await this._enqueuePersist();
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
   * - This always persists the compacted DB via the configured BinaryStorage,
   *   regardless of the `autoSave` setting. Unlike incremental writes,
   *   compaction is an explicit, manual operation intended to reclaim persisted
   *   storage space.
   */
  async compact() {
    // If there is an in-flight persist, wait for it to complete before running
    // `VACUUM` to avoid mutating the DB while `export()` is executing.
    try {
      await this._persistQueue;
    } catch {
      // ignore - a failed persist shouldn't prevent compaction, and we re-enqueue
      // a fresh persist below.
    }
    // `VACUUM` cannot run inside a transaction. All of our public methods free
    // prepared statements and close transactions before returning, so we can
    // safely run it directly here.
    this._db.run("VACUUM;");
    this._dirty = true;
    // Persist the compacted DB image so external storage (LocalStorage/IndexedDB/file)
    // immediately reflects the reclaimed size.
    await this._enqueuePersist();
  }

  /**
   * Alias for {@link compact}. Kept for API discoverability for users familiar
   * with SQLite's `VACUUM` terminology.
   */
  async vacuum() {
    await this.compact();
  }

  /**
   * Delete all records associated with a workbook.
   *
   * @param {string} workbookId
   * @returns {Promise<number>} number of deleted records
   */
  async deleteWorkbook(workbookId) {
    const countStmt = this._db.prepare("SELECT COUNT(*) FROM vectors WHERE workbook_id = ?;");
    countStmt.bind([workbookId]);
    const deleted = countStmt.step() ? Number(countStmt.get()[0]) : 0;
    countStmt.free();
    if (deleted === 0) return 0;

    const stmt = this._db.prepare("DELETE FROM vectors WHERE workbook_id = ?;");
    this._db.run("BEGIN;");
    try {
      stmt.run([workbookId]);
      this._db.run("COMMIT;");
      this._dirty = true;
    } catch (err) {
      this._db.run("ROLLBACK;");
      throw err;
    } finally {
      stmt.free();
    }

    if (this._autoSave) await this._enqueuePersist();
    return deleted;
  }

  /**
   * Remove all records from the store.
   */
  async clear() {
    const stmt = this._db.prepare("DELETE FROM vectors;");
    this._db.run("BEGIN;");
    try {
      stmt.run();
      this._db.run("COMMIT;");
      this._dirty = true;
    } catch (err) {
      this._db.run("ROLLBACK;");
      throw err;
    } finally {
      stmt.free();
    }
    if (this._autoSave) await this._enqueuePersist();
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
        metadata_hash,
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

    const vec = blobToFloat32WithContext(
      row[0],
      `SqliteVectorStore.get(${JSON.stringify(id)}) failed to decode stored vector blob`
    );
    assertVectorDim(vec, this._dimension, `SqliteVectorStore.get(${JSON.stringify(id)}) vector dimension mismatch`);

    const base = {};
    if (row[1] != null) base.workbookId = row[1];
    if (row[2] != null) base.sheetName = row[2];
    if (row[3] != null) base.kind = row[3];
    if (row[4] != null) base.title = row[4];
    if (row[5] != null && row[6] != null && row[7] != null && row[8] != null) {
      base.rect = { r0: row[5], c0: row[6], r1: row[7], c1: row[8] };
    }
    if (row[9] != null) base.contentHash = row[9];
    if (row[10] != null) base.metadataHash = row[10];
    if (row[11] != null) base.tokenCount = row[11];
    if (row[12] != null) base.text = row[12];

    const extra = parseExtraMetadata(row[13]);
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
    const vectorSelect = includeVector ? "vector" : "length(vector) AS vector_bytes";
    const sql = workbookId
      ? `
        SELECT
          id,
          ${vectorSelect},
          workbook_id,
          sheet_name,
          kind,
          title,
          r0,
          c0,
          r1,
          c1,
          content_hash,
          metadata_hash,
          token_count,
          text,
          metadata_json
        FROM vectors
        WHERE workbook_id = ?
      `
      : `
        SELECT
          id,
          ${vectorSelect},
          workbook_id,
          sheet_name,
          kind,
          title,
          r0,
          c0,
          r1,
          c1,
          content_hash,
          metadata_hash,
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
        const offset = 2;
        const vecBlobOrBytes = row[1];

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
        if (row[offset + 9] != null) base.metadataHash = row[offset + 9];
        if (row[offset + 10] != null) base.tokenCount = row[offset + 10];
        if (row[offset + 11] != null) base.text = row[offset + 11];

        const extra = parseExtraMetadata(row[offset + 12]);
        const metadata = { ...extra, ...base };
        if (filter && !filter(metadata, id)) continue;
        let vecOut;
        if (includeVector) {
          const vec = blobToFloat32WithContext(
            vecBlobOrBytes,
            `SqliteVectorStore.list() failed to decode stored vector blob for id=${JSON.stringify(id)}`
          );
          assertVectorDim(
            vec,
            this._dimension,
            `SqliteVectorStore.list() vector dimension mismatch for id=${JSON.stringify(id)}`
          );
          vecOut = new Float32Array(vec);
        } else {
          const bytes = Number(vecBlobOrBytes);
          if (!Number.isFinite(bytes)) {
            throw new Error(
              `SqliteVectorStore.list() failed to read vector blob byte length for id=${JSON.stringify(id)}: got ${String(
                vecBlobOrBytes
              )}`
            );
          }
          if (bytes % 4 !== 0) {
            throw new Error(
              `SqliteVectorStore.list() invalid vector blob length for id=${JSON.stringify(id)}: ${bytes}`
            );
          }
          const expectedBytes = this._dimension * 4;
          if (bytes !== expectedBytes) {
            throw new Error(
              `SqliteVectorStore.list() vector blob byte length mismatch for id=${JSON.stringify(id)}: expected ${expectedBytes}, got ${bytes}`
            );
          }
        }
        out.push({
          id,
          vector: vecOut,
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
   * Return `{ id, contentHash, metadataHash }` for records, avoiding `metadata_json` parsing.
   *
   * This is primarily used by incremental indexing, which only needs the
   * `contentHash` and `metadataHash` to determine whether a chunk or its metadata changed.
   *
   * @param {{ workbookId?: string, signal?: AbortSignal }} [opts]
   * @returns {Promise<Array<{ id: string, contentHash: string | null, metadataHash: string | null }>>}
   */
  async listContentHashes(opts) {
    const signal = opts?.signal;
    const workbookId = opts?.workbookId;
    throwIfAborted(signal);
    const sql = workbookId
      ? "SELECT id, content_hash, metadata_hash, length(vector) AS vector_bytes FROM vectors WHERE workbook_id = ?"
      : "SELECT id, content_hash, metadata_hash, length(vector) AS vector_bytes FROM vectors";
    const stmt = this._db.prepare(sql);
    try {
      if (workbookId) stmt.bind([workbookId]);
      /** @type {Array<{ id: string, contentHash: string | null, metadataHash: string | null }>} */
      const out = [];
      while (true) {
        throwIfAborted(signal);
        if (!stmt.step()) break;
        const row = stmt.get();
        const id = row[0];
        const bytes = Number(row[3]);
        if (!Number.isFinite(bytes)) {
          throw new Error(
            `SqliteVectorStore.listContentHashes() failed to read vector blob byte length for id=${JSON.stringify(id)}: got ${String(
              row[3]
            )}`
          );
        }
        if (bytes % 4 !== 0) {
          throw new Error(
            `SqliteVectorStore.listContentHashes() invalid vector blob length for id=${JSON.stringify(id)}: ${bytes}`
          );
        }
        const expectedBytes = this._dimension * 4;
        if (bytes !== expectedBytes) {
          throw new Error(
            `SqliteVectorStore.listContentHashes() vector blob byte length mismatch for id=${JSON.stringify(id)}: expected ${expectedBytes}, got ${bytes}`
          );
        }
        out.push({ id, contentHash: row[1] ?? null, metadataHash: row[2] ?? null });
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
    if (!Number.isFinite(topK)) {
      throw new Error(`Invalid topK: expected a finite number, got ${String(topK)}`);
    }
    // We deterministically floor floats (e.g. 1.9 -> 1) so SQLite never sees a
    // fractional LIMIT value, and so query semantics match InMemoryVectorStore.
    const k = Math.floor(topK);
    if (k <= 0) return [];
    const qVec = toFloat32Array(vector);
    assertVectorDim(qVec, this._dimension, "SqliteVectorStore.query() vector dimension mismatch");
    const q = normalizeL2(qVec);
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
    const minLimit = filter ? 64 : k;
    let limit = Math.max(k * oversampleFactor, minLimit);

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
          metadata_hash,
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
          metadata_hash,
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
      /** @type {any} */
      let sqlError = null;
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
          if (row[10] != null) base.metadataHash = row[10];
          if (row[11] != null) base.tokenCount = row[11];
          if (row[12] != null) base.text = row[12];

          const extra = parseExtraMetadata(row[13]);
          const metadata = { ...extra, ...base };
          if (filter && !filter(metadata, id)) continue;

          const score = Number(row[14]);
          out.push({ id, score, metadata });
          if (out.length >= k) break;
        }
        throwIfAborted(signal);
      } catch (err) {
        sqlError = err;
      } finally {
        stmt.free();
      }

      if (sqlError) {
        // sql.js can surface exceptions thrown inside user-defined functions (like our `dot`)
        // as a generic `Error` without preserving the original message. When this happens,
        // attempt to detect invalid vector blobs and throw a descriptive error instead.
        throwIfAborted(signal);

        const label = String(sqlError);
        if (label === "Error") {
          const expectedBytes = this._dimension * 4;

          // 1) Ensure stored vector byteLength is a multiple of 4 (Float32 alignment).
          const invalidByteLengthSql = workbookId
            ? "SELECT id, length(vector) FROM vectors WHERE workbook_id = ? AND (length(vector) % 4) != 0 LIMIT 1;"
            : "SELECT id, length(vector) FROM vectors WHERE (length(vector) % 4) != 0 LIMIT 1;";
          const invalidByteLengthStmt = this._db.prepare(invalidByteLengthSql);
          try {
            if (workbookId) invalidByteLengthStmt.bind([workbookId]);
            if (invalidByteLengthStmt.step()) {
              const row = invalidByteLengthStmt.get();
              throw new Error(
                `SqliteVectorStore dot() failed to decode arg0 vector blob: Invalid vector blob length: ${Number(row[1])} (id=${JSON.stringify(row[0])})`
              );
            }
          } finally {
            invalidByteLengthStmt.free();
          }

          // 2) Ensure stored vector length matches the configured embedding dimension.
          const invalidDimSql = workbookId
            ? "SELECT id, length(vector) FROM vectors WHERE workbook_id = ? AND length(vector) != ? LIMIT 1;"
            : "SELECT id, length(vector) FROM vectors WHERE length(vector) != ? LIMIT 1;";
          const invalidDimStmt = this._db.prepare(invalidDimSql);
          try {
            if (workbookId) invalidDimStmt.bind([workbookId, expectedBytes]);
            else invalidDimStmt.bind([expectedBytes]);
            if (invalidDimStmt.step()) {
              const row = invalidDimStmt.get();
              const gotDim = Math.floor(Number(row[1]) / 4);
              throw new Error(
                `SqliteVectorStore dot() dimension mismatch: expected ${this._dimension}, got ${gotDim} (id=${JSON.stringify(row[0])})`
              );
            }
          } finally {
            invalidDimStmt.free();
          }
        }

        throw sqlError;
      }

      if (out.length >= k) return out;
      // If SQLite returned fewer rows than we requested, we've exhausted the
      // candidate set, so return whatever matches we found.
      if (rows < limit) return out;

      // Still not enough matches; widen the window and retry.
      limit *= 2;
    }
  }

  async close() {
    await this._enqueuePersist();
    await this._persistQueue;
    this._db.close();
  }
}
