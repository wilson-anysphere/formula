import initSqlJs from "sql.js";
import type { AIAuditEntry, AuditListFilters, TokenUsage, ToolCallLog, UserFeedback, AIVerificationResult } from "./types.ts";
import type { AIAuditStore } from "./store.ts";
import type { SqliteBinaryStorage } from "./storage.ts";
import { InMemoryBinaryStorage } from "./storage.ts";
import { stableStringify } from "./stable-json.ts";

type SqlJsDatabase = any;

export interface SqliteAIAuditStoreRetention {
  /**
   * Maximum number of rows to keep (newest retained). If unset, rows are unbounded.
   */
  max_entries?: number;
  /**
   * Maximum age in milliseconds. Rows older than (now - max_age_ms) are deleted
   * at write-time. If unset, age-based retention is disabled.
   */
  max_age_ms?: number;
}

export interface SqliteAIAuditStoreOptions {
  storage?: SqliteBinaryStorage;
  locateFile?: (file: string, prefix?: string) => string;
  retention?: SqliteAIAuditStoreRetention;
  /**
   * When true (default), changes are persisted automatically after `logEntry()`.
   *
   * Set to false to buffer writes in-memory and call `flush()` explicitly.
   */
  auto_persist?: boolean;
  /**
   * Optional debounce interval for automatic persistence.
   *
   * When set (and `auto_persist` is true), multiple `logEntry()` calls within the
   * interval will result in a single persistence write.
   */
  auto_persist_interval_ms?: number;
}

/**
 * SQL.js-backed audit store (SQLite in WASM).
 *
 * Note: this store supports retention via row count / max age, but does not
 * enforce a strict per-entry size cap. When persisting via LocalStorage (e.g.
 * `LocalStorageBinaryStorage`), large single entries can still cause quota
 * failures. For defense-in-depth against oversized entries, wrap the store with
 * `BoundedAIAuditStore`.
 */
export class SqliteAIAuditStore implements AIAuditStore {
  private readonly db: SqlJsDatabase;
  private readonly storage: SqliteBinaryStorage;
  private readonly retention: SqliteAIAuditStoreRetention;
  private readonly autoPersist: boolean;
  private readonly autoPersistIntervalMs: number | undefined;
  private schemaDirty: boolean = false;

  private persistTimer: ReturnType<typeof setTimeout> | null = null;
  private persistSerial: Promise<void> = Promise.resolve();
  private writeRevision = 0;
  private persistedRevision = 0;
  private closed = false;

  private constructor(
    db: SqlJsDatabase,
    storage: SqliteBinaryStorage,
    retention: SqliteAIAuditStoreRetention,
    autoPersist: boolean,
    autoPersistIntervalMs: number | undefined,
  ) {
    this.db = db;
    this.storage = storage;
    this.retention = retention;
    this.autoPersist = autoPersist;
    this.autoPersistIntervalMs = autoPersistIntervalMs;
    this.ensureSchema();
  }

  static async create(options: SqliteAIAuditStoreOptions = {}): Promise<SqliteAIAuditStore> {
    const storage = options.storage ?? new InMemoryBinaryStorage();
    const SQL = await initSqlJs({ locateFile: options.locateFile ?? locateSqlJsFile });
    const existing = await storage.load();
    const db = existing ? new SQL.Database(existing) : new SQL.Database();
    const autoPersist = options.auto_persist ?? true;
    const autoPersistIntervalMs = options.auto_persist_interval_ms;
    const store = new SqliteAIAuditStore(db, storage, options.retention ?? {}, autoPersist, autoPersistIntervalMs);
    // Persist schema migrations/backfills eagerly for existing databases so that
    // upgraded clients won't need to redo work on every load.
    if (existing && store.schemaDirty) {
      try {
        await store.persistOnce();
      } catch {
        // Persistence is best-effort; if it fails (e.g. read-only storage), the
        // in-memory migration/backfill still allows workbook filtering to work
        // for the current session.
      }
      store.schemaDirty = false;
    }
    return store;
  }

  async logEntry(entry: AIAuditEntry): Promise<void> {
    this.assertOpen();
    const tokenUsage = normalizeTokenUsage(entry.token_usage);
    const workbookIdRaw =
      entry.workbook_id ?? extractWorkbookIdFromInput(entry.input) ?? extractWorkbookIdFromSessionId(entry.session_id);
    const workbookId = typeof workbookIdRaw === "string" ? workbookIdRaw.trim() : "";
    const stmt = this.db.prepare(
      `INSERT INTO ai_audit_log (
        id,
        timestamp_ms,
        session_id,
        workbook_id,
        user_id,
        mode,
        input_json,
        model,
        prompt_tokens,
        completion_tokens,
        total_tokens,
        latency_ms,
        tool_calls_json,
        verification_json,
        user_feedback
      ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?);`
    );

    stmt.run([
      entry.id,
      entry.timestamp_ms,
      entry.session_id,
      workbookId || null,
      entry.user_id ?? null,
      entry.mode,
      stableStringify(entry.input ?? null),
      entry.model,
      tokenUsage?.prompt_tokens ?? null,
      tokenUsage?.completion_tokens ?? null,
      tokenUsage?.total_tokens ?? null,
      entry.latency_ms ?? null,
      stableStringify(entry.tool_calls ?? []),
      entry.verification ? stableStringify(entry.verification) : null,
      entry.user_feedback ?? null
    ]);
    stmt.free();

    this.enforceRetention();
    this.writeRevision++;

    if (!this.autoPersist) return;

    const interval = this.autoPersistIntervalMs;
    if (typeof interval === "number" && Number.isFinite(interval) && interval > 0) {
      this.schedulePersist(interval);
      return;
    }

    await this.flush();
  }

  async listEntries(filters: AuditListFilters = {}): Promise<AIAuditEntry[]> {
    this.assertOpen();
    const params: any[] = [];
    let sql = "SELECT * FROM ai_audit_log";
    const where: string[] = [];
    const after_timestamp_ms =
      typeof filters.after_timestamp_ms === "number" && Number.isFinite(filters.after_timestamp_ms)
        ? filters.after_timestamp_ms
        : undefined;
    const before_timestamp_ms =
      typeof filters.before_timestamp_ms === "number" && Number.isFinite(filters.before_timestamp_ms)
        ? filters.before_timestamp_ms
        : undefined;
    const cursor =
      filters.cursor && typeof filters.cursor.before_timestamp_ms === "number" && Number.isFinite(filters.cursor.before_timestamp_ms)
        ? filters.cursor
        : undefined;
    const limit = typeof filters.limit === "number" && Number.isFinite(filters.limit) ? Math.max(0, Math.trunc(filters.limit)) : undefined;
    if (filters.session_id) {
      where.push("session_id = ?");
      params.push(filters.session_id);
    }
    const workbookIdFilter = typeof filters.workbook_id === "string" ? filters.workbook_id.trim() : "";
    if (workbookIdFilter) {
      where.push("workbook_id = ?");
      params.push(workbookIdFilter);
    }
    if (filters.mode) {
      const modes = Array.isArray(filters.mode) ? filters.mode : [filters.mode];
      if (modes.length === 1) {
        where.push("mode = ?");
        params.push(modes[0]);
      } else if (modes.length > 1) {
        where.push(`mode IN (${modes.map(() => "?").join(", ")})`);
        params.push(...modes);
      }
    }
    if (typeof after_timestamp_ms === "number") {
      where.push("timestamp_ms >= ?");
      params.push(after_timestamp_ms);
    }
    if (typeof before_timestamp_ms === "number") {
      where.push("timestamp_ms < ?");
      params.push(before_timestamp_ms);
    }
    if (cursor) {
      const beforeId = typeof cursor.before_id === "string" ? cursor.before_id : undefined;
      if (beforeId) {
        where.push("(timestamp_ms < ? OR (timestamp_ms = ? AND id < ?))");
        params.push(cursor.before_timestamp_ms, cursor.before_timestamp_ms, beforeId);
      } else {
        where.push("timestamp_ms < ?");
        params.push(cursor.before_timestamp_ms);
      }
    }
    if (where.length > 0) {
      sql += ` WHERE ${where.join(" AND ")}`;
    }
    sql += " ORDER BY timestamp_ms DESC, id DESC";
    if (typeof limit === "number") {
      sql += " LIMIT ?";
      params.push(limit);
    }

    const stmt = this.db.prepare(sql);
    stmt.bind(params);

    const rows: AIAuditEntry[] = [];
    while (stmt.step()) {
      const row = stmt.getAsObject() as any;
      rows.push(deserializeRow(row));
    }
    stmt.free();
    return rows;
  }

  /**
   * Forces persistence of the in-memory database to the underlying
   * {@link SqliteBinaryStorage}.
   *
   * When `auto_persist=false`, callers must invoke this periodically (or via
   * {@link close}) to durably save audit entries.
   */
  async flush(): Promise<void> {
    this.assertOpen();
    this.clearPersistTimer();

    const targetRevision = this.writeRevision;
    // If a persistence operation is already queued/running, wait for it to
    // complete before deciding whether we still need to write.
    await this.persistSerial;

    if (this.persistedRevision >= targetRevision) return;
    await this.enqueuePersist();
  }

  /**
   * Flushes any pending persistence and releases the sql.js database.
   */
  async close(): Promise<void> {
    if (this.closed) return;
    try {
      await this.flush();
    } finally {
      this.clearPersistTimer();
      try {
        if (this.db && typeof this.db.close === "function") {
          this.db.close();
        }
      } catch {
        // Best-effort close: sql.js may throw if the db is already closed.
      }
      this.closed = true;
    }
  }

  private ensureSchema(): void {
    this.db.run(`
      CREATE TABLE IF NOT EXISTS ai_audit_log (
        id TEXT PRIMARY KEY,
        timestamp_ms INTEGER NOT NULL,
        session_id TEXT NOT NULL,
        workbook_id TEXT,
        user_id TEXT,
        mode TEXT NOT NULL,
        input_json TEXT NOT NULL,
        model TEXT NOT NULL,
        prompt_tokens INTEGER,
        completion_tokens INTEGER,
        total_tokens INTEGER,
        latency_ms INTEGER,
        tool_calls_json TEXT NOT NULL,
        verification_json TEXT,
        user_feedback TEXT
      );
    `);

    // Migrate databases created before verification was tracked.
    if (ensureColumnExists(this.db, "ai_audit_log", "verification_json", "TEXT")) {
      this.schemaDirty = true;
    }
    // Migrate databases created before workbook metadata was tracked.
    if (ensureColumnExists(this.db, "ai_audit_log", "workbook_id", "TEXT")) {
      this.schemaDirty = true;
    }

    if (backfillLegacyWorkbookIds(this.db)) {
      this.schemaDirty = true;
    }

    // Indexes should be created after migrations (so older databases missing columns
    // don't error when we try to index them).
    this.db.run(`
      CREATE INDEX IF NOT EXISTS idx_ai_audit_log_session ON ai_audit_log(session_id);
      CREATE INDEX IF NOT EXISTS idx_ai_audit_log_workbook ON ai_audit_log(workbook_id);
      CREATE INDEX IF NOT EXISTS idx_ai_audit_log_mode ON ai_audit_log(mode);
      CREATE INDEX IF NOT EXISTS idx_ai_audit_log_timestamp ON ai_audit_log(timestamp_ms);
      CREATE INDEX IF NOT EXISTS idx_ai_audit_log_timestamp_id ON ai_audit_log(timestamp_ms, id);
    `);
  }

  private enforceRetention(): void {
    const maxAgeMs = this.retention.max_age_ms;
    if (typeof maxAgeMs === "number" && Number.isFinite(maxAgeMs) && maxAgeMs > 0) {
      const cutoff = Date.now() - maxAgeMs;
      const stmt = this.db.prepare("DELETE FROM ai_audit_log WHERE timestamp_ms < ?;");
      stmt.run([cutoff]);
      stmt.free();
    }

    const maxEntries = this.retention.max_entries;
    if (typeof maxEntries === "number" && Number.isFinite(maxEntries) && maxEntries > 0) {
      const stmt = this.db.prepare(`
        DELETE FROM ai_audit_log
        WHERE id IN (
          SELECT id FROM ai_audit_log
          ORDER BY timestamp_ms DESC, id DESC
          LIMIT -1 OFFSET ?
        );
      `);
      stmt.run([Math.floor(maxEntries)]);
      stmt.free();
    }
  }

  private schedulePersist(intervalMs: number): void {
    this.clearPersistTimer();
    this.persistTimer = setTimeout(() => {
      this.persistTimer = null;
      // Auto-persistence is best-effort; callers can use `flush()`/`close()` if
      // they need error visibility.
      void this.flush().catch(() => {});
    }, intervalMs);
  }

  private clearPersistTimer(): void {
    if (!this.persistTimer) return;
    clearTimeout(this.persistTimer as any);
    this.persistTimer = null;
  }

  private enqueuePersist(): Promise<void> {
    const run = async () => {
      await this.persistOnce();
    };

    const next = this.persistSerial.then(run, run);
    // Keep the serialization chain alive regardless of errors, while still
    // allowing callers awaiting `next` to observe failures.
    this.persistSerial = next.catch(() => {});
    return next;
  }

  private async persistOnce(): Promise<void> {
    const revisionAtStart = this.writeRevision;
    const data = this.db.export() as Uint8Array;
    await this.storage.save(data);
    // The export() snapshot corresponds to the state at the time persistOnce()
    // began; if new writes land while storage.save() is inflight, a subsequent
    // persist will still be required.
    this.persistedRevision = Math.max(this.persistedRevision, revisionAtStart);
  }

  private assertOpen(): void {
    if (!this.closed) return;
    throw new Error("SqliteAIAuditStore is closed");
  }
}

function locateSqlJsFile(file: string, prefix: string = ""): string {
  try {
    if (typeof import.meta.resolve === "function") {
      const resolved = import.meta.resolve(`sql.js/dist/${file}`);
      if (resolved) return resolved;
    }
  } catch {
    // ignore
  }

  // Emscripten calls locateFile(path, prefix). When we can't fully resolve,
  // preserve the default behaviour (prefix + path).
  return prefix ? `${prefix}${file}` : file;
}

function normalizeTokenUsage(usage: TokenUsage | undefined): TokenUsage | undefined {
  if (!usage) return undefined;
  const total = usage.total_tokens ?? usage.prompt_tokens + usage.completion_tokens;
  return { ...usage, total_tokens: total };
}

function deserializeRow(row: any): AIAuditEntry {
  const token_usage = deserializeTokenUsage(row);
  const verification = deserializeVerification(row.verification_json);
  return {
    id: String(row.id),
    timestamp_ms: Number(row.timestamp_ms),
    session_id: String(row.session_id),
    workbook_id: row.workbook_id ? String(row.workbook_id) : undefined,
    user_id: row.user_id ? String(row.user_id) : undefined,
    mode: row.mode as any,
    input: row.input_json ? safeJsonParse(row.input_json) : null,
    model: String(row.model),
    token_usage,
    latency_ms: row.latency_ms === null || row.latency_ms === undefined ? undefined : Number(row.latency_ms),
    tool_calls: deserializeToolCalls(row.tool_calls_json),
    verification,
    user_feedback: row.user_feedback ? (row.user_feedback as UserFeedback) : undefined
  };
}

function deserializeTokenUsage(row: any): TokenUsage | undefined {
  if (row.prompt_tokens === null && row.completion_tokens === null && row.total_tokens === null) return undefined;
  return {
    prompt_tokens: row.prompt_tokens === null ? 0 : Number(row.prompt_tokens),
    completion_tokens: row.completion_tokens === null ? 0 : Number(row.completion_tokens),
    total_tokens: row.total_tokens === null ? undefined : Number(row.total_tokens)
  };
}

function deserializeToolCalls(encoded: string): ToolCallLog[] {
  const parsed = safeJsonParse(encoded);
  if (!Array.isArray(parsed)) return [];
  return parsed as ToolCallLog[];
}

function deserializeVerification(encoded: unknown): AIVerificationResult | undefined {
  if (encoded === null || encoded === undefined || encoded === "") return undefined;
  if (typeof encoded !== "string") return undefined;
  const parsed = safeJsonParse(encoded);
  if (!parsed || typeof parsed !== "object") return undefined;
  return parsed as AIVerificationResult;
}

function safeJsonParse(value: string): unknown {
  try {
    return JSON.parse(value);
  } catch {
    return null;
  }
}

const AI_AUDIT_SCHEMA_VERSION = 1;

function backfillLegacyWorkbookIds(db: SqlJsDatabase): boolean {
  const version = getUserVersion(db);
  if (version >= AI_AUDIT_SCHEMA_VERSION) return false;

  let didChange = false;
  db.run("BEGIN TRANSACTION;");
  const selectStmt = db.prepare(
    "SELECT id, session_id, input_json FROM ai_audit_log WHERE workbook_id IS NULL OR workbook_id = '';",
  );
  const updateStmt = db.prepare("UPDATE ai_audit_log SET workbook_id = ? WHERE id = ?;");
  try {
    while (selectStmt.step()) {
      const row = selectStmt.getAsObject() as any;
      const workbookId =
        extractWorkbookIdFromInputJson(row.input_json) ?? extractWorkbookIdFromSessionId(String(row.session_id ?? ""));
      if (!workbookId) continue;
      updateStmt.run([workbookId, String(row.id)]);
      didChange = true;
    }

    setUserVersion(db, AI_AUDIT_SCHEMA_VERSION);
    didChange = true;
    db.run("COMMIT;");
    return didChange;
  } catch (err) {
    try {
      db.run("ROLLBACK;");
    } catch {
      // ignore rollback failures
    }
    throw err;
  } finally {
    selectStmt.free();
    updateStmt.free();
  }
}

function getUserVersion(db: SqlJsDatabase): number {
  const stmt = db.prepare("PRAGMA user_version;");
  try {
    if (!stmt.step()) return 0;
    const row = stmt.getAsObject() as any;
    const value = row?.user_version ?? 0;
    const num = Number(value);
    return Number.isFinite(num) ? Math.trunc(num) : 0;
  } finally {
    stmt.free();
  }
}

function setUserVersion(db: SqlJsDatabase, version: number): void {
  const normalized = Number.isFinite(version) ? Math.max(0, Math.trunc(version)) : 0;
  db.run(`PRAGMA user_version = ${normalized};`);
}

function extractWorkbookIdFromInput(input: unknown): string | null {
  if (!input || typeof input !== "object") return null;
  const obj = input as Record<string, unknown>;
  const workbookId = obj.workbook_id ?? obj.workbookId;
  const trimmed = typeof workbookId === "string" ? workbookId.trim() : "";
  return trimmed ? trimmed : null;
}

function extractWorkbookIdFromInputJson(encoded: unknown): string | null {
  if (typeof encoded !== "string") return null;
  const parsed = safeJsonParse(encoded);
  return extractWorkbookIdFromInput(parsed);
}

function extractWorkbookIdFromSessionId(sessionId: string): string | null {
  const match = sessionId.match(/^([^:]+):([0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12})$/);
  if (!match) return null;
  const workbookId = match[1];
  const trimmed = workbookId?.trim() ?? "";
  return trimmed ? trimmed : null;
}

function ensureColumnExists(db: SqlJsDatabase, table: string, column: string, type: string): boolean {
  if (tableHasColumn(db, table, column)) return false;
  db.run(`ALTER TABLE ${table} ADD COLUMN ${column} ${type};`);
  return true;
}

function tableHasColumn(db: SqlJsDatabase, table: string, column: string): boolean {
  const stmt = db.prepare(`PRAGMA table_info(${table});`);
  while (stmt.step()) {
    const row = stmt.getAsObject() as any;
    if (row?.name === column) {
      stmt.free();
      return true;
    }
  }
  stmt.free();
  return false;
}
