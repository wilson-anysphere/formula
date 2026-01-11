import initSqlJs from "sql.js";
import type { AIAuditEntry, AuditListFilters, TokenUsage, ToolCallLog, UserFeedback, AIVerificationResult } from "./types.js";
import type { AIAuditStore } from "./store.js";
import type { SqliteBinaryStorage } from "./storage.js";
import { InMemoryBinaryStorage } from "./storage.js";

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
}

export class SqliteAIAuditStore implements AIAuditStore {
  private readonly db: SqlJsDatabase;
  private readonly storage: SqliteBinaryStorage;
  private readonly retention: SqliteAIAuditStoreRetention;

  private constructor(db: SqlJsDatabase, storage: SqliteBinaryStorage, retention: SqliteAIAuditStoreRetention) {
    this.db = db;
    this.storage = storage;
    this.retention = retention;
    this.ensureSchema();
  }

  static async create(options: SqliteAIAuditStoreOptions = {}): Promise<SqliteAIAuditStore> {
    const storage = options.storage ?? new InMemoryBinaryStorage();
    const SQL = await initSqlJs({ locateFile: options.locateFile ?? locateSqlJsFile });
    const existing = await storage.load();
    const db = existing ? new SQL.Database(existing) : new SQL.Database();
    return new SqliteAIAuditStore(db, storage, options.retention ?? {});
  }

  async logEntry(entry: AIAuditEntry): Promise<void> {
    const tokenUsage = normalizeTokenUsage(entry.token_usage);
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
      entry.workbook_id ?? null,
      entry.user_id ?? null,
      entry.mode,
      JSON.stringify(entry.input ?? null),
      entry.model,
      tokenUsage?.prompt_tokens ?? null,
      tokenUsage?.completion_tokens ?? null,
      tokenUsage?.total_tokens ?? null,
      entry.latency_ms ?? null,
      JSON.stringify(entry.tool_calls ?? []),
      entry.verification ? JSON.stringify(entry.verification) : null,
      entry.user_feedback ?? null
    ]);
    stmt.free();

    this.enforceRetention();
    await this.persist();
  }

  async listEntries(filters: AuditListFilters = {}): Promise<AIAuditEntry[]> {
    const params: any[] = [];
    let sql = "SELECT * FROM ai_audit_log";
    const where: string[] = [];
    if (filters.session_id) {
      where.push("session_id = ?");
      params.push(filters.session_id);
    }
    if (filters.workbook_id) {
      where.push("workbook_id = ?");
      params.push(filters.workbook_id);
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
    if (where.length > 0) {
      sql += ` WHERE ${where.join(" AND ")}`;
    }
    sql += " ORDER BY timestamp_ms DESC";
    if (typeof filters.limit === "number") {
      sql += " LIMIT ?";
      params.push(filters.limit);
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
    ensureColumnExists(this.db, "ai_audit_log", "verification_json", "TEXT");
    // Migrate databases created before workbook metadata was tracked.
    ensureColumnExists(this.db, "ai_audit_log", "workbook_id", "TEXT");

    // Indexes should be created after migrations (so older databases missing columns
    // don't error when we try to index them).
    this.db.run(`
      CREATE INDEX IF NOT EXISTS idx_ai_audit_log_session ON ai_audit_log(session_id);
      CREATE INDEX IF NOT EXISTS idx_ai_audit_log_workbook ON ai_audit_log(workbook_id);
      CREATE INDEX IF NOT EXISTS idx_ai_audit_log_mode ON ai_audit_log(mode);
      CREATE INDEX IF NOT EXISTS idx_ai_audit_log_timestamp ON ai_audit_log(timestamp_ms);
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
          ORDER BY timestamp_ms DESC
          LIMIT -1 OFFSET ?
        );
      `);
      stmt.run([Math.floor(maxEntries)]);
      stmt.free();
    }
  }

  private async persist(): Promise<void> {
    const data = this.db.export() as Uint8Array;
    await this.storage.save(data);
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

function ensureColumnExists(db: SqlJsDatabase, table: string, column: string, type: string): void {
  if (tableHasColumn(db, table, column)) return;
  db.run(`ALTER TABLE ${table} ADD COLUMN ${column} ${type};`);
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
