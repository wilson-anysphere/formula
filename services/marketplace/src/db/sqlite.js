const fs = require("node:fs/promises");
const path = require("node:path");
const { createRequire } = require("node:module");

function locateSqlJsFile(file) {
  try {
    const requireFromHere = createRequire(__filename);
    return requireFromHere.resolve(`sql.js/dist/${file}`);
  } catch {
    return file;
  }
}

class SqliteFileDb {
  constructor({ filePath, migrationsDir }) {
    if (!filePath) throw new Error("filePath is required");
    if (!migrationsDir) throw new Error("migrationsDir is required");
    this.filePath = filePath;
    this.migrationsDir = migrationsDir;

    this._db = null;
    this._initPromise = null;

    // Serialize filesystem writes and ensure a request is only acknowledged once persisted.
    this._writeChain = Promise.resolve();
  }

  async _open() {
    if (this._db) return this._db;
    if (this._initPromise) return this._initPromise;
    this._initPromise = this._openInner();
    return this._initPromise;
  }

  async _openInner() {
    await fs.mkdir(path.dirname(this.filePath), { recursive: true });

    // sql.js is ESM, but this service is CJS.
    const initSqlJs = (await import("sql.js")).default;
    const SQL = await initSqlJs({ locateFile: locateSqlJsFile });

    /** @type {Uint8Array | null} */
    let existing = null;
    try {
      existing = await fs.readFile(this.filePath);
    } catch {
      existing = null;
    }

    const db = existing ? new SQL.Database(existing) : new SQL.Database();
    this._db = db;
    await this._applyMigrations();
    return db;
  }

  async _applyMigrations() {
    const db = this._db;
    if (!db) throw new Error("DB not initialized");

    db.run(`
      PRAGMA foreign_keys = ON;
      CREATE TABLE IF NOT EXISTS schema_migrations (
        version INTEGER PRIMARY KEY,
        name TEXT NOT NULL,
        applied_at TEXT NOT NULL
      );
    `);

    const applied = new Set();
    const appliedRows = db.exec("SELECT version FROM schema_migrations ORDER BY version ASC");
    if (appliedRows[0]?.values) {
      for (const [v] of appliedRows[0].values) applied.add(Number(v));
    }

    const entries = await fs.readdir(this.migrationsDir, { withFileTypes: true });
    const migrations = entries
      .filter((e) => e.isFile() && e.name.endsWith(".sql"))
      .map((e) => e.name)
      .sort();

    for (const filename of migrations) {
      const match = /^(\d+)_/.exec(filename);
      if (!match) continue;
      const version = Number(match[1]);
      if (applied.has(version)) continue;

      const sql = await fs.readFile(path.join(this.migrationsDir, filename), "utf8");
      db.run("BEGIN");
      try {
        db.run(sql);
        const stmt = db.prepare(
          "INSERT INTO schema_migrations (version, name, applied_at) VALUES (?, ?, ?)"
        );
        stmt.run([version, filename, new Date().toISOString()]);
        stmt.free();
        db.run("COMMIT");
      } catch (err) {
        db.run("ROLLBACK");
        throw err;
      }
    }
  }

  async _persist() {
    const db = await this._open();
    const out = Buffer.from(db.export());
    const tmp = `${this.filePath}.tmp`;
    await fs.writeFile(tmp, out);
    await fs.rename(tmp, this.filePath);
  }

  async _withWriteLock(fn) {
    let release = null;
    const next = new Promise((resolve) => {
      release = resolve;
    });
    const prev = this._writeChain;
    this._writeChain = next;
    await prev;
    try {
      return await fn();
    } finally {
      release();
    }
  }

  async withTransaction(fn) {
    return this._withWriteLock(async () => {
      const db = await this._open();
      db.run("BEGIN");
      try {
        const result = await fn(db);
        db.run("COMMIT");
        await this._persist();
        return result;
      } catch (err) {
        db.run("ROLLBACK");
        throw err;
      }
    });
  }

  async getDb() {
    return this._open();
  }

  close() {
    if (!this._db) return;
    this._db.close();
    this._db = null;
    this._initPromise = null;
  }
}

module.exports = {
  SqliteFileDb,
};

