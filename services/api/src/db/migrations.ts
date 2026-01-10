import fs from "node:fs/promises";
import path from "node:path";
import type { Pool } from "pg";

export interface MigrationOptions {
  migrationsDir: string;
}

export async function runMigrations(pool: Pool, options: MigrationOptions): Promise<void> {
  await pool.query(`
    CREATE TABLE IF NOT EXISTS schema_migrations (
      id bigserial PRIMARY KEY,
      filename text UNIQUE NOT NULL,
      applied_at timestamptz NOT NULL DEFAULT now()
    );
  `);

  const entries = await fs.readdir(options.migrationsDir, { withFileTypes: true });
  const migrationFiles = entries
    .filter((entry) => entry.isFile() && entry.name.endsWith(".sql"))
    .map((entry) => entry.name)
    .sort();

  for (const filename of migrationFiles) {
    const existing = await pool.query("SELECT 1 FROM schema_migrations WHERE filename = $1", [
      filename
    ]);
    if (existing.rowCount && existing.rowCount > 0) continue;

    const sql = await fs.readFile(path.join(options.migrationsDir, filename), "utf8");
    await pool.query("BEGIN");
    try {
      await pool.query(sql);
      await pool.query("INSERT INTO schema_migrations (filename) VALUES ($1)", [filename]);
      await pool.query("COMMIT");
    } catch (error) {
      await pool.query("ROLLBACK");
      throw error;
    }
  }
}

