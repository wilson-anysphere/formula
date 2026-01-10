use rusqlite::Connection;

pub(crate) fn init(conn: &Connection) -> rusqlite::Result<()> {
    // Ensure foreign keys are enforced (disabled by default in SQLite).
    conn.pragma_update(None, "foreign_keys", "ON")?;

    conn.execute_batch(
        r#"
        -- Core tables
        CREATE TABLE IF NOT EXISTS workbooks (
          id TEXT PRIMARY KEY,
          name TEXT NOT NULL,
          created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
          modified_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
          metadata JSON
        );

        CREATE TABLE IF NOT EXISTS sheets (
          id TEXT PRIMARY KEY,
          workbook_id TEXT REFERENCES workbooks(id),
          name TEXT NOT NULL,
          position INTEGER,
          visibility TEXT NOT NULL DEFAULT 'visible' CHECK (visibility IN ('visible','hidden','veryHidden')),
          tab_color TEXT,
          xlsx_sheet_id INTEGER,
          xlsx_rel_id TEXT,
          frozen_rows INTEGER DEFAULT 0,
          frozen_cols INTEGER DEFAULT 0,
          zoom REAL DEFAULT 1.0,
          metadata JSON
        );

        CREATE TABLE IF NOT EXISTS cells (
          sheet_id TEXT REFERENCES sheets(id),
          row INTEGER,
          col INTEGER,
          value_type TEXT,  -- 'number', 'string', 'boolean', 'error', 'formula'
          value_number REAL,
          value_string TEXT,
          formula TEXT,
          style_id INTEGER,
          PRIMARY KEY (sheet_id, row, col)
        );

        CREATE INDEX IF NOT EXISTS idx_cells_sheet ON cells(sheet_id);
        CREATE INDEX IF NOT EXISTS idx_cells_sheet_row ON cells(sheet_id, row);

        -- Style component tables are not detailed in the design doc, but the
        -- `styles` table references them. We keep them minimal for now so the
        -- schema matches the documented foreign keys.
        CREATE TABLE IF NOT EXISTS fonts (
          id INTEGER PRIMARY KEY,
          data JSON
        );

        CREATE TABLE IF NOT EXISTS fills (
          id INTEGER PRIMARY KEY,
          data JSON
        );

        CREATE TABLE IF NOT EXISTS borders (
          id INTEGER PRIMARY KEY,
          data JSON
        );

        CREATE TABLE IF NOT EXISTS styles (
          id INTEGER PRIMARY KEY,
          font_id INTEGER REFERENCES fonts(id),
          fill_id INTEGER REFERENCES fills(id),
          border_id INTEGER REFERENCES borders(id),
          number_format TEXT,
          alignment JSON,
          protection JSON
        );

        CREATE TABLE IF NOT EXISTS named_ranges (
          workbook_id TEXT REFERENCES workbooks(id),
          name TEXT,
          scope TEXT,
          reference TEXT,
          PRIMARY KEY (workbook_id, name, scope)
        );

        CREATE TABLE IF NOT EXISTS change_log (
          id INTEGER PRIMARY KEY AUTOINCREMENT,
          sheet_id TEXT REFERENCES sheets(id),
          timestamp TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
          user_id TEXT,
          operation TEXT,
          target JSON,
          old_value JSON,
          new_value JSON
        );

        CREATE INDEX IF NOT EXISTS idx_change_log_sheet ON change_log(sheet_id);
        "#,
    )?;

    // Best-effort migrations for older databases that predate sheet tab metadata.
    // SQLite only supports ADD COLUMN migrations, so we opportunistically add
    // missing columns when opening an existing database.
    ensure_sheet_columns(conn)?;

    Ok(())
}

fn ensure_sheet_columns(conn: &Connection) -> rusqlite::Result<()> {
    let mut stmt = conn.prepare("PRAGMA table_info(sheets)")?;
    let rows = stmt.query_map([], |row| row.get::<_, String>(1))?;
    let mut existing = std::collections::HashSet::new();
    for name in rows {
        existing.insert(name?);
    }

    if !existing.contains("visibility") {
        conn.execute(
            "ALTER TABLE sheets ADD COLUMN visibility TEXT NOT NULL DEFAULT 'visible' CHECK (visibility IN ('visible','hidden','veryHidden'))",
            [],
        )?;
    }
    if !existing.contains("tab_color") {
        conn.execute("ALTER TABLE sheets ADD COLUMN tab_color TEXT", [])?;
    }
    if !existing.contains("xlsx_sheet_id") {
        conn.execute("ALTER TABLE sheets ADD COLUMN xlsx_sheet_id INTEGER", [])?;
    }
    if !existing.contains("xlsx_rel_id") {
        conn.execute("ALTER TABLE sheets ADD COLUMN xlsx_rel_id TEXT", [])?;
    }

    Ok(())
}
