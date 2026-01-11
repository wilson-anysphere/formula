use rusqlite::{params, Connection, Transaction};

const LATEST_SCHEMA_VERSION: i64 = 5;

pub(crate) fn init(conn: &mut Connection) -> rusqlite::Result<()> {
    // Ensure foreign keys are enforced (disabled by default in SQLite).
    conn.pragma_update(None, "foreign_keys", "ON")?;

    let tx = conn.transaction()?;
    init_schema_version(&tx)?;

    let mut version: i64 = tx.query_row(
        "SELECT version FROM schema_version WHERE id = 1",
        [],
        |row| row.get(0),
    )?;

    // If a newer client has already migrated the database, fail fast. This
    // avoids silently corrupting state by attempting to operate on an unknown schema.
    if version > LATEST_SCHEMA_VERSION {
        return Err(rusqlite::Error::InvalidQuery);
    }

    while version < LATEST_SCHEMA_VERSION {
        let next = version + 1;
        match next {
            1 => migrate_to_v1(&tx)?,
            2 => migrate_to_v2(&tx)?,
            3 => migrate_to_v3(&tx)?,
            4 => migrate_to_v4(&tx)?,
            5 => migrate_to_v5(&tx)?,
            _ => unreachable!("unknown schema migration target: {next}"),
        }
        tx.execute(
            "UPDATE schema_version SET version = ?1 WHERE id = 1",
            params![next],
        )?;
        version = next;
    }

    tx.commit()
}

fn init_schema_version(tx: &Transaction<'_>) -> rusqlite::Result<()> {
    tx.execute_batch(
        r#"
        CREATE TABLE IF NOT EXISTS schema_version (
          id INTEGER PRIMARY KEY CHECK (id = 1),
          version INTEGER NOT NULL
        );
        INSERT OR IGNORE INTO schema_version (id, version) VALUES (1, 0);
        "#,
    )
}

fn migrate_to_v1(tx: &Transaction<'_>) -> rusqlite::Result<()> {
    tx.execute_batch(
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
          value_type TEXT,  -- 'number', 'string', 'boolean', 'error', 'formula' (legacy)
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

    // Best-effort migrations for legacy databases that predate newer sheet metadata.
    ensure_sheet_columns(tx)?;
    Ok(())
}

fn migrate_to_v2(tx: &Transaction<'_>) -> rusqlite::Result<()> {
    // Persist full `formula-model` cell values (RichText, Array, Spill, typed errors)
    // while keeping the scalar fast-path columns for common cases.
    ensure_column(tx, "cells", "value_json", "value_json TEXT")?;

    Ok(())
}

fn migrate_to_v3(tx: &Transaction<'_>) -> rusqlite::Result<()> {
    // Workbook metadata needed to reconstruct a `formula_model::Workbook`.
    ensure_column(
        tx,
        "workbooks",
        "model_schema_version",
        "model_schema_version INTEGER",
    )?;
    ensure_column(tx, "workbooks", "model_workbook_id", "model_workbook_id INTEGER")?;
    ensure_column(tx, "workbooks", "date_system", "date_system TEXT")?;
    ensure_column(tx, "workbooks", "calc_settings", "calc_settings JSON")?;

    // Sheet metadata needed to reconstruct a `formula_model::Worksheet`.
    ensure_column(tx, "sheets", "model_sheet_id", "model_sheet_id INTEGER")?;
    // Preserve full XLSX tab color payload (rgb/theme/indexed/tint/auto) while keeping the
    // legacy `tab_color` fast-path (ARGB hex string) used by existing APIs.
    ensure_column(tx, "sheets", "tab_color_json", "tab_color_json JSON")?;

    // Style component dedup keys. These tables existed in v1 as placeholders; v3
    // upgrades them into usable round-trip storage for `formula_model::style`
    // components.
    ensure_column(tx, "fonts", "key", "key TEXT")?;
    ensure_column(tx, "fills", "key", "key TEXT")?;
    ensure_column(tx, "borders", "key", "key TEXT")?;

    tx.execute_batch(
        r#"
        CREATE UNIQUE INDEX IF NOT EXISTS idx_fonts_key ON fonts(key);
        CREATE UNIQUE INDEX IF NOT EXISTS idx_fills_key ON fills(key);
        CREATE UNIQUE INDEX IF NOT EXISTS idx_borders_key ON borders(key);

        -- Preserve `formula_model::StyleTable` ordering by mapping per-workbook style indices
        -- to global `styles` rows.
        CREATE TABLE IF NOT EXISTS workbook_styles (
          workbook_id TEXT NOT NULL REFERENCES workbooks(id),
          style_index INTEGER NOT NULL,
          style_id INTEGER NOT NULL REFERENCES styles(id),
          PRIMARY KEY (workbook_id, style_index)
        );

        CREATE INDEX IF NOT EXISTS idx_workbook_styles_workbook ON workbook_styles(workbook_id);
        "#,
    )?;

    Ok(())
}

fn migrate_to_v4(tx: &Transaction<'_>) -> rusqlite::Result<()> {
    // Additional `formula_model::Workbook` metadata for round-trip fidelity.
    ensure_column(tx, "workbooks", "theme", "theme JSON")?;
    ensure_column(
        tx,
        "workbooks",
        "workbook_protection",
        "workbook_protection JSON",
    )?;
    ensure_column(tx, "workbooks", "defined_names", "defined_names JSON")?;
    ensure_column(tx, "workbooks", "print_settings", "print_settings JSON")?;
    ensure_column(tx, "workbooks", "view", "view JSON")?;
    Ok(())
}

fn migrate_to_v5(tx: &Transaction<'_>) -> rusqlite::Result<()> {
    tx.execute_batch(
        r#"
        CREATE TABLE IF NOT EXISTS workbook_images (
          workbook_id TEXT NOT NULL REFERENCES workbooks(id),
          image_id TEXT NOT NULL,
          content_type TEXT,
          bytes BLOB NOT NULL,
          PRIMARY KEY (workbook_id, image_id)
        );
        CREATE INDEX IF NOT EXISTS idx_workbook_images_workbook ON workbook_images(workbook_id);

        CREATE TABLE IF NOT EXISTS sheet_drawings (
          sheet_id TEXT NOT NULL REFERENCES sheets(id),
          position INTEGER NOT NULL,
          data JSON NOT NULL,
          PRIMARY KEY (sheet_id, position)
        );
        CREATE INDEX IF NOT EXISTS idx_sheet_drawings_sheet ON sheet_drawings(sheet_id);
        "#,
    )?;
    Ok(())
}

fn ensure_sheet_columns(tx: &Transaction<'_>) -> rusqlite::Result<()> {
    ensure_column(
        tx,
        "sheets",
        "visibility",
        "visibility TEXT NOT NULL DEFAULT 'visible' CHECK (visibility IN ('visible','hidden','veryHidden'))",
    )?;
    ensure_column(tx, "sheets", "tab_color", "tab_color TEXT")?;
    ensure_column(tx, "sheets", "xlsx_sheet_id", "xlsx_sheet_id INTEGER")?;
    ensure_column(tx, "sheets", "xlsx_rel_id", "xlsx_rel_id TEXT")?;
    ensure_column(tx, "sheets", "frozen_rows", "frozen_rows INTEGER DEFAULT 0")?;
    ensure_column(tx, "sheets", "frozen_cols", "frozen_cols INTEGER DEFAULT 0")?;
    ensure_column(tx, "sheets", "zoom", "zoom REAL DEFAULT 1.0")?;
    ensure_column(tx, "sheets", "metadata", "metadata JSON")?;
    Ok(())
}

fn ensure_column(
    tx: &Transaction<'_>,
    table: &str,
    column: &str,
    column_ddl: &str,
) -> rusqlite::Result<()> {
    if column_exists(tx, table, column)? {
        return Ok(());
    }
    tx.execute(&format!("ALTER TABLE {table} ADD COLUMN {column_ddl}"), [])?;
    Ok(())
}

fn column_exists(tx: &Transaction<'_>, table: &str, column: &str) -> rusqlite::Result<bool> {
    let mut stmt = tx.prepare(&format!("PRAGMA table_info({table})"))?;
    let rows = stmt.query_map([], |row| row.get::<_, String>(1))?;
    for name in rows {
        if name? == column {
            return Ok(true);
        }
    }
    Ok(false)
}
