use rusqlite::{params, Connection, OptionalExtension, Transaction};
use std::collections::HashSet;

const LATEST_SCHEMA_VERSION: i64 = 9;

pub(crate) fn init(conn: &mut Connection) -> rusqlite::Result<()> {
    // Ensure foreign keys are enforced (disabled by default in SQLite).
    conn.pragma_update(None, "foreign_keys", "ON")?;

    let tx = conn.transaction()?;
    init_schema_version(&tx)?;

    let mut version: i64 = tx
        .query_row("SELECT version FROM schema_version WHERE id = 1", [], |row| {
            if let Some(version) = row.get::<_, Option<i64>>(0).ok().flatten() {
                return Ok(version);
            }
            if let Some(raw) = row.get::<_, Option<String>>(0).ok().flatten() {
                if let Ok(parsed) = raw.trim().parse::<i64>() {
                    return Ok(parsed);
                }
            }
            Ok(0)
        })
        .optional()?
        .unwrap_or(0);
    if version < 0 {
        version = 0;
    }

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
            6 => migrate_to_v6(&tx)?,
            7 => migrate_to_v7(&tx)?,
            8 => migrate_to_v8(&tx)?,
            9 => migrate_to_v9(&tx)?,
            _ => {
                debug_assert!(false, "unknown schema migration target: {next}");
                return Err(rusqlite::Error::InvalidQuery);
            }
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

    // If a corrupted database contains duplicate style-component keys (or non-TEXT keys), creating
    // the unique indexes below would fail and prevent the workbook from opening. Clear invalid
    // keys and keep only the first occurrence of each non-NULL key.
    for table in ["fonts", "fills", "borders"] {
        tx.execute(
            &format!("UPDATE {table} SET key = NULL WHERE key IS NOT NULL AND typeof(key) != 'text'"),
            [],
        )?;
        tx.execute(
            &format!(
                r#"
                UPDATE {table}
                SET key = NULL
                WHERE key IS NOT NULL
                  AND id NOT IN (
                    SELECT MIN(id)
                    FROM {table}
                    WHERE key IS NOT NULL
                    GROUP BY key
                  )
                "#
            ),
            [],
        )?;
    }

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

fn migrate_to_v6(tx: &Transaction<'_>) -> rusqlite::Result<()> {
    // Persist Power Pivot / Data Model state (tables, encoded columnar chunks, relationships, measures).
    tx.execute_batch(
        r#"
        CREATE TABLE IF NOT EXISTS data_model_tables (
          id INTEGER PRIMARY KEY AUTOINCREMENT,
          workbook_id TEXT NOT NULL REFERENCES workbooks(id) ON DELETE CASCADE,
          name TEXT NOT NULL,
          schema_json TEXT NOT NULL,
          row_count INTEGER NOT NULL,
          created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
          metadata JSON,
          UNIQUE(workbook_id, name)
        );

        CREATE INDEX IF NOT EXISTS idx_data_model_tables_workbook ON data_model_tables(workbook_id);

        CREATE TABLE IF NOT EXISTS data_model_columns (
          id INTEGER PRIMARY KEY AUTOINCREMENT,
          table_id INTEGER NOT NULL REFERENCES data_model_tables(id) ON DELETE CASCADE,
          ordinal INTEGER NOT NULL,
          name TEXT NOT NULL,
          column_type TEXT NOT NULL,
          encoding_json TEXT NOT NULL,
          stats_json TEXT,
          dictionary BLOB,
          UNIQUE(table_id, ordinal),
          UNIQUE(table_id, name)
        );

        CREATE INDEX IF NOT EXISTS idx_data_model_columns_table ON data_model_columns(table_id);

        CREATE TABLE IF NOT EXISTS data_model_chunks (
          id INTEGER PRIMARY KEY AUTOINCREMENT,
          column_id INTEGER NOT NULL REFERENCES data_model_columns(id) ON DELETE CASCADE,
          chunk_index INTEGER NOT NULL,
          kind TEXT NOT NULL CHECK (kind IN ('int','float','bool','dict')),
          data BLOB NOT NULL,
          created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
          metadata JSON,
          UNIQUE(column_id, chunk_index)
        );

        CREATE INDEX IF NOT EXISTS idx_data_model_chunks_column ON data_model_chunks(column_id);

        CREATE TABLE IF NOT EXISTS data_model_relationships (
          id INTEGER PRIMARY KEY AUTOINCREMENT,
          workbook_id TEXT NOT NULL REFERENCES workbooks(id) ON DELETE CASCADE,
          name TEXT NOT NULL,
          from_table TEXT NOT NULL,
          from_column TEXT NOT NULL,
          to_table TEXT NOT NULL,
          to_column TEXT NOT NULL,
          cardinality TEXT NOT NULL,
          cross_filter_direction TEXT NOT NULL,
          is_active INTEGER NOT NULL DEFAULT 1,
          referential_integrity INTEGER NOT NULL DEFAULT 0
        );

        CREATE INDEX IF NOT EXISTS idx_data_model_relationships_workbook ON data_model_relationships(workbook_id);
        CREATE INDEX IF NOT EXISTS idx_data_model_relationships_from ON data_model_relationships(workbook_id, from_table);
        CREATE INDEX IF NOT EXISTS idx_data_model_relationships_to ON data_model_relationships(workbook_id, to_table);

        CREATE TABLE IF NOT EXISTS data_model_measures (
          id INTEGER PRIMARY KEY AUTOINCREMENT,
          workbook_id TEXT NOT NULL REFERENCES workbooks(id) ON DELETE CASCADE,
          name TEXT NOT NULL,
          expression TEXT NOT NULL,
          created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
          metadata JSON,
          UNIQUE(workbook_id, name)
        );

        CREATE INDEX IF NOT EXISTS idx_data_model_measures_workbook ON data_model_measures(workbook_id);

        CREATE TABLE IF NOT EXISTS data_model_calculated_columns (
          id INTEGER PRIMARY KEY AUTOINCREMENT,
          workbook_id TEXT NOT NULL REFERENCES workbooks(id) ON DELETE CASCADE,
          table_name TEXT NOT NULL,
          name TEXT NOT NULL,
          expression TEXT NOT NULL,
          created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
          metadata JSON,
          UNIQUE(workbook_id, table_name, name)
        );

        CREATE INDEX IF NOT EXISTS idx_data_model_calculated_columns_workbook ON data_model_calculated_columns(workbook_id);
        "#,
    )?;

    Ok(())
}

fn migrate_to_v7(tx: &Transaction<'_>) -> rusqlite::Result<()> {
    // Optional JSON representation of `formula_model::Worksheet` metadata (excluding cells).
    // This enables round-tripping features not yet modeled in first-class columns/tables.
    ensure_column(tx, "sheets", "model_sheet_json", "model_sheet_json JSON")?;
    Ok(())
}

fn migrate_to_v8(tx: &Transaction<'_>) -> rusqlite::Result<()> {
    // Backfill `model_sheet_id` for sheets that were created before we started assigning stable
    // model ids. This makes `export_model_workbook` deterministic even after sheet reordering.
    //
    // We allocate ids per workbook, preserving the current sheet order, and ensuring they fit
    // within the `u32` domain used by `formula_model::WorksheetId`.
    //
    // Corrupted databases can also contain sheet rows whose `workbook_id` is not a valid TEXT
    // foreign key into `workbooks` (e.g. written with foreign keys disabled, or containing BLOB
    // values). Those rows are not addressable via the public storage APIs, but they can still
    // prevent the unique `(workbook_id, model_sheet_id)` index from being created. We clear their
    // `model_sheet_id` values preemptively so the migration can complete.
    tx.execute(
        r#"
        UPDATE sheets
        SET model_sheet_id = NULL
        WHERE model_sheet_id IS NOT NULL
          AND (
            typeof(workbook_id) != 'text'
            OR NOT EXISTS (SELECT 1 FROM workbooks w WHERE w.id = sheets.workbook_id)
          )
        "#,
        [],
    )?;

    let mut workbook_stmt = tx.prepare("SELECT id FROM workbooks")?;
    let workbook_ids = workbook_stmt
        .query_map([], |row| Ok(row.get::<_, Option<String>>(0).ok().flatten()))?;
    for workbook_id in workbook_ids {
        let Some(workbook_id) = workbook_id.ok().flatten() else {
            continue;
        };
        // Treat invalid/out-of-range ids as unset so we can safely backfill them.
        tx.execute(
            r#"
            UPDATE sheets
            SET model_sheet_id = NULL
            WHERE workbook_id = ?1
              AND model_sheet_id IS NOT NULL
              AND (
                typeof(model_sheet_id) != 'integer'
                OR model_sheet_id < 0
                OR model_sheet_id > ?2
              )
            "#,
            params![&workbook_id, u32::MAX as i64],
        )?;

        // If the database already contains duplicate model sheet ids (possible if a prior client
        // wrote invalid state), clear all but the first occurrence so the unique index below can
        // be created successfully.
        {
            let mut seen: HashSet<u32> = HashSet::new();
            let mut stmt = tx.prepare(
                r#"
                SELECT id, model_sheet_id
                FROM sheets
                WHERE workbook_id = ?1 AND model_sheet_id IS NOT NULL
                ORDER BY model_sheet_id, COALESCE(position, 0), id
                "#,
            )?;
            let rows = stmt.query_map(params![&workbook_id], |row| {
                Ok((
                    row.get::<_, Option<String>>(0).ok().flatten(),
                    row.get::<_, Option<i64>>(1).ok().flatten(),
                ))
            })?;
            for row in rows {
                let Ok((sheet_id, raw)) = row else {
                    continue;
                };
                let (Some(sheet_id), Some(raw)) = (sheet_id, raw) else {
                    continue;
                };
                let Ok(id) = u32::try_from(raw) else {
                    continue;
                };
                if !seen.insert(id) {
                    tx.execute(
                        "UPDATE sheets SET model_sheet_id = NULL WHERE id = ?1",
                        params![sheet_id],
                    )?;
                }
            }
        }

        let mut used: HashSet<u32> = HashSet::new();
        let mut max_existing: u32 = 0;
        {
            let mut stmt = tx.prepare(
                "SELECT model_sheet_id FROM sheets WHERE workbook_id = ?1 AND model_sheet_id IS NOT NULL",
            )?;
            let ids = stmt.query_map(params![&workbook_id], |row| Ok(row.get::<_, Option<i64>>(0).ok().flatten()))?;
            for raw in ids {
                let Some(raw) = raw? else {
                    continue;
                };
                if let Ok(id) = u32::try_from(raw) {
                    used.insert(id);
                    max_existing = max_existing.max(id);
                }
            }
        }

        let mut next_id = max_existing.wrapping_add(1);

        let mut sheet_stmt = tx.prepare(
            r#"
            SELECT id
            FROM sheets
            WHERE workbook_id = ?1 AND model_sheet_id IS NULL
            ORDER BY COALESCE(position, 0), id
            "#,
        )?;
        let sheet_ids =
            sheet_stmt.query_map(params![&workbook_id], |row| Ok(row.get::<_, Option<String>>(0).ok().flatten()))?;
        for sheet_id in sheet_ids {
            let Some(sheet_id) = sheet_id.ok().flatten() else {
                continue;
            };
            while used.contains(&next_id) {
                next_id = next_id.wrapping_add(1);
            }
            tx.execute(
                "UPDATE sheets SET model_sheet_id = ?1 WHERE id = ?2",
                params![next_id as i64, sheet_id],
            )?;
            used.insert(next_id);
            next_id = next_id.wrapping_add(1);
        }
    }

    tx.execute_batch(
        r#"
        CREATE UNIQUE INDEX IF NOT EXISTS idx_sheets_workbook_model_sheet_id
        ON sheets(workbook_id, model_sheet_id)
        WHERE model_sheet_id IS NOT NULL;
        "#,
    )?;

    Ok(())
}

fn migrate_to_v9(tx: &Transaction<'_>) -> rusqlite::Result<()> {
    // Persist workbook text codepage (used for legacy DBCS text functions and XLS imports).
    ensure_column(
        tx,
        "workbooks",
        "codepage",
        "codepage INTEGER NOT NULL DEFAULT 1252",
    )?;
    // Persist per-cell phonetic guide text (furigana).
    ensure_column(tx, "cells", "phonetic", "phonetic TEXT")?;
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
    // Best-effort: ignore malformed rows rather than failing migrations. While PRAGMA output
    // should be well-formed for valid schemas, corrupted databases can contain unexpected
    // values that would otherwise prevent the storage from opening.
    let rows = stmt.query_map([], |row| Ok(row.get::<_, Option<String>>(1).ok().flatten()))?;
    for name in rows {
        let Ok(Some(name)) = name else {
            continue;
        };
        if name == column {
            return Ok(true);
        }
    }
    Ok(false)
}
