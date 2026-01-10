use crate::schema;
use crate::types::{CellData, CellSnapshot, CellValue, NamedRange, SheetMeta, Style, WorkbookMeta};
use rusqlite::{params, Connection, OpenFlags, OptionalExtension, Transaction};
use serde_json::json;
use std::path::Path;
use std::sync::{Arc, Mutex};
use std::time::Duration;
use thiserror::Error;
use uuid::Uuid;

#[derive(Debug, Error)]
pub enum StorageError {
    #[error("sqlite error: {0}")]
    Sqlite(#[from] rusqlite::Error),
    #[error("json error: {0}")]
    Json(#[from] serde_json::Error),
    #[error("workbook not found: {0}")]
    WorkbookNotFound(Uuid),
    #[error("sheet not found: {0}")]
    SheetNotFound(Uuid),
}

pub type Result<T> = std::result::Result<T, StorageError>;

#[derive(Debug, Clone)]
pub struct Storage {
    conn: Arc<Mutex<Connection>>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct CellRange {
    pub row_start: i64,
    pub row_end: i64,
    pub col_start: i64,
    pub col_end: i64,
}

impl CellRange {
    pub fn new(row_start: i64, row_end: i64, col_start: i64, col_end: i64) -> Self {
        Self {
            row_start,
            row_end,
            col_start,
            col_end,
        }
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct CellChange {
    pub sheet_id: Uuid,
    pub row: i64,
    pub col: i64,
    pub data: CellData,
    pub user_id: Option<String>,
}

impl Storage {
    pub fn open_path(path: impl AsRef<Path>) -> Result<Self> {
        let conn = Connection::open(path)?;
        conn.busy_timeout(Duration::from_secs(5))?;
        schema::init(&conn)?;
        Ok(Self {
            conn: Arc::new(Mutex::new(conn)),
        })
    }

    pub fn open_in_memory() -> Result<Self> {
        let conn = Connection::open_in_memory()?;
        conn.busy_timeout(Duration::from_secs(5))?;
        schema::init(&conn)?;
        Ok(Self {
            conn: Arc::new(Mutex::new(conn)),
        })
    }

    pub fn open_uri(uri: &str) -> Result<Self> {
        let flags = OpenFlags::SQLITE_OPEN_READ_WRITE
            | OpenFlags::SQLITE_OPEN_CREATE
            | OpenFlags::SQLITE_OPEN_URI;
        let conn = Connection::open_with_flags(uri, flags)?;
        conn.busy_timeout(Duration::from_secs(5))?;
        schema::init(&conn)?;
        Ok(Self {
            conn: Arc::new(Mutex::new(conn)),
        })
    }

    pub fn create_workbook(&self, name: &str, metadata: Option<serde_json::Value>) -> Result<WorkbookMeta> {
        let workbook = WorkbookMeta {
            id: Uuid::new_v4(),
            name: name.to_string(),
            metadata,
        };

        let conn = self.conn.lock().expect("storage mutex poisoned");
        conn.execute(
            "INSERT INTO workbooks (id, name, metadata) VALUES (?1, ?2, ?3)",
            params![
                workbook.id.to_string(),
                &workbook.name,
                workbook.metadata.clone()
            ],
        )?;

        Ok(workbook)
    }

    pub fn get_workbook(&self, id: Uuid) -> Result<WorkbookMeta> {
        let conn = self.conn.lock().expect("storage mutex poisoned");
        let row = conn
            .query_row(
                "SELECT id, name, metadata FROM workbooks WHERE id = ?1",
                params![id.to_string()],
                |r| {
                    let id: String = r.get(0)?;
                    Ok(WorkbookMeta {
                        id: Uuid::parse_str(&id).map_err(|_| rusqlite::Error::InvalidQuery)?,
                        name: r.get(1)?,
                        metadata: r.get(2)?,
                    })
                },
            )
            .optional()?;

        row.ok_or(StorageError::WorkbookNotFound(id))
    }

    pub fn create_sheet(
        &self,
        workbook_id: Uuid,
        name: &str,
        position: i64,
        metadata: Option<serde_json::Value>,
    ) -> Result<SheetMeta> {
        // Ensure workbook exists.
        self.get_workbook(workbook_id)?;

        let sheet = SheetMeta {
            id: Uuid::new_v4(),
            workbook_id,
            name: name.to_string(),
            position,
            frozen_rows: 0,
            frozen_cols: 0,
            zoom: 1.0,
            metadata,
        };

        let conn = self.conn.lock().expect("storage mutex poisoned");
        conn.execute(
            r#"
            INSERT INTO sheets (
              id, workbook_id, name, position, frozen_rows, frozen_cols, zoom, metadata
            ) VALUES (?1, ?2, ?3, ?4, ?5, ?6, ?7, ?8)
            "#,
            params![
                sheet.id.to_string(),
                sheet.workbook_id.to_string(),
                &sheet.name,
                sheet.position,
                sheet.frozen_rows,
                sheet.frozen_cols,
                sheet.zoom,
                sheet.metadata.clone()
            ],
        )?;

        Ok(sheet)
    }

    pub fn list_sheets(&self, workbook_id: Uuid) -> Result<Vec<SheetMeta>> {
        let conn = self.conn.lock().expect("storage mutex poisoned");
        let mut stmt = conn.prepare(
            r#"
            SELECT id, workbook_id, name, position, frozen_rows, frozen_cols, zoom, metadata
            FROM sheets
            WHERE workbook_id = ?1
            ORDER BY position
            "#,
        )?;

        let rows = stmt.query_map(params![workbook_id.to_string()], |r| {
            let id: String = r.get(0)?;
            let workbook_id: String = r.get(1)?;
            Ok(SheetMeta {
                id: Uuid::parse_str(&id).map_err(|_| rusqlite::Error::InvalidQuery)?,
                workbook_id: Uuid::parse_str(&workbook_id).map_err(|_| rusqlite::Error::InvalidQuery)?,
                name: r.get(2)?,
                position: r.get(3)?,
                frozen_rows: r.get(4)?,
                frozen_cols: r.get(5)?,
                zoom: r.get(6)?,
                metadata: r.get(7)?,
            })
        })?;

        let mut sheets = Vec::new();
        for sheet in rows {
            sheets.push(sheet?);
        }
        Ok(sheets)
    }

    pub fn get_sheet_meta(&self, sheet_id: Uuid) -> Result<SheetMeta> {
        let conn = self.conn.lock().expect("storage mutex poisoned");
        let row = conn
            .query_row(
                r#"
                SELECT id, workbook_id, name, position, frozen_rows, frozen_cols, zoom, metadata
                FROM sheets
                WHERE id = ?1
                "#,
                params![sheet_id.to_string()],
                |r| {
                    let id: String = r.get(0)?;
                    let workbook_id: String = r.get(1)?;
                    Ok(SheetMeta {
                        id: Uuid::parse_str(&id).map_err(|_| rusqlite::Error::InvalidQuery)?,
                        workbook_id: Uuid::parse_str(&workbook_id)
                            .map_err(|_| rusqlite::Error::InvalidQuery)?,
                        name: r.get(2)?,
                        position: r.get(3)?,
                        frozen_rows: r.get(4)?,
                        frozen_cols: r.get(5)?,
                        zoom: r.get(6)?,
                        metadata: r.get(7)?,
                    })
                },
            )
            .optional()?;

        row.ok_or(StorageError::SheetNotFound(sheet_id))
    }

    /// Load all non-empty cells within an inclusive range.
    pub fn load_cells_in_range(&self, sheet_id: Uuid, range: CellRange) -> Result<Vec<((i64, i64), CellSnapshot)>> {
        let conn = self.conn.lock().expect("storage mutex poisoned");
        let mut stmt = conn.prepare(
            r#"
            SELECT row, col, value_type, value_number, value_string, formula, style_id
            FROM cells
            WHERE sheet_id = ?1
              AND row >= ?2 AND row <= ?3
              AND col >= ?4 AND col <= ?5
            ORDER BY row, col
            "#,
        )?;

        let rows = stmt.query_map(
            params![
                sheet_id.to_string(),
                range.row_start,
                range.row_end,
                range.col_start,
                range.col_end
            ],
            |r| {
                let row: i64 = r.get(0)?;
                let col: i64 = r.get(1)?;
                let snapshot = snapshot_from_row(r.get(2)?, r.get(3)?, r.get(4)?, r.get(5)?, r.get(6)?)?;
                Ok(((row, col), snapshot))
            },
        )?;

        let mut out = Vec::new();
        for item in rows {
            out.push(item?);
        }
        Ok(out)
    }

    /// Stream cells in row batches. Each batch returns all cells whose row is in
    /// `[start_row, start_row + batch_size)`.
    pub fn load_cells_row_batch(
        &self,
        sheet_id: Uuid,
        start_row: i64,
        batch_size: i64,
    ) -> Result<Vec<((i64, i64), CellSnapshot)>> {
        let range = CellRange {
            row_start: start_row,
            row_end: start_row + batch_size - 1,
            col_start: i64::MIN / 2,
            col_end: i64::MAX / 2,
        };
        // The huge col bounds are okay because `cells` is sparse and the query
        // remains bounded by row range and sheet_id.
        self.load_cells_in_range(sheet_id, range)
    }

    pub fn cell_count(&self, sheet_id: Uuid) -> Result<u64> {
        let conn = self.conn.lock().expect("storage mutex poisoned");
        let count: u64 = conn.query_row(
            "SELECT COUNT(*) FROM cells WHERE sheet_id = ?1",
            params![sheet_id.to_string()],
            |r| r.get(0),
        )?;
        Ok(count)
    }

    pub fn apply_cell_changes(&self, changes: &[CellChange]) -> Result<()> {
        if changes.is_empty() {
            return Ok(());
        }

        let mut conn = self.conn.lock().expect("storage mutex poisoned");
        let tx = conn.transaction()?;

        for change in changes {
            apply_one_change(&tx, change)?;
        }

        tx.commit()?;
        Ok(())
    }

    pub fn get_or_insert_style(&self, style: &Style) -> Result<i64> {
        let mut conn = self.conn.lock().expect("storage mutex poisoned");
        let tx = conn.transaction()?;
        let id = get_or_insert_style_tx(&tx, style)?;
        tx.commit()?;
        Ok(id)
    }

    pub fn get_named_range(&self, workbook_id: Uuid, name: &str, scope: &str) -> Result<Option<NamedRange>> {
        let conn = self.conn.lock().expect("storage mutex poisoned");
        let row = conn
            .query_row(
                r#"
                SELECT workbook_id, name, scope, reference
                FROM named_ranges
                WHERE workbook_id = ?1 AND name = ?2 AND scope = ?3
                "#,
                params![workbook_id.to_string(), name, scope],
                |r| {
                    let workbook_id: String = r.get(0)?;
                    Ok(NamedRange {
                        workbook_id: Uuid::parse_str(&workbook_id)
                            .map_err(|_| rusqlite::Error::InvalidQuery)?,
                        name: r.get(1)?,
                        scope: r.get(2)?,
                        reference: r.get(3)?,
                    })
                },
            )
            .optional()?;
        Ok(row)
    }

    pub fn upsert_named_range(&self, range: &NamedRange) -> Result<()> {
        let conn = self.conn.lock().expect("storage mutex poisoned");
        conn.execute(
            r#"
            INSERT INTO named_ranges (workbook_id, name, scope, reference)
            VALUES (?1, ?2, ?3, ?4)
            ON CONFLICT(workbook_id, name, scope) DO UPDATE SET reference = excluded.reference
            "#,
            params![
                range.workbook_id.to_string(),
                &range.name,
                &range.scope,
                &range.reference
            ],
        )?;
        Ok(())
    }
}

fn snapshot_from_row(
    value_type: Option<String>,
    value_number: Option<f64>,
    value_string: Option<String>,
    formula: Option<String>,
    style_id: Option<i64>,
) -> rusqlite::Result<CellSnapshot> {
    let value = match value_type.as_deref() {
        Some("number") => CellValue::Number(value_number.unwrap_or(0.0)),
        Some("string") => CellValue::Text(value_string.unwrap_or_default()),
        Some("boolean") => CellValue::Boolean(value_number.unwrap_or(0.0) != 0.0),
        Some("error") => CellValue::Error(value_string.unwrap_or_default()),
        Some("formula") => CellValue::Empty,
        // `NULL` value_type means a style-only blank cell.
        None => CellValue::Empty,
        Some(other) => {
            // Unknown value types are treated as strings to preserve data.
            CellValue::Text(other.to_string())
        }
    };

    Ok(CellSnapshot {
        value,
        formula,
        style_id,
    })
}

fn apply_one_change(tx: &Transaction<'_>, change: &CellChange) -> Result<()> {
    // Read previous state for change log.
    let old_snapshot = fetch_cell_snapshot_tx(tx, change.sheet_id, change.row, change.col)?;

    if change.data.is_truly_empty() {
        tx.execute(
            "DELETE FROM cells WHERE sheet_id = ?1 AND row = ?2 AND col = ?3",
            params![change.sheet_id.to_string(), change.row, change.col],
        )?;

        insert_change_log(
            tx,
            change,
            old_snapshot,
            None, // new
            "delete_cell",
        )?;

        touch_workbook_modified_at(tx, change.sheet_id)?;
        return Ok(());
    }

    let style_id = match &change.data.style {
        Some(style) => Some(get_or_insert_style_tx(tx, style)?),
        None => None,
    };

    let (value_type, value_number, value_string, formula) = if let Some(formula) = &change.data.formula {
        (
            Some("formula".to_string()),
            None,
            None,
            Some(formula.to_string()),
        )
    } else {
        match &change.data.value {
            CellValue::Empty => (None, None, None, None),
            CellValue::Number(n) => (Some("number".to_string()), Some(*n), None, None),
            CellValue::Text(s) => (Some("string".to_string()), None, Some(s.to_string()), None),
            CellValue::Boolean(b) => (
                Some("boolean".to_string()),
                Some(if *b { 1.0 } else { 0.0 }),
                None,
                None,
            ),
            CellValue::Error(e) => (Some("error".to_string()), None, Some(e.to_string()), None),
        }
    };

    tx.execute(
        r#"
        INSERT INTO cells (
          sheet_id, row, col, value_type, value_number, value_string, formula, style_id
        ) VALUES (?1, ?2, ?3, ?4, ?5, ?6, ?7, ?8)
        ON CONFLICT(sheet_id, row, col) DO UPDATE SET
          value_type = excluded.value_type,
          value_number = excluded.value_number,
          value_string = excluded.value_string,
          formula = excluded.formula,
          style_id = excluded.style_id
        "#,
        params![
            change.sheet_id.to_string(),
            change.row,
            change.col,
            value_type,
            value_number,
            value_string,
            formula,
            style_id
        ],
    )?;

    let snapshot_value = if change.data.formula.is_some() {
        CellValue::Empty
    } else {
        change.data.value.clone()
    };
    let new_snapshot = Some(CellSnapshot {
        value: snapshot_value,
        formula: change.data.formula.clone(),
        style_id,
    });

    insert_change_log(tx, change, old_snapshot, new_snapshot, "set_cell")?;
    touch_workbook_modified_at(tx, change.sheet_id)?;
    Ok(())
}

fn fetch_cell_snapshot_tx(
    tx: &Transaction<'_>,
    sheet_id: Uuid,
    row: i64,
    col: i64,
) -> Result<Option<CellSnapshot>> {
    let row_opt = tx
        .query_row(
            r#"
            SELECT value_type, value_number, value_string, formula, style_id
            FROM cells
            WHERE sheet_id = ?1 AND row = ?2 AND col = ?3
            "#,
            params![sheet_id.to_string(), row, col],
            |r| {
                snapshot_from_row(r.get(0)?, r.get(1)?, r.get(2)?, r.get(3)?, r.get(4)?)
            },
        )
        .optional()?;

    Ok(row_opt)
}

fn insert_change_log(
    tx: &Transaction<'_>,
    change: &CellChange,
    old_snapshot: Option<CellSnapshot>,
    new_snapshot: Option<CellSnapshot>,
    operation: &str,
) -> Result<()> {
    let target = serde_json::to_value(json!({ "row": change.row, "col": change.col }))?;
    let old_value = serde_json::to_value(&old_snapshot)?;
    let new_value = serde_json::to_value(&new_snapshot)?;

    tx.execute(
        r#"
        INSERT INTO change_log (sheet_id, user_id, operation, target, old_value, new_value)
        VALUES (?1, ?2, ?3, ?4, ?5, ?6)
        "#,
        params![
            change.sheet_id.to_string(),
            change.user_id.as_deref(),
            operation,
            target,
            old_value, // stored as JSON because rusqlite serde_json feature
            new_value
        ],
    )?;

    Ok(())
}

fn touch_workbook_modified_at(tx: &Transaction<'_>, sheet_id: Uuid) -> Result<()> {
    tx.execute(
        r#"
        UPDATE workbooks
        SET modified_at = CURRENT_TIMESTAMP
        WHERE id = (SELECT workbook_id FROM sheets WHERE id = ?1)
        "#,
        params![sheet_id.to_string()],
    )?;
    Ok(())
}

fn get_or_insert_style_tx(tx: &Transaction<'_>, style: &Style) -> Result<i64> {
    let alignment = style.canonical_alignment();
    let protection = style.canonical_protection();

    let existing: Option<i64> = tx
        .query_row(
            r#"
            SELECT id
            FROM styles
            WHERE font_id IS ?1
              AND fill_id IS ?2
              AND border_id IS ?3
              AND number_format IS ?4
              AND alignment IS ?5
              AND protection IS ?6
            LIMIT 1
            "#,
            params![
                style.font_id,
                style.fill_id,
                style.border_id,
                style.number_format.as_deref(),
                alignment.clone(),
                protection.clone()
            ],
            |r| r.get(0),
        )
        .optional()?;

    if let Some(id) = existing {
        return Ok(id);
    }

    tx.execute(
        r#"
        INSERT INTO styles (font_id, fill_id, border_id, number_format, alignment, protection)
        VALUES (?1, ?2, ?3, ?4, ?5, ?6)
        "#,
        params![
            style.font_id,
            style.fill_id,
            style.border_id,
            style.number_format.as_deref(),
            alignment,
            protection
        ],
    )?;

    Ok(tx.last_insert_rowid())
}
