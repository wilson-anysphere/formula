use crate::schema;
use crate::types::{
    CellData, CellSnapshot, CellValue, NamedRange, SheetMeta, SheetVisibility, Style, WorkbookMeta,
};
use formula_model::{validate_sheet_name, ErrorValue, SheetNameError};
use rusqlite::{params, Connection, OpenFlags, OptionalExtension, Transaction};
use serde_json::json;
use std::path::Path;
use std::sync::{Arc, Mutex};
use std::time::Duration;
use thiserror::Error;
use unicode_normalization::UnicodeNormalization;
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
    #[error("sheet name cannot be empty")]
    EmptySheetName,
    #[error(transparent)]
    InvalidSheetName(SheetNameError),
    #[error("sheet name already exists: {0}")]
    DuplicateSheetName(String),
}

pub type Result<T> = std::result::Result<T, StorageError>;

fn map_sheet_name_error(err: SheetNameError) -> StorageError {
    match err {
        SheetNameError::EmptyName => StorageError::EmptySheetName,
        other => StorageError::InvalidSheetName(other),
    }
}

fn sheet_name_eq_case_insensitive(a: &str, b: &str) -> bool {
    a.nfkc()
        .flat_map(|c| c.to_uppercase())
        .eq(b.nfkc().flat_map(|c| c.to_uppercase()))
}

#[derive(Debug, Clone)]
pub struct Storage {
    conn: Arc<Mutex<Connection>>,
}

#[derive(Debug, Clone, PartialEq)]
pub struct ChangeLogEntry {
    pub id: i64,
    pub sheet_id: Uuid,
    pub user_id: Option<String>,
    pub operation: String,
    pub target: serde_json::Value,
    pub old_value: serde_json::Value,
    pub new_value: serde_json::Value,
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
        let mut conn = Connection::open(path)?;
        conn.busy_timeout(Duration::from_secs(5))?;
        schema::init(&mut conn)?;
        Ok(Self {
            conn: Arc::new(Mutex::new(conn)),
        })
    }

    pub fn open_in_memory() -> Result<Self> {
        let mut conn = Connection::open_in_memory()?;
        conn.busy_timeout(Duration::from_secs(5))?;
        schema::init(&mut conn)?;
        Ok(Self {
            conn: Arc::new(Mutex::new(conn)),
        })
    }

    pub fn open_uri(uri: &str) -> Result<Self> {
        let flags = OpenFlags::SQLITE_OPEN_READ_WRITE
            | OpenFlags::SQLITE_OPEN_CREATE
            | OpenFlags::SQLITE_OPEN_URI;
        let mut conn = Connection::open_with_flags(uri, flags)?;
        conn.busy_timeout(Duration::from_secs(5))?;
        schema::init(&mut conn)?;
        Ok(Self {
            conn: Arc::new(Mutex::new(conn)),
        })
    }

    pub fn create_workbook(
        &self,
        name: &str,
        metadata: Option<serde_json::Value>,
    ) -> Result<WorkbookMeta> {
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

        validate_sheet_name(name).map_err(map_sheet_name_error)?;

        let conn = self.conn.lock().expect("storage mutex poisoned");
        let workbook_id_str = workbook_id.to_string();
        {
            let mut stmt = conn.prepare("SELECT name FROM sheets WHERE workbook_id = ?1")?;
            let mut rows = stmt.query(params![&workbook_id_str])?;
            while let Some(row) = rows.next()? {
                let existing: String = row.get(0)?;
                if sheet_name_eq_case_insensitive(&existing, name) {
                    return Err(StorageError::DuplicateSheetName(name.to_string()));
                }
            }
        }

        let sheet = SheetMeta {
            id: Uuid::new_v4(),
            workbook_id,
            name: name.to_string(),
            position,
            visibility: SheetVisibility::Visible,
            tab_color: None,
            xlsx_sheet_id: None,
            xlsx_rel_id: None,
            frozen_rows: 0,
            frozen_cols: 0,
            zoom: 1.0,
            metadata,
        };

        conn.execute(
            r#"
            INSERT INTO sheets (
              id,
              workbook_id,
              name,
              position,
              visibility,
              tab_color,
              xlsx_sheet_id,
              xlsx_rel_id,
              frozen_rows,
              frozen_cols,
              zoom,
              metadata
            ) VALUES (?1, ?2, ?3, ?4, ?5, ?6, ?7, ?8, ?9, ?10, ?11, ?12)
            "#,
            params![
                sheet.id.to_string(),
                workbook_id_str,
                &sheet.name,
                sheet.position,
                sheet.visibility.as_str(),
                sheet.tab_color.clone(),
                sheet.xlsx_sheet_id,
                sheet.xlsx_rel_id.clone(),
                sheet.frozen_rows,
                sheet.frozen_cols,
                sheet.zoom,
                sheet.metadata.clone()
            ],
        )?;

        // A new sheet changes workbook metadata.
        conn.execute(
            "UPDATE workbooks SET modified_at = CURRENT_TIMESTAMP WHERE id = ?1",
            params![sheet.workbook_id.to_string()],
        )?;

        Ok(sheet)
    }

    pub fn list_sheets(&self, workbook_id: Uuid) -> Result<Vec<SheetMeta>> {
        let conn = self.conn.lock().expect("storage mutex poisoned");
        let mut stmt = conn.prepare(
            r#"
            SELECT
              id,
              workbook_id,
              name,
              position,
              visibility,
              tab_color,
              xlsx_sheet_id,
              xlsx_rel_id,
              frozen_rows,
              frozen_cols,
              zoom,
              metadata
            FROM sheets
            WHERE workbook_id = ?1
            ORDER BY position
            "#,
        )?;

        let rows = stmt.query_map(params![workbook_id.to_string()], |r| {
            let id: String = r.get(0)?;
            let workbook_id: String = r.get(1)?;
            let visibility: String = r.get(4)?;
            Ok(SheetMeta {
                id: Uuid::parse_str(&id).map_err(|_| rusqlite::Error::InvalidQuery)?,
                workbook_id: Uuid::parse_str(&workbook_id)
                    .map_err(|_| rusqlite::Error::InvalidQuery)?,
                name: r.get(2)?,
                position: r.get(3)?,
                visibility: SheetVisibility::parse(&visibility),
                tab_color: r.get(5)?,
                xlsx_sheet_id: r.get(6)?,
                xlsx_rel_id: r.get(7)?,
                frozen_rows: r.get(8)?,
                frozen_cols: r.get(9)?,
                zoom: r.get(10)?,
                metadata: r.get(11)?,
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
                SELECT
                  id,
                  workbook_id,
                  name,
                  position,
                  visibility,
                  tab_color,
                  xlsx_sheet_id,
                  xlsx_rel_id,
                  frozen_rows,
                  frozen_cols,
                  zoom,
                  metadata
                FROM sheets
                WHERE id = ?1
                "#,
                params![sheet_id.to_string()],
                |r| {
                    let id: String = r.get(0)?;
                    let workbook_id: String = r.get(1)?;
                    let visibility: String = r.get(4)?;
                    Ok(SheetMeta {
                        id: Uuid::parse_str(&id).map_err(|_| rusqlite::Error::InvalidQuery)?,
                        workbook_id: Uuid::parse_str(&workbook_id)
                            .map_err(|_| rusqlite::Error::InvalidQuery)?,
                        name: r.get(2)?,
                        position: r.get(3)?,
                        visibility: SheetVisibility::parse(&visibility),
                        tab_color: r.get(5)?,
                        xlsx_sheet_id: r.get(6)?,
                        xlsx_rel_id: r.get(7)?,
                        frozen_rows: r.get(8)?,
                        frozen_cols: r.get(9)?,
                        zoom: r.get(10)?,
                        metadata: r.get(11)?,
                    })
                },
            )
            .optional()?;

        row.ok_or(StorageError::SheetNotFound(sheet_id))
    }

    /// Rename a worksheet.
    pub fn rename_sheet(&self, sheet_id: Uuid, name: &str) -> Result<()> {
        validate_sheet_name(name).map_err(map_sheet_name_error)?;

        let mut conn = self.conn.lock().expect("storage mutex poisoned");
        let tx = conn.transaction()?;

        let meta = self.get_sheet_meta_tx(&tx, sheet_id)?;

        // Enforce Excel-style uniqueness (Unicode-aware, case-insensitive) within the workbook.
        {
            let mut stmt = tx.prepare("SELECT name FROM sheets WHERE workbook_id = ?1 AND id != ?2")?;
            let mut rows = stmt.query(params![meta.workbook_id.to_string(), sheet_id.to_string()])?;
            while let Some(row) = rows.next()? {
                let existing: String = row.get(0)?;
                if sheet_name_eq_case_insensitive(&existing, name) {
                    return Err(StorageError::DuplicateSheetName(name.to_string()));
                }
            }
        }

        tx.execute(
            "UPDATE sheets SET name = ?1 WHERE id = ?2",
            params![name, sheet_id.to_string()],
        )?;

        touch_workbook_modified_at_by_workbook_id(&tx, meta.workbook_id)?;
        tx.commit()?;
        Ok(())
    }

    pub fn set_sheet_visibility(&self, sheet_id: Uuid, visibility: SheetVisibility) -> Result<()> {
        let mut conn = self.conn.lock().expect("storage mutex poisoned");
        let tx = conn.transaction()?;
        let meta = self.get_sheet_meta_tx(&tx, sheet_id)?;
        tx.execute(
            "UPDATE sheets SET visibility = ?1 WHERE id = ?2",
            params![visibility.as_str(), sheet_id.to_string()],
        )?;
        touch_workbook_modified_at_by_workbook_id(&tx, meta.workbook_id)?;
        tx.commit()?;
        Ok(())
    }

    pub fn set_sheet_tab_color(&self, sheet_id: Uuid, tab_color: Option<&str>) -> Result<()> {
        let mut conn = self.conn.lock().expect("storage mutex poisoned");
        let tx = conn.transaction()?;
        let meta = self.get_sheet_meta_tx(&tx, sheet_id)?;
        tx.execute(
            "UPDATE sheets SET tab_color = ?1 WHERE id = ?2",
            params![tab_color, sheet_id.to_string()],
        )?;
        touch_workbook_modified_at_by_workbook_id(&tx, meta.workbook_id)?;
        tx.commit()?;
        Ok(())
    }

    pub fn set_sheet_xlsx_metadata(
        &self,
        sheet_id: Uuid,
        xlsx_sheet_id: Option<i64>,
        xlsx_rel_id: Option<&str>,
    ) -> Result<()> {
        let mut conn = self.conn.lock().expect("storage mutex poisoned");
        let tx = conn.transaction()?;
        let meta = self.get_sheet_meta_tx(&tx, sheet_id)?;
        tx.execute(
            "UPDATE sheets SET xlsx_sheet_id = ?1, xlsx_rel_id = ?2 WHERE id = ?3",
            params![xlsx_sheet_id, xlsx_rel_id, sheet_id.to_string()],
        )?;
        touch_workbook_modified_at_by_workbook_id(&tx, meta.workbook_id)?;
        tx.commit()?;
        Ok(())
    }

    /// Reorder a sheet within its workbook by setting its 0-based position.
    ///
    /// This renormalizes positions to be contiguous starting at 0.
    pub fn reorder_sheet(&self, sheet_id: Uuid, new_position: i64) -> Result<()> {
        let mut conn = self.conn.lock().expect("storage mutex poisoned");
        let tx = conn.transaction()?;

        let meta = self.get_sheet_meta_tx(&tx, sheet_id)?;

        let mut sheets = self.list_sheets_tx(&tx, meta.workbook_id)?;
        if sheets.len() <= 1 {
            return Ok(());
        }

        let current_index = sheets
            .iter()
            .position(|s| s.id == sheet_id)
            .ok_or(StorageError::SheetNotFound(sheet_id))?;

        let sheet = sheets.remove(current_index);
        let clamped = new_position.max(0).min(sheets.len() as i64) as usize;
        sheets.insert(clamped, sheet);

        for (idx, sheet) in sheets.iter().enumerate() {
            tx.execute(
                "UPDATE sheets SET position = ?1 WHERE id = ?2",
                params![idx as i64, sheet.id.to_string()],
            )?;
        }

        touch_workbook_modified_at_by_workbook_id(&tx, meta.workbook_id)?;
        tx.commit()?;
        Ok(())
    }

    pub fn delete_sheet(&self, sheet_id: Uuid) -> Result<()> {
        let mut conn = self.conn.lock().expect("storage mutex poisoned");
        let tx = conn.transaction()?;

        let meta = self.get_sheet_meta_tx(&tx, sheet_id)?;

        tx.execute(
            "DELETE FROM cells WHERE sheet_id = ?1",
            params![sheet_id.to_string()],
        )?;
        tx.execute(
            "DELETE FROM change_log WHERE sheet_id = ?1",
            params![sheet_id.to_string()],
        )?;
        tx.execute(
            "DELETE FROM sheets WHERE id = ?1",
            params![sheet_id.to_string()],
        )?;

        // Renormalize remaining sheet positions.
        let sheets = self.list_sheets_tx(&tx, meta.workbook_id)?;
        for (idx, sheet) in sheets.iter().enumerate() {
            tx.execute(
                "UPDATE sheets SET position = ?1 WHERE id = ?2",
                params![idx as i64, sheet.id.to_string()],
            )?;
        }

        touch_workbook_modified_at_by_workbook_id(&tx, meta.workbook_id)?;
        tx.commit()?;
        Ok(())
    }

    /// Load all non-empty cells within an inclusive range.
    pub fn load_cells_in_range(
        &self,
        sheet_id: Uuid,
        range: CellRange,
    ) -> Result<Vec<((i64, i64), CellSnapshot)>> {
        let conn = self.conn.lock().expect("storage mutex poisoned");
        let mut stmt = conn.prepare(
            r#"
            SELECT row, col, value_type, value_number, value_string, value_json, formula, style_id
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
                let snapshot = snapshot_from_row(
                    r.get(2)?,
                    r.get(3)?,
                    r.get(4)?,
                    r.get(5)?,
                    r.get(6)?,
                    r.get(7)?,
                )?;
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

    pub fn change_log_count(&self, sheet_id: Uuid) -> Result<u64> {
        let conn = self.conn.lock().expect("storage mutex poisoned");
        let count: u64 = conn.query_row(
            "SELECT COUNT(*) FROM change_log WHERE sheet_id = ?1",
            params![sheet_id.to_string()],
            |r| r.get(0),
        )?;
        Ok(count)
    }

    pub fn latest_change(&self, sheet_id: Uuid) -> Result<Option<ChangeLogEntry>> {
        let conn = self.conn.lock().expect("storage mutex poisoned");
        let row = conn
            .query_row(
                r#"
                SELECT id, sheet_id, user_id, operation, target, old_value, new_value
                FROM change_log
                WHERE sheet_id = ?1
                ORDER BY id DESC
                LIMIT 1
                "#,
                params![sheet_id.to_string()],
                |r| {
                    let sheet_id: String = r.get(1)?;
                    Ok(ChangeLogEntry {
                        id: r.get(0)?,
                        sheet_id: Uuid::parse_str(&sheet_id)
                            .map_err(|_| rusqlite::Error::InvalidQuery)?,
                        user_id: r.get(2)?,
                        operation: r.get(3)?,
                        target: r.get(4)?,
                        old_value: r.get(5)?,
                        new_value: r.get(6)?,
                    })
                },
            )
            .optional()?;
        Ok(row)
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

    pub fn get_named_range(
        &self,
        workbook_id: Uuid,
        name: &str,
        scope: &str,
    ) -> Result<Option<NamedRange>> {
        let conn = self.conn.lock().expect("storage mutex poisoned");
        let row = conn
            .query_row(
                r#"
                SELECT workbook_id, name, scope, reference
                FROM named_ranges
                WHERE workbook_id = ?1
                  AND name = ?2 COLLATE NOCASE
                  AND scope = ?3 COLLATE NOCASE
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
        let mut conn = self.conn.lock().expect("storage mutex poisoned");
        let tx = conn.transaction()?;

        let existing: Option<(String, String)> = tx
            .query_row(
                r#"
                SELECT name, scope
                FROM named_ranges
                WHERE workbook_id = ?1
                  AND name = ?2 COLLATE NOCASE
                  AND scope = ?3 COLLATE NOCASE
                LIMIT 1
                "#,
                params![range.workbook_id.to_string(), &range.name, &range.scope],
                |r| Ok((r.get(0)?, r.get(1)?)),
            )
            .optional()?;

        match existing {
            Some((name, scope)) => {
                tx.execute(
                    r#"
                    UPDATE named_ranges
                    SET reference = ?1
                    WHERE workbook_id = ?2 AND name = ?3 AND scope = ?4
                    "#,
                    params![&range.reference, range.workbook_id.to_string(), name, scope],
                )?;
            }
            None => {
                tx.execute(
                    r#"
                    INSERT INTO named_ranges (workbook_id, name, scope, reference)
                    VALUES (?1, ?2, ?3, ?4)
                    "#,
                    params![
                        range.workbook_id.to_string(),
                        &range.name,
                        &range.scope,
                        &range.reference
                    ],
                )?;
            }
        }

        tx.commit()?;
        Ok(())
    }

    fn list_sheets_tx(&self, tx: &Transaction<'_>, workbook_id: Uuid) -> Result<Vec<SheetMeta>> {
        let mut stmt = tx.prepare(
            r#"
            SELECT
              id,
              workbook_id,
              name,
              position,
              visibility,
              tab_color,
              xlsx_sheet_id,
              xlsx_rel_id,
              frozen_rows,
              frozen_cols,
              zoom,
              metadata
            FROM sheets
            WHERE workbook_id = ?1
            ORDER BY position
            "#,
        )?;

        let rows = stmt.query_map(params![workbook_id.to_string()], |r| {
            let id: String = r.get(0)?;
            let workbook_id: String = r.get(1)?;
            let visibility: String = r.get(4)?;
            Ok(SheetMeta {
                id: Uuid::parse_str(&id).map_err(|_| rusqlite::Error::InvalidQuery)?,
                workbook_id: Uuid::parse_str(&workbook_id)
                    .map_err(|_| rusqlite::Error::InvalidQuery)?,
                name: r.get(2)?,
                position: r.get(3)?,
                visibility: SheetVisibility::parse(&visibility),
                tab_color: r.get(5)?,
                xlsx_sheet_id: r.get(6)?,
                xlsx_rel_id: r.get(7)?,
                frozen_rows: r.get(8)?,
                frozen_cols: r.get(9)?,
                zoom: r.get(10)?,
                metadata: r.get(11)?,
            })
        })?;

        let mut sheets = Vec::new();
        for sheet in rows {
            sheets.push(sheet?);
        }
        Ok(sheets)
    }

    fn get_sheet_meta_tx(&self, tx: &Transaction<'_>, sheet_id: Uuid) -> Result<SheetMeta> {
        let row = tx
            .query_row(
                r#"
                SELECT
                  id,
                  workbook_id,
                  name,
                  position,
                  visibility,
                  tab_color,
                  xlsx_sheet_id,
                  xlsx_rel_id,
                  frozen_rows,
                  frozen_cols,
                  zoom,
                  metadata
                FROM sheets
                WHERE id = ?1
                "#,
                params![sheet_id.to_string()],
                |r| {
                    let id: String = r.get(0)?;
                    let workbook_id: String = r.get(1)?;
                    let visibility: String = r.get(4)?;
                    Ok(SheetMeta {
                        id: Uuid::parse_str(&id).map_err(|_| rusqlite::Error::InvalidQuery)?,
                        workbook_id: Uuid::parse_str(&workbook_id)
                            .map_err(|_| rusqlite::Error::InvalidQuery)?,
                        name: r.get(2)?,
                        position: r.get(3)?,
                        visibility: SheetVisibility::parse(&visibility),
                        tab_color: r.get(5)?,
                        xlsx_sheet_id: r.get(6)?,
                        xlsx_rel_id: r.get(7)?,
                        frozen_rows: r.get(8)?,
                        frozen_cols: r.get(9)?,
                        zoom: r.get(10)?,
                        metadata: r.get(11)?,
                    })
                },
            )
            .optional()?;

        row.ok_or(StorageError::SheetNotFound(sheet_id))
    }
}

fn snapshot_from_row(
    value_type: Option<String>,
    value_number: Option<f64>,
    value_string: Option<String>,
    value_json: Option<String>,
    formula: Option<String>,
    style_id: Option<i64>,
) -> rusqlite::Result<CellSnapshot> {
    let value = if let Some(raw_json) = value_json.as_deref().filter(|s| !s.trim().is_empty()) {
        serde_json::from_str::<CellValue>(raw_json).map_err(|err| {
            rusqlite::Error::FromSqlConversionFailure(0, rusqlite::types::Type::Text, Box::new(err))
        })?
    } else {
        match value_type.as_deref() {
            Some("number") => CellValue::Number(value_number.unwrap_or(0.0)),
            Some("string") => CellValue::String(value_string.unwrap_or_default()),
            Some("boolean") => CellValue::Boolean(value_number.unwrap_or(0.0) != 0.0),
            Some("error") => {
                let legacy = value_string.unwrap_or_default();
                let parsed = legacy.parse::<ErrorValue>().unwrap_or(ErrorValue::Unknown);
                CellValue::Error(parsed)
            }
            // Legacy sentinel used by older schema versions when a cell contains a formula but no cached value.
            Some("formula") => CellValue::Empty,
            // `NULL` value_type means a style-only blank cell.
            None => CellValue::Empty,
            Some(other) => {
                // Unknown value types are treated as strings to preserve data.
                CellValue::String(other.to_string())
            }
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

    let formula = change.data.formula.clone();
    let (value_type, value_number, value_string, value_json) = match &change.data.value {
        CellValue::Empty => (None, None, None, None),
        CellValue::Number(n) => (Some("number".to_string()), Some(*n), None, None),
        CellValue::String(s) => (Some("string".to_string()), None, Some(s.to_string()), None),
        CellValue::Boolean(b) => (
            Some("boolean".to_string()),
            Some(if *b { 1.0 } else { 0.0 }),
            None,
            None,
        ),
        CellValue::Error(err) => (
            Some("error".to_string()),
            None,
            Some(err.as_str().to_string()),
            None,
        ),
        value @ (CellValue::RichText(_) | CellValue::Array(_) | CellValue::Spill(_)) => {
            (None, None, None, Some(serde_json::to_string(value)?))
        }
    };

    tx.execute(
        r#"
        INSERT INTO cells (
          sheet_id, row, col, value_type, value_number, value_string, value_json, formula, style_id
        ) VALUES (?1, ?2, ?3, ?4, ?5, ?6, ?7, ?8, ?9)
        ON CONFLICT(sheet_id, row, col) DO UPDATE SET
          value_type = excluded.value_type,
          value_number = excluded.value_number,
          value_string = excluded.value_string,
          value_json = excluded.value_json,
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
            value_json,
            formula,
            style_id
        ],
    )?;

    let new_snapshot = Some(CellSnapshot {
        value: change.data.value.clone(),
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
            SELECT value_type, value_number, value_string, value_json, formula, style_id
            FROM cells
            WHERE sheet_id = ?1 AND row = ?2 AND col = ?3
            "#,
            params![sheet_id.to_string(), row, col],
            |r| {
                snapshot_from_row(
                    r.get(0)?,
                    r.get(1)?,
                    r.get(2)?,
                    r.get(3)?,
                    r.get(4)?,
                    r.get(5)?,
                )
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

fn touch_workbook_modified_at_by_workbook_id(
    tx: &Transaction<'_>,
    workbook_id: Uuid,
) -> Result<()> {
    tx.execute(
        "UPDATE workbooks SET modified_at = CURRENT_TIMESTAMP WHERE id = ?1",
        params![workbook_id.to_string()],
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
