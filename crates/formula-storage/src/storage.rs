use crate::schema;
use crate::encryption::{
    decrypt_sqlite_bytes, encrypt_sqlite_bytes, is_encrypted_container, load_or_create_keyring,
    KeyProvider,
};
use crate::types::{
    canonical_json, CellData, CellSnapshot, CellValue, ImportModelWorkbookOptions, NamedRange,
    SheetMeta, SheetVisibility, Style as StorageStyle, WorkbookMeta,
};
use formula_model::{validate_sheet_name, ErrorValue, SheetNameError};
use rusqlite::{params, Connection, DatabaseName, OpenFlags, OptionalExtension, Transaction};
use serde::de::DeserializeOwned;
use serde_json::json;
use std::collections::{HashMap, HashSet};
use std::fs;
use std::io::{Read, Write};
use std::path::{Path, PathBuf};
use std::ptr;
use std::sync::{Arc, Mutex};
use std::time::Duration;
use thiserror::Error;
use tempfile::NamedTempFile;
use unicode_normalization::UnicodeNormalization;
use uuid::Uuid;

#[derive(Debug, Error)]
pub enum StorageError {
    #[error("sqlite error: {0}")]
    Sqlite(#[from] rusqlite::Error),
    #[error("json error: {0}")]
    Json(#[from] serde_json::Error),
    #[error("io error: {0}")]
    Io(#[from] std::io::Error),
    #[error(transparent)]
    Encryption(#[from] crate::encryption::EncryptionError),
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

#[derive(Clone)]
pub struct Storage {
    conn: Arc<Mutex<Connection>>,
    encrypted: Option<Arc<EncryptedStorageContext>>,
}

struct EncryptedStorageContext {
    path: PathBuf,
    key_provider: Arc<dyn KeyProvider>,
}

impl std::fmt::Debug for EncryptedStorageContext {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("EncryptedStorageContext")
            .field("path", &self.path)
            .finish_non_exhaustive()
    }
}

impl std::fmt::Debug for Storage {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        let mut s = f.debug_struct("Storage");
        s.field("conn", &self.conn);
        if let Some(ctx) = &self.encrypted {
            s.field("encrypted_path", &Some(&ctx.path));
        } else {
            s.field("encrypted_path", &Option::<&PathBuf>::None);
        }
        s.finish()
    }
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
            encrypted: None,
        })
    }

    pub fn open_in_memory() -> Result<Self> {
        let mut conn = Connection::open_in_memory()?;
        conn.busy_timeout(Duration::from_secs(5))?;
        schema::init(&mut conn)?;
        Ok(Self {
            conn: Arc::new(Mutex::new(conn)),
            encrypted: None,
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
            encrypted: None,
        })
    }

    /// Open (or create) a workbook at `path`, encrypting the persisted bytes with AES-256-GCM.
    ///
    /// This keeps the live SQLite database in-memory; callers must invoke [`Storage::persist`]
    /// to flush the encrypted container to disk.
    pub fn open_encrypted_path(
        path: impl AsRef<Path>,
        key_provider: Arc<dyn KeyProvider>,
    ) -> Result<Self> {
        let path = path.as_ref().to_path_buf();
        let mut conn = Connection::open_in_memory()?;
        conn.busy_timeout(Duration::from_secs(5))?;

        match fs::metadata(&path) {
            Ok(metadata) => {
                if metadata.len() > 0 {
                    let mut prefix = [0u8; 8];
                    let read_len = {
                        let mut file = fs::File::open(&path)?;
                        file.read(&mut prefix)?
                    };

                    if is_encrypted_container(&prefix[..read_len]) {
                        let bytes = fs::read(&path)?;
                        let keyring = load_or_create_keyring(key_provider.as_ref(), false)?;
                        let sqlite_bytes = decrypt_sqlite_bytes(&bytes, &keyring)?;
                        load_sqlite_bytes_into_connection(&mut conn, &sqlite_bytes)?;
                    } else {
                        let flags = OpenFlags::SQLITE_OPEN_READ_ONLY;
                        let src = Connection::open_with_flags(&path, flags)?;
                        let sqlite_bytes = export_connection_to_sqlite_bytes(&src)?;
                        load_sqlite_bytes_into_connection(&mut conn, &sqlite_bytes)?;
                    }
                }
            }
            Err(err) if err.kind() == std::io::ErrorKind::NotFound => {}
            Err(err) => return Err(err.into()),
        }

        schema::init(&mut conn)?;

        // Ensure a keyring exists for new/plaintext workbooks so the first persist can encrypt.
        let _ = load_or_create_keyring(key_provider.as_ref(), true)?;

        Ok(Self {
            conn: Arc::new(Mutex::new(conn)),
            encrypted: Some(Arc::new(EncryptedStorageContext { path, key_provider })),
        })
    }

    /// Persist the workbook to disk if opened in encrypted mode.
    ///
    /// Plaintext-backed storages (`open_path`, `open_uri`, `open_in_memory`) are already persisted
    /// by SQLite itself, so this is a no-op for them.
    pub fn persist(&self) -> Result<()> {
        let Some(ctx) = self.encrypted.as_ref() else {
            return Ok(());
        };

        let sqlite_bytes = {
            let conn = self.conn.lock().expect("storage mutex poisoned");
            export_connection_to_sqlite_bytes(&conn)?
        };

        let keyring = load_or_create_keyring(ctx.key_provider.as_ref(), true)?;
        let encrypted_bytes = encrypt_sqlite_bytes(&sqlite_bytes, &keyring)?;

        atomic_write(&ctx.path, &encrypted_bytes)?;
        cleanup_sqlite_sidecar_files(&ctx.path)?;
        Ok(())
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

    pub fn import_model_workbook(
        &self,
        workbook: &formula_model::Workbook,
        opts: ImportModelWorkbookOptions,
    ) -> Result<WorkbookMeta> {
        let workbook_meta = WorkbookMeta {
            id: Uuid::new_v4(),
            name: opts.name,
            metadata: opts.metadata,
        };

        let mut conn = self.conn.lock().expect("storage mutex poisoned");
        let tx = conn.transaction()?;

        tx.execute(
            r#"
            INSERT INTO workbooks (
              id,
              name,
              metadata,
              model_schema_version,
              model_workbook_id,
              date_system,
              calc_settings
            ) VALUES (?1, ?2, ?3, ?4, ?5, ?6, ?7)
            "#,
            params![
                workbook_meta.id.to_string(),
                &workbook_meta.name,
                workbook_meta.metadata.clone(),
                workbook.schema_version as i64,
                workbook.id as i64,
                date_system_to_str(workbook.date_system),
                serde_json::to_value(&workbook.calc_settings)?
            ],
        )?;

        // Insert worksheets in the order they appear in the model workbook.
        let mut seen_names: Vec<String> = Vec::with_capacity(workbook.sheets.len());
        let mut model_sheet_to_storage: HashMap<u32, Uuid> = HashMap::new();
        for (position, sheet) in workbook.sheets.iter().enumerate() {
            validate_sheet_name(&sheet.name).map_err(map_sheet_name_error)?;
            if seen_names
                .iter()
                .any(|existing| sheet_name_eq_case_insensitive(existing, &sheet.name))
            {
                return Err(StorageError::DuplicateSheetName(sheet.name.clone()));
            }
            seen_names.push(sheet.name.clone());

            let storage_sheet_id = Uuid::new_v4();
            model_sheet_to_storage.insert(sheet.id, storage_sheet_id);

            let visibility = model_sheet_visibility_to_storage(sheet.visibility);
            let tab_color_fast = sheet.tab_color.as_ref().and_then(|c| c.rgb.clone());
            let tab_color_json = sheet
                .tab_color
                .as_ref()
                .map(|c| serde_json::to_value(c))
                .transpose()?;

            tx.execute(
                r#"
                INSERT INTO sheets (
                  id,
                  workbook_id,
                  name,
                  position,
                  visibility,
                  tab_color,
                  tab_color_json,
                  xlsx_sheet_id,
                  xlsx_rel_id,
                  frozen_rows,
                  frozen_cols,
                  zoom,
                  metadata,
                  model_sheet_id
                ) VALUES (
                  ?1, ?2, ?3, ?4, ?5, ?6, ?7, ?8, ?9, ?10, ?11, ?12, ?13, ?14
                )
                "#,
                params![
                    storage_sheet_id.to_string(),
                    workbook_meta.id.to_string(),
                    &sheet.name,
                    position as i64,
                    visibility.as_str(),
                    tab_color_fast,
                    tab_color_json,
                    sheet.xlsx_sheet_id.map(|v| v as i64),
                    sheet.xlsx_rel_id.clone(),
                    sheet.frozen_rows as i64,
                    sheet.frozen_cols as i64,
                    sheet.zoom as f64,
                    Option::<serde_json::Value>::None,
                    sheet.id as i64
                ],
            )?;
        }

        // Persist styles (deduplicated globally), and preserve model style indices per workbook.
        let mut model_style_to_storage: Vec<Option<i64>> =
            vec![None; workbook.styles.styles.len().max(1)];
        for (style_index, style) in workbook.styles.styles.iter().enumerate().skip(1) {
            let storage_style_id = get_or_insert_model_style_tx(&tx, style)?;
            model_style_to_storage[style_index] = Some(storage_style_id);
            tx.execute(
                r#"
                INSERT INTO workbook_styles (workbook_id, style_index, style_id)
                VALUES (?1, ?2, ?3)
                "#,
                params![
                    workbook_meta.id.to_string(),
                    style_index as i64,
                    storage_style_id
                ],
            )?;
        }

        // Stream cell inserts without building an intermediate change list.
        {
            let mut stmt = tx.prepare(
                r#"
                INSERT INTO cells (
                  sheet_id,
                  row,
                  col,
                  value_type,
                  value_number,
                  value_string,
                  value_json,
                  formula,
                  style_id
                ) VALUES (?1, ?2, ?3, ?4, ?5, ?6, ?7, ?8, ?9)
                "#,
            )?;

            for sheet in &workbook.sheets {
                let Some(storage_sheet_id) = model_sheet_to_storage.get(&sheet.id) else {
                    continue;
                };
                for (cell_ref, cell) in sheet.iter_cells() {
                    if cell.is_truly_empty() {
                        continue;
                    }

                    let storage_style_id = model_style_to_storage
                        .get(cell.style_id as usize)
                        .copied()
                        .flatten();

                    let value_json = serde_json::to_string(&cell.value)?;
                    let (value_type, value_number, value_string) =
                        cell_value_fast_path(&cell.value);

                    stmt.execute(params![
                        storage_sheet_id.to_string(),
                        cell_ref.row as i64,
                        cell_ref.col as i64,
                        value_type,
                        value_number,
                        value_string,
                        value_json,
                        cell.formula.as_deref(),
                        storage_style_id
                    ])?;
                }
            }
        }

        tx.commit()?;
        Ok(workbook_meta)
    }

    pub fn export_model_workbook(&self, workbook_id: Uuid) -> Result<formula_model::Workbook> {
        let conn = self.conn.lock().expect("storage mutex poisoned");

        let row = conn
            .query_row(
                r#"
                SELECT model_schema_version, model_workbook_id, date_system, calc_settings
                FROM workbooks
                WHERE id = ?1
                "#,
                params![workbook_id.to_string()],
                |r| {
                    Ok((
                        r.get::<_, Option<i64>>(0)?,
                        r.get::<_, Option<i64>>(1)?,
                        r.get::<_, Option<String>>(2)?,
                        r.get::<_, Option<serde_json::Value>>(3)?,
                    ))
                },
            )
            .optional()?;

        let Some((model_schema_version, model_workbook_id, date_system, calc_settings)) = row
        else {
            return Err(StorageError::WorkbookNotFound(workbook_id));
        };

        let (styles, storage_style_to_model) = build_model_style_table(&conn, workbook_id)?;

        let mut model_workbook = formula_model::Workbook::new();
        if let Some(schema_version) = model_schema_version {
            model_workbook.schema_version = schema_version as u32;
        }
        if let Some(id) = model_workbook_id {
            model_workbook.id = id as u32;
        }
        if let Some(date_system) = date_system {
            model_workbook.date_system = parse_date_system(&date_system)?;
        }
        if let Some(calc_settings) = calc_settings {
            model_workbook.calc_settings = serde_json::from_value(calc_settings)?;
        }
        model_workbook.styles = styles;

        // Allocate deterministic worksheet ids for sheets that predate model import.
        let mut used_sheet_ids: HashSet<u32> = HashSet::new();
        let max_model_sheet_id: i64 = conn.query_row(
            "SELECT COALESCE(MAX(model_sheet_id), 0) FROM sheets WHERE workbook_id = ?1",
            params![workbook_id.to_string()],
            |r| r.get(0),
        )?;

        let mut sheets_stmt = conn.prepare(
            r#"
            SELECT
              id,
              name,
              position,
              visibility,
              tab_color,
              tab_color_json,
              xlsx_sheet_id,
              xlsx_rel_id,
              frozen_rows,
              frozen_cols,
              zoom,
              model_sheet_id
            FROM sheets
            WHERE workbook_id = ?1
            ORDER BY position
            "#,
        )?;
        let mut sheet_rows = sheets_stmt.query(params![workbook_id.to_string()])?;

        let mut worksheets: Vec<formula_model::Worksheet> = Vec::new();
        let mut next_generated_sheet_id: u32 = (max_model_sheet_id.max(0) as u32).wrapping_add(1);

        while let Some(row) = sheet_rows.next()? {
            let storage_sheet_id: String = row.get(0)?;
            let name: String = row.get(1)?;
            let visibility_raw: String = row.get(3)?;
            let tab_color_fast: Option<String> = row.get(4)?;
            let tab_color_json: Option<serde_json::Value> = row.get(5)?;
            let xlsx_sheet_id: Option<i64> = row.get(6)?;
            let xlsx_rel_id: Option<String> = row.get(7)?;
            let frozen_rows: i64 = row.get(8)?;
            let frozen_cols: i64 = row.get(9)?;
            let zoom: f64 = row.get(10)?;
            let model_sheet_id: Option<i64> = row.get(11)?;

            let sheet_id = match model_sheet_id.map(|id| id.max(0) as u32) {
                Some(explicit) if !used_sheet_ids.contains(&explicit) => explicit,
                _ => {
                    while used_sheet_ids.contains(&next_generated_sheet_id) {
                        next_generated_sheet_id = next_generated_sheet_id.wrapping_add(1);
                    }
                    let out = next_generated_sheet_id;
                    next_generated_sheet_id = next_generated_sheet_id.wrapping_add(1);
                    out
                }
            };
            used_sheet_ids.insert(sheet_id);

            let mut sheet = formula_model::Worksheet::new(sheet_id, name.clone());
            sheet.visibility = storage_sheet_visibility_to_model(&visibility_raw);
            sheet.xlsx_sheet_id = xlsx_sheet_id.map(|v| v as u32);
            sheet.xlsx_rel_id = xlsx_rel_id;
            sheet.frozen_rows = frozen_rows.max(0) as u32;
            sheet.frozen_cols = frozen_cols.max(0) as u32;
            sheet.zoom = zoom as f32;
            sheet.view.pane.frozen_rows = sheet.frozen_rows;
            sheet.view.pane.frozen_cols = sheet.frozen_cols;
            sheet.view.zoom = sheet.zoom;

            sheet.tab_color = if let Some(raw) = tab_color_json {
                Some(serde_json::from_value(raw)?)
            } else {
                tab_color_fast.map(formula_model::TabColor::rgb)
            };

            stream_cells_into_model_sheet(
                &conn,
                &storage_sheet_id,
                &storage_style_to_model,
                &mut sheet,
            )?;

            worksheets.push(sheet);
        }

        model_workbook.sheets = worksheets;
        model_workbook.recompute_runtime_state();
        Ok(model_workbook)
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

    pub fn get_or_insert_style(&self, style: &StorageStyle) -> Result<i64> {
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
    // `cells.value_json` is the canonical representation for all cell values
    // (including scalar types), while the legacy scalar columns are kept as an
    // optional fast path.
    let value_json = Some(serde_json::to_string(&change.data.value)?);
    let (value_type, value_number, value_string) = match &change.data.value {
        CellValue::Empty => (None, None, None),
        CellValue::Number(n) => (Some("number".to_string()), Some(*n), None),
        CellValue::String(s) => (Some("string".to_string()), None, Some(s.to_string())),
        CellValue::Boolean(b) => (
            Some("boolean".to_string()),
            Some(if *b { 1.0 } else { 0.0 }),
            None,
        ),
        CellValue::Error(err) => (Some("error".to_string()), None, Some(err.as_str().to_string())),
        CellValue::RichText(_) | CellValue::Array(_) | CellValue::Spill(_) => (None, None, None),
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

fn get_or_insert_style_tx(tx: &Transaction<'_>, style: &StorageStyle) -> Result<i64> {
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

fn load_sqlite_bytes_into_connection(dst: &mut Connection, sqlite_bytes: &[u8]) -> Result<()> {
    if sqlite_bytes.is_empty() {
        return Ok(());
    }

    let sz: rusqlite::ffi::sqlite3_int64 = sqlite_bytes.len().try_into().map_err(|_| {
        rusqlite::Error::SqliteFailure(rusqlite::ffi::Error::new(rusqlite::ffi::SQLITE_TOOBIG), None)
    })?;

    let ptr = unsafe { rusqlite::ffi::sqlite3_malloc64(sqlite_bytes.len() as u64) }
        as *mut std::os::raw::c_uchar;
    if ptr.is_null() {
        return Err(rusqlite::Error::SqliteFailure(
            rusqlite::ffi::Error::new(rusqlite::ffi::SQLITE_NOMEM),
            None,
        )
        .into());
    }

    unsafe {
        ptr::copy_nonoverlapping(sqlite_bytes.as_ptr(), ptr.cast(), sqlite_bytes.len());
    }

    let schema = std::ffi::CString::new("main").expect("main schema name has no nul bytes");
    let flags = rusqlite::ffi::SQLITE_DESERIALIZE_FREEONCLOSE
        | rusqlite::ffi::SQLITE_DESERIALIZE_RESIZEABLE;
    let handle = unsafe { dst.handle() };
    let rc = unsafe { rusqlite::ffi::sqlite3_deserialize(handle, schema.as_ptr(), ptr, sz, sz, flags) };
    if rc != rusqlite::ffi::SQLITE_OK {
        unsafe {
            rusqlite::ffi::sqlite3_free(ptr.cast());
        }
        return Err(rusqlite::Error::SqliteFailure(rusqlite::ffi::Error::new(rc), None).into());
    }

    Ok(())
}

fn date_system_to_str(value: formula_model::DateSystem) -> &'static str {
    match value {
        formula_model::DateSystem::Excel1900 => "excel1900",
        formula_model::DateSystem::Excel1904 => "excel1904",
    }
}

fn parse_date_system(value: &str) -> Result<formula_model::DateSystem> {
    match value {
        "excel1900" => Ok(formula_model::DateSystem::Excel1900),
        "excel1904" => Ok(formula_model::DateSystem::Excel1904),
        other => Ok(serde_json::from_str::<formula_model::DateSystem>(&format!("\"{other}\""))?),
    }
}

fn model_sheet_visibility_to_storage(value: formula_model::SheetVisibility) -> SheetVisibility {
    match value {
        formula_model::SheetVisibility::Visible => SheetVisibility::Visible,
        formula_model::SheetVisibility::Hidden => SheetVisibility::Hidden,
        formula_model::SheetVisibility::VeryHidden => SheetVisibility::VeryHidden,
    }
}

fn storage_sheet_visibility_to_model(value: &str) -> formula_model::SheetVisibility {
    match value {
        "hidden" => formula_model::SheetVisibility::Hidden,
        "veryHidden" => formula_model::SheetVisibility::VeryHidden,
        _ => formula_model::SheetVisibility::Visible,
    }
}

fn cell_value_fast_path(value: &CellValue) -> (Option<String>, Option<f64>, Option<String>) {
    match value {
        CellValue::Empty => (None, None, None),
        CellValue::Number(n) => (Some("number".to_string()), Some(*n), None),
        CellValue::String(s) => (Some("string".to_string()), None, Some(s.to_string())),
        CellValue::Boolean(b) => (
            Some("boolean".to_string()),
            Some(if *b { 1.0 } else { 0.0 }),
            None,
        ),
        CellValue::Error(err) => (Some("error".to_string()), None, Some(err.as_str().to_string())),
        CellValue::RichText(_) | CellValue::Array(_) | CellValue::Spill(_) => (None, None, None),
    }
}

fn get_or_insert_style_component_tx(
    tx: &Transaction<'_>,
    table: &str,
    data: &serde_json::Value,
) -> Result<i64> {
    let key = canonical_json(data);
    tx.execute(
        &format!("INSERT OR IGNORE INTO {table} (key, data) VALUES (?1, ?2)"),
        params![key, data],
    )?;
    let id: i64 = tx.query_row(
        &format!("SELECT id FROM {table} WHERE key = ?1 LIMIT 1"),
        params![key],
        |r| r.get(0),
    )?;
    Ok(id)
}

fn get_or_insert_font_tx(tx: &Transaction<'_>, font: &formula_model::Font) -> Result<i64> {
    let data = serde_json::to_value(font)?;
    get_or_insert_style_component_tx(tx, "fonts", &data)
}

fn get_or_insert_fill_tx(tx: &Transaction<'_>, fill: &formula_model::Fill) -> Result<i64> {
    let data = serde_json::to_value(fill)?;
    get_or_insert_style_component_tx(tx, "fills", &data)
}

fn get_or_insert_border_tx(tx: &Transaction<'_>, border: &formula_model::Border) -> Result<i64> {
    let data = serde_json::to_value(border)?;
    get_or_insert_style_component_tx(tx, "borders", &data)
}

fn get_or_insert_model_style_tx(tx: &Transaction<'_>, style: &formula_model::Style) -> Result<i64> {
    let font_id = match style.font.as_ref() {
        Some(font) => Some(get_or_insert_font_tx(tx, font)?),
        None => None,
    };
    let fill_id = match style.fill.as_ref() {
        Some(fill) => Some(get_or_insert_fill_tx(tx, fill)?),
        None => None,
    };
    let border_id = match style.border.as_ref() {
        Some(border) => Some(get_or_insert_border_tx(tx, border)?),
        None => None,
    };

    let alignment = style
        .alignment
        .as_ref()
        .map(serde_json::to_value)
        .transpose()?;
    let protection = style
        .protection
        .as_ref()
        .map(serde_json::to_value)
        .transpose()?;

    let stored = StorageStyle {
        font_id,
        fill_id,
        border_id,
        number_format: style.number_format.clone(),
        alignment,
        protection,
    };

    get_or_insert_style_tx(tx, &stored)
}

fn load_style_component<T: DeserializeOwned>(
    conn: &Connection,
    table: &str,
    id: i64,
) -> Result<T> {
    let data: serde_json::Value = conn.query_row(
        &format!("SELECT data FROM {table} WHERE id = ?1"),
        params![id],
        |r| r.get(0),
    )?;
    Ok(serde_json::from_value(data)?)
}

fn load_model_style(conn: &Connection, style_id: i64) -> Result<formula_model::Style> {
    let (font_id, fill_id, border_id, number_format, alignment, protection): (
        Option<i64>,
        Option<i64>,
        Option<i64>,
        Option<String>,
        Option<String>,
        Option<String>,
    ) = conn.query_row(
        r#"
        SELECT font_id, fill_id, border_id, number_format, alignment, protection
        FROM styles
        WHERE id = ?1
        "#,
        params![style_id],
        |r| Ok((r.get(0)?, r.get(1)?, r.get(2)?, r.get(3)?, r.get(4)?, r.get(5)?)),
    )?;

    let font = font_id.and_then(|id| load_style_component::<formula_model::Font>(conn, "fonts", id).ok());
    let fill = fill_id.and_then(|id| load_style_component::<formula_model::Fill>(conn, "fills", id).ok());
    let border =
        border_id.and_then(|id| load_style_component::<formula_model::Border>(conn, "borders", id).ok());

    let alignment = alignment.and_then(|raw| serde_json::from_str::<formula_model::Alignment>(&raw).ok());
    let protection =
        protection.and_then(|raw| serde_json::from_str::<formula_model::Protection>(&raw).ok());

    Ok(formula_model::Style {
        font,
        fill,
        border,
        alignment,
        protection,
        number_format,
    })
}

fn build_model_style_table(
    conn: &Connection,
    workbook_id: Uuid,
) -> Result<(formula_model::StyleTable, HashMap<i64, u32>)> {
    let mut style_table = formula_model::StyleTable::new();
    let mut storage_to_model: HashMap<i64, u32> = HashMap::new();

    let mut mapping_stmt = conn.prepare(
        r#"
        SELECT style_index, style_id
        FROM workbook_styles
        WHERE workbook_id = ?1
        ORDER BY style_index
        "#,
    )?;
    let mut mapping_rows = mapping_stmt.query(params![workbook_id.to_string()])?;

    let mut mapped_any = false;
    while let Some(row) = mapping_rows.next()? {
        mapped_any = true;
        let style_id: i64 = row.get(1)?;
        let style = load_model_style(conn, style_id)?;
        let model_id = style_table.intern(style);
        storage_to_model.insert(style_id, model_id);
    }

    if mapped_any {
        return Ok((style_table, storage_to_model));
    }

    // Older databases (or workbooks created via the legacy APIs) do not have an explicit style
    // table ordering. Build a minimal table from the styles referenced by cells.
    let mut stmt = conn.prepare(
        r#"
        SELECT DISTINCT c.style_id
        FROM cells c
        JOIN sheets s ON s.id = c.sheet_id
        WHERE s.workbook_id = ?1
          AND c.style_id IS NOT NULL
        ORDER BY c.style_id
        "#,
    )?;

    let style_ids = stmt.query_map(params![workbook_id.to_string()], |r| r.get::<_, i64>(0))?;
    for style_id in style_ids {
        let style_id = style_id?;
        let style = load_model_style(conn, style_id)?;
        let model_id = style_table.intern(style);
        storage_to_model.insert(style_id, model_id);
    }

    Ok((style_table, storage_to_model))
}

fn stream_cells_into_model_sheet(
    conn: &Connection,
    sheet_id: &str,
    style_map: &HashMap<i64, u32>,
    sheet: &mut formula_model::Worksheet,
) -> Result<()> {
    let mut stmt = conn.prepare(
        r#"
        SELECT row, col, value_type, value_number, value_string, value_json, formula, style_id
        FROM cells
        WHERE sheet_id = ?1
        ORDER BY row, col
        "#,
    )?;

    let mut rows = stmt.query(params![sheet_id])?;
    while let Some(row) = rows.next()? {
        let row_idx: i64 = row.get(0)?;
        let col_idx: i64 = row.get(1)?;

        if row_idx < 0
            || row_idx > u32::MAX as i64
            || col_idx < 0
            || col_idx >= formula_model::EXCEL_MAX_COLS as i64
        {
            continue;
        }

        let snapshot = snapshot_from_row(
            row.get(2)?,
            row.get(3)?,
            row.get(4)?,
            row.get(5)?,
            row.get(6)?,
            row.get(7)?,
        )?;

        let style_id = snapshot
            .style_id
            .and_then(|id| style_map.get(&id).copied())
            .unwrap_or(0);
        let formula = snapshot
            .formula
            .as_deref()
            .and_then(formula_model::normalize_formula_text);

        let cell = formula_model::Cell {
            value: snapshot.value,
            formula,
            style_id,
        };

        sheet.set_cell(
            formula_model::CellRef::new(row_idx as u32, col_idx as u32),
            cell,
        );
    }

    Ok(())
}
fn export_connection_to_sqlite_bytes(conn: &Connection) -> Result<Vec<u8>> {
    let data = conn.serialize(DatabaseName::Main)?;
    Ok(data.to_vec())
}

fn atomic_write(path: &Path, bytes: &[u8]) -> std::io::Result<()> {
    let dir = path.parent().unwrap_or_else(|| Path::new("."));
    fs::create_dir_all(dir)?;
    let mut temp = NamedTempFile::new_in(dir)?;
    temp.write_all(bytes)?;
    temp.flush()?;
    temp.as_file().sync_all()?;

    match temp.persist(path) {
        Ok(_) => Ok(()),
        Err(err) => match err.error.kind() {
            // Best-effort replacement on platforms where rename doesn't clobber.
            std::io::ErrorKind::AlreadyExists => {
                let _ = fs::remove_file(path);
                err.file.persist(path).map(|_| ()).map_err(|e| e.error)
            }
            _ => Err(err.error),
        },
    }
}

fn cleanup_sqlite_sidecar_files(path: &Path) -> std::io::Result<()> {
    for suffix in ["-wal", "-shm", "-journal"] {
        let mut sidecar = path.as_os_str().to_os_string();
        sidecar.push(suffix);
        let sidecar_path = PathBuf::from(sidecar);
        match fs::remove_file(&sidecar_path) {
            Ok(()) => {}
            Err(err) if err.kind() == std::io::ErrorKind::NotFound => {}
            Err(err) => return Err(err),
        }
    }
    Ok(())
}
