use crate::data_model;
use crate::schema;
use crate::encryption::{
    decrypt_sqlite_bytes, encrypt_sqlite_bytes, is_encrypted_container, load_or_create_keyring,
    KeyProvider,
};
use crate::lock_unpoisoned;
use crate::types::{
    canonical_json, CellData, CellSnapshot, CellValue, ImportModelWorkbookOptions, NamedRange,
    SheetMeta, SheetVisibility, Style as StorageStyle, WorkbookMeta,
};
use formula_model::{
    rewrite_deleted_sheet_references_in_formula, rewrite_sheet_names_in_formula, validate_sheet_name,
    sheet_name_eq_case_insensitive, DefinedName, DefinedNameScope, ErrorValue, SheetNameError,
};
use rusqlite::{params, Connection, DatabaseName, OpenFlags, OptionalExtension, Transaction};
use serde::de::DeserializeOwned;
use serde_json::{json, Value as JsonValue};
use std::collections::{BTreeMap, HashMap, HashSet};
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
    #[error("dax error: {0}")]
    Dax(#[from] formula_dax::DaxError),
}

pub type Result<T> = std::result::Result<T, StorageError>;

fn map_sheet_name_error(err: SheetNameError) -> StorageError {
    match err {
        SheetNameError::EmptyName => StorageError::EmptySheetName,
        other => StorageError::InvalidSheetName(other),
    }
}

fn parse_optional_json_value(raw: Option<String>) -> Option<serde_json::Value> {
    raw.and_then(|raw| {
        let trimmed = raw.trim();
        if trimmed.is_empty() {
            return None;
        }
        serde_json::from_str(trimmed).ok()
    })
}

fn invalid_sheet_name_placeholder(sheet_id: Uuid) -> String {
    // Use a deterministic, Excel-safe placeholder name (<= 31 chars, ASCII, no forbidden chars).
    // This allows corrupted databases (e.g. non-TEXT `sheets.name`) to remain operable and
    // round-trip through export/import without failing.
    let id_str = sheet_id.to_string();
    let prefix = id_str.get(0..8).unwrap_or(&id_str);
    format!("_invalid_{prefix}")
}

fn invalid_workbook_name_placeholder(workbook_id: Uuid) -> String {
    let id_str = workbook_id.to_string();
    let prefix = id_str.get(0..8).unwrap_or(&id_str);
    format!("_invalid_workbook_{prefix}")
}

#[derive(Clone)]
pub struct Storage {
    conn: Arc<Mutex<Connection>>,
    encrypted: Option<Arc<EncryptedStorageContext>>,
}

struct EncryptedStorageContext {
    path: PathBuf,
    key_provider: Arc<dyn KeyProvider>,
    persist_lock: Mutex<()>,
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
    ///
    /// The on-disk format uses the same header as the JS encrypted file helper:
    /// `FMLENC01` + `keyVersion` + `iv` + `tag` + ciphertext.
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
            encrypted: Some(Arc::new(EncryptedStorageContext {
                path,
                key_provider,
                persist_lock: Mutex::new(()),
            })),
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
        let _persist_guard = lock_unpoisoned(&ctx.persist_lock);
        self.persist_inner(ctx)
    }

    /// Rotate the current encryption key (preserving older versions for decryption) and persist.
    ///
    /// Returns `Ok(None)` when this storage is not using encrypted-at-rest mode.
    pub fn rotate_encryption_key(&self) -> Result<Option<u32>> {
        let Some(ctx) = self.encrypted.as_ref() else {
            return Ok(None);
        };
        let _persist_guard = lock_unpoisoned(&ctx.persist_lock);

        let mut keyring = load_or_create_keyring(ctx.key_provider.as_ref(), true)?;
        keyring.rotate();
        ctx.key_provider
            .store_keyring(&keyring)
            .map_err(crate::encryption::EncryptionError::from)?;
        self.persist_inner(ctx)?;
        Ok(Some(keyring.current_version))
    }

    fn persist_inner(&self, ctx: &EncryptedStorageContext) -> Result<()> {
        let sqlite_bytes = {
            let conn = lock_unpoisoned(&self.conn);
            export_connection_to_sqlite_bytes(&conn)?
        };

        let keyring = load_or_create_keyring(ctx.key_provider.as_ref(), true)?;
        let encrypted_bytes = encrypt_sqlite_bytes(&sqlite_bytes, &keyring)?;

        atomic_write(&ctx.path, &encrypted_bytes)?;
        cleanup_sqlite_sidecar_files(&ctx.path)?;
        Ok(())
    }

    pub fn save_data_model(&self, workbook_id: Uuid, model: &formula_dax::DataModel) -> Result<()> {
        // Ensure workbook exists.
        self.get_workbook(workbook_id)?;

        let mut conn = lock_unpoisoned(&self.conn);
        let tx = conn.transaction()?;
        data_model::save_data_model_tx(&tx, workbook_id, model)?;
        touch_workbook_modified_at_by_workbook_id(&tx, workbook_id)?;
        tx.commit()?;
        Ok(())
    }

    pub fn load_data_model(&self, workbook_id: Uuid) -> Result<formula_dax::DataModel> {
        // Ensure workbook exists.
        self.get_workbook(workbook_id)?;

        let conn = lock_unpoisoned(&self.conn);
        data_model::load_data_model(&conn, workbook_id)
    }

    pub fn load_data_model_schema(&self, workbook_id: Uuid) -> Result<data_model::DataModelSchema> {
        // Ensure workbook exists.
        self.get_workbook(workbook_id)?;

        let conn = lock_unpoisoned(&self.conn);
        data_model::load_data_model_schema(&conn, workbook_id)
    }

    pub fn stream_data_model_column_chunks<F>(
        &self,
        workbook_id: Uuid,
        table_name: &str,
        column_name: &str,
        f: F,
    ) -> Result<()>
    where
        F: FnMut(data_model::DataModelChunk) -> Result<()>,
    {
        // Ensure workbook exists.
        self.get_workbook(workbook_id)?;

        let conn = lock_unpoisoned(&self.conn);
        data_model::stream_column_chunks(&conn, workbook_id, table_name, column_name, f)
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

        let conn = lock_unpoisoned(&self.conn);
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
        let conn = lock_unpoisoned(&self.conn);
        let row = conn
            .query_row(
                "SELECT id, name, metadata FROM workbooks WHERE id = ?1",
                params![id.to_string()],
                |r| {
                    let id: String = r.get(0)?;
                    let workbook_id =
                        Uuid::parse_str(&id).map_err(|_| rusqlite::Error::InvalidQuery)?;
                    let name = r
                        .get::<_, Option<String>>(1)
                        .ok()
                        .flatten()
                        .unwrap_or_else(|| invalid_workbook_name_placeholder(workbook_id));
                    let metadata_raw: Option<String> = r.get::<_, Option<String>>(2).ok().flatten();
                    Ok(WorkbookMeta {
                        id: workbook_id,
                        name,
                        metadata: parse_optional_json_value(metadata_raw),
                    })
                },
            )
            .optional()?;

        row.ok_or(StorageError::WorkbookNotFound(id))
    }

    pub fn list_workbooks(&self) -> Result<Vec<WorkbookMeta>> {
        let conn = lock_unpoisoned(&self.conn);
        let mut stmt = conn.prepare("SELECT id, name, metadata FROM workbooks ORDER BY created_at")?;
        let rows = stmt.query_map([], |r| {
            let Some(id_raw) = r.get::<_, Option<String>>(0).ok().flatten() else {
                return Ok(None);
            };
            let Ok(id) = Uuid::parse_str(&id_raw).map_err(|_| rusqlite::Error::InvalidQuery) else {
                return Ok(None);
            };
            let name = r
                .get::<_, Option<String>>(1)
                .ok()
                .flatten()
                .unwrap_or_else(|| invalid_workbook_name_placeholder(id));
            let metadata_raw: Option<String> = r.get::<_, Option<String>>(2).ok().flatten();
            Ok(Some(WorkbookMeta {
                id,
                name,
                metadata: parse_optional_json_value(metadata_raw),
            }))
        })?;

        let mut out = Vec::new();
        for row in rows {
            let Some(row) = row? else {
                continue;
            };
            out.push(row);
        }
        Ok(out)
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

        let mut conn = lock_unpoisoned(&self.conn);
        let tx = conn.transaction()?;

        let theme = (!workbook.theme.is_default())
            .then_some(serde_json::to_value(&workbook.theme)?);
        let workbook_protection = (!formula_model::WorkbookProtection::is_default(
            &workbook.workbook_protection,
        ))
        .then_some(serde_json::to_value(&workbook.workbook_protection)?);
        let defined_names = (!workbook.defined_names.is_empty())
            .then_some(serde_json::to_value(&workbook.defined_names)?);
        let print_settings = (!workbook.print_settings.is_empty())
            .then_some(serde_json::to_value(&workbook.print_settings)?);
        let view = (workbook.view != formula_model::WorkbookView::default())
            .then_some(serde_json::to_value(&workbook.view)?);

         tx.execute(
             r#"
             INSERT INTO workbooks (
               id,
               name,
               metadata,
               model_schema_version,
               model_workbook_id,
               codepage,
               date_system,
               calc_settings,
               theme,
               workbook_protection,
               defined_names,
               print_settings,
               view
             ) VALUES (?1, ?2, ?3, ?4, ?5, ?6, ?7, ?8, ?9, ?10, ?11, ?12, ?13)
             "#,
             params![
                 workbook_meta.id.to_string(),
                 &workbook_meta.name,
                 workbook_meta.metadata.clone(),
                 workbook.schema_version as i64,
                 workbook.id as i64,
                 workbook.codepage as i64,
                 date_system_to_str(workbook.date_system),
                 serde_json::to_value(&workbook.calc_settings)?,
                 theme,
                 workbook_protection,
                 defined_names,
                 print_settings,
                 view
             ],
         )?;

        // Persist workbook-scoped images.
        {
            let mut stmt = tx.prepare(
                r#"
                INSERT INTO workbook_images (workbook_id, image_id, content_type, bytes)
                VALUES (?1, ?2, ?3, ?4)
                "#,
            )?;
            let workbook_id = workbook_meta.id.to_string();
            for (image_id, image) in workbook.images.iter() {
                stmt.execute(params![
                    &workbook_id,
                    image_id.as_str(),
                    image.content_type.as_deref(),
                    &image.bytes
                ])?;
            }
        }

        // Populate the legacy `named_ranges` table for compatibility with existing APIs.
        //
        // The canonical representation of defined names is stored in `workbooks.defined_names`,
        // but callers that still use `Storage::{get_named_range, upsert_named_range}` expect
        // `named_ranges` to reflect the workbook state.
        {
            let mut stmt = tx.prepare(
                r#"
                INSERT INTO named_ranges (workbook_id, name, scope, reference)
                VALUES (?1, ?2, ?3, ?4)
                "#,
            )?;
            let workbook_id = workbook_meta.id.to_string();
            for name in &workbook.defined_names {
                let scope = match name.scope {
                    DefinedNameScope::Workbook => "workbook".to_string(),
                    DefinedNameScope::Sheet(sheet_id) => {
                        let Some(sheet) = workbook.sheets.iter().find(|s| s.id == sheet_id) else {
                            continue;
                        };
                        sheet.name.clone()
                    }
                };

                stmt.execute(params![
                    &workbook_id,
                    &name.name,
                    &scope,
                    &name.refers_to
                ])?;
            }
        }

        // Insert worksheets in the order they appear in the model workbook.
        let mut seen_names: Vec<String> = Vec::new();
        let _ = seen_names.try_reserve_exact(workbook.sheets.len());
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
            let model_sheet_json = worksheet_metadata_json(sheet)?;

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
                  model_sheet_id,
                  model_sheet_json
                ) VALUES (
                  ?1, ?2, ?3, ?4, ?5, ?6, ?7, ?8, ?9, ?10, ?11, ?12, ?13, ?14, ?15
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
                    sheet.id as i64,
                    model_sheet_json
                ],
            )?;
        }

        // Persist sheet drawing objects (images, shapes, etc).
        {
            let mut stmt = tx.prepare(
                r#"
                INSERT INTO sheet_drawings (sheet_id, position, data)
                VALUES (?1, ?2, ?3)
                "#,
            )?;

            for sheet in &workbook.sheets {
                let Some(storage_sheet_id) = model_sheet_to_storage.get(&sheet.id) else {
                    continue;
                };
                for (position, drawing) in sheet.drawings.iter().enumerate() {
                    stmt.execute(params![
                        storage_sheet_id.to_string(),
                        position as i64,
                        serde_json::to_value(drawing)?
                    ])?;
                }
            }
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
                   style_id,
                   phonetic
                 ) VALUES (?1, ?2, ?3, ?4, ?5, ?6, ?7, ?8, ?9, ?10)
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
                         storage_style_id,
                         cell.phonetic.as_deref()
                     ])?;
                 }
             }
         }

        tx.commit()?;
        Ok(workbook_meta)
    }

    pub fn export_model_workbook(&self, workbook_id: Uuid) -> Result<formula_model::Workbook> {
        let conn = lock_unpoisoned(&self.conn);

        let row = conn
            .query_row(
                r#"
                 SELECT
                   model_schema_version,
                   model_workbook_id,
                   codepage,
                   date_system,
                   calc_settings,
                   theme,
                   workbook_protection,
                   defined_names,
                   print_settings,
                   view
                 FROM workbooks
                 WHERE id = ?1
                "#,
                params![workbook_id.to_string()],
                |r| {
                     Ok((
                         r.get::<_, Option<i64>>(0).ok().flatten(),
                         r.get::<_, Option<i64>>(1).ok().flatten(),
                         r.get::<_, Option<i64>>(2).ok().flatten(),
                         r.get::<_, Option<String>>(3).ok().flatten(),
                         r.get::<_, Option<String>>(4).ok().flatten(),
                         r.get::<_, Option<String>>(5).ok().flatten(),
                         r.get::<_, Option<String>>(6).ok().flatten(),
                         r.get::<_, Option<String>>(7).ok().flatten(),
                         r.get::<_, Option<String>>(8).ok().flatten(),
                         r.get::<_, Option<String>>(9).ok().flatten(),
                     ))
                 },
             )
             .optional()?;

        let Some((
             model_schema_version,
             model_workbook_id,
             codepage,
             date_system,
             calc_settings,
             theme,
             workbook_protection,
             defined_names,
             print_settings,
             view,
        )) = row
        else {
            return Err(StorageError::WorkbookNotFound(workbook_id));
        };

        let (styles, storage_style_to_model) = build_model_style_table(&conn, workbook_id)?;

        let mut model_workbook = formula_model::Workbook::new();
        if let Some(schema_version) = model_schema_version.and_then(|v| u32::try_from(v).ok()) {
            model_workbook.schema_version = schema_version;
        }
         if let Some(id) = model_workbook_id.and_then(|v| u32::try_from(v).ok()) {
             model_workbook.id = id;
         }
         if let Some(codepage) = codepage.and_then(|v| u16::try_from(v).ok()) {
             model_workbook.codepage = codepage;
         }
         if let Some(date_system) = date_system {
             if let Ok(date_system) = parse_date_system(&date_system) {
                 model_workbook.date_system = date_system;
             }
         }
        if let Some(calc_settings) = calc_settings {
            if let Ok(calc_settings) = serde_json::from_str(&calc_settings) {
                model_workbook.calc_settings = calc_settings;
            }
        }
        if let Some(theme) = theme {
            if let Ok(theme) = serde_json::from_str(&theme) {
                model_workbook.theme = theme;
            }
        }
        if let Some(workbook_protection) = workbook_protection {
            if let Ok(workbook_protection) = serde_json::from_str(&workbook_protection) {
                model_workbook.workbook_protection = workbook_protection;
            }
        }
        if let Some(defined_names) = defined_names {
            if let Ok(defined_names) = serde_json::from_str(&defined_names) {
                model_workbook.defined_names = defined_names;
            }
        }
        if let Some(print_settings) = print_settings {
            if let Ok(print_settings) = serde_json::from_str(&print_settings) {
                model_workbook.print_settings = print_settings;
            }
        }
        if let Some(view) = view {
            if let Ok(view) = serde_json::from_str(&view) {
                model_workbook.view = view;
            }
        }
        model_workbook.styles = styles;

        // Load workbook images.
        {
            let mut stmt = conn.prepare(
                r#"
                SELECT image_id, content_type, bytes
                FROM workbook_images
                WHERE workbook_id = ?1
                ORDER BY image_id
                "#,
            )?;
            let mut rows = stmt.query(params![workbook_id.to_string()])?;
            while let Some(row) = rows.next()? {
                let Ok(image_id) = row.get::<_, String>(0) else {
                    continue;
                };
                let content_type: Option<String> = row.get(1).ok().flatten();
                let Ok(bytes) = row.get::<_, Vec<u8>>(2) else {
                    continue;
                };
                model_workbook.images.insert(
                    formula_model::drawings::ImageId::new(image_id),
                    formula_model::drawings::ImageData { bytes, content_type },
                );
            }
        }

        // Allocate deterministic worksheet ids for sheets that predate model import.
        //
        // Note: `model_sheet_id` is stored as an SQLite INTEGER (i64) but represents a `u32`
        // worksheet id. We treat out-of-range values as missing and synthesize a stable id.
        let mut used_sheet_ids: HashSet<u32> = HashSet::new();
        let max_model_sheet_id: i64 = conn.query_row(
            r#"
            SELECT COALESCE(
              MAX(
                CASE
                  WHEN typeof(model_sheet_id) = 'integer'
                    AND model_sheet_id >= 0
                    AND model_sheet_id <= ?2
                  THEN model_sheet_id
                  ELSE NULL
                END
              ),
              0
            )
            FROM sheets
            WHERE workbook_id = ?1
            "#,
            params![workbook_id.to_string(), u32::MAX as i64],
            |r| r.get(0),
        )?;

        let mut sheets_stmt = conn.prepare(
            r#"
            SELECT
              id,
              name,
              position,
              COALESCE(visibility, 'visible'),
              tab_color,
              tab_color_json,
              xlsx_sheet_id,
              xlsx_rel_id,
              COALESCE(frozen_rows, 0),
              COALESCE(frozen_cols, 0),
              COALESCE(zoom, 1.0),
              model_sheet_id,
              model_sheet_json
            FROM sheets
            WHERE workbook_id = ?1
            ORDER BY COALESCE(position, 0), id
            "#,
        )?;
        let mut sheet_rows = sheets_stmt.query(params![workbook_id.to_string()])?;

        let mut worksheets: Vec<formula_model::Worksheet> = Vec::new();
        let mut next_generated_sheet_id: u32 = (max_model_sheet_id.max(0) as u32).wrapping_add(1);

        while let Some(row) = sheet_rows.next()? {
            let Ok(storage_sheet_id) = row.get::<_, String>(0) else {
                continue;
            };
            let Ok(storage_sheet_uuid) = Uuid::parse_str(&storage_sheet_id) else {
                continue;
            };
            let name = row
                .get::<_, Option<String>>(1)
                .ok()
                .flatten()
                .unwrap_or_else(|| invalid_sheet_name_placeholder(storage_sheet_uuid));
            let visibility_raw: String = row.get(3).unwrap_or_else(|_| "visible".to_string());
            let tab_color_fast: Option<String> = row.get(4).ok().flatten();
            let tab_color_json: Option<String> = row.get(5).ok().flatten();
            let xlsx_sheet_id: Option<i64> = row.get::<_, Option<i64>>(6).ok().flatten();
            let xlsx_rel_id: Option<String> = row.get::<_, Option<String>>(7).ok().flatten();
            let frozen_rows: i64 = row.get(8).unwrap_or(0);
            let frozen_cols: i64 = row.get(9).unwrap_or(0);
            let zoom: f64 = row.get(10).unwrap_or(1.0);
            let model_sheet_id: Option<i64> = row.get::<_, Option<i64>>(11).ok().flatten();
            let model_sheet_json: Option<String> = row.get(12).ok().flatten();

            let sheet_id = match model_sheet_id.and_then(|id| u32::try_from(id).ok()) {
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

            let mut sheet = model_sheet_json
                .and_then(|raw| serde_json::from_str::<formula_model::Worksheet>(&raw).ok())
                .unwrap_or_else(|| formula_model::Worksheet::new(sheet_id, name.clone()));
            sheet.id = sheet_id;
            sheet.name = name.clone();
            sheet.visibility = storage_sheet_visibility_to_model(&visibility_raw);
            sheet.xlsx_sheet_id = xlsx_sheet_id.and_then(|v| u32::try_from(v).ok());
            sheet.xlsx_rel_id = xlsx_rel_id;
            sheet.frozen_rows = u32::try_from(frozen_rows).unwrap_or(0);
            sheet.frozen_cols = u32::try_from(frozen_cols).unwrap_or(0);
            let zoom_f32 = zoom as f32;
            sheet.zoom = if zoom_f32.is_finite() && zoom_f32 > 0.0 {
                zoom_f32
            } else {
                1.0
            };
            sheet.view.pane.frozen_rows = sheet.frozen_rows;
            sheet.view.pane.frozen_cols = sheet.frozen_cols;
            sheet.view.zoom = sheet.zoom;

            sheet.tab_color = tab_color_json
                .and_then(|raw| serde_json::from_str::<formula_model::TabColor>(&raw).ok())
                .or_else(|| tab_color_fast.map(formula_model::TabColor::rgb));

            // Load sheet drawing objects.
            {
                let mut stmt = conn.prepare(
                    r#"
                    SELECT data
                    FROM sheet_drawings
                    WHERE sheet_id = ?1
                    ORDER BY position
                    "#,
                )?;
                let drawings = stmt.query_map(params![&storage_sheet_id], |r| {
                    Ok(r.get::<_, Option<String>>(0).ok().flatten())
                })?;
                let mut out = Vec::new();
                for drawing in drawings {
                    let Some(raw) = drawing? else {
                        continue;
                    };
                    let Ok(parsed) =
                        serde_json::from_str::<formula_model::drawings::DrawingObject>(&raw)
                    else {
                        continue;
                    };
                    out.push(parsed);
                }
                sheet.drawings = out;
            }

            stream_cells_into_model_sheet(
                &conn,
                &storage_sheet_id,
                &storage_style_to_model,
                &mut sheet,
            )?;

            worksheets.push(sheet);
        }

        model_workbook.sheets = worksheets;
        merge_named_ranges_into_defined_names(
            &conn,
            workbook_id,
            &model_workbook.sheets,
            &mut model_workbook.defined_names,
        )?;
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

        let mut conn = lock_unpoisoned(&self.conn);
        let tx = conn.transaction()?;
        let workbook_id_str = workbook_id.to_string();

        let mut sheets = self.list_sheets_tx(&tx, workbook_id)?;
        if sheets
            .iter()
            .any(|existing| sheet_name_eq_case_insensitive(&existing.name, name))
        {
            return Err(StorageError::DuplicateSheetName(name.to_string()));
        }

        let clamped = position.max(0).min(sheets.len() as i64) as usize;

        let mut used_sheet_ids: HashSet<u32> = HashSet::new();
        let mut max_sheet_id: u32 = 0;
        {
            let mut stmt = tx.prepare(
                "SELECT model_sheet_id FROM sheets WHERE workbook_id = ?1 AND model_sheet_id IS NOT NULL",
            )?;
            let ids = stmt.query_map(params![&workbook_id_str], |row| Ok(row.get::<_, i64>(0).ok()))?;
            for raw in ids {
                let Some(raw) = raw? else {
                    continue;
                };
                if let Ok(id) = u32::try_from(raw) {
                    used_sheet_ids.insert(id);
                    max_sheet_id = max_sheet_id.max(id);
                }
            }
        }
        let mut model_sheet_id: u32 = max_sheet_id.wrapping_add(1);
        while used_sheet_ids.contains(&model_sheet_id) {
            model_sheet_id = model_sheet_id.wrapping_add(1);
        }

        let sheet = SheetMeta {
            id: Uuid::new_v4(),
            workbook_id,
            name: name.to_string(),
            position: clamped as i64,
            visibility: SheetVisibility::Visible,
            tab_color: None,
            xlsx_sheet_id: None,
            xlsx_rel_id: None,
            frozen_rows: 0,
            frozen_cols: 0,
            zoom: 1.0,
            metadata,
        };

        tx.execute(
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
              metadata,
              model_sheet_id
            ) VALUES (?1, ?2, ?3, ?4, ?5, ?6, ?7, ?8, ?9, ?10, ?11, ?12, ?13)
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
                sheet.metadata.clone(),
                model_sheet_id as i64
            ],
        )?;

        sheets.insert(clamped, sheet.clone());
        for (idx, sheet) in sheets.iter().enumerate() {
            tx.execute(
                "UPDATE sheets SET position = ?1 WHERE id = ?2",
                params![idx as i64, sheet.id.to_string()],
            )?;
        }

        touch_workbook_modified_at_by_workbook_id(&tx, workbook_id)?;
        tx.commit()?;

        Ok(sheet)
    }

    pub fn list_sheets(&self, workbook_id: Uuid) -> Result<Vec<SheetMeta>> {
        let conn = lock_unpoisoned(&self.conn);
        let mut stmt = conn.prepare(
            r#"
            SELECT
              id,
              workbook_id,
              name,
              COALESCE(position, 0),
              COALESCE(visibility, 'visible'),
              tab_color,
              xlsx_sheet_id,
              xlsx_rel_id,
              COALESCE(frozen_rows, 0),
              COALESCE(frozen_cols, 0),
              COALESCE(zoom, 1.0),
              metadata
            FROM sheets
            WHERE workbook_id = ?1
            ORDER BY COALESCE(position, 0), id
            "#,
        )?;

        let rows = stmt.query_map(params![workbook_id.to_string()], |r| {
            let Some(id) = r.get::<_, Option<String>>(0).ok().flatten() else {
                return Ok(None);
            };
            let Ok(id) = Uuid::parse_str(&id).map_err(|_| rusqlite::Error::InvalidQuery) else {
                return Ok(None);
            };

            let Some(workbook_id) = r.get::<_, Option<String>>(1).ok().flatten() else {
                return Ok(None);
            };
            let Ok(workbook_id) =
                Uuid::parse_str(&workbook_id).map_err(|_| rusqlite::Error::InvalidQuery)
            else {
                return Ok(None);
            };

            let name = r
                .get::<_, Option<String>>(2)
                .ok()
                .flatten()
                .unwrap_or_else(|| invalid_sheet_name_placeholder(id));
            let visibility: String = r.get(4).unwrap_or_else(|_| "visible".to_string());
            let metadata_raw: Option<String> = r.get::<_, Option<String>>(11).ok().flatten();
            Ok(Some(SheetMeta {
                id,
                workbook_id,
                name,
                position: r.get(3).unwrap_or(0),
                visibility: SheetVisibility::parse(&visibility),
                tab_color: r.get::<_, Option<String>>(5).ok().flatten(),
                xlsx_sheet_id: r.get::<_, Option<i64>>(6).ok().flatten(),
                xlsx_rel_id: r.get::<_, Option<String>>(7).ok().flatten(),
                frozen_rows: r.get(8).unwrap_or(0),
                frozen_cols: r.get(9).unwrap_or(0),
                zoom: r.get(10).unwrap_or(1.0),
                metadata: parse_optional_json_value(metadata_raw),
            }))
        })?;

        let mut sheets = Vec::new();
        for sheet in rows {
            let Some(sheet) = sheet? else {
                continue;
            };
            sheets.push(sheet);
        }
        Ok(sheets)
    }

    pub fn get_sheet_meta(&self, sheet_id: Uuid) -> Result<SheetMeta> {
        let conn = lock_unpoisoned(&self.conn);
        let row = conn
            .query_row(
                r#"
                SELECT
                  id,
                  workbook_id,
                  name,
                  COALESCE(position, 0),
                  COALESCE(visibility, 'visible'),
                  tab_color,
                  xlsx_sheet_id,
                  xlsx_rel_id,
                  COALESCE(frozen_rows, 0),
                  COALESCE(frozen_cols, 0),
                  COALESCE(zoom, 1.0),
                  metadata
                FROM sheets
                WHERE id = ?1
                "#,
                params![sheet_id.to_string()],
                |r| {
                    let id: String = r.get(0)?;
                    let workbook_id: String = r.get(1)?;
                    let visibility: String = r.get(4).unwrap_or_else(|_| "visible".to_string());
                    let metadata_raw: Option<String> = r.get::<_, Option<String>>(11).ok().flatten();
                    let id_parsed =
                        Uuid::parse_str(&id).map_err(|_| rusqlite::Error::InvalidQuery)?;
                    Ok(SheetMeta {
                        id: id_parsed,
                        workbook_id: Uuid::parse_str(&workbook_id)
                            .map_err(|_| rusqlite::Error::InvalidQuery)?,
                        name: r
                            .get::<_, Option<String>>(2)
                            .ok()
                            .flatten()
                            .unwrap_or_else(|| invalid_sheet_name_placeholder(id_parsed)),
                        position: r.get(3).unwrap_or(0),
                        visibility: SheetVisibility::parse(&visibility),
                        tab_color: r.get::<_, Option<String>>(5).ok().flatten(),
                        xlsx_sheet_id: r.get::<_, Option<i64>>(6).ok().flatten(),
                        xlsx_rel_id: r.get::<_, Option<String>>(7).ok().flatten(),
                        frozen_rows: r.get(8).unwrap_or(0),
                        frozen_cols: r.get(9).unwrap_or(0),
                        zoom: r.get(10).unwrap_or(1.0),
                        metadata: parse_optional_json_value(metadata_raw),
                    })
                },
            )
            .optional()?;

        row.ok_or(StorageError::SheetNotFound(sheet_id))
    }

    pub fn get_sheet_model_worksheet(
        &self,
        sheet_id: Uuid,
    ) -> Result<Option<formula_model::Worksheet>> {
        let conn = lock_unpoisoned(&self.conn);
        let raw_row: Option<Option<String>> = conn
            .query_row(
                "SELECT model_sheet_json FROM sheets WHERE id = ?1",
                params![sheet_id.to_string()],
                |r| Ok(r.get::<_, Option<String>>(0).ok().flatten()),
            )
            .optional()?;
        let Some(raw_row) = raw_row else {
            return Err(StorageError::SheetNotFound(sheet_id));
        };
        let Some(raw) = raw_row else {
            return Ok(None);
        };
        Ok(serde_json::from_str::<formula_model::Worksheet>(&raw).ok())
    }

    /// Replace the `sheets.metadata` JSON payload for the given sheet.
    ///
    /// This is intended for application-specific per-sheet state that should be persisted alongside
    /// core workbook data (e.g. UI formatting layers).
    pub fn set_sheet_metadata(&self, sheet_id: Uuid, metadata: Option<JsonValue>) -> Result<()> {
        let mut conn = lock_unpoisoned(&self.conn);
        let tx = conn.transaction()?;
        let updated = tx.execute(
            "UPDATE sheets SET metadata = ?1 WHERE id = ?2",
            params![metadata, sheet_id.to_string()],
        )?;
        if updated == 0 {
            return Err(StorageError::SheetNotFound(sheet_id));
        }
        touch_workbook_modified_at(&tx, sheet_id)?;
        tx.commit()?;
        Ok(())
    }

    /// Read-modify-write helper for `sheets.metadata`.
    ///
    /// The callback receives the current parsed JSON (or `None` if unset/invalid) and returns the
    /// updated JSON (or `None` to clear).
    pub fn update_sheet_metadata<F>(&self, sheet_id: Uuid, f: F) -> Result<()>
    where
        F: FnOnce(Option<JsonValue>) -> Result<Option<JsonValue>>,
    {
        let mut conn = lock_unpoisoned(&self.conn);
        let tx = conn.transaction()?;

        let raw_row: Option<Option<String>> = tx
            .query_row(
                "SELECT metadata FROM sheets WHERE id = ?1",
                params![sheet_id.to_string()],
                |r| Ok(r.get::<_, Option<String>>(0).ok().flatten()),
            )
            .optional()?;
        let Some(raw_row) = raw_row else {
            return Err(StorageError::SheetNotFound(sheet_id));
        };

        let current = parse_optional_json_value(raw_row);
        let next = f(current)?;

        tx.execute(
            "UPDATE sheets SET metadata = ?1 WHERE id = ?2",
            params![next, sheet_id.to_string()],
        )?;
        touch_workbook_modified_at(&tx, sheet_id)?;
        tx.commit()?;
        Ok(())
    }

    /// Rename a worksheet.
    pub fn rename_sheet(&self, sheet_id: Uuid, name: &str) -> Result<()> {
        validate_sheet_name(name).map_err(map_sheet_name_error)?;

        let mut conn = lock_unpoisoned(&self.conn);
        let tx = conn.transaction()?;

        let meta = match self.get_sheet_meta_tx(&tx, sheet_id) {
            Ok(meta) => meta,
            Err(StorageError::SheetNotFound(id)) => return Err(StorageError::SheetNotFound(id)),
            Err(_) => {
                // The sheet row exists but is corrupted (e.g. invalid `workbook_id` type).
                // Best-effort fallback: update the sheet name without workbook-wide rewrites.
                let updated = tx.execute(
                    "UPDATE sheets SET name = ?1 WHERE id = ?2",
                    params![name, sheet_id.to_string()],
                )?;
                if updated == 0 {
                    return Err(StorageError::SheetNotFound(sheet_id));
                }
                update_sheet_model_json_tx(&tx, sheet_id, |sheet| {
                    sheet.name = name.to_string();
                })?;
                touch_workbook_modified_at(&tx, sheet_id)?;
                tx.commit()?;
                return Ok(());
            }
        };
        let old_name = meta.name.clone();

        // Enforce Excel-style uniqueness (Unicode-aware, case-insensitive) within the workbook.
        {
            let mut stmt =
                tx.prepare("SELECT name FROM sheets WHERE workbook_id = ?1 AND id != ?2")?;
            let mut rows =
                stmt.query(params![meta.workbook_id.to_string(), sheet_id.to_string()])?;
            while let Some(row) = rows.next()? {
                let Ok(existing) = row.get::<_, String>(0) else {
                    continue;
                };
                if sheet_name_eq_case_insensitive(&existing, name) {
                    return Err(StorageError::DuplicateSheetName(name.to_string()));
                }
            }
        }

        if old_name != name {
            rewrite_sheet_rename_references_tx(&tx, meta.workbook_id, &old_name, name)?;
        }

        update_named_range_scopes_for_sheet_rename_tx(&tx, meta.workbook_id, &old_name, name)?;

        tx.execute(
            "UPDATE sheets SET name = ?1 WHERE id = ?2",
            params![name, sheet_id.to_string()],
        )?;

        touch_workbook_modified_at_by_workbook_id(&tx, meta.workbook_id)?;
        tx.commit()?;
        Ok(())
    }

    pub fn set_sheet_visibility(&self, sheet_id: Uuid, visibility: SheetVisibility) -> Result<()> {
        let mut conn = lock_unpoisoned(&self.conn);
        let tx = conn.transaction()?;
        let updated = tx.execute(
            "UPDATE sheets SET visibility = ?1 WHERE id = ?2",
            params![visibility.as_str(), sheet_id.to_string()],
        )?;
        if updated == 0 {
            return Err(StorageError::SheetNotFound(sheet_id));
        }
        // Keep `model_sheet_json` aligned so export/import round-trips after legacy metadata edits.
        update_sheet_model_json_tx(&tx, sheet_id, |sheet| {
            sheet.visibility = storage_sheet_visibility_to_model(visibility.as_str());
        })?;
        touch_workbook_modified_at(&tx, sheet_id)?;
        tx.commit()?;
        Ok(())
    }

    pub fn set_sheet_tab_color(
        &self,
        sheet_id: Uuid,
        tab_color: Option<&formula_model::TabColor>,
    ) -> Result<()> {
        let mut conn = lock_unpoisoned(&self.conn);
        let tx = conn.transaction()?;
        let tab_color_fast = tab_color.and_then(|c| c.rgb.as_deref());
        let tab_color_json = tab_color.map(serde_json::to_value).transpose()?;
        let updated = tx.execute(
            "UPDATE sheets SET tab_color = ?1, tab_color_json = ?2 WHERE id = ?3",
            params![tab_color_fast, tab_color_json, sheet_id.to_string()],
        )?;
        if updated == 0 {
            return Err(StorageError::SheetNotFound(sheet_id));
        }
        update_sheet_model_json_tx(&tx, sheet_id, |sheet| {
            sheet.tab_color = tab_color.cloned();
        })?;
        touch_workbook_modified_at(&tx, sheet_id)?;
        tx.commit()?;
        Ok(())
    }

    pub fn set_sheet_xlsx_metadata(
        &self,
        sheet_id: Uuid,
        xlsx_sheet_id: Option<i64>,
        xlsx_rel_id: Option<&str>,
    ) -> Result<()> {
        let mut conn = lock_unpoisoned(&self.conn);
        let tx = conn.transaction()?;
        let updated = tx.execute(
            "UPDATE sheets SET xlsx_sheet_id = ?1, xlsx_rel_id = ?2 WHERE id = ?3",
            params![xlsx_sheet_id, xlsx_rel_id, sheet_id.to_string()],
        )?;
        if updated == 0 {
            return Err(StorageError::SheetNotFound(sheet_id));
        }
        update_sheet_model_json_tx(&tx, sheet_id, |sheet| {
            sheet.xlsx_sheet_id = xlsx_sheet_id.and_then(|v| u32::try_from(v).ok());
            sheet.xlsx_rel_id = xlsx_rel_id.map(|s| s.to_string());
        })?;
        touch_workbook_modified_at(&tx, sheet_id)?;
        tx.commit()?;
        Ok(())
    }

    /// Reorder a sheet within its workbook by setting its 0-based position.
    ///
    /// This renormalizes positions to be contiguous starting at 0.
    pub fn reorder_sheet(&self, sheet_id: Uuid, new_position: i64) -> Result<()> {
        let mut conn = lock_unpoisoned(&self.conn);
        let tx = conn.transaction()?;

        let meta = match self.get_sheet_meta_tx(&tx, sheet_id) {
            Ok(meta) => meta,
            Err(StorageError::SheetNotFound(id)) => return Err(StorageError::SheetNotFound(id)),
            Err(_) => {
                // The sheet row exists but is corrupted (e.g. invalid `workbook_id` type).
                // Best-effort fallback: update the sheet's position without renormalizing the rest.
                let clamped = new_position.max(0);
                let updated = tx.execute(
                    "UPDATE sheets SET position = ?1 WHERE id = ?2",
                    params![clamped, sheet_id.to_string()],
                )?;
                if updated == 0 {
                    return Err(StorageError::SheetNotFound(sheet_id));
                }
                touch_workbook_modified_at(&tx, sheet_id)?;
                tx.commit()?;
                return Ok(());
            }
        };

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

    /// Reorder all sheets in a workbook by setting their 0-based positions.
    ///
    /// This is a batch alternative to calling [`Self::reorder_sheet`] repeatedly. It updates all
    /// sheet positions in a single transaction and renormalizes positions to be contiguous
    /// starting at 0.
    ///
    /// `sheet_ids_in_order` is treated as a partial ordering: any workbook sheets not present in
    /// the list are appended in their current order.
    pub fn reorder_sheets(&self, workbook_id: Uuid, sheet_ids_in_order: &[Uuid]) -> Result<()> {
        let mut conn = lock_unpoisoned(&self.conn);
        let tx = conn.transaction()?;

        let sheets = self.list_sheets_tx(&tx, workbook_id)?;
        if sheets.len() <= 1 {
            return Ok(());
        }

        let workbook_sheet_ids: HashSet<Uuid> = sheets.iter().map(|s| s.id).collect();
        let mut seen: HashSet<Uuid> = HashSet::new();
        let mut desired: Vec<Uuid> = Vec::new();
        let _ = desired.try_reserve_exact(sheets.len());

        for id in sheet_ids_in_order {
            if !workbook_sheet_ids.contains(id) {
                continue;
            }
            if !seen.insert(*id) {
                continue;
            }
            desired.push(*id);
        }

        // Append any remaining sheets in their current order (stable fallback).
        for sheet in &sheets {
            if seen.insert(sheet.id) {
                desired.push(sheet.id);
            }
        }

        for (idx, sheet_id) in desired.iter().enumerate() {
            tx.execute(
                "UPDATE sheets SET position = ?1 WHERE id = ?2",
                params![idx as i64, sheet_id.to_string()],
            )?;
        }

        touch_workbook_modified_at_by_workbook_id(&tx, workbook_id)?;
        tx.commit()?;
        Ok(())
    }

    pub fn delete_sheet(&self, sheet_id: Uuid) -> Result<()> {
        let mut conn = lock_unpoisoned(&self.conn);
        let tx = conn.transaction()?;

        let meta = match self.get_sheet_meta_tx(&tx, sheet_id) {
            Ok(meta) => meta,
            Err(StorageError::SheetNotFound(id)) => return Err(StorageError::SheetNotFound(id)),
            Err(_) => {
                // The sheet row exists but is corrupted (e.g. invalid `workbook_id` type).
                // Best-effort fallback: delete the sheet and its dependent rows without attempting
                // workbook-wide reference rewrites / position renormalization.
                touch_workbook_modified_at(&tx, sheet_id)?;
                tx.execute(
                    "DELETE FROM sheet_drawings WHERE sheet_id = ?1",
                    params![sheet_id.to_string()],
                )?;
                tx.execute(
                    "DELETE FROM cells WHERE sheet_id = ?1",
                    params![sheet_id.to_string()],
                )?;
                tx.execute(
                    "DELETE FROM change_log WHERE sheet_id = ?1",
                    params![sheet_id.to_string()],
                )?;
                let deleted = tx.execute(
                    "DELETE FROM sheets WHERE id = ?1",
                    params![sheet_id.to_string()],
                )?;
                if deleted == 0 {
                    return Err(StorageError::SheetNotFound(sheet_id));
                }
                tx.commit()?;
                return Ok(());
            }
        };
        let (deleted_model_sheet_id, sheet_order, ordered_sheet_ids) = {
            let (model_sheet_id,): (Option<i64>,) = tx.query_row(
                "SELECT model_sheet_id FROM sheets WHERE id = ?1",
                params![sheet_id.to_string()],
                |r| Ok((r.get::<_, Option<i64>>(0).ok().flatten(),)),
            )?;

            let mut sheet_order = Vec::new();
            let mut ordered_sheet_ids = Vec::new();
            let mut stmt = tx.prepare(
                r#"
                SELECT name, model_sheet_id
                FROM sheets
                WHERE workbook_id = ?1
                ORDER BY COALESCE(position, 0), id
                "#,
            )?;
            let mut rows = stmt.query(params![meta.workbook_id.to_string()])?;
            while let Some(row) = rows.next()? {
                let Ok(name) = row.get::<_, String>(0) else {
                    continue;
                };
                let model_sheet_id: Option<i64> = row.get::<_, Option<i64>>(1).ok().flatten();
                let parsed = model_sheet_id.and_then(|id| u32::try_from(id).ok());
                sheet_order.push(name.clone());
                ordered_sheet_ids.push((name, parsed));
            }

            (model_sheet_id.and_then(|id| u32::try_from(id).ok()), sheet_order, ordered_sheet_ids)
        };

        delete_named_ranges_for_sheet_scope_tx(&tx, meta.workbook_id, &meta.name)?;
        tx.execute(
            "DELETE FROM sheet_drawings WHERE sheet_id = ?1",
            params![sheet_id.to_string()],
        )?;
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

        rewrite_sheet_delete_references_tx(
            &tx,
            meta.workbook_id,
            &meta.name,
            &sheet_order,
            deleted_model_sheet_id,
            &ordered_sheet_ids,
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
        let conn = lock_unpoisoned(&self.conn);
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
                let Ok(row) = r.get::<_, i64>(0) else {
                    return Ok(None);
                };
                let Ok(col) = r.get::<_, i64>(1) else {
                    return Ok(None);
                };
                let snapshot = snapshot_from_row(
                    r.get::<_, Option<String>>(2).ok().flatten(),
                    r.get::<_, Option<f64>>(3).ok().flatten(),
                    r.get::<_, Option<String>>(4).ok().flatten(),
                    r.get::<_, Option<String>>(5).ok().flatten(),
                    r.get::<_, Option<String>>(6).ok().flatten(),
                    r.get::<_, Option<i64>>(7).ok().flatten(),
                )?;
                Ok(Some(((row, col), snapshot)))
            },
        )?;

        let mut out = Vec::new();
        for item in rows {
            let Some(item) = item? else {
                continue;
            };
            out.push(item);
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
        let conn = lock_unpoisoned(&self.conn);
        let count: u64 = conn.query_row(
            "SELECT COUNT(*) FROM cells WHERE sheet_id = ?1",
            params![sheet_id.to_string()],
            |r| r.get(0),
        )?;
        Ok(count)
    }

    pub fn change_log_count(&self, sheet_id: Uuid) -> Result<u64> {
        let conn = lock_unpoisoned(&self.conn);
        let count: u64 = conn.query_row(
            "SELECT COUNT(*) FROM change_log WHERE sheet_id = ?1",
            params![sheet_id.to_string()],
            |r| r.get(0),
        )?;
        Ok(count)
    }

    pub fn latest_change(&self, sheet_id: Uuid) -> Result<Option<ChangeLogEntry>> {
        let conn = lock_unpoisoned(&self.conn);
        let mut stmt = conn.prepare(
            r#"
            SELECT id, sheet_id, user_id, operation, target, old_value, new_value
            FROM change_log
            WHERE sheet_id = ?1
            ORDER BY id DESC
            "#,
        )?;
        let mut rows = stmt.query(params![sheet_id.to_string()])?;
        while let Some(row) = rows.next()? {
            let Ok(id) = row.get::<_, i64>(0) else {
                continue;
            };
            let Ok(sheet_id_raw) = row.get::<_, String>(1) else {
                continue;
            };
            let Ok(sheet_id) = Uuid::parse_str(&sheet_id_raw).map_err(|_| rusqlite::Error::InvalidQuery)
            else {
                continue;
            };
            let target_raw: Option<String> = row.get::<_, Option<String>>(4).ok().flatten();
            let old_value_raw: Option<String> = row.get::<_, Option<String>>(5).ok().flatten();
            let new_value_raw: Option<String> = row.get::<_, Option<String>>(6).ok().flatten();
            return Ok(Some(ChangeLogEntry {
                id,
                sheet_id,
                user_id: row.get::<_, Option<String>>(2).ok().flatten(),
                operation: row.get::<_, String>(3).unwrap_or_default(),
                target: parse_optional_json_value(target_raw).unwrap_or(JsonValue::Null),
                old_value: parse_optional_json_value(old_value_raw).unwrap_or(JsonValue::Null),
                new_value: parse_optional_json_value(new_value_raw).unwrap_or(JsonValue::Null),
            }));
        }
        Ok(None)
    }

    pub fn apply_cell_changes(&self, changes: &[CellChange]) -> Result<()> {
        if changes.is_empty() {
            return Ok(());
        }

        let mut conn = lock_unpoisoned(&self.conn);
        let tx = conn.transaction()?;

        for change in changes {
            apply_one_change(&tx, change)?;
        }

        tx.commit()?;
        Ok(())
    }

    pub fn get_or_insert_style(&self, style: &StorageStyle) -> Result<i64> {
        let mut conn = lock_unpoisoned(&self.conn);
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
        let conn = lock_unpoisoned(&self.conn);
        let wants_workbook_scope = scope.eq_ignore_ascii_case("workbook");
        let mut stmt = conn.prepare(
            r#"
            SELECT rowid, workbook_id, name, scope, reference
            FROM named_ranges
            WHERE workbook_id = ?1
              AND name = ?2 COLLATE NOCASE
            ORDER BY rowid DESC
            "#,
        )?;
        let mut rows = stmt.query(params![workbook_id.to_string(), name])?;

        while let Some(row) = rows.next()? {
            let Ok(scope_value) = row.get::<_, String>(3) else {
                continue;
            };

            let scope_matches = if wants_workbook_scope {
                scope_value.eq_ignore_ascii_case("workbook")
            } else {
                !scope_value.eq_ignore_ascii_case("workbook")
                    && sheet_name_eq_case_insensitive(&scope_value, scope)
            };
            if !scope_matches {
                continue;
            }

            let Ok(name) = row.get::<_, String>(2) else {
                continue;
            };
            let Ok(reference) = row.get::<_, String>(4) else {
                continue;
            };
            return Ok(Some(NamedRange {
                workbook_id,
                name,
                scope: scope_value,
                reference,
            }));
        }

        Ok(None)
    }

    pub fn upsert_named_range(&self, range: &NamedRange) -> Result<()> {
        let mut conn = lock_unpoisoned(&self.conn);
        let tx = conn.transaction()?;

        let wants_workbook_scope = range.scope.eq_ignore_ascii_case("workbook");
        let mut existing_rowid: Option<i64> = None;
        let mut duplicate_rowids: Vec<i64> = Vec::new();
        {
            let mut stmt = tx.prepare(
                r#"
                SELECT rowid, name, scope
                FROM named_ranges
                WHERE workbook_id = ?1
                  AND name = ?2 COLLATE NOCASE
                ORDER BY rowid DESC
                "#,
            )?;
            let mut rows = stmt.query(params![range.workbook_id.to_string(), &range.name])?;
            while let Some(row) = rows.next()? {
                let Ok(rowid) = row.get::<_, i64>(0) else {
                    continue;
                };
                let Ok(scope) = row.get::<_, String>(2) else {
                    continue;
                };
                let scope_matches = if wants_workbook_scope {
                    scope.eq_ignore_ascii_case("workbook")
                } else {
                    !scope.eq_ignore_ascii_case("workbook")
                        && sheet_name_eq_case_insensitive(&scope, &range.scope)
                };
                if scope_matches {
                    if existing_rowid.is_none() {
                        existing_rowid = Some(rowid);
                    } else {
                        duplicate_rowids.push(rowid);
                    }
                }
            }
        }

        match existing_rowid {
            Some(rowid) => {
                tx.execute(
                    r#"
                    UPDATE named_ranges
                    SET reference = ?1
                    WHERE rowid = ?2
                    "#,
                    params![&range.reference, rowid],
                )?;
                if !duplicate_rowids.is_empty() {
                    let mut delete_stmt =
                        tx.prepare("DELETE FROM named_ranges WHERE rowid = ?1")?;
                    for duplicate in duplicate_rowids {
                        delete_stmt.execute(params![duplicate])?;
                    }
                }
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

        sync_named_range_into_defined_names_tx(&tx, range)?;
        touch_workbook_modified_at_by_workbook_id(&tx, range.workbook_id)?;
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
              COALESCE(position, 0),
              COALESCE(visibility, 'visible'),
              tab_color,
              xlsx_sheet_id,
              xlsx_rel_id,
              COALESCE(frozen_rows, 0),
              COALESCE(frozen_cols, 0),
              COALESCE(zoom, 1.0),
              metadata
            FROM sheets
            WHERE workbook_id = ?1
            ORDER BY COALESCE(position, 0), id
            "#,
        )?;

        let rows = stmt.query_map(params![workbook_id.to_string()], |r| {
            let Some(id) = r.get::<_, Option<String>>(0).ok().flatten() else {
                return Ok(None);
            };
            let Ok(id) = Uuid::parse_str(&id).map_err(|_| rusqlite::Error::InvalidQuery) else {
                return Ok(None);
            };

            let Some(workbook_id) = r.get::<_, Option<String>>(1).ok().flatten() else {
                return Ok(None);
            };
            let Ok(workbook_id) =
                Uuid::parse_str(&workbook_id).map_err(|_| rusqlite::Error::InvalidQuery)
            else {
                return Ok(None);
            };

            let name = r
                .get::<_, Option<String>>(2)
                .ok()
                .flatten()
                .unwrap_or_else(|| invalid_sheet_name_placeholder(id));
            let visibility: String = r.get(4).unwrap_or_else(|_| "visible".to_string());
            let metadata_raw: Option<String> = r.get::<_, Option<String>>(11).ok().flatten();
            Ok(Some(SheetMeta {
                id,
                workbook_id,
                name,
                position: r.get(3).unwrap_or(0),
                visibility: SheetVisibility::parse(&visibility),
                tab_color: r.get::<_, Option<String>>(5).ok().flatten(),
                xlsx_sheet_id: r.get::<_, Option<i64>>(6).ok().flatten(),
                xlsx_rel_id: r.get::<_, Option<String>>(7).ok().flatten(),
                frozen_rows: r.get(8).unwrap_or(0),
                frozen_cols: r.get(9).unwrap_or(0),
                zoom: r.get(10).unwrap_or(1.0),
                metadata: parse_optional_json_value(metadata_raw),
            }))
        })?;

        let mut sheets = Vec::new();
        for sheet in rows {
            let Some(sheet) = sheet? else {
                continue;
            };
            sheets.push(sheet);
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
                  COALESCE(position, 0),
                  COALESCE(visibility, 'visible'),
                  tab_color,
                  xlsx_sheet_id,
                  xlsx_rel_id,
                  COALESCE(frozen_rows, 0),
                  COALESCE(frozen_cols, 0),
                  COALESCE(zoom, 1.0),
                  metadata
                FROM sheets
                WHERE id = ?1
                "#,
                params![sheet_id.to_string()],
                |r| {
                    let id: String = r.get(0)?;
                    let workbook_id: String = r.get(1)?;
                    let visibility: String = r.get(4).unwrap_or_else(|_| "visible".to_string());
                    let metadata_raw: Option<String> = r.get::<_, Option<String>>(11).ok().flatten();
                    let id_parsed =
                        Uuid::parse_str(&id).map_err(|_| rusqlite::Error::InvalidQuery)?;
                    Ok(SheetMeta {
                        id: id_parsed,
                        workbook_id: Uuid::parse_str(&workbook_id)
                            .map_err(|_| rusqlite::Error::InvalidQuery)?,
                        name: r
                            .get::<_, Option<String>>(2)
                            .ok()
                            .flatten()
                            .unwrap_or_else(|| invalid_sheet_name_placeholder(id_parsed)),
                        position: r.get(3).unwrap_or(0),
                        visibility: SheetVisibility::parse(&visibility),
                        tab_color: r.get::<_, Option<String>>(5).ok().flatten(),
                        xlsx_sheet_id: r.get::<_, Option<i64>>(6).ok().flatten(),
                        xlsx_rel_id: r.get::<_, Option<String>>(7).ok().flatten(),
                        frozen_rows: r.get(8).unwrap_or(0),
                        frozen_cols: r.get(9).unwrap_or(0),
                        zoom: r.get(10).unwrap_or(1.0),
                        metadata: parse_optional_json_value(metadata_raw),
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
    let value_type_str = value_type.as_deref();
    let raw_json = value_json.as_deref().filter(|s| !s.trim().is_empty());
    let parse_json = |raw: &str| serde_json::from_str::<CellValue>(raw).ok();

    // Prefer the scalar columns when they are available, even if `value_json` is present.
    // This keeps `cells.value_json` as the canonical stored representation (we always write it),
    // while allowing reads to avoid JSON parsing on hot paths.
    let value = match value_type_str {
        Some("number") => match value_number {
            Some(n) => CellValue::Number(n),
            None => raw_json.and_then(parse_json).unwrap_or(CellValue::Number(0.0)),
        },
        Some("string") => match value_string {
            Some(s) => CellValue::String(s),
            None => raw_json
                .and_then(parse_json)
                .unwrap_or_else(|| CellValue::String(String::new())),
        },
        Some("boolean") => match value_number {
            Some(n) => CellValue::Boolean(n != 0.0),
            None => raw_json.and_then(parse_json).unwrap_or(CellValue::Boolean(false)),
        },
        Some("error") => match value_string {
            Some(s) => {
                let parsed = s.parse::<ErrorValue>().unwrap_or(ErrorValue::Unknown);
                CellValue::Error(parsed)
            }
            None => raw_json
                .and_then(parse_json)
                .unwrap_or(CellValue::Error(ErrorValue::Unknown)),
        },
        // Legacy sentinel used by older schema versions when a cell contains a formula but no cached value.
        Some("formula") => CellValue::Empty,
        // Unknown/NULL types fall back to canonical JSON when present.
        _ => raw_json.and_then(parse_json).unwrap_or_else(|| match value_type_str {
            // `NULL` value_type means a style-only blank cell.
            None => CellValue::Empty,
            // Unknown value types are treated as strings to preserve data.
            Some(other) => CellValue::String(other.to_string()),
        }),
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
    let previous_style_id = old_snapshot.as_ref().and_then(|s| s.style_id);

    // If no explicit style payload is provided, preserve the existing style. This
    // matches Excel semantics where editing/clearing a cell's contents typically
    // does not clear formatting (cell styles remain unless explicitly changed).
    let style_id = match &change.data.style {
        Some(style) => Some(get_or_insert_style_tx(tx, style)?),
        None => previous_style_id,
    };

    let is_empty = change.data.value.is_empty() && change.data.formula.is_none();

    if is_empty && style_id.is_none() {
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
        CellValue::RichText(_)
        | CellValue::Entity(_)
        | CellValue::Record(_)
        | CellValue::Image(_)
        | CellValue::Array(_)
        | CellValue::Spill(_) => (None, None, None),
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
                    r.get::<_, Option<String>>(0).ok().flatten(),
                    r.get::<_, Option<f64>>(1).ok().flatten(),
                    r.get::<_, Option<String>>(2).ok().flatten(),
                    r.get::<_, Option<String>>(3).ok().flatten(),
                    r.get::<_, Option<String>>(4).ok().flatten(),
                    r.get::<_, Option<i64>>(5).ok().flatten(),
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

    let schema = std::ffi::CString::new("main").map_err(|_| rusqlite::Error::InvalidQuery)?;
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
        CellValue::RichText(_)
        | CellValue::Entity(_)
        | CellValue::Record(_)
        | CellValue::Image(_)
        | CellValue::Array(_)
        | CellValue::Spill(_) => (None, None, None),
    }
}

fn worksheet_metadata_json(sheet: &formula_model::Worksheet) -> Result<Option<JsonValue>> {
    let mut comments: BTreeMap<formula_model::CellKey, Vec<formula_model::Comment>> =
        BTreeMap::new();
    for (cell_ref, comment) in sheet.iter_comments() {
        comments
            .entry(formula_model::CellKey::from(cell_ref))
            .or_default()
            .push(comment.clone());
    }

    let mut map = serde_json::Map::new();
    map.insert("id".to_string(), serde_json::to_value(sheet.id)?);
    map.insert("name".to_string(), serde_json::to_value(&sheet.name)?);
    map.insert(
        "xlsx_sheet_id".to_string(),
        serde_json::to_value(sheet.xlsx_sheet_id)?,
    );
    map.insert(
        "xlsx_rel_id".to_string(),
        serde_json::to_value(&sheet.xlsx_rel_id)?,
    );
    map.insert("visibility".to_string(), serde_json::to_value(sheet.visibility)?);
    map.insert("tab_color".to_string(), serde_json::to_value(&sheet.tab_color)?);
    // `drawings` are persisted separately in `sheet_drawings` to avoid duplicating
    // large drawing payloads in this JSON blob.

    map.insert("tables".to_string(), serde_json::to_value(&sheet.tables)?);
    map.insert(
        "auto_filter".to_string(),
        serde_json::to_value(&sheet.auto_filter)?,
    );
    map.insert(
        "conditional_formatting_rules".to_string(),
        serde_json::to_value(&sheet.conditional_formatting_rules)?,
    );
    map.insert(
        "conditional_formatting_dxfs".to_string(),
        serde_json::to_value(&sheet.conditional_formatting_dxfs)?,
    );
    map.insert("row_count".to_string(), serde_json::to_value(sheet.row_count)?);
    map.insert("col_count".to_string(), serde_json::to_value(sheet.col_count)?);
    map.insert(
        "merged_regions".to_string(),
        serde_json::to_value(&sheet.merged_regions)?,
    );
    map.insert(
        "row_properties".to_string(),
        serde_json::to_value(&sheet.row_properties)?,
    );
    map.insert(
        "col_properties".to_string(),
        serde_json::to_value(&sheet.col_properties)?,
    );
    map.insert("outline".to_string(), serde_json::to_value(&sheet.outline)?);
    map.insert(
        "frozen_rows".to_string(),
        serde_json::to_value(sheet.frozen_rows)?,
    );
    map.insert(
        "frozen_cols".to_string(),
        serde_json::to_value(sheet.frozen_cols)?,
    );
    map.insert("zoom".to_string(), serde_json::to_value(sheet.zoom)?);
    map.insert("view".to_string(), serde_json::to_value(&sheet.view)?);
    map.insert(
        "hyperlinks".to_string(),
        serde_json::to_value(&sheet.hyperlinks)?,
    );
    map.insert(
        "data_validations".to_string(),
        serde_json::to_value(&sheet.data_validations)?,
    );
    map.insert("comments".to_string(), serde_json::to_value(comments)?);
    map.insert(
        "sheet_protection".to_string(),
        serde_json::to_value(&sheet.sheet_protection)?,
    );

    Ok(Some(JsonValue::Object(map)))
}

fn normalize_refers_to(value: &str) -> String {
    let trimmed = value.trim();
    trimmed.strip_prefix('=').unwrap_or(trimmed).to_string()
}

fn canonical_sheet_name_key(name: &str) -> String {
    name.nfkc().flat_map(|c| c.to_uppercase()).collect()
}

fn merge_named_ranges_into_defined_names(
    conn: &Connection,
    workbook_id: Uuid,
    sheets: &[formula_model::Worksheet],
    defined_names: &mut Vec<DefinedName>,
) -> Result<()> {
    let mut sheet_ids: HashMap<String, u32> = HashMap::new();
    for sheet in sheets {
        sheet_ids.insert(canonical_sheet_name_key(&sheet.name), sheet.id);
    }

    let mut sheet_sort_keys: HashMap<u32, String> = HashMap::new();
    for sheet in sheets {
        sheet_sort_keys.insert(sheet.id, canonical_sheet_name_key(&sheet.name));
    }

    let mut stmt = conn.prepare(
        r#"
        SELECT rowid, name, scope, reference
        FROM named_ranges
        WHERE workbook_id = ?1
        ORDER BY rowid DESC
        "#,
    )?;
    let mut rows = stmt.query(params![workbook_id.to_string()])?;

    struct LegacyNamedRangeRow {
        name: String,
        scope: DefinedNameScope,
        reference: String,
    }

    // Legacy databases may contain duplicate rows that differ only by case in the scope or name
    // fields (especially for non-ASCII sheet names). SQLite's `COLLATE NOCASE` is ASCII-only, so
    // the relative ordering of those duplicates is not necessarily insertion-ordered.
    //
    // Prefer the newest row (highest `rowid`) for each `(scope, name)` pair and then merge the
    // surviving rows into the workbook in a deterministic order (scope/name).
    let mut winners: BTreeMap<(u8, String, String), LegacyNamedRangeRow> = BTreeMap::new();
    while let Some(row) = rows.next()? {
        let Ok(name) = row.get::<_, String>(1) else {
            continue;
        };
        let Ok(scope_raw) = row.get::<_, String>(2) else {
            continue;
        };
        let Ok(reference) = row.get::<_, String>(3) else {
            continue;
        };

        let (scope, scope_rank, scope_sort_key) = if scope_raw.eq_ignore_ascii_case("workbook") {
            (DefinedNameScope::Workbook, 0u8, String::new())
        } else if let Some(sheet_id) = sheet_ids.get(&canonical_sheet_name_key(&scope_raw)).copied()
        {
            let sort_key = sheet_sort_keys
                .get(&sheet_id)
                .cloned()
                .unwrap_or_else(String::new);
            (DefinedNameScope::Sheet(sheet_id), 1u8, sort_key)
        } else {
            // Unknown scope (likely stale sheet name) - skip rather than generating broken entries.
            continue;
        };

        let name_key = name.to_ascii_lowercase();
        winners
            .entry((scope_rank, scope_sort_key, name_key))
            .or_insert(LegacyNamedRangeRow {
                name,
                scope,
                reference,
            });
    }

    let mut next_id = defined_names
        .iter()
        .map(|n| n.id)
        .max()
        .unwrap_or(0)
        .wrapping_add(1);

    for row in winners.values() {
        let refers_to = normalize_refers_to(&row.reference);

        if let Some(existing) = defined_names
            .iter_mut()
            .find(|n| n.scope == row.scope && n.name.eq_ignore_ascii_case(&row.name))
        {
            existing.refers_to = refers_to;
            continue;
        }

        defined_names.push(DefinedName {
            id: next_id,
            name: row.name.clone(),
            scope: row.scope,
            refers_to,
            comment: None,
            hidden: false,
            xlsx_local_sheet_id: None,
        });
        next_id = next_id.wrapping_add(1);
    }

    Ok(())
}

fn sync_named_range_into_defined_names_tx(tx: &Transaction<'_>, range: &NamedRange) -> Result<()> {
    let workbook_id_str = range.workbook_id.to_string();
    let wants_workbook_scope = range.scope.eq_ignore_ascii_case("workbook");

    let scope = if wants_workbook_scope {
        DefinedNameScope::Workbook
    } else {
        let sheet_key = canonical_sheet_name_key(&range.scope);
        let mut stmt =
            tx.prepare("SELECT name, model_sheet_id FROM sheets WHERE workbook_id = ?1")?;
        let mut rows = stmt.query(params![&workbook_id_str])?;
        let mut sheet_id: Option<u32> = None;
        while let Some(row) = rows.next()? {
            let Ok(name) = row.get::<_, String>(0) else {
                continue;
            };
            let model_sheet_id: Option<i64> = row.get::<_, Option<i64>>(1).ok().flatten();
            if canonical_sheet_name_key(&name) == sheet_key {
                sheet_id = model_sheet_id.and_then(|id| u32::try_from(id).ok());
                break;
            }
        }

        let Some(sheet_id) = sheet_id else {
            // Unknown scope (likely stale sheet name); keep the legacy row but don't try to map it
            // into `workbooks.defined_names`.
            return Ok(());
        };

        DefinedNameScope::Sheet(sheet_id)
    };

    let Some(defined_names) = tx
        .query_row(
            "SELECT defined_names FROM workbooks WHERE id = ?1",
            params![&workbook_id_str],
            |r| Ok(r.get::<_, Option<String>>(0).ok().flatten()),
        )
        .optional()?
    else {
        // Orphaned named range row; avoid failing the legacy API just because the workbook row is
        // missing.
        return Ok(());
    };

    let mut names = match defined_names {
        Some(raw) => match serde_json::from_str::<Vec<DefinedName>>(&raw) {
            Ok(names) => names,
            // Corrupt JSON blob - avoid overwriting it from the legacy compatibility layer.
            Err(_) => return Ok(()),
        },
        None => Vec::new(),
    };

    let refers_to = normalize_refers_to(&range.reference);
    if let Some(existing) = names
        .iter_mut()
        .find(|n| n.scope == scope && n.name.eq_ignore_ascii_case(&range.name))
    {
        existing.refers_to = refers_to;
    } else {
        let mut used: HashSet<u32> = names.iter().map(|n| n.id).collect();
        let mut next_id = names.iter().map(|n| n.id).max().unwrap_or(0).wrapping_add(1);
        while used.contains(&next_id) {
            next_id = next_id.wrapping_add(1);
        }
        used.insert(next_id);

        names.push(DefinedName {
            id: next_id,
            name: range.name.clone(),
            scope,
            refers_to,
            comment: None,
            hidden: false,
            xlsx_local_sheet_id: None,
        });
    }

    let updated = (!names.is_empty()).then_some(serde_json::to_value(&names)?);
    tx.execute(
        "UPDATE workbooks SET defined_names = ?1 WHERE id = ?2",
        params![updated, &workbook_id_str],
    )?;
    Ok(())
}

fn update_named_range_scopes_for_sheet_rename_tx(
    tx: &Transaction<'_>,
    workbook_id: Uuid,
    old_name: &str,
    new_name: &str,
) -> Result<()> {
    let workbook_id_str = workbook_id.to_string();
    let old_key = canonical_sheet_name_key(old_name);

    let mut stmt = tx.prepare(
        r#"
        SELECT rowid, name, scope
        FROM named_ranges
        WHERE workbook_id = ?1
        ORDER BY rowid DESC
        "#,
    )?;
    let mut rows = stmt.query(params![&workbook_id_str])?;

    let mut to_update: Vec<(i64, String)> = Vec::new();
    while let Some(row) = rows.next()? {
        let Ok(rowid) = row.get::<_, i64>(0) else {
            continue;
        };
        let Ok(name) = row.get::<_, String>(1) else {
            continue;
        };
        let Ok(scope) = row.get::<_, String>(2) else {
            continue;
        };
        if scope.eq_ignore_ascii_case("workbook") {
            continue;
        }
        if canonical_sheet_name_key(&scope) == old_key {
            to_update.push((rowid, name));
        }
    }

    if to_update.is_empty() {
        return Ok(());
    }

    let mut update_stmt = tx.prepare("UPDATE named_ranges SET scope = ?1 WHERE rowid = ?2")?;
    let mut delete_stmt = tx.prepare("DELETE FROM named_ranges WHERE rowid = ?1")?;
    let mut conflict_stmt = tx.prepare(
        r#"
        SELECT rowid
        FROM named_ranges
        WHERE workbook_id = ?1
          AND name = ?2
          AND scope = ?3
        LIMIT 1
        "#,
    )?;

    for (rowid, name) in to_update {
        let update_result = update_stmt.execute(params![new_name, rowid]);
        match update_result {
            Ok(_) => {}
            Err(rusqlite::Error::SqliteFailure(err, _))
                if err.code == rusqlite::ffi::ErrorCode::ConstraintViolation =>
            {
                let conflict_rowid: i64 = conflict_stmt.query_row(
                    params![&workbook_id_str, &name, new_name],
                    |r| r.get(0),
                )?;

                if conflict_rowid > rowid {
                    delete_stmt.execute(params![rowid])?;
                } else {
                    delete_stmt.execute(params![conflict_rowid])?;
                    update_stmt.execute(params![new_name, rowid])?;
                }
            }
            Err(err) => return Err(err.into()),
        }
    }

    Ok(())
}

fn delete_named_ranges_for_sheet_scope_tx(
    tx: &Transaction<'_>,
    workbook_id: Uuid,
    sheet_name: &str,
) -> Result<()> {
    let workbook_id_str = workbook_id.to_string();
    let sheet_key = canonical_sheet_name_key(sheet_name);

    let mut stmt = tx.prepare("SELECT name, scope FROM named_ranges WHERE workbook_id = ?1")?;
    let mut rows = stmt.query(params![&workbook_id_str])?;

    let mut to_delete: Vec<(String, String)> = Vec::new();
    while let Some(row) = rows.next()? {
        let Ok(name) = row.get::<_, String>(0) else {
            continue;
        };
        let Ok(scope) = row.get::<_, String>(1) else {
            continue;
        };
        if scope.eq_ignore_ascii_case("workbook") {
            continue;
        }
        if canonical_sheet_name_key(&scope) == sheet_key {
            to_delete.push((name, scope));
        }
    }

    if to_delete.is_empty() {
        return Ok(());
    }

    let mut delete_stmt = tx.prepare(
        "DELETE FROM named_ranges WHERE workbook_id = ?1 AND name = ?2 AND scope = ?3",
    )?;
    for (name, scope) in to_delete {
        delete_stmt.execute(params![&workbook_id_str, name, scope])?;
    }

    Ok(())
}

fn update_sheet_model_json_tx<F>(
    tx: &Transaction<'_>,
    sheet_id: Uuid,
    f: F,
) -> Result<()>
where
    F: FnOnce(&mut formula_model::Worksheet),
{
    let model_sheet_json: Option<String> = tx.query_row(
        "SELECT model_sheet_json FROM sheets WHERE id = ?1",
        params![sheet_id.to_string()],
        |r| Ok(r.get::<_, Option<String>>(0).ok().flatten()),
    )?;
    let Some(raw) = model_sheet_json else {
        return Ok(());
    };

    let Ok(mut sheet) = serde_json::from_str::<formula_model::Worksheet>(&raw) else {
        return Ok(());
    };
    f(&mut sheet);
    let updated = worksheet_metadata_json(&sheet)?;
    tx.execute(
        "UPDATE sheets SET model_sheet_json = ?1 WHERE id = ?2",
        params![updated, sheet_id.to_string()],
    )?;
    Ok(())
}

fn rewrite_sheet_rename_references_tx(
    tx: &Transaction<'_>,
    workbook_id: Uuid,
    old_name: &str,
    new_name: &str,
) -> Result<()> {
    let workbook_id_str = workbook_id.to_string();

    // Update formulas stored in the cell grid (including formulas on other sheets that reference
    // the renamed sheet).
    {
        let mut select_stmt = tx.prepare(
            r#"
            SELECT c.sheet_id, c.row, c.col, c.formula
            FROM cells c
            JOIN sheets s ON s.id = c.sheet_id
            WHERE s.workbook_id = ?1
              AND c.formula IS NOT NULL
            "#,
        )?;
        let mut update_stmt =
            tx.prepare("UPDATE cells SET formula = ?1 WHERE sheet_id = ?2 AND row = ?3 AND col = ?4")?;

        let mut rows = select_stmt.query(params![&workbook_id_str])?;
        while let Some(row) = rows.next()? {
            let Ok(sheet_id) = row.get::<_, String>(0) else {
                continue;
            };
            let Ok(row_idx) = row.get::<_, i64>(1) else {
                continue;
            };
            let Ok(col_idx) = row.get::<_, i64>(2) else {
                continue;
            };
            let Ok(formula) = row.get::<_, String>(3) else {
                continue;
            };
            let rewritten = rewrite_sheet_names_in_formula(&formula, old_name, new_name);
            if rewritten != formula {
                update_stmt.execute(params![rewritten, sheet_id, row_idx, col_idx])?;
            }
        }
    }

    // Update named range references (legacy table) so `get_named_range` remains correct.
    {
        let mut select_stmt = tx.prepare(
            r#"
            SELECT name, scope, reference
            FROM named_ranges
            WHERE workbook_id = ?1
            "#,
        )?;
        let mut update_stmt = tx.prepare(
            "UPDATE named_ranges SET reference = ?1 WHERE workbook_id = ?2 AND name = ?3 AND scope = ?4",
        )?;

        let mut rows = select_stmt.query(params![&workbook_id_str])?;
        while let Some(row) = rows.next()? {
            let Ok(name) = row.get::<_, String>(0) else {
                continue;
            };
            let Ok(scope) = row.get::<_, String>(1) else {
                continue;
            };
            let Ok(reference) = row.get::<_, String>(2) else {
                continue;
            };
            let rewritten = rewrite_sheet_names_in_formula(&reference, old_name, new_name);
            if rewritten != reference {
                update_stmt.execute(params![rewritten, &workbook_id_str, name, scope])?;
            }
        }
    }

    // Keep workbook-level JSON columns in sync so `export_model_workbook` round-trips correctly.
    {
        let defined_names: Option<String> = tx
            .query_row(
                "SELECT defined_names FROM workbooks WHERE id = ?1",
                params![&workbook_id_str],
                |r| Ok(r.get::<_, Option<String>>(0).ok().flatten()),
            )
            .optional()?
            .flatten();
        if let Some(raw) = defined_names {
            if let Ok(mut names) = serde_json::from_str::<Vec<DefinedName>>(&raw) {
                let mut changed = false;
                for name in &mut names {
                    let rewritten =
                        rewrite_sheet_names_in_formula(&name.refers_to, old_name, new_name);
                    if rewritten != name.refers_to {
                        name.refers_to = rewritten;
                        changed = true;
                    }
                }
                if changed {
                    let updated = (!names.is_empty())
                        .then_some(serde_json::to_value(&names)?);
                    tx.execute(
                        "UPDATE workbooks SET defined_names = ?1 WHERE id = ?2",
                        params![updated, &workbook_id_str],
                    )?;
                }
            }
        }
    }

    {
        let print_settings: Option<String> = tx
            .query_row(
                "SELECT print_settings FROM workbooks WHERE id = ?1",
                params![&workbook_id_str],
                |r| Ok(r.get::<_, Option<String>>(0).ok().flatten()),
            )
            .optional()?
            .flatten();
        if let Some(raw) = print_settings {
            if let Ok(mut settings) =
                serde_json::from_str::<formula_model::WorkbookPrintSettings>(&raw)
            {
                let mut changed = false;
                for sheet_settings in &mut settings.sheets {
                    if sheet_name_eq_case_insensitive(&sheet_settings.sheet_name, old_name) {
                        sheet_settings.sheet_name = new_name.to_string();
                        changed = true;
                    }
                }
                if changed {
                    let updated = (!settings.is_empty())
                        .then_some(serde_json::to_value(&settings)?);
                    tx.execute(
                        "UPDATE workbooks SET print_settings = ?1 WHERE id = ?2",
                        params![updated, &workbook_id_str],
                    )?;
                }
            }
        }
    }

    rewrite_sheet_metadata_json_for_rename_tx(tx, &workbook_id_str, old_name, new_name)?;

    Ok(())
}

fn rewrite_sheet_metadata_json_for_rename_tx(
    tx: &Transaction<'_>,
    workbook_id: &str,
    old_name: &str,
    new_name: &str,
) -> Result<()> {
    let mut select_stmt = tx.prepare(
        r#"
        SELECT id, model_sheet_json
        FROM sheets
        WHERE workbook_id = ?1
          AND model_sheet_json IS NOT NULL
        "#,
    )?;
    let mut update_stmt = tx.prepare("UPDATE sheets SET model_sheet_json = ?1 WHERE id = ?2")?;

    let mut rows = select_stmt.query(params![workbook_id])?;
    while let Some(row) = rows.next()? {
        let Ok(sheet_id) = row.get::<_, String>(0) else {
            continue;
        };
        let Ok(json) = row.get::<_, String>(1) else {
            continue;
        };
        let Ok(mut sheet) = serde_json::from_str::<formula_model::Worksheet>(&json) else {
            continue;
        };

        if sheet_name_eq_case_insensitive(&sheet.name, old_name) {
            sheet.name = new_name.to_string();
        }

        rewrite_sheet_references_in_sheet_metadata_for_rename(&mut sheet, old_name, new_name);
        let updated = worksheet_metadata_json(&sheet)?;
        update_stmt.execute(params![updated, sheet_id])?;
    }

    Ok(())
}

fn rewrite_sheet_references_in_sheet_metadata_for_rename(
    sheet: &mut formula_model::Worksheet,
    old_name: &str,
    new_name: &str,
) {
    for table in &mut sheet.tables {
        for column in &mut table.columns {
            if let Some(formula) = column.formula.as_mut() {
                *formula = rewrite_sheet_names_in_formula(formula, old_name, new_name);
            }
            if let Some(formula) = column.totals_formula.as_mut() {
                *formula = rewrite_sheet_names_in_formula(formula, old_name, new_name);
            }
        }
    }

    for rule in &mut sheet.conditional_formatting_rules {
        rewrite_cf_rule_kind_for_rename(&mut rule.kind, old_name, new_name);
    }

    for link in &mut sheet.hyperlinks {
        if let formula_model::HyperlinkTarget::Internal { sheet: target, .. } = &mut link.target {
            if sheet_name_eq_case_insensitive(target, old_name) {
                *target = new_name.to_string();
            }
        }
    }

    for assignment in &mut sheet.data_validations {
        rewrite_data_validation_for_rename(&mut assignment.validation, old_name, new_name);
    }
}

fn rewrite_cf_rule_kind_for_rename(
    kind: &mut formula_model::CfRuleKind,
    old_name: &str,
    new_name: &str,
) {
    match kind {
        formula_model::CfRuleKind::CellIs { formulas, .. } => {
            for formula in formulas {
                *formula = rewrite_sheet_names_in_formula(formula, old_name, new_name);
            }
        }
        formula_model::CfRuleKind::Expression { formula } => {
            *formula = rewrite_sheet_names_in_formula(formula, old_name, new_name);
        }
        formula_model::CfRuleKind::DataBar(rule) => {
            rewrite_cfvo_for_rename(&mut rule.min, old_name, new_name);
            rewrite_cfvo_for_rename(&mut rule.max, old_name, new_name);
        }
        formula_model::CfRuleKind::ColorScale(rule) => {
            for cfvo in &mut rule.cfvos {
                rewrite_cfvo_for_rename(cfvo, old_name, new_name);
            }
        }
        formula_model::CfRuleKind::IconSet(rule) => {
            for cfvo in &mut rule.cfvos {
                rewrite_cfvo_for_rename(cfvo, old_name, new_name);
            }
        }
        formula_model::CfRuleKind::TopBottom(_)
        | formula_model::CfRuleKind::UniqueDuplicate(_)
        | formula_model::CfRuleKind::Unsupported { .. } => {}
    }
}

fn rewrite_cfvo_for_rename(cfvo: &mut formula_model::Cfvo, old_name: &str, new_name: &str) {
    if cfvo.type_ != formula_model::CfvoType::Formula {
        return;
    }
    let Some(value) = cfvo.value.as_mut() else {
        return;
    };
    *value = rewrite_sheet_names_in_formula(value, old_name, new_name);
}

fn normalize_validation_formula(formula: &str) -> &str {
    let trimmed = formula.trim();
    trimmed.strip_prefix('=').unwrap_or(trimmed).trim()
}

fn validation_formula_is_literal_list(formula: &str) -> bool {
    let s = normalize_validation_formula(formula);
    let bytes = s.as_bytes();
    if bytes.len() >= 2 && bytes.first() == Some(&b'"') && bytes.last() == Some(&b'"') {
        return true;
    }
    s.contains(',') || s.contains(';')
}

fn rewrite_data_validation_for_rename(
    validation: &mut formula_model::DataValidation,
    old_name: &str,
    new_name: &str,
) {
    let formula1_is_literal_list = validation.kind == formula_model::DataValidationKind::List
        && validation_formula_is_literal_list(&validation.formula1);

    if !formula1_is_literal_list && !validation.formula1.is_empty() {
        validation.formula1 =
            rewrite_sheet_names_in_formula(&validation.formula1, old_name, new_name);
    }
    if let Some(formula2) = validation.formula2.as_mut() {
        *formula2 = rewrite_sheet_names_in_formula(formula2, old_name, new_name);
    }
}

fn rewrite_sheet_delete_references_tx(
    tx: &Transaction<'_>,
    workbook_id: Uuid,
    deleted_name: &str,
    sheet_order: &[String],
    deleted_model_sheet_id: Option<u32>,
    ordered_sheet_ids: &[(String, Option<u32>)],
) -> Result<()> {
    let workbook_id_str = workbook_id.to_string();

    // Update formulas stored in the cell grid (including formulas on other sheets that reference
    // the deleted sheet).
    {
        let mut select_stmt = tx.prepare(
            r#"
            SELECT c.sheet_id, c.row, c.col, c.formula
            FROM cells c
            JOIN sheets s ON s.id = c.sheet_id
            WHERE s.workbook_id = ?1
              AND c.formula IS NOT NULL
            "#,
        )?;
        let mut update_stmt =
            tx.prepare("UPDATE cells SET formula = ?1 WHERE sheet_id = ?2 AND row = ?3 AND col = ?4")?;

        let mut rows = select_stmt.query(params![&workbook_id_str])?;
        while let Some(row) = rows.next()? {
            let Ok(sheet_id) = row.get::<_, String>(0) else {
                continue;
            };
            let Ok(row_idx) = row.get::<_, i64>(1) else {
                continue;
            };
            let Ok(col_idx) = row.get::<_, i64>(2) else {
                continue;
            };
            let Ok(formula) = row.get::<_, String>(3) else {
                continue;
            };
            let rewritten =
                rewrite_deleted_sheet_references_in_formula(&formula, deleted_name, sheet_order);
            if rewritten != formula {
                update_stmt.execute(params![rewritten, sheet_id, row_idx, col_idx])?;
            }
        }
    }

    // Update named range references (legacy table) so `get_named_range` remains correct.
    {
        let mut select_stmt = tx.prepare(
            r#"
            SELECT name, scope, reference
            FROM named_ranges
            WHERE workbook_id = ?1
            "#,
        )?;
        let mut update_stmt = tx.prepare(
            "UPDATE named_ranges SET reference = ?1 WHERE workbook_id = ?2 AND name = ?3 AND scope = ?4",
        )?;

        let mut rows = select_stmt.query(params![&workbook_id_str])?;
        while let Some(row) = rows.next()? {
            let Ok(name) = row.get::<_, String>(0) else {
                continue;
            };
            let Ok(scope) = row.get::<_, String>(1) else {
                continue;
            };
            let Ok(reference) = row.get::<_, String>(2) else {
                continue;
            };
            let rewritten = rewrite_deleted_sheet_references_in_formula(
                &reference,
                deleted_name,
                sheet_order,
            );
            if rewritten != reference {
                update_stmt.execute(params![rewritten, &workbook_id_str, name, scope])?;
            }
        }
    }

    // Keep workbook-level JSON columns in sync so `export_model_workbook` round-trips correctly.
    {
        let defined_names: Option<String> = tx
            .query_row(
                "SELECT defined_names FROM workbooks WHERE id = ?1",
                params![&workbook_id_str],
                |r| Ok(r.get::<_, Option<String>>(0).ok().flatten()),
            )
            .optional()?
            .flatten();
        if let Some(raw) = defined_names {
            if let Ok(mut names) = serde_json::from_str::<Vec<DefinedName>>(&raw) {
                let mut changed = false;
                if let Some(deleted_id) = deleted_model_sheet_id {
                    let before = names.len();
                    names.retain(|n| n.scope != DefinedNameScope::Sheet(deleted_id));
                    changed |= names.len() != before;
                }

                for name in &mut names {
                    let rewritten = rewrite_deleted_sheet_references_in_formula(
                        &name.refers_to,
                        deleted_name,
                        sheet_order,
                    );
                    if rewritten != name.refers_to {
                        name.refers_to = rewritten;
                        changed = true;
                    }
                }

                if changed {
                    let updated = (!names.is_empty()).then_some(serde_json::to_value(&names)?);
                    tx.execute(
                        "UPDATE workbooks SET defined_names = ?1 WHERE id = ?2",
                        params![updated, &workbook_id_str],
                    )?;
                }
            }
        }
    }

    {
        let print_settings: Option<String> = tx
            .query_row(
                "SELECT print_settings FROM workbooks WHERE id = ?1",
                params![&workbook_id_str],
                |r| Ok(r.get::<_, Option<String>>(0).ok().flatten()),
            )
            .optional()?
            .flatten();
        if let Some(raw) = print_settings {
            if let Ok(mut settings) =
                serde_json::from_str::<formula_model::WorkbookPrintSettings>(&raw)
            {
                let before = settings.sheets.len();
                settings
                    .sheets
                    .retain(|s| !sheet_name_eq_case_insensitive(&s.sheet_name, deleted_name));
                if settings.sheets.len() != before {
                    let updated =
                        (!settings.is_empty()).then_some(serde_json::to_value(&settings)?);
                    tx.execute(
                        "UPDATE workbooks SET print_settings = ?1 WHERE id = ?2",
                        params![updated, &workbook_id_str],
                    )?;
                }
            }
        }
    }

    {
        let view: Option<String> = tx
            .query_row(
                "SELECT view FROM workbooks WHERE id = ?1",
                params![&workbook_id_str],
                |r| Ok(r.get::<_, Option<String>>(0).ok().flatten()),
            )
            .optional()?
            .flatten();
        if let Some(raw) = view {
            if let Ok(mut view) = serde_json::from_str::<formula_model::WorkbookView>(&raw) {
                if let Some(deleted_id) = deleted_model_sheet_id {
                    if view.active_sheet_id == Some(deleted_id) {
                        let idx = ordered_sheet_ids
                            .iter()
                            .position(|(name, id)| {
                                sheet_name_eq_case_insensitive(name, deleted_name)
                                    || id == &Some(deleted_id)
                            });
                        if let Some(idx) = idx {
                            let replacement = if idx + 1 < ordered_sheet_ids.len() {
                                ordered_sheet_ids[idx + 1].1
                            } else if idx > 0 {
                                ordered_sheet_ids[idx - 1].1
                            } else {
                                None
                            };
                            view.active_sheet_id = replacement;
                        } else {
                            view.active_sheet_id = None;
                        }

                        let updated = (view != formula_model::WorkbookView::default())
                            .then_some(serde_json::to_value(&view)?);
                        tx.execute(
                            "UPDATE workbooks SET view = ?1 WHERE id = ?2",
                            params![updated, &workbook_id_str],
                        )?;
                    }
                }
            }
        }
    }

    rewrite_sheet_metadata_json_for_delete_tx(tx, &workbook_id_str, deleted_name, sheet_order)?;

    Ok(())
}

fn rewrite_sheet_metadata_json_for_delete_tx(
    tx: &Transaction<'_>,
    workbook_id: &str,
    deleted_name: &str,
    sheet_order: &[String],
) -> Result<()> {
    let mut select_stmt = tx.prepare(
        r#"
        SELECT id, model_sheet_json
        FROM sheets
        WHERE workbook_id = ?1
          AND model_sheet_json IS NOT NULL
        "#,
    )?;
    let mut update_stmt = tx.prepare("UPDATE sheets SET model_sheet_json = ?1 WHERE id = ?2")?;

    let mut rows = select_stmt.query(params![workbook_id])?;
    while let Some(row) = rows.next()? {
        let Ok(sheet_id) = row.get::<_, String>(0) else {
            continue;
        };
        let Ok(json) = row.get::<_, String>(1) else {
            continue;
        };
        let Ok(mut sheet) = serde_json::from_str::<formula_model::Worksheet>(&json) else {
            continue;
        };

        rewrite_sheet_references_in_sheet_metadata_for_delete(&mut sheet, deleted_name, sheet_order);
        let updated = worksheet_metadata_json(&sheet)?;
        update_stmt.execute(params![updated, sheet_id])?;
    }

    Ok(())
}

fn rewrite_sheet_references_in_sheet_metadata_for_delete(
    sheet: &mut formula_model::Worksheet,
    deleted_name: &str,
    sheet_order: &[String],
) {
    for table in &mut sheet.tables {
        for column in &mut table.columns {
            if let Some(formula) = column.formula.as_mut() {
                *formula =
                    rewrite_deleted_sheet_references_in_formula(formula, deleted_name, sheet_order);
            }
            if let Some(formula) = column.totals_formula.as_mut() {
                *formula =
                    rewrite_deleted_sheet_references_in_formula(formula, deleted_name, sheet_order);
            }
        }
    }

    for rule in &mut sheet.conditional_formatting_rules {
        rewrite_cf_rule_kind_for_delete(&mut rule.kind, deleted_name, sheet_order);
    }

    for assignment in &mut sheet.data_validations {
        rewrite_data_validation_for_delete(&mut assignment.validation, deleted_name, sheet_order);
    }
}

fn rewrite_cf_rule_kind_for_delete(
    kind: &mut formula_model::CfRuleKind,
    deleted_name: &str,
    sheet_order: &[String],
) {
    match kind {
        formula_model::CfRuleKind::CellIs { formulas, .. } => {
            for formula in formulas {
                *formula =
                    rewrite_deleted_sheet_references_in_formula(formula, deleted_name, sheet_order);
            }
        }
        formula_model::CfRuleKind::Expression { formula } => {
            *formula =
                rewrite_deleted_sheet_references_in_formula(formula, deleted_name, sheet_order);
        }
        formula_model::CfRuleKind::DataBar(rule) => {
            rewrite_cfvo_for_delete(&mut rule.min, deleted_name, sheet_order);
            rewrite_cfvo_for_delete(&mut rule.max, deleted_name, sheet_order);
        }
        formula_model::CfRuleKind::ColorScale(rule) => {
            for cfvo in &mut rule.cfvos {
                rewrite_cfvo_for_delete(cfvo, deleted_name, sheet_order);
            }
        }
        formula_model::CfRuleKind::IconSet(rule) => {
            for cfvo in &mut rule.cfvos {
                rewrite_cfvo_for_delete(cfvo, deleted_name, sheet_order);
            }
        }
        formula_model::CfRuleKind::TopBottom(_)
        | formula_model::CfRuleKind::UniqueDuplicate(_)
        | formula_model::CfRuleKind::Unsupported { .. } => {}
    }
}

fn rewrite_cfvo_for_delete(cfvo: &mut formula_model::Cfvo, deleted_name: &str, sheet_order: &[String]) {
    if cfvo.type_ != formula_model::CfvoType::Formula {
        return;
    }
    let Some(value) = cfvo.value.as_mut() else {
        return;
    };
    *value = rewrite_deleted_sheet_references_in_formula(value, deleted_name, sheet_order);
}

fn rewrite_data_validation_for_delete(
    validation: &mut formula_model::DataValidation,
    deleted_name: &str,
    sheet_order: &[String],
) {
    let formula1_is_literal_list = validation.kind == formula_model::DataValidationKind::List
        && validation_formula_is_literal_list(&validation.formula1);

    if !formula1_is_literal_list && !validation.formula1.is_empty() {
        validation.formula1 = rewrite_deleted_sheet_references_in_formula(
            &validation.formula1,
            deleted_name,
            sheet_order,
        );
    }
    if let Some(formula2) = validation.formula2.as_mut() {
        *formula2 = rewrite_deleted_sheet_references_in_formula(formula2, deleted_name, sheet_order);
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

    while let Some(row) = mapping_rows.next()? {
        let Ok(style_id) = row.get::<_, i64>(1) else {
            continue;
        };
        let Ok(style) = load_model_style(conn, style_id) else {
            continue;
        };
        let model_id = style_table.intern(style);
        storage_to_model.insert(style_id, model_id);
    }

    // Include any additional styles referenced by cells that are not present in `workbook_styles`.
    //
    // This matters for backwards compatibility: workbooks imported from a `formula_model::Workbook`
    // have `workbook_styles` populated, but callers can still create new styles through the legacy
    // storage APIs (e.g. `apply_cell_changes` with a `Style` payload). Those style rows need to be
    // emitted in exports so formatting does not get dropped.
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
        let style_id = match style_id {
            Ok(id) => id,
            Err(_) => continue,
        };
        if storage_to_model.contains_key(&style_id) {
            continue;
        }
        let Ok(style) = load_model_style(conn, style_id) else {
            continue;
        };
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
        SELECT row, col, value_type, value_number, value_string, value_json, formula, style_id, phonetic
        FROM cells
        WHERE sheet_id = ?1
        ORDER BY row, col
        "#,
    )?;

    let mut rows = stmt.query(params![sheet_id])?;
    while let Some(row) = rows.next()? {
        let Ok(row_idx) = row.get::<_, i64>(0) else {
            continue;
        };
        let Ok(col_idx) = row.get::<_, i64>(1) else {
            continue;
        };

        if row_idx < 0
            || row_idx > u32::MAX as i64
            || col_idx < 0
            || col_idx >= formula_model::EXCEL_MAX_COLS as i64
        {
            continue;
        }

        let snapshot = snapshot_from_row(
            row.get::<_, Option<String>>(2).ok().flatten(),
            row.get::<_, Option<f64>>(3).ok().flatten(),
            row.get::<_, Option<String>>(4).ok().flatten(),
            row.get::<_, Option<String>>(5).ok().flatten(),
            row.get::<_, Option<String>>(6).ok().flatten(),
            row.get::<_, Option<i64>>(7).ok().flatten(),
        )?;
        let phonetic = row.get::<_, Option<String>>(8).ok().flatten();

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
            phonetic,
            style_id,
            ..Default::default()
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
