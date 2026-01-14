use crate::atomic_write::write_file_atomic;
use crate::power_query_validation::MAX_POWER_QUERY_XML_BYTES;
use crate::resource_limits::{
    MAX_PRESERVED_DRAWING_PART_BYTES, MAX_PRESERVED_DRAWING_TOTAL_BYTES,
    MAX_PRESERVED_PIVOT_PART_BYTES, MAX_PRESERVED_PIVOT_TOTAL_BYTES, MAX_VBA_PROJECT_BIN_BYTES,
    MAX_VBA_PROJECT_SIGNATURE_BIN_BYTES,
};
use crate::sheet_name::sheet_name_eq_case_insensitive;
use crate::state::{Cell, CellScalar};
use anyhow::Context;
use calamine::{open_workbook_auto, Data, Reader};
use formula_columnar::{ColumnType as ColumnarType, ColumnarTable, Value as ColumnarValue};
use formula_fs::{atomic_write_with_path, AtomicWriteError};
use formula_model::{
    import::{import_csv_to_columnar_table, CsvOptions, CsvTextEncoding},
    sanitize_sheet_name, CellValue as ModelCellValue, ColProperties as ModelColProperties,
    DateSystem as WorkbookDateSystem, SheetVisibility, TabColor, WorksheetId,
};
use formula_office_crypto::OfficeCryptoError;
use formula_xlsb::biff12_varint;
use formula_xlsb::{
    CellEdit as XlsbCellEdit, CellValue as XlsbCellValue, OpenOptions as XlsbOpenOptions,
    XlsbWorkbook,
};
use formula_xlsx::drawingml::PreservedDrawingParts;
use formula_xlsx::print::{write_workbook_print_settings, WorkbookPrintSettings};
#[cfg(test)]
use formula_xlsx::print::read_workbook_print_settings;
use formula_xlsx::{
    parse_sheet_tab_color, parse_workbook_sheets, patch_xlsx_streaming_workbook_cell_patches,
    patch_xlsx_streaming_workbook_cell_patches_with_part_overrides, strip_vba_project_streaming,
    write_sheet_tab_color, write_workbook_sheets, CellPatch as XlsxCellPatch, PartOverride,
    PreservedPivotParts, WorkbookCellPatches, WorkbookKind, XlsxPackage, XlsxPackageLimits,
};
use std::collections::{BTreeMap, HashMap, HashSet};
use std::io::{BufReader, Cursor, Read};
use std::path::Path;
#[cfg(feature = "desktop")]
use std::path::PathBuf;
use std::sync::Arc;

use crate::macro_trust::compute_macro_fingerprint;

trait ReadSeek: Read + std::io::Seek {}
impl<T: Read + std::io::Seek> ReadSeek for T {}

const FORMULA_POWER_QUERY_PART: &str = "xl/formula/power-query.xml";
const OLE_MAGIC: [u8; 8] = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];
const PASSWORD_REQUIRED_PREFIX: &str = "PASSWORD_REQUIRED:";
const INVALID_PASSWORD_PREFIX: &str = "INVALID_PASSWORD:";
const XLSX_WORKBOOK_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
const XLSM_WORKBOOK_CONTENT_TYPE: &str = "application/vnd.ms-excel.sheet.macroEnabled.main+xml";
const XLTX_WORKBOOK_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml";
const XLTM_WORKBOOK_CONTENT_TYPE: &str = "application/vnd.ms-excel.template.macroEnabled.main+xml";
const XLAM_WORKBOOK_CONTENT_TYPE: &str = "application/vnd.ms-excel.addin.macroEnabled.main+xml";

/// Encrypt a raw OOXML ZIP package into an Office `EncryptedPackage` OLE/CFB wrapper.
///
/// This helper enforces a non-empty password and uses defaults that match Excel's common "Agile"
/// encryption settings.
//
// This is only referenced by the Tauri command layer (behind the `desktop` feature) and by unit
// tests, so keep it behind the same gate to avoid dead_code warnings in headless builds.
#[cfg(any(feature = "desktop", test))]
pub(crate) fn encrypt_package_to_ole_bytes(
    zip_bytes: &[u8],
    password: &str,
) -> anyhow::Result<Vec<u8>> {
    if password.is_empty() {
        anyhow::bail!("{INVALID_PASSWORD_PREFIX} password must not be empty");
    }
    formula_office_crypto::encrypt_package_to_ole(
        zip_bytes,
        password,
        formula_office_crypto::EncryptOptions::default(),
    )
    .map_err(anyhow::Error::new)
}

// Limits for extracting individual parts from untrusted XLSX/XLSM ZIP containers. These should be
// generous enough for legitimate workbooks but finite to prevent ZIP-bomb OOM via untrusted
// uncompressed size metadata.
const XLSX_CONTENT_TYPES_MAX_BYTES: u64 = 4 * 1024 * 1024;
const XLSX_WORKBOOK_XML_MAX_BYTES: u64 = 16 * 1024 * 1024;
const XLSX_WORKSHEET_XML_MAX_BYTES: u64 = 128 * 1024 * 1024;
const XLSX_THEME_XML_MAX_BYTES: u64 = 8 * 1024 * 1024;

#[derive(Clone, Debug)]
pub struct Sheet {
    pub id: String,
    pub name: String,
    /// Excel-style sheet visibility state (visible/hidden/veryHidden).
    pub visibility: SheetVisibility,
    /// Excel-style tab color (OpenXML CT_Color).
    pub tab_color: Option<TabColor>,
    /// Sheet default column width (Excel character units).
    ///
    /// This corresponds to the worksheet's OOXML `<sheetFormatPr defaultColWidth="...">`.
    /// It is used by worksheet information functions like `CELL("width")` when a column does not
    /// have an explicit per-column width override.
    pub default_col_width: Option<f32>,
    /// Sparse per-column formatting/visibility overrides (0-based column index).
    ///
    /// This is populated when importing XLSX/XLSM/XLS files via `formula_xlsx`/`formula_xls` so
    /// workbook metadata (eg column widths/hidden flags) can be applied to the desktop formula
    /// engine (eg for `CELL("width")`).
    pub col_properties: BTreeMap<u32, formula_model::ColProperties>,
    /// Stable worksheet identifier for XLSX/XLSM inputs (`xl/worksheets/sheetN.xml`).
    ///
    /// We prefer this over `name` when writing cell patches so in-app sheet renames don't break
    /// patch application (the source `workbook.xml` might not be rewritten yet).
    pub xlsx_worksheet_part: Option<String>,
    pub(crate) origin_ordinal: Option<usize>,
    pub(crate) cells: HashMap<(usize, usize), Cell>,
    pub(crate) dirty_cells: HashSet<(usize, usize)>,
    pub(crate) columnar: Option<Arc<ColumnarTable>>,
}

#[derive(Clone, Debug)]
pub struct DefinedName {
    pub name: String,
    /// Definition formula stored **without** leading '='.
    pub refers_to: String,
    /// Sheet id when the name is sheet-scoped; `None` for workbook-scoped names.
    pub sheet_id: Option<String>,
    pub hidden: bool,
}

#[derive(Clone, Debug)]
pub struct Table {
    pub name: String,
    pub sheet_id: String,
    pub start_row: usize,
    pub start_col: usize,
    pub end_row: usize,
    pub end_col: usize,
    pub columns: Vec<String>,
}

impl Sheet {
    pub fn new(id: String, name: String) -> Self {
        Self {
            id,
            name,
            visibility: SheetVisibility::Visible,
            tab_color: None,
            default_col_width: None,
            col_properties: BTreeMap::new(),
            xlsx_worksheet_part: None,
            origin_ordinal: None,
            cells: HashMap::new(),
            dirty_cells: HashSet::new(),
            columnar: None,
        }
    }

    pub fn get_cell(&self, row: usize, col: usize) -> Cell {
        let overlay = self.cells.get(&(row, col));
        if let Some(cell) = overlay {
            // For columnar-backed sheets, allow format-only overlay cells without clobbering the
            // underlying table value (format edits shouldn't materialize the full dataset into
            // the sparse overlay).
            let is_format_only = cell.formula.is_none() && cell.input_value.is_none();
            if !(is_format_only && cell.number_format.is_some() && self.columnar.is_some()) {
                return cell.clone();
            }
        }

        if let Some(table) = &self.columnar {
            if row < table.row_count() && col < table.column_count() {
                let col_type = table
                    .schema()
                    .get(col)
                    .map(|c| c.column_type)
                    .unwrap_or(ColumnarType::String);
                let value = table.get_cell(row, col);
                let scalar = columnar_to_scalar(value, col_type);
                let mut base = match scalar {
                    CellScalar::Empty => Cell::empty(),
                    other => Cell::from_literal(Some(other)),
                };
                if let Some(cell) = overlay {
                    if cell.number_format.is_some() {
                        base.number_format = cell.number_format.clone();
                    }
                }
                return base;
            }
        }

        if let Some(cell) = overlay {
            return cell.clone();
        }

        Cell::empty()
    }

    pub fn set_cell(&mut self, row: usize, col: usize, cell: Cell) {
        self.dirty_cells.insert((row, col));
        if cell.formula.is_none() && cell.input_value.is_none() && cell.number_format.is_none() {
            self.cells.remove(&(row, col));
        } else {
            self.cells.insert((row, col), cell);
        }
    }

    pub fn cells_iter(&self) -> impl Iterator<Item = ((usize, usize), &Cell)> {
        self.cells.iter().map(|(k, v)| (*k, v))
    }

    pub fn set_columnar_table(&mut self, table: Arc<ColumnarTable>) {
        self.columnar = Some(table);
    }

    pub fn clear_dirty_cells(&mut self) {
        self.dirty_cells.clear();
    }

    pub fn get_range_cells(
        &self,
        start_row: usize,
        start_col: usize,
        end_row: usize,
        end_col: usize,
    ) -> Vec<Vec<Cell>> {
        let rows = end_row.saturating_sub(start_row) + 1;
        let cols = end_col.saturating_sub(start_col) + 1;
        let mut out = vec![vec![Cell::empty(); cols]; rows];

        if let Some(table) = &self.columnar {
            let fetched = table.get_range(
                start_row,
                end_row.saturating_add(1),
                start_col,
                end_col.saturating_add(1),
            );
            let dest_row_off = fetched.row_start.saturating_sub(start_row);
            let dest_col_off = fetched.col_start.saturating_sub(start_col);
            let fetched_col_start = fetched.col_start;

            for (local_col, values) in fetched.columns.into_iter().enumerate() {
                let table_col_idx = fetched_col_start + local_col;
                let col_type = table
                    .schema()
                    .get(table_col_idx)
                    .map(|c| c.column_type)
                    .unwrap_or(ColumnarType::String);
                let out_col = dest_col_off + local_col;
                for (local_row, v) in values.into_iter().enumerate() {
                    let out_row = dest_row_off + local_row;
                    if let Some(row_vec) = out.get_mut(out_row) {
                        if let Some(cell) = row_vec.get_mut(out_col) {
                            let scalar = columnar_to_scalar(v, col_type);
                            *cell = match scalar {
                                CellScalar::Empty => Cell::empty(),
                                other => Cell::from_literal(Some(other)),
                            };
                        }
                    }
                }
            }
        } else {
            for r in 0..rows {
                for c in 0..cols {
                    out[r][c] = self.get_cell(start_row + r, start_col + c);
                }
            }
        }

        // Overlay sparse edits/formulas.
        for ((row, col), cell) in &self.cells {
            if *row < start_row || *row > end_row || *col < start_col || *col > end_col {
                continue;
            }
            let is_format_only = cell.formula.is_none() && cell.input_value.is_none();
            if is_format_only && cell.number_format.is_some() && self.columnar.is_some() {
                // Preserve the underlying columnar value and apply only the format.
                out[row - start_row][col - start_col].number_format = cell.number_format.clone();
                continue;
            }
            out[row - start_row][col - start_col] = cell.clone();
        }

        out
    }
}

fn columnar_to_scalar(value: ColumnarValue, column_type: ColumnarType) -> CellScalar {
    match value {
        ColumnarValue::Null => CellScalar::Empty,
        ColumnarValue::Number(v) => CellScalar::Number(v),
        ColumnarValue::Boolean(v) => CellScalar::Bool(v),
        ColumnarValue::String(v) => CellScalar::Text(v.as_ref().to_string()),
        ColumnarValue::DateTime(v) => CellScalar::Number(v as f64),
        ColumnarValue::Currency(v) => match column_type {
            ColumnarType::Currency { scale } => {
                let denom = 10f64.powi(scale as i32);
                CellScalar::Number(v as f64 / denom)
            }
            _ => CellScalar::Number(v as f64),
        },
        ColumnarValue::Percentage(v) => match column_type {
            ColumnarType::Percentage { scale } => {
                let denom = 10f64.powi(scale as i32);
                CellScalar::Number(v as f64 / denom)
            }
            _ => CellScalar::Number(v as f64),
        },
    }
}

#[derive(Clone, Debug)]
pub struct Workbook {
    pub path: Option<String>,
    /// Path the workbook was opened from, even if we later save under a different
    /// name/extension (e.g. opening legacy `.xls` defaults to saving as `.xlsx`).
    pub origin_path: Option<String>,
    /// Raw bytes for the workbook we opened (XLSX/XLSM only). When present we use it as the base
    /// package and patch only the edited worksheet cell XML (+ print settings) on save.
    pub origin_xlsx_bytes: Option<Arc<[u8]>>,
    /// Formula Power Query query definitions persisted inside the workbook package (XLSX/XLSM).
    ///
    /// This is stored as a dedicated OPC part (`xl/formula/power-query.xml`) containing an XML
    /// wrapper with JSON inside a CDATA section.
    pub power_query_xml: Option<Vec<u8>>,
    /// When the workbook was opened from an XLSB file, this stores the origin path so we can
    /// re-open the source package and write back using `formula-xlsb`'s lossless OPC writer.
    pub origin_xlsb_path: Option<String>,
    pub vba_project_bin: Option<Vec<u8>>,
    /// Optional VBA project signature payload stored as a separate OPC part
    /// (`xl/vbaProjectSignature.bin`) in some macro-enabled workbooks.
    pub vba_project_signature_bin: Option<Vec<u8>>,
    /// Stable identifier used for macro trust decisions (hash of workbook identity + `vbaProject.bin`).
    pub macro_fingerprint: Option<String>,
    pub preserved_drawing_parts: Option<PreservedDrawingParts>,
    /// Preserved pivot tables/caches/slicers/timelines for regeneration-based XLSX round-trips.
    ///
    /// Pivot attachments are re-applied by sheet name, falling back to the original
    /// sheet ordinal position in the workbook when a sheet is renamed in-app.
    pub preserved_pivot_parts: Option<PreservedPivotParts>,
    pub theme_palette: Option<formula_xlsx::theme::ThemePalette>,
    /// Excel workbook date system (1900 vs 1904) used to interpret serial dates.
    pub date_system: WorkbookDateSystem,
    /// Workbook-level defined names (named ranges / constants / formulas).
    pub defined_names: Vec<DefinedName>,
    /// Excel tables (structured ranges) across all worksheets.
    pub tables: Vec<Table>,
    pub sheets: Vec<Sheet>,
    pub print_settings: WorkbookPrintSettings,
    pub(crate) original_print_settings: WorkbookPrintSettings,
    pub(crate) original_power_query_xml: Option<Vec<u8>>,
    /// Baseline input snapshot for cells that have been edited since the last save/open.
    ///
    /// Keyed by `(sheet_id, row, col)` storing `(value, formula)` from the first time the cell
    /// was touched. On save we patch only cells whose current input differs from this baseline,
    /// so edit→revert cycles don't churn the XLSX package.
    pub(crate) cell_input_baseline:
        HashMap<(String, usize, usize), (Option<CellScalar>, Option<String>)>,
}

impl Workbook {
    pub fn new_empty(path: Option<String>) -> Self {
        Self {
            origin_path: path.clone(),
            path,
            origin_xlsx_bytes: None,
            power_query_xml: None,
            origin_xlsb_path: None,
            vba_project_bin: None,
            vba_project_signature_bin: None,
            macro_fingerprint: None,
            preserved_drawing_parts: None,
            preserved_pivot_parts: None,
            theme_palette: None,
            date_system: WorkbookDateSystem::Excel1900,
            defined_names: Vec::new(),
            tables: Vec::new(),
            sheets: Vec::new(),
            print_settings: WorkbookPrintSettings::default(),
            original_print_settings: WorkbookPrintSettings::default(),
            original_power_query_xml: None,
            cell_input_baseline: HashMap::new(),
        }
    }

    pub fn add_sheet(&mut self, name: String) {
        let id = name.clone();
        self.sheets.push(Sheet::new(id, name));
    }

    pub fn ensure_sheet_ids(&mut self) {
        // Ensure sheet ids are stable and unique. For now, use the name and
        // disambiguate with a suffix when needed.
        let mut seen = std::collections::HashSet::new();
        for sheet in &mut self.sheets {
            if sheet.id.trim().is_empty() {
                sheet.id = sheet.name.clone();
            }
            let mut candidate = sheet.id.clone();
            let mut counter = 1usize;
            while !seen.insert(candidate.clone()) {
                counter += 1;
                candidate = format!("{}-{}", sheet.id, counter);
            }
            sheet.id = candidate;
        }
    }

    pub fn sheet(&self, sheet_id: &str) -> Option<&Sheet> {
        self.sheets.iter().find(|s| s.id == sheet_id)
    }

    pub fn sheet_mut(&mut self, sheet_id: &str) -> Option<&mut Sheet> {
        self.sheets.iter_mut().find(|s| s.id == sheet_id)
    }

    /// Remove a sheet from the workbook by id (case-insensitive), returning the removed sheet.
    ///
    /// This also clears any per-cell input baselines tracked for the removed sheet so future
    /// patch-based saves don't retain stale edit history.
    pub fn remove_sheet(&mut self, sheet_id: &str) -> Option<Sheet> {
        let idx = self.sheets.iter().position(|s| {
            s.id.eq_ignore_ascii_case(sheet_id)
                || crate::sheet_name::sheet_name_eq_case_insensitive(&s.name, sheet_id)
        })?;
        let removed = self.sheets.remove(idx);
        self.cell_input_baseline
            .retain(|(id, _, _), _| id != &removed.id);
        Some(removed)
    }

    pub fn cell_has_formula(&self, sheet_id: &str, row: usize, col: usize) -> bool {
        self.sheet(sheet_id)
            .and_then(|sheet| sheet.cells.get(&(row, col)))
            .and_then(|cell| cell.formula.as_ref())
            .is_some()
    }

    pub fn cell_formula(&self, sheet_id: &str, row: usize, col: usize) -> Option<String> {
        self.sheet(sheet_id)
            .and_then(|sheet| sheet.cells.get(&(row, col)))
            .and_then(|cell| cell.formula.clone())
    }

    pub fn cell_value(&self, sheet_id: &str, row: usize, col: usize) -> CellScalar {
        self.sheet(sheet_id)
            .and_then(|sheet| sheet.cells.get(&(row, col)))
            .map(|cell| cell.computed_value.clone())
            .unwrap_or(CellScalar::Empty)
    }

    pub fn set_computed_value(
        &mut self,
        sheet_id: &str,
        row: usize,
        col: usize,
        value: CellScalar,
    ) {
        if let Some(sheet) = self.sheet_mut(sheet_id) {
            let cell = sheet.cells.entry((row, col)).or_insert_with(Cell::empty);
            cell.computed_value = value;
        }
    }
}

#[cfg(feature = "desktop")]
pub async fn read_xlsx(path: impl Into<PathBuf> + Send + 'static) -> anyhow::Result<Workbook> {
    let path = path.into();
    tauri::async_runtime::spawn_blocking(move || read_xlsx_blocking(&path))
        .await
        .map_err(|e| anyhow::anyhow!(e.to_string()))?
}

#[cfg(feature = "desktop")]
pub async fn read_csv(path: impl Into<PathBuf> + Send + 'static) -> anyhow::Result<Workbook> {
    let path = path.into();
    tauri::async_runtime::spawn_blocking(move || read_csv_blocking(&path))
        .await
        .map_err(|e| anyhow::anyhow!(e.to_string()))?
}

#[cfg(all(feature = "desktop", feature = "parquet"))]
pub async fn read_parquet(path: impl Into<PathBuf> + Send + 'static) -> anyhow::Result<Workbook> {
    let path = path.into();
    tauri::async_runtime::spawn_blocking(move || read_parquet_blocking(&path))
        .await
        .map_err(|e| anyhow::anyhow!(e.to_string()))?
}

#[cfg(feature = "desktop")]
pub async fn read_workbook(
    path: impl Into<PathBuf> + Send + 'static,
    password: Option<String>,
) -> anyhow::Result<Workbook> {
    let path = path.into();
    tauri::async_runtime::spawn_blocking(move || {
        read_workbook_blocking_with_password(&path, password.as_deref())
    })
    .await
    .map_err(|e| anyhow::anyhow!(e.to_string()))?
}

fn cfb_stream_exists<R: std::io::Read + std::io::Write + std::io::Seek>(
    ole: &mut cfb::CompoundFile<R>,
    name: &str,
) -> bool {
    open_stream_best_effort(ole, name).is_some()
}

fn open_stream_best_effort<R: std::io::Read + std::io::Write + std::io::Seek>(
    ole: &mut cfb::CompoundFile<R>,
    name: &str,
) -> Option<cfb::Stream<R>> {
    if let Ok(stream) = ole.open_stream(name) {
        return Some(stream);
    }

    let trimmed = name.strip_prefix('/').unwrap_or(name);
    let with_leading_slash = format!("/{trimmed}");
    if let Ok(stream) = ole.open_stream(&with_leading_slash) {
        return Some(stream);
    }

    // The `cfb` crate's `open_stream` implementation is case-sensitive; some real-world producers
    // appear to vary the casing of the `EncryptionInfo`/`EncryptedPackage` streams. Walk the
    // directory and locate the stream path case-insensitively, then open the exact discovered
    // entry name.
    let mut found_path: Option<String> = None;
    for entry in ole.walk() {
        if !entry.is_stream() {
            continue;
        }
        let path = entry.path().to_string_lossy();
        let normalized = path.as_ref().strip_prefix('/').unwrap_or(path.as_ref());
        if normalized.eq_ignore_ascii_case(trimmed) {
            found_path = Some(path.into_owned());
            break;
        }
    }

    let found_path = found_path?;
    ole.open_stream(&found_path).ok()
}

fn is_encrypted_ooxml_workbook(path: &Path) -> std::io::Result<bool> {
    use std::io::{Read, Seek, SeekFrom};

    let mut file = std::fs::File::open(path)?;
    let mut magic = [0u8; 8];
    match file.read_exact(&mut magic) {
        Ok(()) => {}
        Err(err) if err.kind() == std::io::ErrorKind::UnexpectedEof => return Ok(false),
        Err(err) => return Err(err),
    }
    if magic != OLE_MAGIC {
        return Ok(false);
    }
    file.seek(SeekFrom::Start(0))?;

    let mut ole = match cfb::CompoundFile::open(file) {
        Ok(ole) => ole,
        Err(_) => return Ok(false),
    };

    Ok(cfb_stream_exists(&mut ole, "EncryptionInfo")
        && cfb_stream_exists(&mut ole, "EncryptedPackage"))
}

fn is_xls_ole_workbook(path: &Path) -> std::io::Result<bool> {
    use std::io::{Read, Seek, SeekFrom};

    let mut file = std::fs::File::open(path)?;
    let mut magic = [0u8; 8];
    match file.read_exact(&mut magic) {
        Ok(()) => {}
        Err(err) if err.kind() == std::io::ErrorKind::UnexpectedEof => return Ok(false),
        Err(err) => return Err(err),
    }
    if magic != OLE_MAGIC {
        return Ok(false);
    }
    file.seek(SeekFrom::Start(0))?;

    let mut ole = match cfb::CompoundFile::open(file) {
        Ok(ole) => ole,
        Err(_) => return Ok(false),
    };

    Ok(cfb_stream_exists(&mut ole, "Workbook") || cfb_stream_exists(&mut ole, "Book"))
}

#[derive(Copy, Clone, Debug, Eq, PartialEq)]
enum SniffedWorkbookFormat {
    Xls,
    Xlsx,
    Xlsb,
}

const PARQUET_MAGIC: [u8; 4] = *b"PAR1";

fn zip_entry_name_matches(candidate: &str, target: &str) -> bool {
    let target = target.trim_start_matches('/').replace('\\', "/");

    let mut normalized = candidate.trim_start_matches('/');
    let replaced;
    if normalized.contains('\\') {
        replaced = normalized.replace('\\', "/");
        normalized = &replaced;
    }

    normalized.eq_ignore_ascii_case(&target)
}

fn zip_archive_has_entry<R: Read + std::io::Seek>(
    archive: &zip::ZipArchive<R>,
    name: &str,
) -> bool {
    archive
        .file_names()
        .any(|candidate| zip_entry_name_matches(candidate, name))
}

fn sniff_workbook_format(path: &Path) -> Option<SniffedWorkbookFormat> {
    use std::io::{Read, Seek, SeekFrom};

    const ZIP_LOCAL_FILE_HEADER: [u8; 4] = [0x50, 0x4B, 0x03, 0x04];
    const ZIP_CENTRAL_DIRECTORY: [u8; 4] = [0x50, 0x4B, 0x05, 0x06];
    const ZIP_SPANNING_SIGNATURE: [u8; 4] = [0x50, 0x4B, 0x07, 0x08];

    let mut file = std::fs::File::open(path).ok()?;
    let mut header = [0u8; 8];
    let read = file.read(&mut header).ok()?;

    if read >= OLE_MAGIC.len() && header == OLE_MAGIC {
        // Many file formats share the OLE header, so avoid blindly classifying everything as an
        // `.xls`. Confirm this looks like either:
        // - a legacy BIFF workbook (Workbook/Book stream), or
        // - an encrypted OOXML container (EncryptionInfo + EncryptedPackage streams).
        if is_encrypted_ooxml_workbook(path).ok()? || is_xls_ole_workbook(path).ok()? {
            return Some(SniffedWorkbookFormat::Xls);
        }
        return None;
    }

    let is_zip = read >= 4
        && (header[..4] == ZIP_LOCAL_FILE_HEADER
            || header[..4] == ZIP_CENTRAL_DIRECTORY
            || header[..4] == ZIP_SPANNING_SIGNATURE);
    if !is_zip {
        return None;
    }

    let _ = file.seek(SeekFrom::Start(0));
    let archive = zip::ZipArchive::new(file).ok()?;
    if zip_archive_has_entry(&archive, "xl/workbook.bin") {
        return Some(SniffedWorkbookFormat::Xlsb);
    }
    if zip_archive_has_entry(&archive, "xl/workbook.xml") {
        return Some(SniffedWorkbookFormat::Xlsx);
    }

    None
}

pub(crate) fn looks_like_workbook(path: &Path) -> bool {
    #[cfg(feature = "parquet")]
    {
        use std::io::Read;

        let mut file = match std::fs::File::open(path) {
            Ok(file) => file,
            Err(_) => return false,
        };
        let mut magic = [0u8; 4];
        if file.read_exact(&mut magic).is_ok() && magic == PARQUET_MAGIC {
            return true;
        }
    }

    sniff_workbook_format(path).is_some()
}

fn read_xlsx_or_xlsm_blocking(path: &Path) -> anyhow::Result<Workbook> {
    read_xlsx_or_xlsm_blocking_with_limits(
        path,
        MAX_VBA_PROJECT_BIN_BYTES as u64,
        MAX_VBA_PROJECT_SIGNATURE_BIN_BYTES as u64,
        MAX_POWER_QUERY_XML_BYTES as u64,
    )
}

fn read_xlsx_or_xlsm_blocking_with_limits(
    path: &Path,
    max_vba_project_bin_bytes: u64,
    max_vba_project_signature_bin_bytes: u64,
    max_power_query_xml_bytes: u64,
) -> anyhow::Result<Workbook> {
    let max_origin_bytes = crate::resource_limits::max_origin_xlsx_bytes();
    let file_size = std::fs::metadata(path)
        .with_context(|| format!("stat workbook {:?}", path))?
        .len();

    let origin_xlsx_bytes = if file_size <= max_origin_bytes as u64 {
        Some(Arc::<[u8]>::from(std::fs::read(path).with_context(
            || format!("read workbook bytes {:?}", path),
        )?))
    } else {
        None
    };

    // Helper that returns a fresh `Read+Seek` handle for the workbook package.
    //
    // When `origin_xlsx_bytes` is retained this avoids touching the filesystem again; when it is
    // not retained we fall back to `File::open` so large workbooks don't require a full `read()`
    // into memory.
    let origin_xlsx_bytes_for_reader = origin_xlsx_bytes.clone();
    let open_reader = || -> anyhow::Result<Box<dyn ReadSeek>> {
        if let Some(bytes) = origin_xlsx_bytes_for_reader.as_ref() {
            Ok(Box::new(Cursor::new(bytes.clone())))
        } else {
            Ok(Box::new(
                std::fs::File::open(path).with_context(|| format!("open workbook {:?}", path))?,
            ))
        }
    };

    read_xlsx_or_xlsm_from_open_reader(
        path,
        origin_xlsx_bytes,
        open_reader,
        max_vba_project_bin_bytes,
        max_vba_project_signature_bin_bytes,
        max_power_query_xml_bytes,
    )
}

fn read_xlsx_or_xlsm_from_open_reader<F>(
    path: &Path,
    origin_xlsx_bytes: Option<Arc<[u8]>>,
    open_reader: F,
    max_vba_project_bin_bytes: u64,
    max_vba_project_signature_bin_bytes: u64,
    max_power_query_xml_bytes: u64,
) -> anyhow::Result<Workbook>
where
    F: Fn() -> anyhow::Result<Box<dyn ReadSeek>>,
{
    let workbook_model = formula_xlsx::read_workbook_from_reader(open_reader()?)
        .with_context(|| format!("parse xlsx {:?}", path))?;

    // Print settings are already parsed into the `formula-model` workbook as part of the main XLSX
    // import. Avoid re-reading and inflating worksheet XML a second time.
    let print_settings = WorkbookPrintSettings {
        sheets: workbook_model
            .sheets
            .iter()
            .map(|sheet| {
                let model = workbook_model.sheet_print_settings(sheet.id);
                formula_xlsx::print::SheetPrintSettings::from(&model)
            })
            .collect(),
    };

    let mut out = Workbook {
        path: Some(path.to_string_lossy().to_string()),
        origin_path: Some(path.to_string_lossy().to_string()),
        origin_xlsx_bytes: origin_xlsx_bytes.clone(),
        power_query_xml: None,
        origin_xlsb_path: None,
        vba_project_bin: None,
        vba_project_signature_bin: None,
        macro_fingerprint: None,
        preserved_drawing_parts: None,
        preserved_pivot_parts: None,
        theme_palette: None,
        date_system: workbook_model.date_system,
        defined_names: Vec::new(),
        tables: Vec::new(),
        sheets: Vec::new(),
        print_settings: print_settings.clone(),
        original_print_settings: print_settings,
        original_power_query_xml: None,
        cell_input_baseline: HashMap::new(),
    };
    // Preserve macros: if the source file contains `xl/vbaProject.bin`, stash it so that
    // `write_xlsx_blocking` can re-inject it when saving as `.xlsm`.
    //
    // Note: formula-xlsx only understands XLSX/XLSM ZIP containers (not legacy XLS).
    let mut worksheet_parts_by_name: HashMap<String, String> = HashMap::new();
    fn read_optional_part_limited(
        open_reader: &dyn Fn() -> anyhow::Result<Box<dyn ReadSeek>>,
        part_name: &str,
        max_bytes: u64,
    ) -> Option<Vec<u8>> {
        let reader = open_reader().ok()?;
        match formula_xlsx::read_part_from_reader_limited(reader, part_name, max_bytes) {
            Ok(bytes) => bytes,
            Err(formula_xlsx::XlsxError::PartTooLarge { part, size, max }) => {
                eprintln!(
                    "warning: dropped oversized xlsx part `{part}` ({size} bytes, max {max})"
                );
                None
            }
            Err(_) => None,
        }
    }

    out.vba_project_bin = read_optional_part_limited(
        &open_reader,
        "xl/vbaProject.bin",
        max_vba_project_bin_bytes,
    );
    out.vba_project_signature_bin = read_optional_part_limited(
        &open_reader,
        "xl/vbaProjectSignature.bin",
        max_vba_project_signature_bin_bytes,
    );
    if let Some(power_query_xml) = read_optional_part_limited(
        &open_reader,
        FORMULA_POWER_QUERY_PART,
        max_power_query_xml_bytes,
    ) {
        out.power_query_xml = Some(power_query_xml.clone());
        out.original_power_query_xml = Some(power_query_xml);
    }
    if let (Some(origin), Some(vba)) = (out.origin_path.as_deref(), out.vba_project_bin.as_deref())
    {
        out.macro_fingerprint = Some(compute_macro_fingerprint(origin, vba));
    }
    if let Ok(reader) = open_reader() {
        if let Ok(parts) =
            formula_xlsx::worksheet_parts_from_reader_limited(reader, XLSX_WORKBOOK_XML_MAX_BYTES)
        {
            for part in parts {
                worksheet_parts_by_name.insert(part.name, part.worksheet_part);
            }
        }
    }
    if let Ok(reader) = open_reader() {
        if let Ok(preserved) = formula_xlsx::drawingml::preserve_drawing_parts_from_reader_limited(
            reader,
            MAX_PRESERVED_DRAWING_PART_BYTES as u64,
            MAX_PRESERVED_DRAWING_TOTAL_BYTES as u64,
        ) {
            if !preserved.is_empty() {
                out.preserved_drawing_parts = Some(preserved);
            }
        }
    }
    if let Ok(reader) = open_reader() {
        if let Ok(preserved) = formula_xlsx::pivots::preserve_pivot_parts_from_reader_limited(
            reader,
            MAX_PRESERVED_PIVOT_PART_BYTES as u64,
            MAX_PRESERVED_PIVOT_TOTAL_BYTES as u64,
        ) {
            if !preserved.is_empty() {
                out.preserved_pivot_parts = Some(preserved);
            }
        }
    }
    if let Ok(reader) = open_reader() {
        if let Ok(palette) =
            formula_xlsx::theme_palette_from_reader_limited(reader, XLSX_THEME_XML_MAX_BYTES)
        {
            out.theme_palette = palette;
        }
    }

    out.sheets = workbook_model
        .sheets
        .iter()
        .map(|sheet| formula_model_sheet_to_app_sheet(sheet, &workbook_model.styles))
        .collect::<anyhow::Result<Vec<_>>>()?;
    for sheet in &mut out.sheets {
        sheet.xlsx_worksheet_part = worksheet_parts_by_name.get(&sheet.name).cloned();
    }

    let sheet_names_by_id: HashMap<WorksheetId, String> = workbook_model
        .sheets
        .iter()
        .map(|sheet| (sheet.id, sheet.name.clone()))
        .collect();

    out.defined_names = workbook_model
        .defined_names
        .iter()
        .map(|dn| {
            let sheet_id = match dn.scope {
                formula_model::DefinedNameScope::Workbook => None,
                formula_model::DefinedNameScope::Sheet(id) => sheet_names_by_id.get(&id).cloned(),
            };

            DefinedName {
                name: dn.name.clone(),
                refers_to: dn.refers_to.clone(),
                sheet_id,
                hidden: dn.hidden,
            }
        })
        .collect();

    out.tables = workbook_model
        .sheets
        .iter()
        .flat_map(|sheet| {
            let sheet_id = sheet.name.clone();
            sheet.tables.iter().map(move |table| Table {
                name: table.display_name.clone(),
                sheet_id: sheet_id.clone(),
                start_row: table.range.start.row as usize,
                start_col: table.range.start.col as usize,
                end_row: table.range.end.row as usize,
                end_col: table.range.end.col as usize,
                columns: table.columns.iter().map(|c| c.name.clone()).collect(),
            })
        })
        .collect();

    out.ensure_sheet_ids();
    for sheet in &mut out.sheets {
        sheet.clear_dirty_cells();
    }

    Ok(out)
}

fn read_xls_blocking_with_password(path: &Path, password: Option<&str>) -> anyhow::Result<Workbook> {
    let imported = match formula_xls::import_xls_path_with_password(path, password) {
        Ok(imported) => imported,
        // Keep the desktop-facing error message user-actionable (the UI can prompt for a
        // password) without exposing internal Rust API names.
        Err(formula_xls::ImportError::EncryptedWorkbook) => anyhow::bail!(
            "{PASSWORD_REQUIRED_PREFIX} workbook `{}` is password-protected; supply a password to open it",
            path.display()
        ),
        Err(formula_xls::ImportError::InvalidPassword) => {
            anyhow::bail!("{INVALID_PASSWORD_PREFIX} invalid password")
        }
        Err(other) => {
            return Err(anyhow::anyhow!(other)).with_context(|| format!("import xls {:?}", path));
        }
    };
    let workbook_model = imported.workbook;

    let mut out = Workbook {
        path: Some(path.to_string_lossy().to_string()),
        origin_path: Some(path.to_string_lossy().to_string()),
        origin_xlsx_bytes: None,
        power_query_xml: None,
        origin_xlsb_path: None,
        vba_project_bin: None,
        vba_project_signature_bin: None,
        macro_fingerprint: None,
        preserved_drawing_parts: None,
        preserved_pivot_parts: None,
        theme_palette: None,
        date_system: workbook_model.date_system,
        defined_names: Vec::new(),
        tables: Vec::new(),
        sheets: Vec::new(),
        print_settings: WorkbookPrintSettings::default(),
        original_print_settings: WorkbookPrintSettings::default(),
        original_power_query_xml: None,
        cell_input_baseline: HashMap::new(),
    };

    out.sheets = workbook_model
        .sheets
        .iter()
        .map(|sheet| formula_model_sheet_to_app_sheet(sheet, &workbook_model.styles))
        .collect::<anyhow::Result<Vec<_>>>()?;

    let sheet_names_by_id: HashMap<WorksheetId, String> = workbook_model
        .sheets
        .iter()
        .map(|sheet| (sheet.id, sheet.name.clone()))
        .collect();

    out.defined_names = workbook_model
        .defined_names
        .iter()
        .map(|dn| {
            let sheet_id = match dn.scope {
                formula_model::DefinedNameScope::Workbook => None,
                formula_model::DefinedNameScope::Sheet(id) => sheet_names_by_id.get(&id).cloned(),
            };

            DefinedName {
                name: dn.name.clone(),
                refers_to: dn.refers_to.clone(),
                sheet_id,
                hidden: dn.hidden,
            }
        })
        .collect();

    out.tables = workbook_model
        .sheets
        .iter()
        .flat_map(|sheet| {
            let sheet_id = sheet.name.clone();
            sheet.tables.iter().map(move |table| Table {
                name: table.display_name.clone(),
                sheet_id: sheet_id.clone(),
                start_row: table.range.start.row as usize,
                start_col: table.range.start.col as usize,
                end_row: table.range.end.row as usize,
                end_col: table.range.end.col as usize,
                columns: table.columns.iter().map(|c| c.name.clone()).collect(),
            })
        })
        .collect();

    out.ensure_sheet_ids();
    for sheet in &mut out.sheets {
        sheet.clear_dirty_cells();
    }

    Ok(out)
}

fn read_encrypted_ooxml_workbook_blocking(
    path: &Path,
    password: Option<&str>,
) -> anyhow::Result<Workbook> {
    let Some(password) = password else {
        anyhow::bail!(
            "{PASSWORD_REQUIRED_PREFIX} workbook `{}` is password-protected; supply a password to open it",
            path.display()
        );
    };

    let raw_ole = std::fs::read(path).with_context(|| format!("read workbook {:?}", path))?;

    let decrypted_zip = match formula_office_crypto::decrypt_encrypted_package(&raw_ole, password) {
        Ok(bytes) => bytes,
        Err(err) => match err {
            OfficeCryptoError::InvalidPassword | OfficeCryptoError::IntegrityCheckFailed => {
                anyhow::bail!("{INVALID_PASSWORD_PREFIX} {err}")
            }
            other => anyhow::bail!("encrypted workbook not supported: {other}"),
        },
    };

    let decrypted_zip = Arc::<[u8]>::from(decrypted_zip);

    // The decrypted payload is a normal ZIP/OPC package. Sniff whether it is XLSX/XLSM
    // (`workbook.xml`) or XLSB (`workbook.bin`) and route to the appropriate reader.
    let workbook_format = {
        let cursor = Cursor::new(decrypted_zip.as_ref());
        let archive = zip::ZipArchive::new(cursor).ok();
        archive.and_then(|archive| {
            if zip_archive_has_entry(&archive, "xl/workbook.bin") {
                Some(SniffedWorkbookFormat::Xlsb)
            } else if zip_archive_has_entry(&archive, "xl/workbook.xml") {
                Some(SniffedWorkbookFormat::Xlsx)
            } else {
                None
            }
        })
    };

    match workbook_format {
        Some(SniffedWorkbookFormat::Xlsb) => {
            // Open XLSB from in-memory ZIP bytes. Avoid preserving unknown parts to keep memory
            // usage bounded.
            let decrypted_zip_for_cols = decrypted_zip.clone();
            let wb = XlsbWorkbook::from_bytes(
                decrypted_zip,
                XlsbOpenOptions {
                    preserve_unknown_parts: false,
                    preserve_parsed_parts: false,
                    preserve_worksheets: false,
                    decode_formulas: true,
                },
            )
            .with_context(|| format!("open decrypted xlsb workbook {:?}", path))?;

            let mut workbook = read_xlsb_from_open_workbook(
                path,
                wb,
                move |part_path| {
                    read_xlsb_col_properties_for_sheet_from_bytes(
                        decrypted_zip_for_cols.as_ref(),
                        part_path,
                    )
                },
                 // This workbook is encrypted on disk (OLE container), so we don't have an on-disk XLSB
                 // ZIP package we can re-open for lossless `.xlsb` saves.
                 None,
            )?;
            // We cannot losslessly save back into the encrypted OLE container, and `.xlsb` writes
            // require `origin_xlsb_path` (an on-disk XLSB zip). Force Save As.
            workbook.path = None;
            Ok(workbook)
        }
        Some(SniffedWorkbookFormat::Xlsx) => {
            // Mirror the non-encrypted XLSX/XLSM open logic: only retain the decrypted ZIP bytes baseline when
            // it is within the configured limit. This keeps memory usage bounded even if a user opens a
            // large password-protected workbook.
            let max_origin_bytes = crate::resource_limits::max_origin_xlsx_bytes();
            let origin_xlsx_bytes = if decrypted_zip.len() <= max_origin_bytes {
                Some(decrypted_zip.clone())
            } else {
                None
            };

            let decrypted_zip_for_reader = decrypted_zip;
            let open_reader = move || -> anyhow::Result<Box<dyn ReadSeek>> {
                Ok(Box::new(Cursor::new(decrypted_zip_for_reader.clone())))
            };

            read_xlsx_or_xlsm_from_open_reader(
                path,
                origin_xlsx_bytes,
                open_reader,
                MAX_VBA_PROJECT_BIN_BYTES as u64,
                MAX_VBA_PROJECT_SIGNATURE_BIN_BYTES as u64,
                MAX_POWER_QUERY_XML_BYTES as u64,
            )
        }
        _ => anyhow::bail!(
            "encrypted workbook not supported: decrypted package is not a recognized XLSX/XLSM/XLSB ZIP container"
        ),
    }
}

fn validate_workbook_open_size(path: &Path) -> anyhow::Result<u64> {
    let size = std::fs::metadata(path)
        .with_context(|| format!("stat workbook {:?}", path))?
        .len();
    let max = crate::resource_limits::max_workbook_open_bytes();
    if size > max {
        anyhow::bail!(
            "File size {size} exceeds maximum allowed workbook open size ({max} bytes): `{}`",
            path.display()
        );
    }
    Ok(size)
}

pub fn read_workbook_blocking(path: &Path) -> anyhow::Result<Workbook> {
    read_workbook_blocking_with_password(path, None)
}

pub fn read_workbook_blocking_with_password(
    path: &Path,
    password: Option<&str>,
) -> anyhow::Result<Workbook> {
    use std::io::Read;

    const TEXT_SNIFF_BYTES: usize = 4096;

    fn looks_like_text(buf: &[u8]) -> bool {
        if buf.is_empty() {
            return false;
        }
        // Text/CSV inputs should never contain NUL bytes; treat those as a strong signal that the
        // file is binary data and we should not route it to the CSV importer.
        if buf.iter().any(|&b| b == 0) {
            return false;
        }

        // Reject buffers with a meaningful amount of control bytes other than whitespace.
        let mut suspicious = 0usize;
        for &b in buf {
            match b {
                b'\t' | b'\n' | b'\r' => {}
                0x20..=0x7E => {}
                _ if b >= 0x80 => {}
                _ => suspicious += 1,
            }
        }
        suspicious <= buf.len() / 50
    }

    validate_workbook_open_size(path)?;

    let mut file =
        std::fs::File::open(path).with_context(|| format!("open workbook {:?}", path))?;
    let mut prefix = [0u8; 16];
    let read = file
        .read(&mut prefix)
        .with_context(|| format!("read workbook header {:?}", path))?;
    let prefix = &prefix[..read];

    if prefix.starts_with(&PARQUET_MAGIC) {
        #[cfg(feature = "parquet")]
        {
            return read_parquet_blocking(path);
        }
        #[cfg(not(feature = "parquet"))]
        {
            anyhow::bail!("parquet support is not enabled in this build");
        }
    }

    // Encrypted OOXML workbooks live in an OLE container. Decrypt them when a password is provided,
    // otherwise surface a password-required error that the frontend can prompt for.
    if let Ok(true) = is_encrypted_ooxml_workbook(path) {
        return read_encrypted_ooxml_workbook_blocking(path, password);
    }

    if let Some(format) = sniff_workbook_format(path) {
        return match format {
            SniffedWorkbookFormat::Xls => read_xls_blocking_with_password(path, password),
            SniffedWorkbookFormat::Xlsx => read_xlsx_or_xlsm_blocking(path),
            SniffedWorkbookFormat::Xlsb => read_xlsb_blocking(path),
        };
    }

    let is_zip = prefix.len() >= 4
        && prefix[0] == b'P'
        && prefix[1] == b'K'
        && matches!(prefix[2], 0x03 | 0x05 | 0x07)
        && matches!(prefix[3], 0x04 | 0x06 | 0x08);
    if is_zip {
        // ZIP containers should route to the XLSX/XLSM/XLSB sniffing logic inside `read_xlsx_blocking`.
        return read_xlsx_blocking(path);
    }

    // Best-effort: sniff for text/CSV and only route to the CSV importer when it doesn't look
    // like a binary file.
    let mut file =
        std::fs::File::open(path).with_context(|| format!("open workbook {:?}", path))?;
    let mut buf = vec![0u8; TEXT_SNIFF_BYTES];
    let read = file
        .read(&mut buf)
        .with_context(|| format!("read workbook header {:?}", path))?;
    buf.truncate(read);
    if looks_like_text(&buf) {
        return read_csv_blocking(path);
    }

    // Fall back to Calamine's extension-based readers for other spreadsheet formats (e.g. `.ods`).
    // If it isn't a supported spreadsheet format, this should surface a clear error rather than
    // trying to interpret arbitrary binary data as CSV.
    read_xlsx_blocking(path)
}

pub fn read_xlsx_blocking(path: &Path) -> anyhow::Result<Workbook> {
    read_xlsx_blocking_with_password(path, None)
}

pub fn read_xlsx_blocking_with_password(
    path: &Path,
    password: Option<&str>,
) -> anyhow::Result<Workbook> {
    validate_workbook_open_size(path)?;
    if let Ok(true) = is_encrypted_ooxml_workbook(path) {
        return read_encrypted_ooxml_workbook_blocking(path, password);
    }

    let extension = path
        .extension()
        .and_then(|s| s.to_str())
        .map(|s| s.to_ascii_lowercase());

    // Sniff the workbook content so we can open valid workbooks even when the extension is
    // missing or incorrect (e.g. a `.xlsx` saved as `.bin` or a legacy `.xls` with a `.xlsx`
    // suffix).
    if !matches!(extension.as_deref(), Some("csv")) {
        if let Some(format) = sniff_workbook_format(path) {
            match format {
                SniffedWorkbookFormat::Xls => return read_xls_blocking_with_password(path, password),
                SniffedWorkbookFormat::Xlsx => return read_xlsx_or_xlsm_blocking(path),
                SniffedWorkbookFormat::Xlsb => return read_xlsb_blocking(path),
            }
        }
    }

    if matches!(extension.as_deref(), Some("xlsb")) {
        return read_xlsb_blocking(path);
    }

    if matches!(
        extension.as_deref(),
        Some("xlsx") | Some("xlsm") | Some("xltx") | Some("xltm") | Some("xlam")
    ) {
        return read_xlsx_or_xlsm_blocking(path);
    }

    if matches!(
        extension.as_deref(),
        Some("xls") | Some("xlt") | Some("xla")
    ) {
        return read_xls_blocking_with_password(path, password);
    }

    let mut workbook =
        open_workbook_auto(path).with_context(|| format!("open workbook {:?}", path))?;
    let sheet_names = workbook.sheet_names().to_owned();

    let mut out = Workbook {
        path: Some(path.to_string_lossy().to_string()),
        origin_path: Some(path.to_string_lossy().to_string()),
        origin_xlsx_bytes: None,
        power_query_xml: None,
        origin_xlsb_path: None,
        vba_project_bin: None,
        vba_project_signature_bin: None,
        macro_fingerprint: None,
        preserved_drawing_parts: None,
        preserved_pivot_parts: None,
        theme_palette: None,
        // Calamine doesn't surface workbook date system metadata; default to Excel 1900.
        date_system: WorkbookDateSystem::Excel1900,
        defined_names: Vec::new(),
        tables: Vec::new(),
        sheets: Vec::new(),
        print_settings: WorkbookPrintSettings::default(),
        original_print_settings: WorkbookPrintSettings::default(),
        original_power_query_xml: None,
        cell_input_baseline: HashMap::new(),
    };

    for sheet_name in sheet_names {
        let range = workbook
            .worksheet_range(&sheet_name)
            .with_context(|| format!("read worksheet range {sheet_name}"))?;

        let mut sheet = Sheet::new(sheet_name.clone(), sheet_name.clone());

        let (row_offset, col_offset) = range.start().unwrap_or((0, 0));
        let row_offset = row_offset as usize;
        let col_offset = col_offset as usize;
        for (row, col, cell) in range.cells() {
            let value = match cell {
                Data::Empty => CellScalar::Empty,
                Data::String(s) => CellScalar::Text(s.clone()),
                Data::Float(f) => CellScalar::Number(*f),
                Data::Int(i) => CellScalar::Number(*i as f64),
                Data::Bool(b) => CellScalar::Bool(*b),
                Data::Error(e) => CellScalar::Error(format!("{e:?}")),
                other => CellScalar::Text(other.to_string()),
            };

            if matches!(value, CellScalar::Empty) {
                continue;
            }

            sheet.set_cell(
                row_offset + row,
                col_offset + col,
                Cell::from_literal(Some(value)),
            );
        }

        // Formula ranges may be absent for some formats; treat as optional.
        if let Ok(formulas) = workbook.worksheet_formula(&sheet_name) {
            let (row_offset, col_offset) = formulas.start().unwrap_or((0, 0));
            let row_offset = row_offset as usize;
            let col_offset = col_offset as usize;
            for (row, col, formula) in formulas.cells() {
                let normalized = formula_model::display_formula_text(formula);
                if normalized.is_empty() {
                    continue;
                }

                let row = row_offset + row;
                let col = col_offset + col;

                match sheet.cells.get_mut(&(row, col)) {
                    Some(existing) => {
                        if existing.formula.is_some() {
                            continue;
                        }
                        let cached = existing.computed_value.clone();
                        let mut cell = Cell::from_formula(normalized);
                        cell.computed_value = cached;
                        *existing = cell;
                    }
                    None => {
                        sheet.set_cell(row, col, Cell::from_formula(normalized));
                    }
                }
            }
        }

        out.sheets.push(sheet);
    }

    out.ensure_sheet_ids();
    for sheet in &mut out.sheets {
        sheet.clear_dirty_cells();
    }
    Ok(out)
}

fn formula_model_sheet_to_app_sheet(
    sheet: &formula_model::Worksheet,
    styles: &formula_model::StyleTable,
) -> anyhow::Result<Sheet> {
    let mut out = Sheet::new(sheet.name.clone(), sheet.name.clone());
    out.visibility = sheet.visibility;
    out.tab_color = sheet.tab_color.clone();
    out.default_col_width = sheet.default_col_width;
    out.col_properties = sheet.col_properties.clone();

    for (cell_ref, cell) in sheet.iter_cells() {
        let row = cell_ref.row as usize;
        let col = cell_ref.col as usize;

        let cached_value = formula_model_value_to_scalar(&cell.value);
        let number_format = (cell.style_id != 0)
            .then(|| {
                styles
                    .get(cell.style_id)
                    .and_then(|style| style.number_format.clone())
            })
            .flatten();

        if let Some(formula) = cell.formula.as_deref() {
            let normalized = formula_model::display_formula_text(formula);
            if !normalized.is_empty() {
                let mut c = Cell::from_formula(normalized);
                c.computed_value = cached_value;
                c.number_format = number_format;
                out.set_cell(row, col, c);
                continue;
            }
            // Treat empty formulas as blank/no-formula cells, matching our XLSB import behavior.
        }

        if matches!(cached_value, CellScalar::Empty) {
            if let Some(number_format) = number_format {
                let mut c = Cell::empty();
                c.number_format = Some(number_format);
                out.set_cell(row, col, c);
            }
            continue;
        }

        let mut c = Cell::from_literal(Some(cached_value));
        c.number_format = number_format;
        out.set_cell(row, col, c);
    }

    Ok(out)
}

fn formula_model_value_to_scalar(value: &ModelCellValue) -> CellScalar {
    match value {
        ModelCellValue::Empty => CellScalar::Empty,
        ModelCellValue::Number(n) => CellScalar::Number(*n),
        ModelCellValue::String(s) => CellScalar::Text(s.clone()),
        ModelCellValue::Boolean(b) => CellScalar::Bool(*b),
        ModelCellValue::Error(e) => CellScalar::Error(e.to_string()),
        ModelCellValue::RichText(rt) => CellScalar::Text(rt.text.clone()),
        ModelCellValue::Entity(entity) => CellScalar::Text(entity.display_value.clone()),
        ModelCellValue::Record(record) => CellScalar::Text(record.to_string()),
        ModelCellValue::Image(image) => CellScalar::Text(
            image
                .alt_text
                .clone()
                .filter(|s| !s.is_empty())
                .unwrap_or_else(|| "[Image]".to_string()),
        ),
        other => match other {
            ModelCellValue::Array(arr) => CellScalar::Text(format!("{:?}", arr.data)),
            ModelCellValue::Spill(_) => CellScalar::Error("#SPILL!".to_string()),
            _ => rich_model_cell_value_to_scalar(other)
                .unwrap_or_else(|| CellScalar::Text(format!("{other:?}"))),
        },
    }
}

fn rich_model_cell_value_to_scalar(value: &ModelCellValue) -> Option<CellScalar> {
    fn json_get_str<'a>(value: &'a serde_json::Value, keys: &[&str]) -> Option<&'a str> {
        for key in keys {
            if let Some(s) = value.get(key).and_then(|v| v.as_str()) {
                return Some(s);
            }
        }
        None
    }

    fn cell_value_json_to_display_string(value: &serde_json::Value) -> Option<String> {
        let value_type = value.get("type")?.as_str()?;
        match value_type {
            "number" => Some(value.get("value")?.as_f64()?.to_string()),
            "string" => Some(value.get("value")?.as_str()?.to_string()),
            "boolean" => Some(if value.get("value")?.as_bool()? {
                "TRUE".to_string()
            } else {
                "FALSE".to_string()
            }),
            "error" => Some(value.get("value")?.as_str()?.to_string()),
            "rich_text" => Some(value.get("value")?.get("text")?.as_str()?.to_string()),
            _ => None,
        }
    }

    let serialized = serde_json::to_value(value).ok()?;
    let value_type = serialized.get("type")?.as_str()?;

    match value_type {
        "entity" => {
            let entity = serialized.get("value")?;
            let display_value =
                json_get_str(entity, &["displayValue", "display_value", "display"])?.to_string();
            Some(CellScalar::Text(display_value))
        }
        "record" => {
            let record = serialized.get("value")?;
            if let Some(display_field) = json_get_str(record, &["displayField", "display_field"]) {
                if let Some(fields) = record.get("fields").and_then(|v| v.as_object()) {
                    if let Some(display_value) = fields.get(display_field) {
                        if let Some(display) = cell_value_json_to_display_string(display_value) {
                            return Some(CellScalar::Text(display));
                        }
                    }
                }
            }

            let display_value =
                json_get_str(record, &["displayValue", "display_value", "display"])?.to_string();
            Some(CellScalar::Text(display_value))
        }
        "image" => {
            let image = serialized.get("value")?;
            let alt_text = image
                .get("altText")
                .or_else(|| image.get("alt_text"))
                .and_then(|v| v.as_str())
                .unwrap_or("");
            if alt_text.is_empty() {
                Some(CellScalar::Text("[Image]".to_string()))
            } else {
                Some(CellScalar::Text(alt_text.to_string()))
            }
        }
        _ => None,
    }
}

pub fn read_csv_blocking(path: &Path) -> anyhow::Result<Workbook> {
    validate_workbook_open_size(path)?;

    let file = std::fs::File::open(path).with_context(|| format!("open csv {:?}", path))?;
    let reader = BufReader::new(file);
    // Default to Excel-like behavior: attempt UTF-8 first, then fall back to Windows-1252.
    let table = import_csv_to_columnar_table(
        reader,
        CsvOptions {
            encoding: CsvTextEncoding::Auto,
            ..CsvOptions::default()
        },
    )
    .map_err(|e| anyhow::anyhow!(e.to_string()))
    .with_context(|| format!("import csv {:?}", path))?;

    let sheet_name = sanitize_sheet_name(path.file_stem().and_then(|s| s.to_str()).unwrap_or(""));
    let mut sheet = Sheet::new(sheet_name.clone(), sheet_name);
    sheet.set_columnar_table(Arc::new(table));

    let mut out = Workbook {
        path: Some(path.to_string_lossy().to_string()),
        origin_path: Some(path.to_string_lossy().to_string()),
        origin_xlsx_bytes: None,
        power_query_xml: None,
        origin_xlsb_path: None,
        vba_project_bin: None,
        vba_project_signature_bin: None,
        macro_fingerprint: None,
        preserved_drawing_parts: None,
        preserved_pivot_parts: None,
        theme_palette: None,
        date_system: WorkbookDateSystem::Excel1900,
        defined_names: Vec::new(),
        tables: Vec::new(),
        sheets: vec![sheet],
        print_settings: WorkbookPrintSettings::default(),
        original_print_settings: WorkbookPrintSettings::default(),
        original_power_query_xml: None,
        cell_input_baseline: HashMap::new(),
    };
    out.ensure_sheet_ids();
    for sheet in &mut out.sheets {
        sheet.clear_dirty_cells();
    }
    Ok(out)
}

#[cfg(feature = "parquet")]
pub fn read_parquet_blocking(path: &Path) -> anyhow::Result<Workbook> {
    validate_workbook_open_size(path)?;

    let table = formula_columnar::parquet::read_parquet_to_columnar(path)
        .map_err(|e| anyhow::anyhow!(e.to_string()))
        .with_context(|| format!("import parquet {:?}", path))?;

    let sheet_name = sanitize_sheet_name(path.file_stem().and_then(|s| s.to_str()).unwrap_or(""));
    let mut sheet = Sheet::new(sheet_name.clone(), sheet_name);
    sheet.set_columnar_table(Arc::new(table));

    let mut out = Workbook {
        path: Some(path.to_string_lossy().to_string()),
        origin_path: Some(path.to_string_lossy().to_string()),
        origin_xlsx_bytes: None,
        power_query_xml: None,
        origin_xlsb_path: None,
        vba_project_bin: None,
        vba_project_signature_bin: None,
        macro_fingerprint: None,
        preserved_drawing_parts: None,
        preserved_pivot_parts: None,
        theme_palette: None,
        date_system: WorkbookDateSystem::Excel1900,
        defined_names: Vec::new(),
        tables: Vec::new(),
        sheets: vec![sheet],
        print_settings: WorkbookPrintSettings::default(),
        original_print_settings: WorkbookPrintSettings::default(),
        original_power_query_xml: None,
        cell_input_baseline: HashMap::new(),
    };
    out.ensure_sheet_ids();
    for sheet in &mut out.sheets {
        sheet.clear_dirty_cells();
    }
    Ok(out)
}

fn discard_xlsb_record_payload(reader: &mut impl Read, mut len: u32) -> anyhow::Result<()> {
    // Avoid allocations when skipping BIFF12 payloads; XLSB worksheet records are attacker-controlled.
    let mut buf = [0u8; 8192];
    while len > 0 {
        let n = (len as usize).min(buf.len());
        reader.read_exact(&mut buf[..n])?;
        len -= n as u32;
    }
    Ok(())
}

fn parse_xlsb_col_info_payload(payload: &[u8]) -> Option<(u32, u32, formula_model::ColProperties)> {
    // Column indices in XLSB are 0-based, matching cell records and our in-memory model.
    //
    // BrtColInfo layouts observed in the wild:
    // - Common (MS-XLSB): [colFirst: u32][colLast: u32][width: u32 (1/256 chars)][xf: u32][grbit: u16][reserved: u16]
    // - Some writers may use u16 column indices + u16 width (BIFF8-like).

    if payload.len() < 4 {
        return None;
    }

    let (col_first, col_last, base_offset) = if payload.len() >= 8 {
        let c1 = u32::from_le_bytes(payload[0..4].try_into().ok()?);
        let c2 = u32::from_le_bytes(payload[4..8].try_into().ok()?);
        if c1 < formula_model::EXCEL_MAX_COLS && c2 < formula_model::EXCEL_MAX_COLS && c1 <= c2 {
            (c1, c2, 8)
        } else {
            let c1 = u16::from_le_bytes(payload[0..2].try_into().ok()?) as u32;
            let c2 = u16::from_le_bytes(payload[2..4].try_into().ok()?) as u32;
            if c1 < formula_model::EXCEL_MAX_COLS
                && c2 < formula_model::EXCEL_MAX_COLS
                && c1 <= c2
            {
                (c1, c2, 4)
            } else {
                return None;
            }
        }
    } else {
        let c1 = u16::from_le_bytes(payload[0..2].try_into().ok()?) as u32;
        let c2 = u16::from_le_bytes(payload[2..4].try_into().ok()?) as u32;
        if c1 < formula_model::EXCEL_MAX_COLS && c2 < formula_model::EXCEL_MAX_COLS && c1 <= c2 {
            (c1, c2, 4)
        } else {
            return None;
        }
    };

    let hidden = if payload.len() >= 4 {
        // Many BIFF layouts encode the column visibility in a `grbit` field near the end.
        let off = payload.len().saturating_sub(4);
        let options = u16::from_le_bytes(payload[off..off + 2].try_into().ok()?);
        (options & 0x0001) != 0
    } else {
        false
    };

    let mut width: Option<f32> = None;

    if payload.len() >= base_offset + 4 {
        let raw = u32::from_le_bytes(payload[base_offset..base_offset + 4].try_into().ok()?);
        let candidate = (raw as f32) / 256.0;
        if candidate.is_finite() && candidate > 0.0 && candidate <= 255.0 {
            width = Some(candidate);
        }
    }

    if width.is_none() && payload.len() >= base_offset + 2 {
        let raw = u16::from_le_bytes(payload[base_offset..base_offset + 2].try_into().ok()?);
        let candidate = (raw as f32) / 256.0;
        if candidate.is_finite() && candidate > 0.0 && candidate <= 255.0 {
            width = Some(candidate);
        }
    }

    if width.is_none() && payload.len() >= base_offset + 8 {
        // Best-effort fallback for alternative encodings; treat the next 8 bytes as f64 chars.
        let raw = f64::from_le_bytes(payload[base_offset..base_offset + 8].try_into().ok()?);
        if raw.is_finite() && raw > 0.0 && raw <= 255.0 {
            width = Some(raw as f32);
        }
    }

    if width.is_none() && !hidden {
        return None;
    }

    let mut props = formula_model::ColProperties::default();
    props.width = width;
    props.hidden = hidden;

    Some((col_first, col_last, props))
}

fn read_xlsb_col_properties_for_sheet_from_zip<R: Read + std::io::Seek>(
    zip: &mut zip::ZipArchive<R>,
    part_path: &str,
) -> anyhow::Result<BTreeMap<u32, formula_model::ColProperties>> {
    // Stream the worksheet's BIFF12 record stream only until we hit the sheetData block.
    // Column info records appear before sheetData and can be extracted without scanning the
    // potentially enormous cell grid.
    //
    // We cap the amount scanned to avoid pathological packages that omit sheetData and force a
    // full scan of a large worksheet stream.
    const SHEETDATA_RECORD_ID: u32 = 0x0091;
    const BEGIN_COL_INFOS_RECORD_ID: u32 = 0x0186;
    const END_COL_INFOS_RECORD_ID: u32 = 0x0187;

    const MAX_SCAN_PAYLOAD_BYTES: u64 = 8 * 1024 * 1024; // 8 MiB
    const MAX_COL_INFO_RECORD_BYTES: u32 = 1024; // Defensive; BrtColInfo records are tiny.

    let Some(entry_name) = zip
        .file_names()
        .find(|candidate| zip_entry_name_matches(candidate, part_path))
        .map(str::to_owned)
    else {
        return Ok(BTreeMap::new());
    };

    let mut sheet = zip
        .by_name(&entry_name)
        .with_context(|| format!("open xlsb worksheet part {entry_name}"))?;

    let mut out: BTreeMap<u32, formula_model::ColProperties> = BTreeMap::new();
    let mut in_col_infos = false;
    let mut scanned_payload_bytes: u64 = 0;

    while let Some(id) = biff12_varint::read_record_id(&mut sheet)? {
        let Some(len) = biff12_varint::read_record_len(&mut sheet)? else {
            break;
        };

        let next_scanned = scanned_payload_bytes.saturating_add(len as u64);
        if next_scanned > MAX_SCAN_PAYLOAD_BYTES {
            break;
        }
        scanned_payload_bytes = next_scanned;

        match id {
            SHEETDATA_RECORD_ID => {
                discard_xlsb_record_payload(&mut sheet, len)?;
                break;
            }
            BEGIN_COL_INFOS_RECORD_ID => {
                in_col_infos = true;
                discard_xlsb_record_payload(&mut sheet, len)?;
            }
            END_COL_INFOS_RECORD_ID => {
                in_col_infos = false;
                discard_xlsb_record_payload(&mut sheet, len)?;
            }
            _ if in_col_infos => {
                if len > MAX_COL_INFO_RECORD_BYTES {
                    discard_xlsb_record_payload(&mut sheet, len)?;
                    continue;
                }

                let mut payload = vec![0u8; len as usize];
                sheet.read_exact(&mut payload)?;

                // BrtColInfo is typically record id 0x003C, but be tolerant and attempt to parse
                // any record inside the colInfos block when it matches the expected shape.
                if let Some((start_col, end_col, props)) = parse_xlsb_col_info_payload(&payload) {
                    for col in start_col..=end_col {
                        if col >= formula_model::EXCEL_MAX_COLS {
                            break;
                        }
                        let entry = out
                            .entry(col)
                            .or_insert_with(formula_model::ColProperties::default);
                        if props.width.is_some() {
                            entry.width = props.width;
                        }
                        entry.hidden = props.hidden;
                    }
                }
            }
            _ => {
                discard_xlsb_record_payload(&mut sheet, len)?;
            }
        }
    }

    // Keep the map sparse.
    out.retain(|_, p| p.width.is_some() || p.hidden || p.style_id.is_some());

    Ok(out)
}

fn read_xlsb_col_properties_for_sheet(
    path: &Path,
    part_path: &str,
) -> anyhow::Result<BTreeMap<u32, formula_model::ColProperties>> {
    let file = std::fs::File::open(path)?;
    let mut zip = zip::ZipArchive::new(file).with_context(|| format!("open xlsb zip {:?}", path))?;
    read_xlsb_col_properties_for_sheet_from_zip(&mut zip, part_path)
}

fn read_xlsb_col_properties_for_sheet_from_bytes(
    bytes: &[u8],
    part_path: &str,
) -> anyhow::Result<BTreeMap<u32, formula_model::ColProperties>> {
    let cursor = Cursor::new(bytes);
    let mut zip = zip::ZipArchive::new(cursor).context("open xlsb zip from bytes")?;
    read_xlsb_col_properties_for_sheet_from_zip(&mut zip, part_path)
}

fn read_xlsb_from_open_workbook<F>(
    path: &Path,
    wb: XlsbWorkbook,
    mut read_col_properties: F,
    origin_xlsb_path: Option<String>,
) -> anyhow::Result<Workbook>
where
    F: FnMut(&str) -> anyhow::Result<BTreeMap<u32, formula_model::ColProperties>>,
{
    let date_system = if wb.workbook_properties().date_system_1904 {
        WorkbookDateSystem::Excel1904
    } else {
        WorkbookDateSystem::Excel1900
    };

    let mut out = Workbook {
        path: Some(path.to_string_lossy().to_string()),
        origin_path: Some(path.to_string_lossy().to_string()),
        origin_xlsx_bytes: None,
        power_query_xml: None,
        origin_xlsb_path,
        vba_project_bin: None,
        vba_project_signature_bin: None,
        macro_fingerprint: None,
        preserved_drawing_parts: None,
        preserved_pivot_parts: None,
        theme_palette: None,
        date_system,
        defined_names: Vec::new(),
        tables: Vec::new(),
        sheets: Vec::new(),
        print_settings: WorkbookPrintSettings::default(),
        original_print_settings: WorkbookPrintSettings::default(),
        original_power_query_xml: None,
        cell_input_baseline: HashMap::new(),
    };

    // If formula-xlsb can detect a formula record but can't decode its RGCE token stream yet,
    // it surfaces the formula without any text. In that case we keep the cached value, but also
    // track the cell so we can fall back to Calamine's formula text extraction.
    let mut missing_formulas: Vec<(
        usize,
        String,
        Vec<(usize, usize, CellScalar, Option<String>)>,
    )> = Vec::new();

    for (idx, sheet_meta) in wb.sheet_metas().iter().enumerate() {
        let mut sheet = Sheet::new(sheet_meta.name.clone(), sheet_meta.name.clone());
        sheet.origin_ordinal = Some(idx);
        // Best-effort: extract per-column metadata (width/hidden) without materializing the full
        // worksheet part. This enables Excel-compatible `CELL("width")` behavior for XLSB inputs.
        if let Ok(props) = read_col_properties(&sheet_meta.part_path) {
            sheet.col_properties = props;
        }
        let styles = wb.styles();
        let mut undecoded_formula_cells: Vec<(usize, usize, CellScalar, Option<String>)> =
            Vec::new();

        wb.for_each_cell(idx, |cell| {
            let row = cell.row as usize;
            let col = cell.col as usize;

            let number_format = styles
                .get(cell.style)
                .filter(|info| info.is_date_time)
                .and_then(|info| info.number_format.clone());

            let value = match cell.value {
                XlsbCellValue::Blank => CellScalar::Empty,
                XlsbCellValue::Number(n) => CellScalar::Number(n),
                XlsbCellValue::Bool(b) => CellScalar::Bool(b),
                XlsbCellValue::Error(code) => CellScalar::Error(xlsb_error_display(code)),
                XlsbCellValue::Text(s) => CellScalar::Text(s),
            };

            match cell.formula {
                Some(formula) => match formula.text {
                    Some(formula) => {
                        let normalized = formula_model::display_formula_text(&formula);
                        if normalized.is_empty() {
                            // Treat empty formulas as blank/no-formula cells.
                            if matches!(value, CellScalar::Empty) {
                                return;
                            }
                            let mut c = Cell::from_literal(Some(value));
                            c.number_format = number_format;
                            sheet.set_cell(row, col, c);
                            return;
                        }
                        let mut c = Cell::from_formula(normalized);
                        c.computed_value = value;
                        c.number_format = number_format;
                        sheet.set_cell(row, col, c);
                    }
                    None => {
                        // Preserve the cached value and fill the formula text later (best-effort).
                        undecoded_formula_cells.push((row, col, value, number_format));
                    }
                },
                None => {
                    if matches!(value, CellScalar::Empty) {
                        return;
                    }
                    let mut c = Cell::from_literal(Some(value));
                    c.number_format = number_format;
                    sheet.set_cell(row, col, c);
                }
            }
        })
        .with_context(|| format!("read xlsb sheet {}", sheet_meta.name))?;

        if !undecoded_formula_cells.is_empty() {
            missing_formulas.push((
                out.sheets.len(),
                sheet_meta.name.clone(),
                undecoded_formula_cells,
            ));
        }
        out.sheets.push(sheet);
    }

    out.defined_names = wb
        .defined_names()
        .iter()
        .filter_map(|dn| {
            let formula = dn.formula.as_ref()?;
            let refers_to = formula.text.as_ref()?;
            if refers_to.trim().is_empty() {
                return None;
            }
            let sheet_id = dn.scope_sheet.and_then(|idx| {
                wb.sheet_metas()
                    .get(idx as usize)
                    .map(|meta| meta.name.clone())
            });
            Some(DefinedName {
                name: dn.name.clone(),
                refers_to: refers_to.clone(),
                sheet_id,
                hidden: dn.hidden,
            })
        })
        .collect();

    if !missing_formulas.is_empty() {
        // Best-effort: if Calamine can't open the workbook (or doesn't expose formulas), fall back
        // to the cached values only, matching the previous behavior.
        //
        // Note: this intentionally uses the on-disk `path` even when the workbook was opened from
        // decrypted in-memory bytes (encrypted `.xlsb`). In that case Calamine will fail to open
        // the OLE container and we'll fall back to cached values only.
        let mut calamine_wb = open_workbook_auto(path).ok();
        let mut formula_cache: HashMap<String, HashMap<(usize, usize), String>> = HashMap::new();
        let empty_lookup: HashMap<(usize, usize), String> = HashMap::new();

        for (sheet_idx, sheet_name, missing_cells) in missing_formulas {
            let formula_lookup = if let Some(wb) = calamine_wb.as_mut() {
                if !formula_cache.contains_key(&sheet_name) {
                    let lookup = calamine_formula_lookup_for_sheet(wb, &sheet_name);
                    formula_cache.insert(sheet_name.clone(), lookup);
                }
                formula_cache.get(&sheet_name).unwrap_or(&empty_lookup)
            } else {
                &empty_lookup
            };

            if let Some(sheet) = out.sheets.get_mut(sheet_idx) {
                apply_xlsb_formula_fallback(sheet, missing_cells, formula_lookup);
            }
        }
    }

    out.ensure_sheet_ids();
    for sheet in &mut out.sheets {
        sheet.clear_dirty_cells();
    }
    Ok(out)
}

fn read_xlsb_blocking(path: &Path) -> anyhow::Result<Workbook> {
    // Open XLSB with minimal preservation to reduce memory usage and exposure to
    // potentially-large/unsupported OPC parts (e.g. embedded media). The desktop app's XLSB save
    // path re-opens `Workbook::origin_xlsb_path` with preservation disabled as well, so we do not
    // need to keep `formula-xlsb`'s `preserved_parts` in-memory during read.
    let wb = XlsbWorkbook::open_with_options(
        path,
        XlsbOpenOptions {
            preserve_unknown_parts: false,
            preserve_parsed_parts: false,
            preserve_worksheets: false,
            // Keep formula decoding enabled to preserve historical behavior (UI-visible formulas).
            decode_formulas: true,
        },
    )
    .with_context(|| format!("open xlsb workbook {:?}", path))?;

    read_xlsb_from_open_workbook(
        path,
        wb,
        |part_path| read_xlsb_col_properties_for_sheet(path, part_path),
        Some(path.to_string_lossy().to_string()),
    )
}

fn calamine_formula_lookup_for_sheet<R, RS>(
    workbook: &mut R,
    sheet_name: &str,
) -> HashMap<(usize, usize), String>
where
    RS: std::io::Read + std::io::Seek,
    R: Reader<RS>,
{
    let mut out = HashMap::new();
    let Ok(formulas) = workbook.worksheet_formula(sheet_name) else {
        return out;
    };

    let (row_offset, col_offset) = formulas.start().unwrap_or((0, 0));
    let row_offset = row_offset as usize;
    let col_offset = col_offset as usize;

    for (row, col, formula) in formulas.cells() {
        let normalized = formula_model::display_formula_text(formula);
        if normalized.is_empty() {
            continue;
        }

        out.insert((row_offset + row, col_offset + col), normalized);
    }

    out
}

fn apply_xlsb_formula_fallback(
    sheet: &mut Sheet,
    missing_cells: Vec<(usize, usize, CellScalar, Option<String>)>,
    formula_lookup: &HashMap<(usize, usize), String>,
) {
    for (row, col, cached_value, number_format) in missing_cells {
        if let Some(formula) = formula_lookup.get(&(row, col)) {
            let mut cell = Cell::from_formula(formula.clone());
            cell.computed_value = cached_value;
            cell.number_format = number_format;
            sheet.set_cell(row, col, cell);
        } else if !matches!(cached_value, CellScalar::Empty) {
            let mut cell = Cell::from_literal(Some(cached_value));
            cell.number_format = number_format;
            sheet.set_cell(row, col, cell);
        }
    }
}

fn xlsb_error_display(code: u8) -> String {
    // Keep the string form compatible with what calamine uses for `Data::Error` so UI/engine
    // layers behave consistently.
    formula_xlsb::errors::xlsb_error_display(code)
}

#[cfg(feature = "desktop")]
pub async fn write_xlsx(
    path: impl Into<PathBuf> + Send + 'static,
    workbook: Workbook,
) -> anyhow::Result<Arc<[u8]>> {
    let path = path.into();
    tauri::async_runtime::spawn_blocking(move || write_xlsx_blocking(&path, &workbook))
        .await
        .map_err(|e| anyhow::anyhow!(e.to_string()))?
}

pub(crate) fn is_xlsx_family_extension(ext: &str) -> bool {
    matches!(ext, "xlsx" | "xlsm" | "xltx" | "xltm" | "xlam")
}

pub(crate) fn is_macro_free_xlsx_extension(ext: &str) -> bool {
    matches!(ext, "xlsx" | "xltx")
}

pub(crate) fn is_macro_enabled_xlsx_extension(ext: &str) -> bool {
    matches!(ext, "xlsm" | "xltm" | "xlam")
}

pub(crate) fn workbook_main_content_type_for_extension(ext: &str) -> Option<&'static str> {
    match ext {
        "xlsx" => Some(XLSX_WORKBOOK_CONTENT_TYPE),
        "xlsm" => Some(XLSM_WORKBOOK_CONTENT_TYPE),
        "xltx" => Some(XLTX_WORKBOOK_CONTENT_TYPE),
        "xltm" => Some(XLTM_WORKBOOK_CONTENT_TYPE),
        "xlam" => Some(XLAM_WORKBOOK_CONTENT_TYPE),
        _ => None,
    }
}

fn zip_part_exists(bytes: &[u8], part_name: &str) -> bool {
    let mut cursor = Cursor::new(bytes);
    let Ok(archive) = zip::ZipArchive::new(&mut cursor) else {
        return false;
    };
    zip_archive_has_entry(&archive, part_name)
}

fn workbook_override_matches_content_type(content_types_xml: &str, desired: &str) -> bool {
    let idx = match content_types_xml.find("/xl/workbook.xml") {
        Some(idx) => idx,
        None => return false,
    };
    let start = match content_types_xml[..idx].rfind('<') {
        Some(idx) => idx,
        None => return false,
    };
    let end = match content_types_xml[idx..].find('>') {
        Some(off) => idx + off,
        None => return false,
    };
    content_types_xml[start..=end].contains(desired)
}

pub(crate) fn patch_workbook_main_content_type_in_package(
    bytes: &[u8],
    desired: &str,
) -> anyhow::Result<Option<Vec<u8>>> {
    let Some(content_types_bytes) = formula_xlsx::read_part_from_reader_limited(
        Cursor::new(bytes),
        "[Content_Types].xml",
        XLSX_CONTENT_TYPES_MAX_BYTES,
    )
    .ok()
    .flatten() else {
        return Ok(None);
    };

    let Some(patched_xml) =
        formula_xlsx::rewrite_content_types_workbook_content_type(&content_types_bytes, desired)
            .context("rewrite workbook content type in [Content_Types].xml")?
    else {
        return Ok(None);
    };

    let mut part_overrides: HashMap<String, PartOverride> = HashMap::new();
    part_overrides.insert(
        "[Content_Types].xml".to_string(),
        PartOverride::Replace(patched_xml),
    );

    let mut cursor = Cursor::new(Vec::new());
    patch_xlsx_streaming_workbook_cell_patches_with_part_overrides(
        Cursor::new(bytes),
        &mut cursor,
        &WorkbookCellPatches::default(),
        &part_overrides,
    )
    .context("patch [Content_Types].xml workbook override content type")?;
    Ok(Some(cursor.into_inner()))
}

fn patch_vba_and_date_system_in_package_streaming(
    bytes: &[u8],
    workbook: &Workbook,
    needs_inject_vba: bool,
    needs_date_system_update: bool,
    xlsx_date_system: formula_xlsx::DateSystem,
) -> anyhow::Result<Vec<u8>> {
    use std::io::Write;
    use zip::write::FileOptions;
    use zip::{CompressionMethod, ZipWriter};

    if !needs_inject_vba && !needs_date_system_update {
        return Ok(bytes.to_vec());
    }

    let Some(content_types_xml) = formula_xlsx::read_part_from_reader_limited(
        Cursor::new(bytes),
        "[Content_Types].xml",
        XLSX_CONTENT_TYPES_MAX_BYTES,
    )
    .context("read [Content_Types].xml")?
    else {
        anyhow::bail!("missing [Content_Types].xml part");
    };

    let Some(workbook_xml) = formula_xlsx::read_part_from_reader_limited(
        Cursor::new(bytes),
        "xl/workbook.xml",
        XLSX_WORKBOOK_XML_MAX_BYTES,
    )
    .context("read xl/workbook.xml")?
    else {
        anyhow::bail!("missing xl/workbook.xml part");
    };

    let Some(workbook_rels_xml) = formula_xlsx::read_part_from_reader_limited(
        Cursor::new(bytes),
        "xl/_rels/workbook.xml.rels",
        XLSX_WORKBOOK_XML_MAX_BYTES,
    )
    .context("read xl/_rels/workbook.xml.rels")?
    else {
        anyhow::bail!("missing xl/_rels/workbook.xml.rels part");
    };

    // Include `xl/_rels/vbaProject.bin.rels` when present so macro repair can patch/extend it
    // (and so we preserve any existing relationships in that part).
    let vba_project_rels_xml = formula_xlsx::read_part_from_reader_limited(
        Cursor::new(bytes),
        "xl/_rels/vbaProject.bin.rels",
        XLSX_WORKBOOK_XML_MAX_BYTES,
    )
    .ok()
    .flatten();

    // Build a "subset" package containing just the parts we need to update. This allows us to
    // rely on `XlsxPackage::write_to_bytes()`'s macro-repair logic (to update `[Content_Types].xml`
    // and `.rels` files) without inflating the full workbook (which may contain very large sheet
    // XML parts).
    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(CompressionMethod::Deflated);

    zip.start_file("[Content_Types].xml", options)
        .context("start subset [Content_Types].xml")?;
    zip.write_all(&content_types_xml)
        .context("write subset [Content_Types].xml")?;

    zip.start_file("xl/workbook.xml", options)
        .context("start subset xl/workbook.xml")?;
    zip.write_all(&workbook_xml)
        .context("write subset xl/workbook.xml")?;

    zip.start_file("xl/_rels/workbook.xml.rels", options)
        .context("start subset xl/_rels/workbook.xml.rels")?;
    zip.write_all(&workbook_rels_xml)
        .context("write subset xl/_rels/workbook.xml.rels")?;

    if let Some(xml) = vba_project_rels_xml {
        zip.start_file("xl/_rels/vbaProject.bin.rels", options)
            .context("start subset xl/_rels/vbaProject.bin.rels")?;
        zip.write_all(&xml)
            .context("write subset xl/_rels/vbaProject.bin.rels")?;
    }

    let subset_bytes = zip
        .finish()
        .context("finalize subset package zip")?
        .into_inner();

    let subset_limits = XlsxPackageLimits {
        max_part_bytes: MAX_VBA_PROJECT_BIN_BYTES
            .max(MAX_VBA_PROJECT_SIGNATURE_BIN_BYTES)
            .max(XLSX_WORKBOOK_XML_MAX_BYTES as usize)
            .max(XLSX_CONTENT_TYPES_MAX_BYTES as usize) as u64,
        // The subset contains at most a handful of XML parts plus the VBA payload (<=16MiB); keep
        // the budget tight to avoid surprising allocations.
        max_total_bytes: 64 * 1024 * 1024, // 64MiB
    };

    let mut pkg = XlsxPackage::from_bytes_limited(&subset_bytes, subset_limits)
        .context("parse subset package for VBA/date system patching")?;

    if needs_inject_vba {
        pkg.set_part(
            "xl/vbaProject.bin",
            workbook.vba_project_bin.clone().expect("checked is_some"),
        );
        if let Some(sig) = workbook.vba_project_signature_bin.clone() {
            pkg.set_part("xl/vbaProjectSignature.bin", sig);
        }
    }

    if needs_date_system_update {
        pkg.set_workbook_date_system(xlsx_date_system)
            .context("set workbook date system")?;
    }

    let repacked_subset = pkg
        .write_to_bytes()
        .context("write subset package with VBA/date system updates")?;

    // Re-parse so we capture macro repair outputs (content types + relationship repairs).
    let repacked_pkg = XlsxPackage::from_bytes_limited(&repacked_subset, subset_limits)
        .context("parse repacked subset package")?;

    let mut part_overrides: HashMap<String, PartOverride> = HashMap::new();
    for (name, bytes) in repacked_pkg.parts() {
        let canonical = name.strip_prefix('/').unwrap_or(name).to_string();
        part_overrides.insert(canonical, PartOverride::Replace(bytes.to_vec()));
    }

    let mut cursor = Cursor::new(Vec::new());
    patch_xlsx_streaming_workbook_cell_patches_with_part_overrides(
        Cursor::new(bytes),
        &mut cursor,
        &WorkbookCellPatches::default(),
        &part_overrides,
    )
    .context("apply VBA/date system part overrides (streaming)")?;

    Ok(cursor.into_inner())
}

fn workbook_xml_sheet_order_override(
    origin_bytes: &[u8],
    workbook: &Workbook,
) -> anyhow::Result<Option<Vec<u8>>> {
    let Some(workbook_xml_bytes) = formula_xlsx::read_part_from_reader_limited(
        Cursor::new(origin_bytes),
        "xl/workbook.xml",
        XLSX_WORKBOOK_XML_MAX_BYTES,
    )
    .context("read xl/workbook.xml")?
    else {
        return Ok(None);
    };
    let workbook_xml =
        std::str::from_utf8(&workbook_xml_bytes).context("decode xl/workbook.xml")?;

    let worksheet_parts = formula_xlsx::worksheet_parts_from_reader_limited(
        Cursor::new(origin_bytes),
        XLSX_WORKBOOK_XML_MAX_BYTES,
    )
    .context("resolve worksheet parts")?;
    if worksheet_parts.is_empty() {
        return Ok(None);
    }

    // Only attempt to rewrite the sheet list when the in-memory workbook and the origin package
    // agree on the set of sheets (reordering should not add/remove sheets).
    if workbook.sheets.len() != worksheet_parts.len() {
        return Ok(None);
    }

    let mut info_by_part: HashMap<String, formula_xlsx::WorkbookSheetInfo> =
        HashMap::with_capacity(worksheet_parts.len());
    for part in &worksheet_parts {
        info_by_part.insert(
            part.worksheet_part.clone(),
            formula_xlsx::WorkbookSheetInfo {
                name: part.name.clone(),
                sheet_id: part.sheet_id,
                rel_id: part.rel_id.clone(),
                visibility: part.visibility,
            },
        );
    }

    let mut reordered_infos: Vec<formula_xlsx::WorkbookSheetInfo> =
        Vec::with_capacity(workbook.sheets.len());
    let mut reordered_parts: Vec<String> = Vec::with_capacity(workbook.sheets.len());
    let mut seen_parts: HashSet<String> = HashSet::new();

    for sheet in &workbook.sheets {
        let resolved_part = match sheet.xlsx_worksheet_part.as_deref() {
            Some(part) => Some(part.to_string()),
            None => worksheet_parts
                .iter()
                .find(|p| sheet_name_eq_case_insensitive(&p.name, &sheet.name))
                .map(|p| p.worksheet_part.clone()),
        };

        let Some(part) = resolved_part else {
            return Ok(None);
        };

        if !seen_parts.insert(part.clone()) {
            // Duplicate worksheet part resolution; bail out to avoid dropping sheets.
            return Ok(None);
        }

        let Some(info) = info_by_part.get(&part) else {
            return Ok(None);
        };

        reordered_parts.push(part);
        reordered_infos.push(info.clone());
    }

    let original_parts: Vec<&str> = worksheet_parts
        .iter()
        .map(|p| p.worksheet_part.as_str())
        .collect();
    let next_parts: Vec<&str> = reordered_parts.iter().map(|p| p.as_str()).collect();
    if next_parts == original_parts {
        return Ok(None);
    }

    let rewritten = formula_xlsx::write_workbook_sheets(workbook_xml, &reordered_infos)
        .map_err(|e| anyhow::anyhow!(e.to_string()))
        .context("rewrite workbook.xml sheets")?;
    if rewritten == workbook_xml {
        return Ok(None);
    }

    Ok(Some(rewritten.into_bytes()))
}

fn parse_xml_bool(val: &str) -> bool {
    val == "1" || val.eq_ignore_ascii_case("true")
}

fn worksheet_spreadsheetml_prefix(xml: &str) -> anyhow::Result<Option<String>> {
    use quick_xml::events::Event;
    use quick_xml::Reader;

    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();

    loop {
        match reader
            .read_event_into(&mut buf)
            .map_err(|e| anyhow::anyhow!(e.to_string()))?
        {
            Event::Start(e) | Event::Empty(e) if e.local_name().as_ref() == b"worksheet" => {
                let name = e.name();
                let name = name.as_ref();
                let prefix = name
                    .iter()
                    .rposition(|b| *b == b':')
                    .map(|idx| &name[..idx])
                    .and_then(|p| std::str::from_utf8(p).ok())
                    .map(|s| s.to_string());
                return Ok(prefix);
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    Ok(None)
}

fn normalized_sheet_col_properties(
    props: &BTreeMap<u32, ModelColProperties>,
) -> BTreeMap<u32, (Option<f32>, bool)> {
    let mut out = BTreeMap::new();
    for (&col, prop) in props {
        let width = prop.width;
        let hidden = prop.hidden;
        if width.is_none() && !hidden {
            continue;
        }
        out.insert(col, (width, hidden));
    }
    out
}

fn parse_col_properties_width_hidden(
    sheet_xml: &str,
) -> anyhow::Result<BTreeMap<u32, (Option<f32>, bool)>> {
    use quick_xml::events::Event;
    use quick_xml::Reader;

    let mut reader = Reader::from_str(sheet_xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut in_cols = false;
    let mut map: BTreeMap<u32, (Option<f32>, bool)> = BTreeMap::new();

    loop {
        match reader
            .read_event_into(&mut buf)
            .map_err(|e| anyhow::anyhow!(e.to_string()))?
        {
            Event::Eof => break,
            Event::Start(e) if e.local_name().as_ref() == b"cols" => in_cols = true,
            Event::End(e) if e.local_name().as_ref() == b"cols" => {
                in_cols = false;
                drop(e);
            }
            Event::Empty(e) if e.local_name().as_ref() == b"cols" => {
                in_cols = false;
                drop(e);
            }
            Event::Start(e) | Event::Empty(e) if in_cols && e.local_name().as_ref() == b"col" => {
                let mut min: Option<u32> = None;
                let mut max: Option<u32> = None;
                let mut width: Option<f32> = None;
                let mut custom_width: Option<bool> = None;
                let mut hidden = false;

                for attr in e.attributes().with_checks(false) {
                    let attr = attr.map_err(|e| anyhow::anyhow!(e.to_string()))?;
                    let val = attr
                        .unescape_value()
                        .map_err(|e| anyhow::anyhow!(e.to_string()))?
                        .into_owned();
                    match attr.key.as_ref() {
                        b"min" => min = val.parse().ok(),
                        b"max" => max = val.parse().ok(),
                        b"width" => width = val.parse().ok(),
                        b"customWidth" => custom_width = Some(parse_xml_bool(&val)),
                        b"hidden" => hidden = parse_xml_bool(&val),
                        _ => {}
                    }
                }

                let Some(min_1) = min else { continue };
                let max_1 = max.unwrap_or(min_1).min(formula_model::EXCEL_MAX_COLS);
                if min_1 == 0 || max_1 == 0 || min_1 > formula_model::EXCEL_MAX_COLS {
                    continue;
                }

                let width = if custom_width == Some(false) {
                    None
                } else {
                    width
                };

                for col_1 in min_1..=max_1 {
                    let col_0 = col_1 - 1;
                    if col_0 >= formula_model::EXCEL_MAX_COLS {
                        continue;
                    }
                    if width.is_none() && !hidden {
                        continue;
                    }
                    let entry = map.entry(col_0).or_insert((None, false));
                    if width.is_some() {
                        entry.0 = width;
                    }
                    if hidden {
                        entry.1 = true;
                    }
                }
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(map)
}

fn escape_xml_attr(value: &str) -> String {
    value
        .replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
        .replace('\'', "&apos;")
}

fn parse_cols_attribute_map(
    sheet_xml: &str,
) -> anyhow::Result<BTreeMap<u32, BTreeMap<String, String>>> {
    use quick_xml::events::Event;
    use quick_xml::Reader;

    let mut reader = Reader::from_str(sheet_xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut in_cols = false;
    let mut attrs_by_col: BTreeMap<u32, BTreeMap<String, String>> = BTreeMap::new();

    loop {
        match reader
            .read_event_into(&mut buf)
            .map_err(|e| anyhow::anyhow!(e.to_string()))?
        {
            Event::Eof => break,
            Event::Start(e) if e.local_name().as_ref() == b"cols" => in_cols = true,
            Event::End(e) if e.local_name().as_ref() == b"cols" => {
                in_cols = false;
                drop(e);
            }
            Event::Empty(e) if e.local_name().as_ref() == b"cols" => {
                in_cols = false;
                drop(e);
            }
            Event::Start(e) | Event::Empty(e) if in_cols && e.local_name().as_ref() == b"col" => {
                let mut min: Option<u32> = None;
                let mut max: Option<u32> = None;
                let mut element_attrs: Vec<(String, String)> = Vec::new();

                for attr in e.attributes().with_checks(false) {
                    let attr = attr.map_err(|e| anyhow::anyhow!(e.to_string()))?;
                    let key = std::str::from_utf8(attr.key.as_ref())
                        .map_err(|e| anyhow::anyhow!(e.to_string()))?
                        .to_string();
                    let val = attr
                        .unescape_value()
                        .map_err(|e| anyhow::anyhow!(e.to_string()))?
                        .into_owned();
                    match key.as_str() {
                        "min" => min = val.parse().ok(),
                        "max" => max = val.parse().ok(),
                        _ => element_attrs.push((key, val)),
                    }
                }

                let Some(min_1) = min else { continue };
                let max_1 = max.unwrap_or(min_1).min(formula_model::EXCEL_MAX_COLS);
                if min_1 == 0 || max_1 == 0 || min_1 > formula_model::EXCEL_MAX_COLS {
                    continue;
                }

                for col_1 in min_1..=max_1 {
                    let entry = attrs_by_col.entry(col_1).or_default();
                    for (k, v) in &element_attrs {
                        entry.insert(k.clone(), v.clone());
                    }
                }
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(attrs_by_col)
}

fn render_cols_section(
    prefix: Option<&str>,
    attrs_by_col: &BTreeMap<u32, BTreeMap<String, String>>,
) -> String {
    if attrs_by_col.is_empty() {
        return String::new();
    }

    let cols_tag = match prefix {
        Some(p) => format!("{p}:cols"),
        None => "cols".to_string(),
    };
    let col_tag = match prefix {
        Some(p) => format!("{p}:col"),
        None => "col".to_string(),
    };

    let mut out = String::new();
    out.push('<');
    out.push_str(&cols_tag);
    out.push('>');

    let mut current: Option<(u32, u32, BTreeMap<String, String>)> = None;
    for (&col_1, attrs) in attrs_by_col {
        let attrs = attrs.clone();
        match current.take() {
            None => current = Some((col_1, col_1, attrs)),
            Some((start, end, cur)) if col_1 == end + 1 && attrs == cur => {
                current = Some((start, col_1, cur));
            }
            Some((start, end, cur)) => {
                out.push_str(&render_col_range(&col_tag, start, end, &cur));
                current = Some((col_1, col_1, attrs));
            }
        }
    }
    if let Some((start, end, cur)) = current {
        out.push_str(&render_col_range(&col_tag, start, end, &cur));
    }

    out.push_str("</");
    out.push_str(&cols_tag);
    out.push('>');
    out
}

fn render_col_range(
    col_tag: &str,
    start_col_1: u32,
    end_col_1: u32,
    attrs: &BTreeMap<String, String>,
) -> String {
    let mut s = String::new();
    s.push('<');
    s.push_str(col_tag);
    s.push_str(&format!(r#" min="{start_col_1}" max="{end_col_1}""#));
    for (k, v) in attrs {
        s.push(' ');
        s.push_str(k);
        s.push_str("=\"");
        s.push_str(&escape_xml_attr(v));
        s.push('"');
    }
    s.push_str("/>");
    s
}

fn update_cols_xml(sheet_xml: &str, cols_section: &str) -> anyhow::Result<String> {
    use quick_xml::events::Event;
    use quick_xml::Reader;
    use quick_xml::Writer;

    let mut reader = Reader::from_str(sheet_xml);
    reader.config_mut().trim_text(false);

    let mut writer = Writer::new(Vec::new());
    let mut buf = Vec::new();

    let mut skip_depth: usize = 0;
    let mut replaced = false;
    let mut inserted = false;

    loop {
        let event = reader
            .read_event_into(&mut buf)
            .map_err(|e| anyhow::anyhow!(e.to_string()))?;
        match event {
            Event::Eof => break,
            _ if skip_depth > 0 => match event {
                Event::Start(_) => skip_depth += 1,
                Event::End(_) => skip_depth = skip_depth.saturating_sub(1),
                Event::Empty(_) => {}
                _ => {}
            },
            Event::Start(ref e) if e.local_name().as_ref() == b"cols" => {
                replaced = true;
                if !cols_section.is_empty() {
                    writer.get_mut().extend_from_slice(cols_section.as_bytes());
                }
                skip_depth = 1;
            }
            Event::Empty(ref e) if e.local_name().as_ref() == b"cols" => {
                replaced = true;
                if !cols_section.is_empty() {
                    writer.get_mut().extend_from_slice(cols_section.as_bytes());
                }
            }
            Event::Start(ref e) if e.local_name().as_ref() == b"sheetData" => {
                if !replaced && !inserted && !cols_section.is_empty() {
                    writer.get_mut().extend_from_slice(cols_section.as_bytes());
                    inserted = true;
                }
                writer
                    .write_event(Event::Start(e.to_owned()))
                    .map_err(|e| anyhow::anyhow!(e.to_string()))?;
            }
            Event::Empty(ref e) if e.local_name().as_ref() == b"sheetData" => {
                if !replaced && !inserted && !cols_section.is_empty() {
                    writer.get_mut().extend_from_slice(cols_section.as_bytes());
                    inserted = true;
                }
                writer
                    .write_event(Event::Empty(e.to_owned()))
                    .map_err(|e| anyhow::anyhow!(e.to_string()))?;
            }
            _ => {
                writer
                    .write_event(event.to_owned())
                    .map_err(|e| anyhow::anyhow!(e.to_string()))?;
            }
        }
        buf.clear();
    }

    String::from_utf8(writer.into_inner())
        .map_err(|e| anyhow::anyhow!(std::io::Error::new(std::io::ErrorKind::InvalidData, e)))
}

fn patch_cols_in_worksheet_xml(
    sheet_xml: &str,
    desired: &BTreeMap<u32, ModelColProperties>,
) -> anyhow::Result<Option<String>> {
    let current = parse_col_properties_width_hidden(sheet_xml)?;
    let desired_norm = normalized_sheet_col_properties(desired);
    if current == desired_norm {
        return Ok(None);
    }

    let prefix = worksheet_spreadsheetml_prefix(sheet_xml)?;

    let mut attrs_by_col = parse_cols_attribute_map(sheet_xml)?;

    // Apply desired width/hidden overrides, preserving unsupported attributes from the source.
    // Columns not present in `desired` are treated as default (clear width/hidden).
    let mut touched_cols: std::collections::BTreeSet<u32> = attrs_by_col.keys().copied().collect();
    touched_cols.extend(desired.keys().copied().map(|c0| c0.saturating_add(1)));

    for col_1 in touched_cols {
        if col_1 == 0 || col_1 > formula_model::EXCEL_MAX_COLS {
            continue;
        }
        let col_0 = col_1 - 1;
        let desired_props = desired.get(&col_0);
        let desired_width = desired_props.and_then(|p| p.width);
        let desired_hidden = desired_props.map(|p| p.hidden).unwrap_or(false);

        let mut entry = attrs_by_col.remove(&col_1).unwrap_or_default();

        match desired_width {
            Some(width) => {
                entry.insert("width".to_string(), width.to_string());
                entry.insert("customWidth".to_string(), "1".to_string());
            }
            None => {
                entry.remove("width");
                entry.remove("customWidth");
            }
        }

        if desired_hidden {
            entry.insert("hidden".to_string(), "1".to_string());
        } else {
            entry.remove("hidden");
        }

        if !entry.is_empty() {
            attrs_by_col.insert(col_1, entry);
        }
    }

    let cols_section = render_cols_section(prefix.as_deref(), &attrs_by_col);
    let updated = update_cols_xml(sheet_xml, &cols_section)?;
    Ok(Some(updated))
}

fn sheet_metadata_part_overrides(
    bytes: &[u8],
    workbook: &Workbook,
) -> anyhow::Result<HashMap<String, PartOverride>> {
    let mut part_overrides: HashMap<String, PartOverride> = HashMap::new();

    // --- workbook.xml sheet visibility (state="hidden"/"veryHidden") ---
    let workbook_xml_bytes = formula_xlsx::read_part_from_reader_limited(
        Cursor::new(bytes),
        "xl/workbook.xml",
        XLSX_WORKBOOK_XML_MAX_BYTES,
    )
    .map_err(|e| anyhow::anyhow!(e.to_string()))
    .context("read xl/workbook.xml")?
    .ok_or_else(|| anyhow::anyhow!("missing xl/workbook.xml"))?;
    let workbook_xml =
        std::str::from_utf8(&workbook_xml_bytes).context("parse xl/workbook.xml as utf8")?;

    let mut sheets = parse_workbook_sheets(workbook_xml)
        .map_err(|e| anyhow::anyhow!(e.to_string()))
        .context("parse workbook sheet metadata")?;

    // Resolve worksheet part <-> relId so we can update visibility even when the sheet has been
    // renamed in-app (the origin workbook.xml might not be rewritten yet).
    let mut worksheet_parts_by_name: HashMap<String, String> = HashMap::new();
    let mut rel_id_by_part: HashMap<String, String> = HashMap::new();
    if let Ok(parts) = formula_xlsx::worksheet_parts_from_reader_limited(
        Cursor::new(bytes),
        XLSX_WORKBOOK_XML_MAX_BYTES,
    ) {
        for part in parts {
            rel_id_by_part.insert(part.worksheet_part.clone(), part.rel_id.clone());
            worksheet_parts_by_name.insert(part.name, part.worksheet_part);
        }
    }

    let mut desired_visibility_by_name: HashMap<&str, SheetVisibility> = HashMap::new();
    let mut desired_visibility_by_rel_id: HashMap<String, SheetVisibility> = HashMap::new();
    for sheet in &workbook.sheets {
        if let Some(part) = sheet.xlsx_worksheet_part.as_deref() {
            if let Some(rel_id) = rel_id_by_part.get(part) {
                desired_visibility_by_rel_id.insert(rel_id.clone(), sheet.visibility);
                continue;
            }
        }
        desired_visibility_by_name.insert(sheet.name.as_str(), sheet.visibility);
    }

    let mut workbook_xml_changed = false;
    for sheet in &mut sheets {
        let desired = desired_visibility_by_rel_id
            .get(&sheet.rel_id)
            .copied()
            .or_else(|| desired_visibility_by_name.get(sheet.name.as_str()).copied());
        if let Some(visibility) = desired {
            if sheet.visibility != visibility {
                sheet.visibility = visibility;
                workbook_xml_changed = true;
            }
        }
    }

    if workbook_xml_changed {
        let updated = write_workbook_sheets(workbook_xml, &sheets)
            .map_err(|e| anyhow::anyhow!(e.to_string()))
            .context("rewrite xl/workbook.xml sheet metadata")?;
        part_overrides.insert(
            "xl/workbook.xml".to_string(),
            PartOverride::Replace(updated.into_bytes()),
        );
    }

    // --- worksheet XML metadata ---
    // Use the cached worksheet part when available; fall back to re-discovering parts from the package.
    for sheet in &workbook.sheets {
        let part_name = sheet
            .xlsx_worksheet_part
            .as_deref()
            .or_else(|| worksheet_parts_by_name.get(&sheet.name).map(|s| s.as_str()));
        let Some(part_name) = part_name else {
            continue;
        };

        let sheet_xml_bytes = match formula_xlsx::read_part_from_reader_limited(
            Cursor::new(bytes),
            part_name,
            XLSX_WORKSHEET_XML_MAX_BYTES,
        ) {
            Ok(Some(bytes)) => bytes,
            Ok(None) => continue,
            Err(formula_xlsx::XlsxError::PartTooLarge { part, size, max }) => {
                // Best-effort: if the worksheet XML is too large to load safely, skip patching
                // optional worksheet-level metadata (tabColor/cols) rather than failing the entire
                // save. The original worksheet part will be preserved byte-for-byte by the
                // streaming patcher.
                eprintln!(
                    "warning: dropped oversized worksheet part `{part}` while patching sheet metadata ({size} bytes, max {max})"
                );
                continue;
            }
            Err(err) => {
                // Best-effort: invalid/unsupported worksheet XML should not fail save.
                eprintln!(
                    "warning: failed to read worksheet part `{part_name}` while patching sheet metadata: {err}"
                );
                continue;
            }
        };

        let sheet_xml = match std::str::from_utf8(&sheet_xml_bytes) {
            Ok(xml) => xml,
            Err(err) => {
                eprintln!(
                    "warning: failed to parse worksheet part `{part_name}` as utf8 while patching sheet metadata: {err}"
                );
                continue;
            }
        };

        let mut updated_xml = sheet_xml.to_string();
        let mut changed = false;

        // tabColor (sheetPr/tabColor)
        let current_tab_color = parse_sheet_tab_color(sheet_xml)
            .map_err(|e| anyhow::anyhow!(e.to_string()))
            .with_context(|| format!("parse tabColor in {part_name}"))?;
        if current_tab_color != sheet.tab_color {
            updated_xml = write_sheet_tab_color(&updated_xml, sheet.tab_color.as_ref())
                .map_err(|e| anyhow::anyhow!(e.to_string()))
                .with_context(|| format!("rewrite tabColor in {part_name}"))?;
            changed = true;
        }

        // Column widths / hidden columns (<cols>/<col>)
        if let Some(patched) = patch_cols_in_worksheet_xml(&updated_xml, &sheet.col_properties)? {
            updated_xml = patched;
            changed = true;
        }

        if changed {
            part_overrides.insert(
                part_name.to_string(),
                PartOverride::Replace(updated_xml.into_bytes()),
            );
        }
    }

    Ok(part_overrides)
}

fn patch_sheet_metadata_in_package(
    bytes: &[u8],
    workbook: &Workbook,
) -> anyhow::Result<Option<Vec<u8>>> {
    let part_overrides = sheet_metadata_part_overrides(bytes, workbook)?;
    if part_overrides.is_empty() {
        return Ok(None);
    }

    let mut cursor = Cursor::new(Vec::new());
    let empty_patches = WorkbookCellPatches::default();
    patch_xlsx_streaming_workbook_cell_patches_with_part_overrides(
        Cursor::new(bytes),
        &mut cursor,
        &empty_patches,
        &part_overrides,
    )
    .context("apply sheet metadata overrides (streaming)")?;

    Ok(Some(cursor.into_inner()))
}

/// Build an XLSX-family package as an in-memory ZIP/OPC byte buffer.
///
/// This performs the same XLSX patch/export logic as [`write_xlsx_blocking`] but does **not** write
/// anything to disk.
///
/// This is used by password-protected save flows, which must never write plaintext ZIP bytes to
/// disk before wrapping them in an Office encryption container.
pub fn build_xlsx_bytes_blocking(path: &Path, workbook: &Workbook) -> anyhow::Result<Arc<[u8]>> {
    let extension = path
        .extension()
        .and_then(|s| s.to_str())
        .map(|s| s.to_ascii_lowercase());

    if matches!(extension.as_deref(), Some("xlsb")) {
        anyhow::bail!("building .xlsb bytes in memory is not supported");
    }

    let workbook_kind = extension
        .as_deref()
        .and_then(WorkbookKind::from_extension)
        .unwrap_or(WorkbookKind::Workbook);

    let xlsx_date_system = match workbook.date_system {
        WorkbookDateSystem::Excel1900 => formula_xlsx::DateSystem::V1900,
        WorkbookDateSystem::Excel1904 => formula_xlsx::DateSystem::V1904,
    };

    fn origin_package_has_macro_content(origin_bytes: &[u8]) -> bool {
        // Determining whether we need to run the (potentially expensive) macro-stripping rewrite
        // must not depend on `Workbook::vba_project_bin`, because macro-enabled workbooks can
        // contain non-VBA macro surfaces (e.g. XLM macrosheets / legacy dialog sheets).
        //
        // Keep this check cheap: only scan ZIP entry names and (when present) inspect
        // `[Content_Types].xml` for macro-enabled workbook types.
        let mut archive = match zip::ZipArchive::new(Cursor::new(origin_bytes)) {
            Ok(archive) => archive,
            Err(_) => return false,
        };

        let mut saw_content_types = false;
        for i in 0..archive.len() {
            let file = match archive.by_index(i) {
                Ok(file) => file,
                Err(_) => continue,
            };
            if file.is_dir() {
                continue;
            }

            let name = file.name();
            let name = name.strip_prefix('/').unwrap_or(name);
            let lower = name.to_ascii_lowercase();
            if lower == "xl/vbaproject.bin"
                || lower == "xl/vbadata.xml"
                || lower == "xl/vbaprojectsignature.bin"
                || lower.starts_with("xl/macrosheets/")
                || lower.starts_with("xl/dialogsheets/")
            {
                return true;
            }

            if lower == "[content_types].xml" {
                saw_content_types = true;
            }
        }

        if saw_content_types {
            if let Ok(Some(bytes)) = formula_xlsx::read_part_from_reader_limited(
                Cursor::new(origin_bytes),
                "[Content_Types].xml",
                XLSX_CONTENT_TYPES_MAX_BYTES,
            ) {
                let content_types = String::from_utf8_lossy(&bytes);
                if content_types.contains("macroEnabled.main+xml") {
                    return true;
                }
            }
        }

        false
    }

    fn origin_xlsx_bytes_support_patch_based_save(origin_bytes: &[u8]) -> bool {
        // The patch-based save path uses `formula_xlsx`'s streaming patcher, which must inflate
        // workbook-level XML parts like `xl/workbook.xml` to resolve sheet relationships.
        //
        // For extremely large (or ZIP-bomb) parts, prefer falling back to the regeneration-based
        // save path rather than risking OOM.
        match formula_xlsx::read_part_from_reader_limited(
            Cursor::new(origin_bytes),
            "xl/workbook.xml",
            XLSX_WORKBOOK_XML_MAX_BYTES,
        ) {
            Ok(Some(_)) => true,
            Ok(None) => false,
            Err(formula_xlsx::XlsxError::PartTooLarge { part, size, max }) => {
                eprintln!(
                    "[save] skipping patch-based save: xlsx part `{part}` is too large ({size} bytes, max {max})"
                );
                false
            }
            Err(err) => {
                eprintln!(
                    "[save] skipping patch-based save: failed to read xl/workbook.xml: {err}"
                );
                false
            }
        }
    }

    if let Some(origin_bytes) = workbook
        .origin_xlsx_bytes
        .as_deref()
        .filter(|bytes| origin_xlsx_bytes_support_patch_based_save(bytes))
    {
        // NOTE: This patch-based save path intentionally preserves most workbook-level parts
        // from the original package. This keeps unsupported XLSX parts (theme, comments,
        // conditional formatting, etc) intact by patching only the modified worksheet XML.
        let print_settings_changed = workbook.print_settings != workbook.original_print_settings;
        let power_query_changed = workbook.power_query_xml != workbook.original_power_query_xml;
        let sheet_order_override = workbook_xml_sheet_order_override(origin_bytes, workbook)
            .ok()
            .flatten();

        let mut patches = WorkbookCellPatches::default();
        for sheet in &workbook.sheets {
            let sheet_selector = sheet
                .xlsx_worksheet_part
                .clone()
                .unwrap_or_else(|| sheet.name.clone());
            for (row, col) in &sheet.dirty_cells {
                if let Some((baseline_value, baseline_formula)) = workbook
                    .cell_input_baseline
                    .get(&(sheet.id.clone(), *row, *col))
                {
                    let (current_value, current_formula) = match sheet.cells.get(&(*row, *col)) {
                        Some(cell) => (cell.input_value.clone(), cell.formula.clone()),
                        None => (None, None),
                    };

                    if &current_value == baseline_value && &current_formula == baseline_formula {
                        continue;
                    }
                }

                let cell_ref = formula_model::CellRef::new(*row as u32, *col as u32);
                let Some(cell) = sheet.cells.get(&(*row, *col)) else {
                    patches.set_cell(sheet_selector.clone(), cell_ref, XlsxCellPatch::clear());
                    continue;
                };

                let (formula, scalar) = match (&cell.formula, &cell.input_value) {
                    (Some(f), _) => (Some(f.clone()), cell.computed_value.clone()),
                    (None, Some(v)) => (None, v.clone()),
                    (None, None) => (None, CellScalar::Empty),
                };

                let patch = match formula {
                    Some(formula) => XlsxCellPatch::set_value_with_formula(
                        scalar_to_model_value(&scalar),
                        formula,
                    ),
                    None => XlsxCellPatch::set_value(scalar_to_model_value(&scalar)),
                };

                patches.set_cell(sheet_selector.clone(), cell_ref, patch);
            }
        }

        let ext = extension.as_deref().unwrap_or_default();
        let is_xlsx_family = !ext.is_empty() && is_xlsx_family_extension(ext);
        let desired_workbook_content_type = (!ext.is_empty())
            .then_some(ext)
            .and_then(workbook_main_content_type_for_extension);

        let origin_has_vba = zip_part_exists(origin_bytes, "xl/vbaProject.bin");
        let needs_strip_vba = !ext.is_empty()
            && is_macro_free_xlsx_extension(ext)
            && (workbook.vba_project_bin.is_some()
                || origin_package_has_macro_content(origin_bytes));
        let needs_inject_vba = !ext.is_empty()
            && is_macro_enabled_xlsx_extension(ext)
            && workbook.vba_project_bin.is_some()
            && !origin_has_vba;

        let needs_workbook_content_type_update = match desired_workbook_content_type {
            Some(desired) => formula_xlsx::read_part_from_reader_limited(
                Cursor::new(origin_bytes),
                "[Content_Types].xml",
                XLSX_CONTENT_TYPES_MAX_BYTES,
            )
            .ok()
            .flatten()
            .and_then(|bytes| {
                std::str::from_utf8(&bytes)
                    .ok()
                    .map(|xml| !workbook_override_matches_content_type(xml, desired))
            })
            .unwrap_or(true),
            None => false,
        };

        let needs_date_system_update = is_xlsx_family
            && matches!(workbook.date_system, WorkbookDateSystem::Excel1904)
            && formula_xlsx::read_part_from_reader_limited(
                Cursor::new(origin_bytes),
                "xl/workbook.xml",
                XLSX_WORKBOOK_XML_MAX_BYTES,
            )
            .ok()
            .flatten()
            .and_then(|bytes| {
                std::str::from_utf8(&bytes)
                    .ok()
                    .map(|xml| !xml.contains("date1904=\"1\"") && !xml.contains("date1904='1'"))
            })
            .unwrap_or(true);

        let fast_path_possible = patches.is_empty()
            && !print_settings_changed
            && !power_query_changed
            && sheet_order_override.is_none()
            && !needs_strip_vba
            && !needs_inject_vba
            && !needs_workbook_content_type_update
            && !needs_date_system_update;

        if fast_path_possible {
            let part_overrides = sheet_metadata_part_overrides(origin_bytes, workbook)?;
            if part_overrides.is_empty() {
                return Ok(workbook
                    .origin_xlsx_bytes
                    .as_ref()
                    .expect("origin_xlsx_bytes should be Some when origin_bytes is Some")
                    .clone());
            }

            let empty_patches = WorkbookCellPatches::default();
            let mut cursor = Cursor::new(Vec::new());
            patch_xlsx_streaming_workbook_cell_patches_with_part_overrides(
                Cursor::new(origin_bytes),
                &mut cursor,
                &empty_patches,
                &part_overrides,
            )
            .context("patch sheet metadata (streaming)")?;
            let bytes = Arc::<[u8]>::from(cursor.into_inner());
            return Ok(bytes);
        }

        let mut part_overrides: HashMap<String, PartOverride> = HashMap::new();
        if let Some(xml) = sheet_order_override {
            part_overrides.insert("xl/workbook.xml".to_string(), PartOverride::Replace(xml));
        }
        if power_query_changed {
            match workbook.power_query_xml.as_ref() {
                Some(bytes) => {
                    let override_op = if workbook.original_power_query_xml.is_some() {
                        PartOverride::Replace(bytes.clone())
                    } else {
                        PartOverride::Add(bytes.clone())
                    };
                    part_overrides.insert(FORMULA_POWER_QUERY_PART.to_string(), override_op);
                }
                None => {
                    part_overrides
                        .insert(FORMULA_POWER_QUERY_PART.to_string(), PartOverride::Remove);
                }
            }
        }

        let mut bytes = if patches.is_empty() && part_overrides.is_empty() {
            origin_bytes.to_vec()
        } else {
            let mut cursor = Cursor::new(Vec::new());
            if part_overrides.is_empty() {
                patch_xlsx_streaming_workbook_cell_patches(
                    Cursor::new(origin_bytes),
                    &mut cursor,
                    &patches,
                )
                .context("apply worksheet cell patches (streaming)")?;
            } else {
                patch_xlsx_streaming_workbook_cell_patches_with_part_overrides(
                    Cursor::new(origin_bytes),
                    &mut cursor,
                    &patches,
                    &part_overrides,
                )
                .context("apply worksheet cell patches + part overrides (streaming)")?;
            }
            cursor.into_inner()
        };

        if needs_strip_vba {
            let mut stripped = Cursor::new(Vec::new());
            strip_vba_project_streaming(Cursor::new(bytes), &mut stripped)
                .context("strip VBA project (streaming)")?;
            bytes = stripped.into_inner();
        }

        if needs_inject_vba {
            // Apply macro injection via a streaming part-override rewrite. This avoids inflating
            // the entire package (which may contain very large worksheet XML parts) just to update
            // `[Content_Types].xml` / relationships / and add `vbaProject.bin`.
            bytes = patch_vba_and_date_system_in_package_streaming(
                &bytes,
                workbook,
                needs_inject_vba,
                needs_date_system_update,
                xlsx_date_system,
            )?;
        }

        if !needs_inject_vba && needs_date_system_update {
            // Apply date system updates via streaming part overrides for the same reason as VBA
            // injection: avoid loading large packages into memory unnecessarily.
            bytes = patch_vba_and_date_system_in_package_streaming(
                &bytes,
                workbook,
                needs_inject_vba,
                needs_date_system_update,
                xlsx_date_system,
            )?;
        }

        if print_settings_changed {
            bytes = write_workbook_print_settings(&bytes, &workbook.print_settings)
                .map_err(|e| anyhow::anyhow!(e.to_string()))?;
        }

        if let Some(desired) = desired_workbook_content_type {
            if let Some(updated) = patch_workbook_main_content_type_in_package(&bytes, desired)? {
                bytes = updated;
            }
        }

        if let Some(updated) = patch_sheet_metadata_in_package(&bytes, workbook)? {
            bytes = updated;
        }

        return Ok(Arc::<[u8]>::from(bytes));
    }

    let model = app_workbook_to_formula_model(workbook).context("convert workbook to model")?;
    let mut cursor = Cursor::new(Vec::new());
    formula_xlsx::write_workbook_to_writer_with_kind(&model, &mut cursor, workbook_kind)
        .map_err(|e| anyhow::anyhow!(e.to_string()))
        .with_context(|| "serialize workbook to buffer")?;
    let mut bytes = cursor.into_inner();
    let wants_vba = workbook.vba_project_bin.is_some()
        && extension
            .as_deref()
            .is_some_and(|ext| is_macro_enabled_xlsx_extension(ext));
    let wants_vba_signature = wants_vba && workbook.vba_project_signature_bin.is_some();
    let wants_preserved_drawings = workbook.preserved_drawing_parts.is_some();
    let wants_preserved_pivots = workbook.preserved_pivot_parts.is_some();
    let needs_date_system_update = extension
        .as_deref()
        .is_some_and(|ext| is_xlsx_family_extension(ext))
        && matches!(workbook.date_system, WorkbookDateSystem::Excel1904);
    let wants_power_query = workbook.power_query_xml.is_some();

    if wants_vba
        || wants_preserved_drawings
        || wants_preserved_pivots
        || wants_power_query
        || needs_date_system_update
    {
        let mut pkg = XlsxPackage::from_bytes_limited(&bytes, XlsxPackageLimits::default())
            .context("parse generated workbook package")?;

        if wants_vba {
            pkg.set_part(
                "xl/vbaProject.bin",
                workbook.vba_project_bin.clone().expect("checked is_some"),
            );
        }
        if wants_vba_signature {
            pkg.set_part(
                "xl/vbaProjectSignature.bin",
                workbook
                    .vba_project_signature_bin
                    .clone()
                    .expect("checked is_some"),
            );
        }

        if let Some(preserved) = workbook.preserved_drawing_parts.as_ref() {
            pkg.apply_preserved_drawing_parts(preserved)
                .context("apply preserved drawing/chart parts")?;
        }

        if let Some(preserved) = workbook.preserved_pivot_parts.as_ref() {
            pkg.apply_preserved_pivot_parts(preserved)
                .context("apply preserved pivot parts")?;
        }

        match workbook.power_query_xml.as_ref() {
            Some(bytes) => pkg.set_part(FORMULA_POWER_QUERY_PART, bytes.clone()),
            None => {
                pkg.parts_map_mut().remove(FORMULA_POWER_QUERY_PART);
            }
        }

        if needs_date_system_update {
            pkg.set_workbook_date_system(xlsx_date_system)
                .context("set workbook date system")?;
        }

        bytes = pkg
            .write_to_bytes()
            .context("repack workbook package with preserved parts")?;
    }

    if extension
        .as_deref()
        .is_some_and(|ext| is_xlsx_family_extension(ext))
    {
        bytes = write_workbook_print_settings(&bytes, &workbook.print_settings)
            .map_err(|e| anyhow::anyhow!(e.to_string()))?;
    }

    if let Some(desired) = extension
        .as_deref()
        .and_then(workbook_main_content_type_for_extension)
    {
        if let Some(updated) = patch_workbook_main_content_type_in_package(&bytes, desired)? {
            bytes = updated;
        }
    }

    Ok(Arc::<[u8]>::from(bytes))
}

pub fn write_xlsx_blocking(path: &Path, workbook: &Workbook) -> anyhow::Result<Arc<[u8]>> {
    let extension = path
        .extension()
        .and_then(|s| s.to_str())
        .map(|s| s.to_ascii_lowercase());

    if matches!(extension.as_deref(), Some("xlsb")) {
        return write_xlsb_blocking(path, workbook);
    }

    let bytes = build_xlsx_bytes_blocking(path, workbook)?;
    write_file_atomic(path, bytes.as_ref())
        .with_context(|| format!("write workbook {:?}", path))?;
    Ok(bytes)
}

/// Write an XLSB workbook to disk without reading the output file back into memory.
///
/// This is useful for the Tauri IPC save path: the app does not keep origin bytes for `.xlsb`
/// workbooks, so returning `Arc<[u8]>` would force an unnecessary full-file read (and can OOM on
/// large workbooks).
pub fn write_xlsb_to_disk_blocking(path: &Path, workbook: &Workbook) -> anyhow::Result<()> {
    write_xlsb_to_disk_impl(path, workbook)
}

fn write_xlsb_to_disk_impl(path: &Path, workbook: &Workbook) -> anyhow::Result<()> {
    let Some(origin_path) = workbook.origin_xlsb_path.as_deref() else {
        return Err(anyhow::anyhow!(
            "Saving as .xlsb is only supported for workbooks opened from an .xlsb file. Save As .xlsx instead."
        ));
    };
    let origin_path = Path::new(origin_path);

    let print_settings_changed = workbook.print_settings != workbook.original_print_settings;
    if print_settings_changed {
        anyhow::bail!("saving print settings to .xlsb is not supported yet");
    }

    let any_dirty_cells = workbook
        .sheets
        .iter()
        .any(|sheet| !sheet.dirty_cells.is_empty());

    let mut temp_paths: Vec<std::path::PathBuf> = Vec::new();
    let res = atomic_write_with_path(path, |tmp_out_path| -> anyhow::Result<()> {
        let final_out_path = tmp_out_path.to_path_buf();

        let xlsb = XlsbWorkbook::open_with_options(
            origin_path,
            XlsbOpenOptions {
                preserve_unknown_parts: false,
                preserve_parsed_parts: false,
                preserve_worksheets: false,
                decode_formulas: true,
            },
        )
        .with_context(|| format!("open xlsb {:?}", origin_path))?;

        if !any_dirty_cells {
            xlsb.save_as(&final_out_path)
                .with_context(|| format!("save xlsb {:?}", final_out_path))?;
            return Ok(());
        }

        let sheet_index_by_name: HashMap<String, usize> = xlsb
            .sheet_metas()
            .iter()
            .enumerate()
            .map(|(idx, meta)| (meta.name.clone(), idx))
            .collect();

        let mut edits_by_sheet: BTreeMap<usize, Vec<XlsbCellEdit>> = BTreeMap::new();
        for sheet in &workbook.sheets {
            if sheet.dirty_cells.is_empty() {
                continue;
            }

            let sheet_index = sheet_index_by_name
                .get(&sheet.name)
                .copied()
                .or(sheet.origin_ordinal)
                .with_context(|| {
                    format!(
                        "cannot map sheet {:?} to XLSB sheet index (no name match, no origin ordinal)",
                        sheet.name
                    )
                })?;
            if sheet_index >= xlsb.sheet_metas().len() {
                anyhow::bail!(
                    "sheet index {sheet_index} out of bounds for XLSB workbook ({} sheets)",
                    xlsb.sheet_metas().len()
                );
            }

            let edits = edits_by_sheet.entry(sheet_index).or_default();
            for (row, col) in &sheet.dirty_cells {
                let cell_key = (sheet.id.clone(), *row, *col);
                let (current_input, current_formula) = match sheet.cells.get(&(*row, *col)) {
                    Some(cell) => (cell.input_value.clone(), cell.formula.clone()),
                    None => (None, None),
                };

                let baseline = workbook.cell_input_baseline.get(&cell_key);
                let baseline_had_formula = baseline.map_or(false, |baseline| baseline.1.is_some());

                if let Some((baseline_value, baseline_formula)) = baseline {
                    if &current_input == baseline_value && &current_formula == baseline_formula {
                        continue;
                    }
                } else if current_input.is_none() && current_formula.is_none() {
                    // No baseline and no stored cell: avoid inserting explicit blank records for
                    // untouched cells (e.g. edit→revert cycles or formatting-only edits).
                    continue;
                }

                let (value, formula) = match sheet.cells.get(&(*row, *col)) {
                    Some(cell) => (
                        cell.input_value
                            .as_ref()
                            .unwrap_or(&cell.computed_value)
                            .clone(),
                        cell.formula.clone(),
                    ),
                    None => (CellScalar::Empty, None),
                };

                let row_u32 = u32::try_from(*row)
                    .with_context(|| format!("row index {row} is too large for XLSB"))?;
                let col_u32 = u32::try_from(*col)
                    .with_context(|| format!("col index {col} is too large for XLSB"))?;

                let new_value = scalar_to_xlsb_value(&value);
                let edit = match formula.as_deref() {
                    Some(formula) => {
                        let normalized = if formula.starts_with('=') {
                            formula.to_string()
                        } else {
                            format!("={formula}")
                        };
                        // Prefer the context-aware encoder so we can emit BIFF12 `rgcb` bytes for
                        // formulas that need them (e.g. array constants / PtgArray). Fall back to
                        // the older `formula_biff` encoder for compatibility.
                        let edit = XlsbCellEdit::with_formula_text_with_context(
                            row_u32,
                            col_u32,
                            new_value.clone(),
                            &normalized,
                            xlsb.workbook_context(),
                        );
                        let edit: anyhow::Result<XlsbCellEdit> = match edit {
                            Ok(edit) => Ok(edit),
                            Err(ctx_err) => XlsbCellEdit::with_formula_text(
                                row_u32,
                                col_u32,
                                new_value,
                                &normalized,
                            )
                            .map_err(|biff_err| {
                                anyhow::anyhow!(
                                    "cannot save .xlsb: unsupported formula edit at {}!({}, {}): {ctx_err}; fallback encoder also failed ({biff_err}). Save As .xlsx instead",
                                    sheet.name,
                                    *row + 1,
                                    *col + 1
                                )
                            }),
                        };
                        edit.with_context(|| {
                            format!("encode RGCE for formula cell at ({row}, {col})")
                        })?
                    }
                    None => XlsbCellEdit {
                        row: row_u32,
                        col: col_u32,
                        new_value,
                        new_style: None,
                        // When a cell has no formula in our model, ensure any existing formula
                        // record in the source XLSB is removed (\"paste values\" semantics).
                        clear_formula: baseline_had_formula,
                        new_formula: None,
                        new_rgcb: None,
                        new_formula_flags: None,
                        shared_string_index: None,
                    },
                };
                edits.push(edit);
            }
        }

        edits_by_sheet.retain(|_, edits| !edits.is_empty());
        if edits_by_sheet.is_empty() {
            xlsb.save_as(&final_out_path)
                .with_context(|| format!("save xlsb {:?}", final_out_path))?;
            return Ok(());
        }

        if edits_by_sheet.len() == 1 {
            let (&sheet_index, edits) = edits_by_sheet.iter().next().expect("non-empty map");
            let has_text_edits = edits.iter().any(|edit| {
                matches!(edit.new_value, XlsbCellValue::Text(_))
                    && edit.new_formula.is_none()
                    && edit.new_rgcb.is_none()
            });
            if has_text_edits {
                xlsb.save_with_cell_edits_streaming_shared_strings(
                    &final_out_path,
                    sheet_index,
                    edits,
                )
                .with_context(|| format!("save edited xlsb {:?}", final_out_path))?;
            } else {
                xlsb.save_with_cell_edits_streaming(&final_out_path, sheet_index, edits)
                    .with_context(|| format!("save edited xlsb {:?}", final_out_path))?;
            }
            return Ok(());
        }

        let has_text_edits = edits_by_sheet.values().flatten().any(|edit| {
            matches!(edit.new_value, XlsbCellValue::Text(_))
                && edit.new_formula.is_none()
                && edit.new_rgcb.is_none()
        });

        // Prefer a single-pass multi-sheet streaming save. Keep the older "patch through temp
        // workbooks" approach only as a fallback if the multi-sheet writer errors.
        let multi_res = if has_text_edits {
            xlsb.save_with_cell_edits_streaming_multi_shared_strings(
                &final_out_path,
                &edits_by_sheet,
            )
        } else {
            xlsb.save_with_cell_edits_streaming_multi(&final_out_path, &edits_by_sheet)
        };
        if multi_res
            .with_context(|| format!("save edited xlsb {:?}", final_out_path))
            .is_ok()
        {
            return Ok(());
        }

        // Multi-sheet streaming writer failed. Fall back to the older sequential writer so we can
        // still save (at the cost of additional ZIP rewrites).
        let dir = path.parent().unwrap_or_else(|| Path::new("."));
        let pid = std::process::id();
        let nanos = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .unwrap_or_default()
            .as_nanos();

        let mut source_path = origin_path.to_path_buf();
        for (step, (&sheet_index, sheet_edits)) in edits_by_sheet.iter().enumerate() {
            let is_last = step + 1 == edits_by_sheet.len();
            let out_path = if is_last {
                final_out_path.clone()
            } else {
                let mut candidate =
                    dir.join(format!(".formula-xlsb-save-{pid}-{nanos}-step{step}.xlsb"));
                let mut bump = 0u32;
                while candidate.exists() {
                    bump += 1;
                    candidate = dir.join(format!(
                        ".formula-xlsb-save-{pid}-{nanos}-step{step}-{bump}.xlsb"
                    ));
                }
                temp_paths.push(candidate.clone());
                candidate
            };

            let wb = XlsbWorkbook::open_with_options(
                &source_path,
                XlsbOpenOptions {
                    preserve_unknown_parts: false,
                    preserve_parsed_parts: false,
                    preserve_worksheets: false,
                    decode_formulas: true,
                },
            )
            .with_context(|| format!("open xlsb {:?}", source_path))?;
            let has_text_edits = sheet_edits.iter().any(|edit| {
                matches!(edit.new_value, XlsbCellValue::Text(_))
                    && edit.new_formula.is_none()
                    && edit.new_rgcb.is_none()
            });
            if has_text_edits {
                wb.save_with_cell_edits_streaming_shared_strings(
                    &out_path,
                    sheet_index,
                    sheet_edits,
                )
                .with_context(|| format!("save edited xlsb {:?}", out_path))?;
            } else {
                wb.save_with_cell_edits_streaming(&out_path, sheet_index, sheet_edits)
                    .with_context(|| format!("save edited xlsb {:?}", out_path))?;
            }

            source_path = out_path;
        }

        Ok(())
    });

    for tmp in &temp_paths {
        let _ = std::fs::remove_file(tmp);
    }

    match res {
        Ok(()) => Ok(()),
        Err(AtomicWriteError::Io(err)) => Err(Err::<(), _>(err)
            .with_context(|| format!("write xlsb {:?}", path))
            .unwrap_err()),
        Err(AtomicWriteError::Writer(err)) => Err(err),
    }
}

fn write_xlsb_blocking(path: &Path, workbook: &Workbook) -> anyhow::Result<Arc<[u8]>> {
    write_xlsb_to_disk_impl(path, workbook)?;
    let bytes =
        Arc::<[u8]>::from(std::fs::read(path).with_context(|| format!("read xlsb {:?}", path))?);
    Ok(bytes)
}

fn scalar_to_xlsb_value(value: &CellScalar) -> XlsbCellValue {
    match value {
        CellScalar::Empty => XlsbCellValue::Blank,
        CellScalar::Number(n) => XlsbCellValue::Number(*n),
        CellScalar::Bool(b) => XlsbCellValue::Bool(*b),
        CellScalar::Text(s) => XlsbCellValue::Text(s.clone()),
        CellScalar::Error(e) => XlsbCellValue::Error(xlsb_error_code(e)),
    }
}

fn xlsb_error_code(display: &str) -> u8 {
    match display.trim() {
        "#NULL!" => 0x00,
        "#DIV/0!" => 0x07,
        "#VALUE!" => 0x0F,
        "#REF!" => 0x17,
        "#NAME?" => 0x1D,
        "#NUM!" => 0x24,
        "#N/A" => 0x2A,
        "#GETTING_DATA" => 0x2B,
        other => {
            if let Some(inner) = other
                .strip_prefix("#ERR(")
                .and_then(|s| s.strip_suffix(')'))
                .map(str::trim)
            {
                if let Some(hex) = inner
                    .strip_prefix("0x")
                    .or_else(|| inner.strip_prefix("0X"))
                {
                    if let Ok(code) = u8::from_str_radix(hex, 16) {
                        return code;
                    }
                }
            }
            // Best-effort fallback.
            0x0F
        }
    }
}

fn scalar_to_model_value(value: &CellScalar) -> formula_model::CellValue {
    match value {
        CellScalar::Empty => formula_model::CellValue::Empty,
        CellScalar::Number(n) => formula_model::CellValue::Number(*n),
        CellScalar::Text(s) => formula_model::CellValue::String(s.clone()),
        CellScalar::Bool(b) => formula_model::CellValue::Boolean(*b),
        CellScalar::Error(e) => formula_model::CellValue::Error(
            e.parse::<formula_model::ErrorValue>()
                .unwrap_or(formula_model::ErrorValue::Unknown),
        ),
    }
}

fn app_workbook_to_formula_model(workbook: &Workbook) -> anyhow::Result<formula_model::Workbook> {
    let mut out = formula_model::Workbook::new();
    out.date_system = workbook.date_system;
    let mut number_format_style_ids: HashMap<String, u32> = HashMap::new();

    let mut sheet_id_by_app_id: HashMap<String, WorksheetId> = HashMap::new();
    let mut sheet_id_by_name: HashMap<String, WorksheetId> = HashMap::new();
    for sheet in &workbook.sheets {
        let sheet_id = out
            .add_sheet(sheet.name.clone())
            .map_err(|e| anyhow::anyhow!(e.to_string()))
            .with_context(|| format!("add sheet {}", sheet.name))?;
        if let Some(model_sheet) = out.sheet_mut(sheet_id) {
            model_sheet.visibility = sheet.visibility;
            model_sheet.tab_color = sheet.tab_color.clone();
            model_sheet.default_col_width = sheet.default_col_width;
            model_sheet.col_properties = sheet.col_properties.clone();
        }
        sheet_id_by_app_id.insert(sheet.id.clone(), sheet_id);
        sheet_id_by_name.insert(sheet.name.clone(), sheet_id);
    }

    for sheet in &workbook.sheets {
        let model_sheet_id = sheet_id_by_app_id
            .get(&sheet.id)
            .copied()
            .ok_or_else(|| anyhow::anyhow!("missing sheet id for {}", sheet.id))?;

        if let Some(columnar) = sheet.columnar.as_ref() {
            // Preserve columnar-backed worksheets without materializing the full dataset
            // into the sparse cell map. The XLSX writer can stream from the columnar
            // table, while `sheet.cells` acts as an overlay for edits/formulas.
            let model_sheet = out
                .sheet_mut(model_sheet_id)
                .ok_or_else(|| anyhow::anyhow!("sheet missing from model: {}", sheet.id))?;
            model_sheet.set_columnar_table(formula_model::CellRef::new(0, 0), columnar.clone());
        }

        for ((row, col), cell) in sheet.cells_iter() {
            let row_u32 =
                u32::try_from(row).map_err(|_| anyhow::anyhow!("row out of bounds: {row}"))?;
            let col_u32 =
                u32::try_from(col).map_err(|_| anyhow::anyhow!("col out of bounds: {col}"))?;
            let Some(cell_ref) = formula_model::CellRef::try_new(row_u32, col_u32) else {
                continue;
            };

            let (formula, scalar) = match cell.formula.as_ref() {
                Some(formula) => (Some(formula.clone()), cell.computed_value.clone()),
                None => (
                    None,
                    cell.input_value
                        .clone()
                        .unwrap_or_else(|| cell.computed_value.clone()),
                ),
            };

            let mut model_cell = formula_model::Cell::new(scalar_to_model_value(&scalar));
            model_cell.formula = formula;
            if let Some(fmt) = cell
                .number_format
                .as_deref()
                .map(str::trim)
                .filter(|s| !s.is_empty())
            {
                if let Some(existing) = number_format_style_ids.get(fmt) {
                    model_cell.style_id = *existing;
                } else {
                    let fmt = fmt.to_string();
                    let style_id = out.intern_style(formula_model::Style {
                        number_format: Some(fmt.clone()),
                        ..Default::default()
                    });
                    number_format_style_ids.insert(fmt, style_id);
                    model_cell.style_id = style_id;
                }
            }

            let model_sheet = out
                .sheet_mut(model_sheet_id)
                .ok_or_else(|| anyhow::anyhow!("sheet missing from model: {}", sheet.id))?;
            model_sheet.set_cell(cell_ref, model_cell);
        }
    }

    for defined in &workbook.defined_names {
        let scope = match defined.sheet_id.as_deref() {
            None => formula_model::DefinedNameScope::Workbook,
            Some(sheet_key) => {
                let sheet_id = sheet_id_by_app_id
                    .get(sheet_key)
                    .or_else(|| sheet_id_by_name.get(sheet_key))
                    .copied()
                    .ok_or_else(|| {
                        anyhow::anyhow!(
                            "defined name {} references unknown sheet {}",
                            defined.name,
                            sheet_key
                        )
                    })?;
                formula_model::DefinedNameScope::Sheet(sheet_id)
            }
        };

        out.create_defined_name(
            scope,
            defined.name.clone(),
            defined.refers_to.clone(),
            None,
            defined.hidden,
            None,
        )
        .map_err(|e| anyhow::anyhow!(e.to_string()))
        .with_context(|| format!("create defined name {}", defined.name))?;
    }

    let mut next_table_id: u32 = 1;
    for table in &workbook.tables {
        let sheet_id = sheet_id_by_app_id
            .get(&table.sheet_id)
            .or_else(|| sheet_id_by_name.get(&table.sheet_id))
            .copied()
            .ok_or_else(|| {
                anyhow::anyhow!(
                    "table {} references unknown sheet {}",
                    table.name,
                    table.sheet_id
                )
            })?;

        let start_row = u32::try_from(table.start_row)
            .map_err(|_| anyhow::anyhow!("table start_row out of bounds: {}", table.start_row))?;
        let start_col = u32::try_from(table.start_col)
            .map_err(|_| anyhow::anyhow!("table start_col out of bounds: {}", table.start_col))?;
        let end_row = u32::try_from(table.end_row)
            .map_err(|_| anyhow::anyhow!("table end_row out of bounds: {}", table.end_row))?;
        let end_col = u32::try_from(table.end_col)
            .map_err(|_| anyhow::anyhow!("table end_col out of bounds: {}", table.end_col))?;

        let columns = table
            .columns
            .iter()
            .enumerate()
            .map(|(idx, name)| formula_model::TableColumn {
                id: (idx + 1) as u32,
                name: name.clone(),
                formula: None,
                totals_formula: None,
            })
            .collect();

        let model_table = formula_model::Table {
            id: next_table_id,
            name: table.name.clone(),
            display_name: table.name.clone(),
            range: formula_model::Range::new(
                formula_model::CellRef::new(start_row, start_col),
                formula_model::CellRef::new(end_row, end_col),
            ),
            header_row_count: 1,
            totals_row_count: 0,
            columns,
            style: None,
            auto_filter: None,
            relationship_id: None,
            part_path: None,
        };
        next_table_id = next_table_id.wrapping_add(1);

        out.add_table(sheet_id, model_table)
            .map_err(|e| anyhow::anyhow!(e.to_string()))
            .with_context(|| format!("add table {}", table.name))?;
    }

    Ok(out)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::state::AppState;
    use formula_format::{format_value, FormatOptions, Value as FormatValue};
    use formula_xlsb::biff12_varint;
    use std::collections::BTreeSet;
    use std::io::{Cursor, Read, Write};
    use xlsx_diff::{
        diff_workbooks, diff_workbooks_with_options, DiffOptions, Severity, WorkbookArchive,
    };
    use zip::write::FileOptions;
    use zip::{CompressionMethod, ZipArchive, ZipWriter};

    fn rewrite_zip_with_leading_slash_entry_names(bytes: &[u8]) -> Vec<u8> {
        let mut input = ZipArchive::new(Cursor::new(bytes)).expect("open zip");

        let mut output = ZipWriter::new(Cursor::new(Vec::<u8>::new()));
        let options = FileOptions::<()>::default().compression_method(CompressionMethod::Stored);

        for i in 0..input.len() {
            let mut entry = input.by_index(i).expect("open zip entry");
            let name = entry.name().to_string();
            let new_name = if name.starts_with('/') {
                name
            } else {
                format!("/{name}")
            };

            // Do not trust `ZipFile::size()` for allocation; ZIP metadata is untrusted and can
            // advertise enormous uncompressed sizes (zip-bomb style OOM).
            let mut contents = Vec::new();
            entry.read_to_end(&mut contents).expect("read zip entry");

            if entry.is_dir() {
                output
                    .add_directory(new_name, options)
                    .expect("write directory");
            } else {
                output.start_file(new_name, options).expect("start file");
                output.write_all(&contents).expect("write file");
            }
        }

        output.finish().expect("finish zip").into_inner()
    }

    fn upsert_zip_entry(bytes: &[u8], target: &str, replacement: &[u8]) -> Vec<u8> {
        let mut input = ZipArchive::new(Cursor::new(bytes)).expect("open zip");

        let mut output = ZipWriter::new(Cursor::new(Vec::<u8>::new()));
        let options = FileOptions::<()>::default().compression_method(CompressionMethod::Stored);

        let mut wrote_target = false;
        for i in 0..input.len() {
            let mut entry = input.by_index(i).expect("open zip entry");
            let name = entry.name().to_string();

            if entry.is_dir() {
                output
                    .add_directory(name, options)
                    .expect("write directory");
                continue;
            }

            output.start_file(name.clone(), options).expect("start file");
            if zip_entry_name_matches(&name, target) {
                output.write_all(replacement).expect("write replacement");
                wrote_target = true;
            } else {
                std::io::copy(&mut entry, &mut output).expect("copy zip entry");
            }
        }

        if !wrote_target {
            output
                .start_file(target, options)
                .expect("start missing target file");
            output.write_all(replacement).expect("write missing target");
        }

        output.finish().expect("finish zip").into_inner()
    }

    // Patch a ZIP file's uncompressed size metadata for `entry_name` in both the central
    // directory header and the local file header.
    //
    // This lets us create a tiny ZIP that *claims* one entry inflates beyond our package limits,
    // without allocating huge buffers in the test.
    fn patch_zip_entry_uncompressed_size(
        mut zip_bytes: Vec<u8>,
        entry_name: &str,
        new_uncompressed_size: u32,
    ) -> Vec<u8> {
        // Locate the end-of-central-directory record (EOCD) by scanning backwards from the end of
        // the file. The ZIP spec allows up to 64KiB of trailing comment.
        const EOCD_SIG: [u8; 4] = [0x50, 0x4B, 0x05, 0x06];
        let min_eocd = zip_bytes.len().saturating_sub(22);
        let search_min = zip_bytes.len().saturating_sub(22 + 0xFFFF);

        let mut eocd_offset = None;
        for i in (search_min..=min_eocd).rev() {
            if zip_bytes.get(i..i + 4) == Some(&EOCD_SIG) {
                eocd_offset = Some(i);
                break;
            }
        }
        let eocd_offset = eocd_offset.expect("expected EOCD record in test zip");

        let central_dir_size =
            u32::from_le_bytes(zip_bytes[eocd_offset + 12..eocd_offset + 16].try_into().unwrap())
                as usize;
        let central_dir_offset =
            u32::from_le_bytes(zip_bytes[eocd_offset + 16..eocd_offset + 20].try_into().unwrap())
                as usize;

        const CEN_SIG: [u8; 4] = [0x50, 0x4B, 0x01, 0x02];
        const LFH_SIG: [u8; 4] = [0x50, 0x4B, 0x03, 0x04];
        let mut cursor = central_dir_offset;
        let end = central_dir_offset + central_dir_size;
        while cursor < end {
            assert_eq!(
                zip_bytes.get(cursor..cursor + 4),
                Some(CEN_SIG.as_slice()),
                "expected central directory header signature"
            );

            let name_len =
                u16::from_le_bytes(zip_bytes[cursor + 28..cursor + 30].try_into().unwrap()) as usize;
            let extra_len =
                u16::from_le_bytes(zip_bytes[cursor + 30..cursor + 32].try_into().unwrap()) as usize;
            let comment_len =
                u16::from_le_bytes(zip_bytes[cursor + 32..cursor + 34].try_into().unwrap()) as usize;
            let local_header_offset =
                u32::from_le_bytes(zip_bytes[cursor + 42..cursor + 46].try_into().unwrap()) as usize;

            let name_start = cursor + 46;
            let name_end = name_start + name_len;
            let name =
                std::str::from_utf8(&zip_bytes[name_start..name_end]).expect("expected utf-8 name");

            if name == entry_name {
                // Patch central directory header's uncompressed size (offset 24, 4 bytes).
                zip_bytes[cursor + 24..cursor + 28]
                    .copy_from_slice(&new_uncompressed_size.to_le_bytes());

                // Patch local file header's uncompressed size too (offset 22, 4 bytes).
                assert_eq!(
                    zip_bytes.get(local_header_offset..local_header_offset + 4),
                    Some(LFH_SIG.as_slice()),
                    "expected local file header signature"
                );
                zip_bytes[local_header_offset + 22..local_header_offset + 26]
                    .copy_from_slice(&new_uncompressed_size.to_le_bytes());
                return zip_bytes;
            }

            cursor += 46 + name_len + extra_len + comment_len;
        }

        panic!("test zip did not contain expected entry: {entry_name}");
    }

    fn assert_no_critical_diffs(expected: &Path, actual: &Path) {
        let report = diff_workbooks(expected, actual).expect("diff workbooks");
        let critical = report.count(Severity::Critical);
        if critical == 0 {
            return;
        }

        let mut details = String::new();
        for diff in report
            .differences
            .iter()
            .filter(|d| d.severity == Severity::Critical)
        {
            details.push_str(&diff.to_string());
            details.push('\n');
        }

        panic!(
            "expected no CRITICAL diffs between {} and {}, found {critical}:\n{details}",
            expected.display(),
            actual.display()
        );
    }

    fn assert_non_worksheet_parts_preserved(original: &[u8], written: &[u8]) {
        let original_pkg = XlsxPackage::from_bytes(original).expect("parse original package");
        let written_pkg = XlsxPackage::from_bytes(written).expect("parse written package");

        let original_names: HashSet<String> =
            original_pkg.part_names().map(str::to_owned).collect();
        let written_names: HashSet<String> = written_pkg.part_names().map(str::to_owned).collect();
        assert_eq!(
            original_names, written_names,
            "expected no parts to be added/removed when patching worksheets"
        );

        for (name, bytes) in original_pkg.parts() {
            let is_worksheet_xml =
                name.starts_with("xl/worksheets/") && !name.starts_with("xl/worksheets/_rels/");
            if is_worksheet_xml {
                continue;
            }
            assert_eq!(
                Some(bytes),
                written_pkg.part(name),
                "expected part {name} to be preserved byte-for-byte"
            );
        }
    }

    #[test]
    fn sniff_workbook_format_tolerates_leading_slash_zip_entry_names() {
        let tmp = tempfile::tempdir().expect("temp dir");

        let xlsx_fixture = concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../../fixtures/xlsx/basic/basic.xlsx"
        );
        let xlsx_bytes = std::fs::read(xlsx_fixture).expect("read xlsx fixture");
        let rewritten = rewrite_zip_with_leading_slash_entry_names(&xlsx_bytes);
        let xlsx_path = tmp.path().join("leading_slash.xlsx");
        std::fs::write(&xlsx_path, rewritten).expect("write rewritten xlsx");
        assert_eq!(
            sniff_workbook_format(&xlsx_path),
            Some(SniffedWorkbookFormat::Xlsx)
        );

        let xlsb_fixture = concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../../crates/formula-xlsb/tests/fixtures/simple.xlsb"
        );
        let xlsb_bytes = std::fs::read(xlsb_fixture).expect("read xlsb fixture");
        let rewritten = rewrite_zip_with_leading_slash_entry_names(&xlsb_bytes);
        let xlsb_path = tmp.path().join("leading_slash.xlsb");
        std::fs::write(&xlsb_path, rewritten).expect("write rewritten xlsb");
        assert_eq!(
            sniff_workbook_format(&xlsb_path),
            Some(SniffedWorkbookFormat::Xlsb)
        );
    }

    #[test]
    fn read_xlsb_blocking_opens_fixture_and_reads_cells() {
        let fixture_path = concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../../crates/formula-xlsb/tests/fixtures/simple.xlsb"
        );
        let workbook = read_xlsb_blocking(Path::new(fixture_path)).expect("read xlsb workbook");

        assert_eq!(workbook.origin_xlsb_path.as_deref(), Some(fixture_path));
        assert_eq!(workbook.date_system, WorkbookDateSystem::Excel1900);
        assert_eq!(workbook.sheets.len(), 1);
        assert_eq!(workbook.sheets[0].name, "Sheet1");

        let sheet = workbook.sheet("Sheet1").expect("Sheet1 present");

        assert_eq!(
            sheet.get_cell(0, 0).computed_value,
            CellScalar::Text("Hello".to_string())
        );
        assert_eq!(
            sheet.get_cell(0, 1).computed_value,
            CellScalar::Number(42.5)
        );

        let formula_cell = sheet.get_cell(0, 2);
        assert_eq!(formula_cell.formula.as_deref(), Some("=B1*2"));
        assert_eq!(formula_cell.computed_value, CellScalar::Number(85.0));
    }

    #[test]
    fn xlsb_column_properties_propagate_to_cell_width() {
        const SHEETDATA_RECORD_ID: u32 = 0x0091;
        const BEGIN_COL_INFOS_RECORD_ID: u32 = 0x0186;
        const END_COL_INFOS_RECORD_ID: u32 = 0x0187;
        const COL_INFO_RECORD_ID: u32 = 0x003C;

        fn write_record(out: &mut Vec<u8>, id: u32, payload: &[u8]) {
            biff12_varint::write_record_id(out, id).expect("write record id");
            biff12_varint::write_record_len(out, payload.len() as u32).expect("write record len");
            out.extend_from_slice(payload);
        }

        fn col_info_payload(col_first: u32, col_last: u32, width_chars: f32, hidden: bool) -> Vec<u8> {
            // MS-XLSB BrtColInfo (common layout):
            // [colFirst: u32][colLast: u32][width: u32 (1/256 chars)][xf: u32][grbit: u16][reserved: u16]
            let width_raw = (width_chars * 256.0) as u32;
            let mut payload = Vec::new();
            payload.extend_from_slice(&col_first.to_le_bytes());
            payload.extend_from_slice(&col_last.to_le_bytes());
            payload.extend_from_slice(&width_raw.to_le_bytes());
            payload.extend_from_slice(&0u32.to_le_bytes()); // xf/style
            let options: u16 = if hidden { 0x0001 } else { 0 };
            payload.extend_from_slice(&options.to_le_bytes());
            payload.extend_from_slice(&0u16.to_le_bytes()); // reserved
            payload
        }

        fn insert_col_infos(sheet_bin: &[u8]) -> Vec<u8> {
            let mut cursor = Cursor::new(sheet_bin);
            loop {
                let record_start = cursor.position() as usize;
                let Some(id) = biff12_varint::read_record_id(&mut cursor)
                    .expect("read record id")
                else {
                    panic!("expected to find BrtBeginSheetData");
                };
                let Some(len) = biff12_varint::read_record_len(&mut cursor)
                    .expect("read record len")
                else {
                    panic!("expected to read record len");
                };

                if id == SHEETDATA_RECORD_ID {
                    let mut patch = Vec::new();
                    write_record(&mut patch, BEGIN_COL_INFOS_RECORD_ID, &[]);
                    // Column A: width override 20.
                    write_record(
                        &mut patch,
                        COL_INFO_RECORD_ID,
                        &col_info_payload(0, 0, 20.0, false),
                    );
                    // Column B: hidden, then later unhidden with a different width. The later record
                    // should win (record order matters).
                    write_record(
                        &mut patch,
                        COL_INFO_RECORD_ID,
                        &col_info_payload(1, 1, 8.0, true),
                    );
                    write_record(
                        &mut patch,
                        COL_INFO_RECORD_ID,
                        &col_info_payload(1, 1, 15.0, false),
                    );
                    // Column C: hidden.
                    write_record(
                        &mut patch,
                        COL_INFO_RECORD_ID,
                        &col_info_payload(2, 2, 8.0, true),
                    );
                    write_record(&mut patch, END_COL_INFOS_RECORD_ID, &[]);

                    let mut out = Vec::with_capacity(sheet_bin.len() + patch.len());
                    out.extend_from_slice(&sheet_bin[..record_start]);
                    out.extend_from_slice(&patch);
                    out.extend_from_slice(&sheet_bin[record_start..]);
                    return out;
                }

                cursor.set_position(cursor.position() + len as u64);
            }
        }

        let fixture_path = concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../../crates/formula-xlsb/tests/fixtures/simple.xlsb"
        );
        let fixture_bytes = std::fs::read(fixture_path).expect("read xlsb fixture bytes");

        // Patch the worksheet stream to include BrtColInfo records.
        let mut input = ZipArchive::new(Cursor::new(fixture_bytes)).expect("open xlsb zip");
        let mut output = ZipWriter::new(Cursor::new(Vec::<u8>::new()));
        let options = FileOptions::<()>::default().compression_method(CompressionMethod::Stored);

        for i in 0..input.len() {
            let mut entry = input.by_index(i).expect("open zip entry");
            let name = entry.name().to_string();

            // Avoid pre-allocating based on attacker-controlled ZIP metadata.
            let mut contents = Vec::new();
            entry.read_to_end(&mut contents).expect("read zip entry");

            let contents = if zip_entry_name_matches(&name, "xl/worksheets/sheet1.bin") {
                insert_col_infos(&contents)
            } else {
                contents
            };

            if entry.is_dir() {
                output
                    .add_directory(name, options)
                    .expect("write directory");
            } else {
                output.start_file(name, options).expect("start file");
                output.write_all(&contents).expect("write file");
            }
        }

        let patched_bytes = output.finish().expect("finish zip").into_inner();

        let mut tmp = tempfile::Builder::new()
            .suffix(".xlsb")
            .tempfile()
            .expect("temp xlsb");
        tmp.write_all(&patched_bytes).expect("write patched xlsb");
        tmp.flush().expect("flush patched xlsb");

        let workbook = read_xlsb_blocking(tmp.path()).expect("read patched xlsb workbook");
        let sheet_id = workbook.sheets[0].id.clone();

        let mut state = AppState::new();
        state.load_workbook(workbook);

        // Column A width override is propagated into the engine.
        state
            .set_cell(
                &sheet_id,
                0,
                2,
                None,
                Some("=CELL(\"width\",A1)".to_string()),
            )
            .expect("set CELL(width) formula for A1");
        let c1 = state.get_cell(&sheet_id, 0, 2).expect("read C1");
        // Excel's CELL("width") uses the first decimal digit as a "custom width" flag.
        match c1.value {
            CellScalar::Number(v) => assert!(
                (v - 20.1).abs() < 0.2,
                "expected column width ~20 for A1, got {v}"
            ),
            other => panic!("expected numeric column width, got {other:?}"),
        }

        // Column B was unhidden by a later record and should report its final width.
        state
            .set_cell(
                &sheet_id,
                1,
                2,
                None,
                Some("=CELL(\"width\",B1)".to_string()),
            )
            .expect("set CELL(width) formula for B1");
        let c2 = state.get_cell(&sheet_id, 1, 2).expect("read C2");
        match c2.value {
            CellScalar::Number(v) => assert!(
                (v - 15.1).abs() < 0.2,
                "expected column width ~15 for B1, got {v}"
            ),
            other => panic!("expected numeric column width, got {other:?}"),
        }

        // Hidden columns report width=0.
        state
            .set_cell(
                &sheet_id,
                2,
                2,
                None,
                Some("=CELL(\"width\",C1)".to_string()),
            )
            .expect("set CELL(width) formula for C1");
        let c3 = state.get_cell(&sheet_id, 2, 2).expect("read C3");
        assert_eq!(c3.value, CellScalar::Number(0.0));
    }

    #[test]
    fn write_xlsb_to_disk_blocking_roundtrips_fixture() {
        let tmp = tempfile::tempdir().expect("temp dir");
        let fixture_path = concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../../crates/formula-xlsb/tests/fixtures/simple.xlsb"
        );
        let workbook = read_xlsb_blocking(Path::new(fixture_path)).expect("read xlsb workbook");

        let out_path = tmp.path().join("roundtrip.xlsb");
        write_xlsb_to_disk_blocking(&out_path, &workbook).expect("write xlsb");

        let meta = std::fs::metadata(&out_path).expect("metadata");
        assert!(meta.len() > 0, "expected non-empty xlsb output file");

        let mut file = std::fs::File::open(&out_path).expect("open xlsb");
        let mut signature = [0u8; 2];
        file.read_exact(&mut signature).expect("read signature");
        assert_eq!(&signature, b"PK");
    }

    #[test]
    fn zip_part_exists_tolerates_leading_slash_zip_entry_names() {
        let fixture = concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../../fixtures/xlsx/macros/basic.xlsm"
        );
        let bytes = std::fs::read(fixture).expect("read xlsm fixture");
        let rewritten = rewrite_zip_with_leading_slash_entry_names(&bytes);
        assert!(
            zip_part_exists(&rewritten, "xl/vbaProject.bin"),
            "expected xl/vbaProject.bin to be discovered even when ZIP entry names have leading '/'"
        );
    }

    #[test]
    fn read_xlsx_or_xlsm_drops_oversized_vba_project_bin() {
        let fixture = concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../../fixtures/xlsx/macros/basic.xlsm"
        );
        let bytes = std::fs::read(fixture).expect("read xlsm fixture");

        let vba_limit = 64u64;
        let oversized_vba = vec![0u8; (vba_limit + 1) as usize];
        let rewritten = upsert_zip_entry(&bytes, "xl/vbaProject.bin", &oversized_vba);

        let tmp = tempfile::tempdir().expect("temp dir");
        let workbook_path = tmp.path().join("oversized.xlsm");
        std::fs::write(&workbook_path, rewritten).expect("write workbook");

        let workbook = read_xlsx_or_xlsm_blocking_with_limits(
            &workbook_path,
            vba_limit,
            1024 * 1024,
            1024 * 1024,
        )
        .expect("open workbook");

        assert!(
            workbook.vba_project_bin.is_none(),
            "expected oversized VBA project bin to be dropped"
        );
        assert!(
            workbook.macro_fingerprint.is_none(),
            "expected macro fingerprint to be absent when VBA project is absent"
        );
    }

    #[test]
    fn read_xlsx_or_xlsm_drops_oversized_power_query_xml() {
        let fixture = concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../../fixtures/xlsx/macros/basic.xlsm"
        );
        let bytes = std::fs::read(fixture).expect("read xlsm fixture");

        let power_query_limit = 64u64;
        let oversized_xml = vec![b'a'; (power_query_limit + 1) as usize];
        let rewritten = upsert_zip_entry(&bytes, FORMULA_POWER_QUERY_PART, &oversized_xml);

        let tmp = tempfile::tempdir().expect("temp dir");
        let workbook_path = tmp.path().join("oversized-power-query.xlsm");
        std::fs::write(&workbook_path, rewritten).expect("write workbook");

        let workbook = read_xlsx_or_xlsm_blocking_with_limits(
            &workbook_path,
            1024 * 1024,
            1024 * 1024,
            power_query_limit,
        )
        .expect("open workbook");

        assert!(
            workbook.power_query_xml.is_none(),
            "expected oversized Power Query XML to be dropped"
        );
        assert!(
            workbook.original_power_query_xml.is_none(),
            "expected original Power Query XML to be unset when part was dropped"
        );
    }

    #[test]
    fn read_xlsx_or_xlsm_drops_oversized_vba_project_signature_bin() {
        let fixture = concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../../fixtures/xlsx/macros/basic.xlsm"
        );
        let bytes = std::fs::read(fixture).expect("read xlsm fixture");

        let signature_limit = 64u64;
        let oversized_sig = vec![0u8; (signature_limit + 1) as usize];
        let rewritten = upsert_zip_entry(&bytes, "xl/vbaProjectSignature.bin", &oversized_sig);

        let tmp = tempfile::tempdir().expect("temp dir");
        let workbook_path = tmp.path().join("oversized-signature.xlsm");
        std::fs::write(&workbook_path, rewritten).expect("write workbook");

        let workbook = read_xlsx_or_xlsm_blocking_with_limits(
            &workbook_path,
            1024 * 1024,
            signature_limit,
            1024 * 1024,
        )
        .expect("open workbook");

        assert!(
            workbook.vba_project_bin.is_some(),
            "expected VBA project to remain present"
        );
        assert!(
            workbook.vba_project_signature_bin.is_none(),
            "expected oversized VBA signature part to be dropped"
        );
        assert!(
            workbook.macro_fingerprint.is_some(),
            "expected macro fingerprint to still be computed when VBA project is present"
        );
    }

    #[test]
    fn read_xlsx_drops_oversized_preserved_drawing_parts() {
        let fixture = concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../../fixtures/xlsx/basic/image.xlsx"
        );
        let bytes = std::fs::read(fixture).expect("read xlsx fixture");

        // Forge the drawing part's ZIP metadata to claim it exceeds the desktop preservation cap.
        // This should cause preservation to fail fast (best-effort) without impacting workbook open.
        let oversized_len = crate::resource_limits::MAX_PRESERVED_DRAWING_PART_BYTES as u32 + 1;
        let patched =
            patch_zip_entry_uncompressed_size(bytes, "xl/drawings/drawing1.xml", oversized_len);

        // Sanity check: the default (library) preservation limits are larger, so preservation should
        // still succeed on this patched file.
        let preserved = formula_xlsx::drawingml::preserve_drawing_parts_from_reader(Cursor::new(
            patched.as_slice(),
        ))
        .expect("preserve drawing parts with default limits");
        assert!(
            !preserved.is_empty(),
            "expected drawing preservation to find parts in fixture"
        );

        let tmp = tempfile::tempdir().expect("temp dir");
        let path = tmp.path().join("oversized-drawing.xlsx");
        std::fs::write(&path, &patched).expect("write patched workbook");

        let workbook = read_xlsx_blocking(&path).expect("open patched workbook");
        assert!(
            workbook.preserved_drawing_parts.is_none(),
            "expected oversized drawing parts to be dropped during open"
        );
    }

    #[test]
    fn read_xlsx_drops_oversized_preserved_pivot_parts() {
        let fixture = concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../../fixtures/xlsx/pivots/pivot-fixture.xlsx"
        );
        let bytes = std::fs::read(fixture).expect("read pivot xlsx fixture");

        // Forge pivot cache records metadata to exceed the preservation cap.
        let oversized_len = crate::resource_limits::MAX_PRESERVED_PIVOT_PART_BYTES as u32 + 1;
        let patched = patch_zip_entry_uncompressed_size(
            bytes,
            "xl/pivotCache/pivotCacheRecords1.xml",
            oversized_len,
        );

        // Sanity check: default limits should still allow preservation to succeed.
        let preserved = formula_xlsx::pivots::preserve_pivot_parts_from_reader(Cursor::new(
            patched.as_slice(),
        ))
        .expect("preserve pivot parts with default limits");
        assert!(
            !preserved.is_empty(),
            "expected pivot preservation to find parts in fixture"
        );

        let tmp = tempfile::tempdir().expect("temp dir");
        let path = tmp.path().join("oversized-pivots.xlsx");
        std::fs::write(&path, &patched).expect("write patched workbook");

        let workbook = read_xlsx_blocking(&path).expect("open patched workbook");
        assert!(
            workbook.preserved_pivot_parts.is_none(),
            "expected oversized pivot parts to be dropped during open"
        );
    }

    #[test]
    fn read_xlsx_print_settings_fall_back_to_model_when_print_extractor_fails() {
        let fixture_path = Path::new(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../../fixtures/xlsx/basic/print-settings.xlsx"
        ));
        let bytes = std::fs::read(fixture_path).expect("read print settings fixture");

        // Corrupt the Print_Area defined name so `read_workbook_print_settings` fails (it requires
        // all print-related defined names to parse), while the main XLSX reader can still open the
        // workbook and salvage the remaining print settings best-effort.
        let workbook_xml = formula_xlsx::read_part_from_reader_limited(
            Cursor::new(bytes.as_slice()),
            "xl/workbook.xml",
            XLSX_WORKBOOK_XML_MAX_BYTES,
        )
        .expect("read xl/workbook.xml")
        .expect("expected xl/workbook.xml to exist");
        let mut workbook_xml =
            String::from_utf8(workbook_xml).expect("workbook.xml should be valid utf-8");
        workbook_xml = workbook_xml.replace(
            "Sheet1!$A$1:$D$10",
            "NotARef",
        );
        let rewritten = upsert_zip_entry(&bytes, "xl/workbook.xml", workbook_xml.as_bytes());
        assert!(
            read_workbook_print_settings(&rewritten).is_err(),
            "expected print-settings extractor to fail when a print-related defined name is invalid"
        );

        let tmp = tempfile::tempdir().expect("temp dir");
        let path = tmp.path().join("invalid-print-area.xlsx");
        std::fs::write(&path, &rewritten).expect("write patched workbook");

        let workbook = read_xlsx_blocking(&path).expect("open workbook with invalid print area");
        assert_eq!(workbook.print_settings.sheets.len(), 1);
        let sheet = &workbook.print_settings.sheets[0];
        assert_eq!(sheet.sheet_name, "Sheet1");
        assert!(sheet.print_area.is_none(), "expected invalid print area to be dropped");
        assert_eq!(
            sheet.print_titles,
            Some(formula_xlsx::print::PrintTitles {
                repeat_rows: Some(formula_xlsx::print::RowRange { start: 1, end: 1 }),
                repeat_cols: Some(formula_xlsx::print::ColRange { start: 1, end: 2 }),
            })
        );
        assert_eq!(sheet.page_setup.orientation, formula_xlsx::print::Orientation::Landscape);
    }

    #[test]
    fn patch_save_rewrites_workbook_xml_sheet_order_when_reordered() {
        let tmp = tempfile::tempdir().expect("temp dir");
        let origin_path = tmp.path().join("origin.xlsx");

        // Build a simple 3-sheet XLSX so `write_xlsx_blocking` takes the patch-based save path.
        let mut model = formula_model::Workbook::new();
        model.add_sheet("Sheet1").expect("add Sheet1");
        model.add_sheet("Sheet2").expect("add Sheet2");
        model.add_sheet("Sheet3").expect("add Sheet3");

        let mut cursor = Cursor::new(Vec::new());
        formula_xlsx::write_workbook_to_writer(&model, &mut cursor).expect("write xlsx bytes");
        std::fs::write(&origin_path, cursor.into_inner()).expect("write origin xlsx");

        let mut workbook = read_xlsx_blocking(&origin_path).expect("read origin xlsx");
        assert!(
            workbook.origin_xlsx_bytes.is_some(),
            "expected origin bytes baseline"
        );
        assert_eq!(
            workbook
                .sheets
                .iter()
                .map(|s| s.name.as_str())
                .collect::<Vec<_>>(),
            vec!["Sheet1", "Sheet2", "Sheet3"]
        );

        // Reorder: move Sheet3 to the front.
        let sheet = workbook.sheets.remove(2);
        workbook.sheets.insert(0, sheet);

        let out_path = tmp.path().join("reordered.xlsx");
        let written = write_xlsx_blocking(&out_path, &workbook).expect("write reordered xlsx");

        let parts = formula_xlsx::worksheet_parts_from_reader(Cursor::new(written.as_ref()))
            .expect("read worksheet parts");
        let names: Vec<String> = parts.iter().map(|p| p.name.clone()).collect();
        assert_eq!(names, vec!["Sheet3", "Sheet1", "Sheet2"]);
    }

    #[test]
    fn patch_save_can_inject_vba_without_inflating_oversized_parts() {
        let tmp = tempfile::tempdir().expect("temp dir");
        let origin_path = tmp.path().join("origin.xlsx");
        let out_path = tmp.path().join("saved.xlsm");

        // Build a small macro-free XLSX so `write_xlsx_blocking` takes the patch-based save path.
        let mut model = formula_model::Workbook::new();
        model.add_sheet("Sheet1").expect("add Sheet1");

        let mut cursor = Cursor::new(Vec::new());
        formula_xlsx::write_workbook_to_writer(&model, &mut cursor).expect("write xlsx bytes");
        std::fs::write(&origin_path, cursor.into_inner()).expect("write origin xlsx");

        let mut workbook = read_xlsx_blocking(&origin_path).expect("read origin xlsx");
        let origin_bytes = workbook
            .origin_xlsx_bytes
            .as_ref()
            .expect("expected origin bytes baseline");

        // Forge the ZIP metadata to claim the worksheet XML inflates beyond `XlsxPackage`'s
        // per-part limit. The patch-based save path should still be able to inject VBA without
        // inflating the worksheet part into memory.
        let oversized_len = formula_xlsx::MAX_XLSX_PACKAGE_PART_BYTES as u32 + 1;
        let patched =
            patch_zip_entry_uncompressed_size(origin_bytes.to_vec(), "xl/worksheets/sheet1.xml", oversized_len);
        workbook.origin_xlsx_bytes = Some(Arc::<[u8]>::from(patched));

        workbook.vba_project_bin = Some(b"fake-vba-project".to_vec());

        let written_bytes = write_xlsx_blocking(&out_path, &workbook).expect("write xlsm workbook");

        let archive = ZipArchive::new(Cursor::new(written_bytes.as_ref())).expect("open written zip");
        assert!(
            zip_archive_has_entry(&archive, "xl/vbaProject.bin"),
            "expected vbaProject.bin to be injected"
        );

        let content_types = formula_xlsx::read_part_from_reader_limited(
            Cursor::new(written_bytes.as_ref()),
            "[Content_Types].xml",
            XLSX_CONTENT_TYPES_MAX_BYTES,
        )
        .expect("read [Content_Types].xml")
        .expect("expected [Content_Types].xml");
        let ct_xml = String::from_utf8_lossy(&content_types);
        assert!(
            ct_xml.contains("application/vnd.ms-office.vbaProject"),
            "expected vbaProject content type override"
        );

        let workbook_rels = formula_xlsx::read_part_from_reader_limited(
            Cursor::new(written_bytes.as_ref()),
            "xl/_rels/workbook.xml.rels",
            XLSX_WORKBOOK_XML_MAX_BYTES,
        )
        .expect("read workbook.xml.rels")
        .expect("expected workbook.xml.rels");
        let rels_xml = String::from_utf8_lossy(&workbook_rels);
        assert!(
            rels_xml.contains("relationships/vbaProject"),
            "expected workbook.xml.rels to contain a vbaProject relationship"
        );
    }

    #[test]
    fn patch_save_can_update_date_system_without_inflating_oversized_parts() {
        let tmp = tempfile::tempdir().expect("temp dir");
        let origin_path = tmp.path().join("origin.xlsx");
        let out_path = tmp.path().join("saved.xlsx");

        let mut model = formula_model::Workbook::new();
        model.add_sheet("Sheet1").expect("add Sheet1");

        let mut cursor = Cursor::new(Vec::new());
        formula_xlsx::write_workbook_to_writer(&model, &mut cursor).expect("write xlsx bytes");
        std::fs::write(&origin_path, cursor.into_inner()).expect("write origin xlsx");

        let mut workbook = read_xlsx_blocking(&origin_path).expect("read origin xlsx");
        let origin_bytes = workbook
            .origin_xlsx_bytes
            .as_ref()
            .expect("expected origin bytes baseline");

        let oversized_len = formula_xlsx::MAX_XLSX_PACKAGE_PART_BYTES as u32 + 1;
        let patched =
            patch_zip_entry_uncompressed_size(origin_bytes.to_vec(), "xl/worksheets/sheet1.xml", oversized_len);
        workbook.origin_xlsx_bytes = Some(Arc::<[u8]>::from(patched));

        // Force a date system mismatch so the patch-based save path needs to update `xl/workbook.xml`.
        workbook.date_system = WorkbookDateSystem::Excel1904;

        let written_bytes = write_xlsx_blocking(&out_path, &workbook).expect("write xlsx workbook");

        let workbook_xml = formula_xlsx::read_part_from_reader_limited(
            Cursor::new(written_bytes.as_ref()),
            "xl/workbook.xml",
            XLSX_WORKBOOK_XML_MAX_BYTES,
        )
        .expect("read xl/workbook.xml")
        .expect("expected xl/workbook.xml");
        let xml = String::from_utf8_lossy(&workbook_xml);
        assert!(
            xml.contains("date1904=\"1\"") || xml.contains("date1904='1'"),
            "expected workbook.xml to have date1904=1, got: {xml}"
        );
    }

    #[test]
    fn patch_save_falls_back_to_regeneration_when_workbook_xml_exceeds_limit() {
        let tmp = tempfile::tempdir().expect("temp dir");

        let mut model = formula_model::Workbook::new();
        model.add_sheet("Sheet1").expect("add Sheet1");

        let mut cursor = Cursor::new(Vec::new());
        formula_xlsx::write_workbook_to_writer(&model, &mut cursor).expect("write xlsx bytes");
        let base_bytes = cursor.into_inner();

        // Add an unknown part so we can detect whether the patch-based save path was used.
        const CUSTOM_PART: &str = "xl/formula-custom-part.xml";
        let mut pkg = XlsxPackage::from_bytes(&base_bytes).expect("parse base package");
        pkg.set_part(CUSTOM_PART, b"<custom>Hello</custom>".to_vec());
        let bytes_with_custom = pkg.write_to_bytes().expect("repack base package");

        // Forge ZIP metadata to claim `xl/workbook.xml` exceeds our patch-based save workbook.xml cap.
        // This should force save to fall back to regeneration (dropping unknown parts) rather than
        // attempting to patch from the origin bytes.
        let oversized_len = XLSX_WORKBOOK_XML_MAX_BYTES as u32 + 1;
        let patched =
            patch_zip_entry_uncompressed_size(bytes_with_custom, "xl/workbook.xml", oversized_len);

        let origin_path = tmp.path().join("origin.xlsx");
        std::fs::write(&origin_path, &patched).expect("write patched origin");

        let workbook = read_xlsx_blocking(&origin_path).expect("read origin workbook");
        assert!(
            workbook.origin_xlsx_bytes.is_some(),
            "expected origin bytes to be retained for this small fixture"
        );

        let out_path = tmp.path().join("saved.xlsx");
        let saved_bytes = write_xlsx_blocking(&out_path, &workbook).expect("save workbook");

        // The regeneration-based save path will drop unknown parts.
        let preserved_custom = formula_xlsx::read_part_from_reader(
            Cursor::new(saved_bytes.as_ref()),
            CUSTOM_PART,
        )
        .expect("read custom part");
        assert!(
            preserved_custom.is_none(),
            "expected regeneration-based save to drop unknown part"
        );
    }

    #[test]
    fn sheet_metadata_part_overrides_rejects_oversize_workbook_xml() {
        let cursor = Cursor::new(Vec::new());
        let mut zip = ZipWriter::new(cursor);
        let options = FileOptions::<()>::default().compression_method(CompressionMethod::Deflated);

        zip.start_file("xl/workbook.xml", options)
            .expect("start workbook.xml");

        // Write a highly compressible payload slightly above the read limit. The ZIP entry's
        // uncompressed size metadata will exceed the cap, triggering a deterministic error without
        // requiring a huge allocation during extraction.
        let target_len = XLSX_WORKBOOK_XML_MAX_BYTES as usize + 1;
        let chunk = vec![b'a'; 8192];
        let mut remaining = target_len;
        while remaining > 0 {
            let n = remaining.min(chunk.len());
            zip.write_all(&chunk[..n]).expect("write chunk");
            remaining -= n;
        }

        let bytes = zip.finish().expect("finish zip").into_inner();

        let workbook = Workbook::new_empty(None);
        let err = sheet_metadata_part_overrides(&bytes, &workbook)
            .expect_err("expected oversize workbook.xml to fail");
        let msg = format!("{err:?}");
        assert!(
            err.chain()
                .any(|cause| cause.to_string().contains("too large")
                    && cause.to_string().contains("xl/workbook.xml")),
            "expected PartTooLarge error, got: {msg}"
        );
    }

    #[test]
    fn record_value_to_scalar_prefers_display_field_over_display_value() {
        let record = formula_model::RecordValue::default()
            .with_display_field("Name")
            .with_field("Name", "Alice");

        let value = ModelCellValue::Record(record);
        assert_eq!(
            formula_model_value_to_scalar(&value),
            CellScalar::Text("Alice".to_string())
        );

        // Our JSON-introspection fallback should also understand the modern
        // camelCase schema (`displayField`, `displayValue`).
        assert_eq!(
            rich_model_cell_value_to_scalar(&value),
            Some(CellScalar::Text("Alice".to_string()))
        );
    }

    #[test]
    fn read_xlsx_blocking_requires_password_for_encrypted_ooxml_container() {
        let tmp = tempfile::tempdir().expect("temp dir");

        fn encrypted_ooxml_bytes(encryption_info: &str, encrypted_package: &str) -> Vec<u8> {
            let cursor = Cursor::new(Vec::new());
            let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
            ole.create_stream(encryption_info)
                .unwrap_or_else(|_| panic!("create {encryption_info} stream"));
            ole.create_stream(encrypted_package)
                .unwrap_or_else(|_| panic!("create {encrypted_package} stream"));
            ole.into_inner().into_inner()
        }

        for (info, package) in [
            ("EncryptionInfo", "EncryptedPackage"),
            ("encryptioninfo", "encryptedpackage"),
            ("/encryptioninfo", "/encryptedpackage"),
        ] {
            let bytes = encrypted_ooxml_bytes(info, package);
            for filename in ["encrypted.xlsx", "encrypted.xls", "encrypted.xlsb"] {
                let path = tmp.path().join(filename);
                std::fs::write(&path, &bytes).expect("write encrypted fixture");

                let err =
                    read_xlsx_blocking(&path).expect_err("expected encrypted workbook to error");
                let msg = err.to_string();
                assert!(
                    msg.starts_with(PASSWORD_REQUIRED_PREFIX),
                    "expected password-required error prefix, got: {msg}"
                );
            }
        }
    }

    #[test]
    fn read_xls_blocking_errors_on_encrypted_filepass_container() {
        // Minimal OLE/CFB container with a BIFF workbook stream containing FILEPASS.
        //
        // The desktop app should surface a structured "password required" prefix so the UI can
        // prompt for a password without leaking internal Rust API guidance.
        let tmp = tempfile::tempdir().expect("temp dir");

        fn record(record_id: u16, payload: &[u8]) -> Vec<u8> {
            let mut out = Vec::with_capacity(4 + payload.len());
            out.extend_from_slice(&record_id.to_le_bytes());
            out.extend_from_slice(&(payload.len() as u16).to_le_bytes());
            out.extend_from_slice(payload);
            out
        }

        const RECORD_BOF_BIFF8: u16 = 0x0809;
        const RECORD_FILEPASS: u16 = 0x002F;
        const RECORD_EOF: u16 = 0x000A;

        let workbook_stream = [
            record(RECORD_BOF_BIFF8, &[0u8; 16]),
            record(RECORD_FILEPASS, &[]),
            record(RECORD_EOF, &[]),
        ]
        .concat();

        let cursor = Cursor::new(Vec::new());
        let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
        {
            let mut stream = ole
                .create_stream("Workbook")
                .expect("create Workbook stream");
            stream
                .write_all(&workbook_stream)
                .expect("write Workbook stream bytes");
        }
        let bytes = ole.into_inner().into_inner();

        let path = tmp.path().join("encrypted.xls");
        std::fs::write(&path, &bytes).expect("write encrypted fixture");

        let err = read_xls_blocking_with_password(&path, None)
            .expect_err("expected encrypted workbook to error");
        let msg = err.to_string();
        assert!(
            msg.starts_with(PASSWORD_REQUIRED_PREFIX),
            "expected password-required error prefix, got: {msg}"
        );
        let msg = msg.to_lowercase();
        assert!(
            !msg.contains("import_xls_path_with_password"),
            "desktop error should not mention internal Rust APIs, got: {msg}"
        );
    }

    #[test]
    fn read_xlsx_blocking_opens_password_protected_ooxml_with_password() {
        use ms_offcrypto_writer::Ecma376AgileWriter;
        use rand::rngs::StdRng;
        use rand::SeedableRng;

        let fixture = concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../../fixtures/xlsx/basic/basic.xlsx"
        );
        let plain_path = std::path::Path::new(fixture);
        let plain = std::fs::read(plain_path).expect("read basic.xlsx fixture bytes");

        let cursor = Cursor::new(Vec::new());
        let mut rng = StdRng::seed_from_u64(0);
        let mut writer =
            Ecma376AgileWriter::create(&mut rng, "password", cursor).expect("create writer");
        writer.write_all(&plain).expect("write plaintext");
        let cursor = writer.into_inner().expect("finalize");
        let encrypted_ole = cursor.into_inner();

        let tmp = tempfile::tempdir().expect("temp dir");
        let path = tmp.path().join("encrypted.xlsx");
        std::fs::write(&path, &encrypted_ole).expect("write encrypted fixture");

        let err = read_xlsx_blocking_with_password(&path, None)
            .expect_err("expected missing password to error");
        assert!(
            err.to_string().starts_with(PASSWORD_REQUIRED_PREFIX),
            "expected password-required error, got: {err}"
        );

        let err = read_xlsx_blocking_with_password(&path, Some("wrong-password"))
            .expect_err("expected wrong password to error");
        assert!(
            err.to_string().starts_with(INVALID_PASSWORD_PREFIX),
            "expected invalid-password prefix, got: {err}"
        );

        let workbook = read_xlsx_blocking_with_password(&path, Some("password"))
            .expect("open encrypted workbook");
        assert!(
            workbook.origin_xlsx_bytes.is_some(),
            "expected origin_xlsx_bytes to be populated with decrypted ZIP bytes"
        );
        let origin = workbook
            .origin_xlsx_bytes
            .as_ref()
            .expect("origin_xlsx_bytes")
            .as_ref();
        assert!(
            origin.starts_with(b"PK"),
            "expected decrypted bytes to be a ZIP/OPC package"
        );
        assert!(!workbook.sheets.is_empty(), "expected sheets to be parsed");
    }

    #[test]
    fn read_workbook_blocking_opens_password_protected_xlsb_fixture_with_password() {
        let fixture = concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../../fixtures/encrypted/encrypted.xlsb"
        );
        let path = std::path::Path::new(fixture);

        let err = read_workbook_blocking_with_password(path, None)
            .expect_err("expected missing password to error");
        assert!(
            err.to_string().starts_with(PASSWORD_REQUIRED_PREFIX),
            "expected password-required error, got: {err}"
        );

        let err = read_workbook_blocking_with_password(path, Some("wrong-password"))
            .expect_err("expected wrong password to error");
        assert!(
            err.to_string().starts_with(INVALID_PASSWORD_PREFIX),
            "expected invalid-password prefix, got: {err}"
        );

        // See `fixtures/encrypted/README.md` for provenance and expected contents.
        let workbook =
            read_workbook_blocking_with_password(path, Some("tika")).expect("open encrypted xlsb");
        assert_eq!(
            workbook.path, None,
            "encrypted XLSB opens should force Save As (no lossless .xlsb save path)"
        );
        assert_eq!(workbook.origin_xlsb_path, None);
        let sheet = workbook
            .sheets
            .iter()
            .find(|s| s.name == "Sheet1")
            .expect("Sheet1 should exist");
        let cell = sheet.get_cell(0, 0);
        assert_eq!(
            cell.computed_value,
            CellScalar::Text("You can't see me".to_string())
        );
    }

    #[test]
    fn read_workbook_blocking_opens_password_protected_xls_fixture_with_password() {
        let fixture = concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../../crates/formula-xls/tests/fixtures/encrypted/biff8_rc4_standard_pw_open.xls"
        );
        let path = std::path::Path::new(fixture);

        let err = read_workbook_blocking_with_password(path, None)
            .expect_err("expected missing password to error");
        assert!(
            err.to_string().starts_with(PASSWORD_REQUIRED_PREFIX),
            "expected password-required error, got: {err}"
        );

        let err = read_workbook_blocking_with_password(path, Some("wrong-password"))
            .expect_err("expected wrong password to error");
        assert!(
            err.to_string().starts_with(INVALID_PASSWORD_PREFIX),
            "expected invalid-password prefix, got: {err}"
        );

        let workbook = read_workbook_blocking_with_password(path, Some("password"))
            .expect("open encrypted xls");
        let sheet = workbook
            .sheets
            .iter()
            .find(|s| s.name == "Sheet1")
            .expect("Sheet1 should exist");
        let cell_a1 = sheet.get_cell(0, 0);
        assert_eq!(cell_a1.computed_value, CellScalar::Number(42.0));
    }

    #[test]
    fn read_workbook_blocking_opens_password_protected_xls_cryptoapi_fixture_with_password() {
        let fixture = concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../../crates/formula-xls/tests/fixtures/encrypted/biff8_rc4_cryptoapi_pw_open.xls"
        );
        let path = std::path::Path::new(fixture);

        let err = read_workbook_blocking_with_password(path, None)
            .expect_err("expected missing password to error");
        assert!(
            err.to_string().starts_with(PASSWORD_REQUIRED_PREFIX),
            "expected password-required error, got: {err}"
        );

        let err = read_workbook_blocking_with_password(path, Some("wrong-password"))
            .expect_err("expected wrong password to error");
        assert!(
            err.to_string().starts_with(INVALID_PASSWORD_PREFIX),
            "expected invalid-password prefix, got: {err}"
        );

        let workbook =
            read_workbook_blocking_with_password(path, Some("correct horse battery staple"))
                .expect("open encrypted xls");
        let sheet = workbook
            .sheets
            .iter()
            .find(|s| s.name == "Sheet1")
            .expect("Sheet1 should exist");
        let cell_a1 = sheet.get_cell(0, 0);
        assert_eq!(cell_a1.computed_value, CellScalar::Number(42.0));
    }

    #[test]
    fn read_workbook_blocking_opens_password_protected_xls_cryptoapi_boundary_fixture_with_password() {
        // This is a real Excel-generated workbook that uses the legacy FILEPASS CryptoAPI layout and
        // crosses the 1024-byte RC4 re-key boundary.
        let fixture = concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../../fixtures/encrypted/encrypted.xls"
        );
        let path = std::path::Path::new(fixture);

        let workbook = read_workbook_blocking_with_password(path, Some("password"))
            .expect("open encrypted boundary .xls");
        let sheet = workbook
            .sheets
            .iter()
            .find(|s| s.name == "Sheet1")
            .expect("Sheet1 should exist");

        // See `fixtures/encrypted/README.md` for provenance and expected contents.
        let cell_a400 = sheet.get_cell(399, 0);
        assert_eq!(cell_a400.computed_value, CellScalar::Number(399.0));
        let cell_b400 = sheet.get_cell(399, 1);
        assert_eq!(
            cell_b400.computed_value,
            CellScalar::Text("RC4_BOUNDARY_OK".to_string())
        );
    }

    #[test]
    fn encrypt_package_to_ole_bytes_rejects_empty_password() {
        let err =
            encrypt_package_to_ole_bytes(b"PK", "").expect_err("expected empty password to fail");
        assert!(
            err.to_string().starts_with(INVALID_PASSWORD_PREFIX),
            "expected invalid-password prefix, got: {err}"
        );
    }

    #[test]
    fn encrypt_package_to_ole_bytes_preserves_unicode_and_whitespace_password() {
        let fixture = concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../../fixtures/xlsx/basic/basic.xlsx"
        );
        let plain_path = std::path::Path::new(fixture);
        let plain = std::fs::read(plain_path).expect("read basic.xlsx fixture bytes");

        // Trailing whitespace is significant and must not be trimmed away.
        // (This also exercises non-BMP UTF-16 surrogate pair encoding via the emoji.)
        let password = "pässwörd🔒 ";

        let ole_bytes =
            encrypt_package_to_ole_bytes(&plain, password).expect("encrypt package to ole");
        let decrypted = formula_office_crypto::decrypt_encrypted_package(&ole_bytes, password)
            .expect("decrypt with exact password");
        assert_eq!(decrypted, plain, "expected decrypt(encrypt(plain)) == plain");

        let trimmed_password = password.trim();
        let err = formula_office_crypto::decrypt_encrypted_package(&ole_bytes, trimmed_password)
            .expect_err("expected trimmed password to fail");
        assert!(
            matches!(err, OfficeCryptoError::InvalidPassword),
            "expected InvalidPassword for trimmed password, got: {err:?}"
        );
    }

    #[test]
    fn encrypted_save_produces_ole_file_and_roundtrips_via_open_with_password() {
        let tmp = tempfile::tempdir().expect("temp dir");
        let out_path = tmp.path().join("encrypted.xlsx");

        let mut workbook = Workbook::new_empty(None);
        workbook.add_sheet("Sheet1".to_string());
        workbook.ensure_sheet_ids();
        workbook.sheets[0].set_cell(0, 0, Cell::from_literal(Some(CellScalar::Number(123.0))));

        // Build plaintext ZIP bytes strictly in-memory, then encrypt into an OLE/CFB wrapper.
        let zip_bytes = build_xlsx_bytes_blocking(&out_path, &workbook).expect("build xlsx bytes");
        let ole_bytes =
            encrypt_package_to_ole_bytes(zip_bytes.as_ref(), "password").expect("encrypt package");

        write_file_atomic(&out_path, &ole_bytes).expect("write encrypted workbook");

        let saved = std::fs::read(&out_path).expect("read saved bytes");
        assert!(
            saved.starts_with(&OLE_MAGIC),
            "expected encrypted save to start with OLE magic"
        );
        assert!(
            !saved.starts_with(b"PK"),
            "expected encrypted save to not start with ZIP magic"
        );

        let err = read_xlsx_blocking_with_password(&out_path, Some("wrong-password"))
            .expect_err("expected wrong password to fail");
        assert!(
            err.to_string().starts_with(INVALID_PASSWORD_PREFIX),
            "expected invalid-password prefix, got: {err}"
        );

        let opened = read_xlsx_blocking_with_password(&out_path, Some("password"))
            .expect("open encrypted workbook");
        let sheet = opened
            .sheets
            .iter()
            .find(|s| s.name == "Sheet1")
            .expect("Sheet1 should exist");
        let cell = sheet.get_cell(0, 0);
        assert_eq!(cell.computed_value, CellScalar::Number(123.0));
    }

    #[test]
    fn encrypted_open_maps_integrity_check_failed_to_invalid_password_error() {
        let tmp = tempfile::tempdir().expect("temp dir");
        let out_path = tmp.path().join("tampered.xlsx");
        let password = "password";

        let mut workbook = Workbook::new_empty(None);
        workbook.add_sheet("Sheet1".to_string());
        workbook.ensure_sheet_ids();
        workbook.sheets[0].set_cell(0, 0, Cell::from_literal(Some(CellScalar::Number(123.0))));

        // Encrypt a valid ZIP package, then tamper with the ciphertext to trigger an HMAC mismatch.
        let zip_bytes = build_xlsx_bytes_blocking(&out_path, &workbook).expect("build xlsx bytes");
        let ole_bytes = encrypt_package_to_ole_bytes(zip_bytes.as_ref(), password)
            .expect("encrypt package to ole");

        let mut ole = cfb::CompoundFile::open(Cursor::new(ole_bytes)).expect("open cfb");
        let mut encryption_info = Vec::new();
        ole.open_stream("EncryptionInfo")
            .expect("open EncryptionInfo stream")
            .read_to_end(&mut encryption_info)
            .expect("read EncryptionInfo bytes");
        let mut encrypted_package = Vec::new();
        ole.open_stream("EncryptedPackage")
            .expect("open EncryptedPackage stream")
            .read_to_end(&mut encrypted_package)
            .expect("read EncryptedPackage bytes");
        assert!(
            encrypted_package.len() > 8,
            "expected EncryptedPackage to include ciphertext bytes"
        );
        encrypted_package[8] ^= 0x01;

        let cursor = Cursor::new(Vec::new());
        let mut out = cfb::CompoundFile::create(cursor).expect("create cfb");
        out.create_stream("EncryptionInfo")
            .expect("create EncryptionInfo stream")
            .write_all(&encryption_info)
            .expect("write EncryptionInfo bytes");
        out.create_stream("EncryptedPackage")
            .expect("create EncryptedPackage stream")
            .write_all(&encrypted_package)
            .expect("write EncryptedPackage bytes");
        let tampered_ole = out.into_inner().into_inner();

        write_file_atomic(&out_path, &tampered_ole).expect("write tampered encrypted workbook");

        let err = read_xlsx_blocking_with_password(&out_path, Some(password))
            .expect_err("expected integrity check failure to error");
        assert!(
            err.to_string().starts_with(INVALID_PASSWORD_PREFIX),
            "expected integrity failures to map to invalid-password prefix, got: {err}"
        );
    }

    fn find_xlsb_cell_record(
        sheet_bin: &[u8],
        target_row: u32,
        target_col: u32,
    ) -> Option<(u32, Vec<u8>)> {
        const SHEETDATA: u32 = 0x0091;
        const SHEETDATA_END: u32 = 0x0092;
        const ROW: u32 = 0x0000;

        let mut cursor = Cursor::new(sheet_bin);
        let mut in_sheet_data = false;
        let mut current_row = 0u32;

        loop {
            let id = match biff12_varint::read_record_id(&mut cursor).ok().flatten() {
                Some(id) => id,
                None => break,
            };
            let len = match biff12_varint::read_record_len(&mut cursor).ok().flatten() {
                Some(len) => len as usize,
                None => return None,
            };
            let mut payload = vec![0u8; len];
            cursor.read_exact(&mut payload).ok()?;

            match id {
                SHEETDATA => in_sheet_data = true,
                SHEETDATA_END => in_sheet_data = false,
                ROW if in_sheet_data => {
                    if payload.len() >= 4 {
                        current_row = u32::from_le_bytes(payload[0..4].try_into().unwrap());
                    }
                }
                _ if in_sheet_data => {
                    if payload.len() < 8 {
                        continue;
                    }
                    let col = u32::from_le_bytes(payload[0..4].try_into().unwrap());
                    if current_row == target_row && col == target_col {
                        return Some((id, payload));
                    }
                }
                _ => {}
            }
        }

        None
    }

    #[test]
    fn read_csv_blocking_sniffs_semicolon_delimiter() {
        let tmp = tempfile::tempdir().expect("temp dir");
        let path = tmp.path().join("data.csv");
        std::fs::write(&path, "a;b\n1;2\n").expect("write csv");

        let workbook = read_csv_blocking(&path).expect("read csv");
        assert_eq!(workbook.sheets.len(), 1);

        let sheet = &workbook.sheets[0];
        let table = sheet.columnar.as_deref().expect("expected columnar table");
        assert_eq!(table.column_count(), 2);
        assert_eq!(table.row_count(), 1);

        assert_eq!(sheet.get_cell(0, 0).computed_value, CellScalar::Number(1.0));
        assert_eq!(sheet.get_cell(0, 1).computed_value, CellScalar::Number(2.0));
    }

    #[test]
    fn read_csv_blocking_respects_excel_sep_directive() {
        let tmp = tempfile::tempdir().expect("temp dir");
        let path = tmp.path().join("data.csv");
        std::fs::write(&path, "sep=;\na;b\n1;2\n").expect("write csv");

        let workbook = read_csv_blocking(&path).expect("read csv");
        assert_eq!(workbook.sheets.len(), 1);

        let sheet = &workbook.sheets[0];
        let table = sheet.columnar.as_deref().expect("expected columnar table");
        assert_eq!(table.column_count(), 2);
        assert_eq!(table.row_count(), 1);

        assert_eq!(sheet.get_cell(0, 0).computed_value, CellScalar::Number(1.0));
        assert_eq!(sheet.get_cell(0, 1).computed_value, CellScalar::Number(2.0));
    }

    #[test]
    fn read_csv_blocking_sniffs_tab_delimiter() {
        let tmp = tempfile::tempdir().expect("temp dir");
        let path = tmp.path().join("data.csv");
        std::fs::write(&path, "a\tb\n1\t2\n").expect("write csv");

        let workbook = read_csv_blocking(&path).expect("read csv");
        assert_eq!(workbook.sheets.len(), 1);

        let sheet = &workbook.sheets[0];
        let table = sheet.columnar.as_deref().expect("expected columnar table");
        assert_eq!(table.column_count(), 2);
        assert_eq!(table.row_count(), 1);

        assert_eq!(sheet.get_cell(0, 0).computed_value, CellScalar::Number(1.0));
        assert_eq!(sheet.get_cell(0, 1).computed_value, CellScalar::Number(2.0));
    }

    #[test]
    fn read_csv_blocking_supports_utf16le_tab_delimited_text() {
        let tmp = tempfile::tempdir().expect("temp dir");
        let path = tmp.path().join("data.csv");

        // Excel's "Unicode Text" export is UTF-16LE with a BOM and (typically) tab-delimited.
        let tsv = "col1\tcol2\r\n1\thello\r\n2\tworld\r\n";
        let mut bytes = vec![0xFF, 0xFE];
        for unit in tsv.encode_utf16() {
            bytes.extend_from_slice(&unit.to_le_bytes());
        }
        std::fs::write(&path, &bytes).expect("write utf16 tsv");

        let workbook = read_csv_blocking(&path).expect("read csv");
        assert_eq!(workbook.sheets.len(), 1);

        let sheet = &workbook.sheets[0];
        let table = sheet.columnar.as_deref().expect("expected columnar table");
        assert_eq!(table.column_count(), 2);
        assert_eq!(table.row_count(), 2);

        assert_eq!(sheet.get_cell(0, 0).computed_value, CellScalar::Number(1.0));
        assert_eq!(
            sheet.get_cell(0, 1).computed_value,
            CellScalar::Text("hello".to_string())
        );
        assert_eq!(sheet.get_cell(1, 0).computed_value, CellScalar::Number(2.0));
        assert_eq!(
            sheet.get_cell(1, 1).computed_value,
            CellScalar::Text("world".to_string())
        );
    }

    #[test]
    fn read_csv_blocking_sniffs_pipe_delimiter() {
        let tmp = tempfile::tempdir().expect("temp dir");
        let path = tmp.path().join("data.csv");
        std::fs::write(&path, "a|b\n1|2\n").expect("write csv");

        let workbook = read_csv_blocking(&path).expect("read csv");
        assert_eq!(workbook.sheets.len(), 1);

        let sheet = &workbook.sheets[0];
        let table = sheet.columnar.as_deref().expect("expected columnar table");
        assert_eq!(table.column_count(), 2);
        assert_eq!(table.row_count(), 1);

        assert_eq!(sheet.get_cell(0, 0).computed_value, CellScalar::Number(1.0));
        assert_eq!(sheet.get_cell(0, 1).computed_value, CellScalar::Number(2.0));
    }

    #[test]
    fn reads_xlsb_fixture() {
        let fixture_path = Path::new(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../../crates/formula-xlsb/tests/fixtures/simple.xlsb"
        ));

        let workbook = read_xlsx_blocking(fixture_path).expect("read xlsb workbook");
        assert_eq!(workbook.sheets.len(), 1);

        let sheet = &workbook.sheets[0];
        assert_eq!(sheet.name, "Sheet1");

        assert_eq!(
            sheet.get_cell(0, 0).computed_value,
            CellScalar::Text("Hello".to_string())
        );
        assert_eq!(
            sheet.get_cell(0, 1).computed_value,
            CellScalar::Number(42.5)
        );
        assert_eq!(
            sheet.get_cell(0, 2).computed_value,
            CellScalar::Number(85.0)
        );
        assert_eq!(sheet.get_cell(0, 2).formula.as_deref(), Some("=B1*2"));
    }

    #[test]
    fn reads_xlsx_fixture_with_unknown_extension() {
        let fixture_path = Path::new(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../../fixtures/xlsx/basic/basic.xlsx"
        ));
        let tmp = tempfile::tempdir().expect("temp dir");
        let renamed_path = tmp.path().join("basic.bin");
        std::fs::copy(fixture_path, &renamed_path).expect("copy fixture");

        let workbook =
            read_xlsx_blocking(&renamed_path).expect("read xlsx workbook with unknown extension");
        assert!(
            workbook.origin_xlsx_bytes.is_some(),
            "expected XLSX/XLSM reader path"
        );

        assert!(!workbook.sheets.is_empty(), "expected at least one sheet");
        let sheet = &workbook.sheets[0];
        assert_eq!(sheet.name, "Sheet1");
        assert_eq!(sheet.get_cell(0, 0).computed_value, CellScalar::Number(1.0));
        assert_eq!(
            sheet.get_cell(0, 1).computed_value,
            CellScalar::Text("Hello".to_string())
        );
    }

    #[test]
    fn reads_xlsb_fixture_with_unknown_extension() {
        let fixture_path = Path::new(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../../crates/formula-xlsb/tests/fixtures/simple.xlsb"
        ));
        let tmp = tempfile::tempdir().expect("temp dir");
        let renamed_path = tmp.path().join("simple.data");
        std::fs::copy(fixture_path, &renamed_path).expect("copy fixture");

        let workbook =
            read_xlsx_blocking(&renamed_path).expect("read xlsb workbook with unknown extension");
        let renamed_str = renamed_path.to_string_lossy().to_string();
        assert_eq!(
            workbook.origin_xlsb_path.as_deref(),
            Some(renamed_str.as_str())
        );

        assert_eq!(workbook.sheets.len(), 1);
        let sheet = &workbook.sheets[0];
        assert_eq!(sheet.name, "Sheet1");
        assert_eq!(
            sheet.get_cell(0, 0).computed_value,
            CellScalar::Text("Hello".to_string())
        );
        assert_eq!(
            sheet.get_cell(0, 1).computed_value,
            CellScalar::Number(42.5)
        );
        assert_eq!(
            sheet.get_cell(0, 2).computed_value,
            CellScalar::Number(85.0)
        );
        assert_eq!(sheet.get_cell(0, 2).formula.as_deref(), Some("=B1*2"));
    }

    #[test]
    fn reads_xls_fixture_with_unknown_extension() {
        let fixture_path = Path::new(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../../crates/formula-xls/tests/fixtures/basic.xls"
        ));
        let expected_date_system = formula_xls::import_xls_path(fixture_path)
            .expect("import xls")
            .workbook
            .date_system;

        let tmp = tempfile::tempdir().expect("temp dir");
        let renamed_path = tmp.path().join("basic.data");
        std::fs::copy(fixture_path, &renamed_path).expect("copy fixture");

        let workbook =
            read_xlsx_blocking(&renamed_path).expect("read xls workbook with unknown extension");
        assert_eq!(workbook.date_system, expected_date_system);

        let sheet1 = workbook
            .sheets
            .iter()
            .find(|s| s.name == "Sheet1")
            .expect("Sheet1 exists");
        assert_eq!(
            sheet1.get_cell(0, 0).computed_value,
            CellScalar::Text("Hello".to_string())
        );
        assert_eq!(
            sheet1.get_cell(1, 1).computed_value,
            CellScalar::Number(123.0)
        );
        assert_eq!(sheet1.get_cell(2, 2).formula.as_deref(), Some("=B2*2"));

        let second = workbook
            .sheets
            .iter()
            .find(|s| s.name == "Second")
            .expect("Second exists");
        assert_eq!(
            second.get_cell(0, 0).computed_value,
            CellScalar::Text("Second sheet".to_string())
        );
    }

    #[test]
    fn reads_xls_fixture_with_xlt_and_xla_extensions() {
        let fixture_path = Path::new(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../../crates/formula-xls/tests/fixtures/basic.xls"
        ));
        let expected_date_system = formula_xls::import_xls_path(fixture_path)
            .expect("import xls")
            .workbook
            .date_system;
        let tmp = tempfile::tempdir().expect("temp dir");

        for ext in ["xlt", "xla"] {
            let renamed_path = tmp.path().join(format!("basic.{ext}"));
            std::fs::copy(fixture_path, &renamed_path).expect("copy fixture");

            let workbook = read_xlsx_blocking(&renamed_path)
                .expect("read legacy template/add-in as xls workbook");
            assert!(workbook.origin_xlsx_bytes.is_none());
            assert_eq!(workbook.date_system, expected_date_system);

            let sheet1 = workbook
                .sheets
                .iter()
                .find(|s| s.name == "Sheet1")
                .expect("Sheet1 exists");
            assert_eq!(
                sheet1.get_cell(0, 0).computed_value,
                CellScalar::Text("Hello".to_string())
            );
        }
    }

    #[test]
    fn reads_xlsx_fixture_with_xls_extension() {
        let fixture_path = Path::new(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../../fixtures/xlsx/basic/basic.xlsx"
        ));
        let tmp = tempfile::tempdir().expect("temp dir");
        let renamed_path = tmp.path().join("basic.xls");
        std::fs::copy(fixture_path, &renamed_path).expect("copy fixture");

        let workbook = read_xlsx_blocking(&renamed_path)
            .expect("read xlsx workbook with wrong .xls extension");
        assert!(
            workbook.origin_xlsx_bytes.is_some(),
            "expected XLSX/XLSM reader path"
        );

        assert!(!workbook.sheets.is_empty(), "expected at least one sheet");
        let sheet = &workbook.sheets[0];
        assert_eq!(sheet.name, "Sheet1");
        assert_eq!(sheet.get_cell(0, 0).computed_value, CellScalar::Number(1.0));
        assert_eq!(
            sheet.get_cell(0, 1).computed_value,
            CellScalar::Text("Hello".to_string())
        );
    }

    #[test]
    fn reads_xls_fixture_with_xlsx_extension() {
        let fixture_path = Path::new(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../../crates/formula-xls/tests/fixtures/basic.xls"
        ));
        let expected_date_system = formula_xls::import_xls_path(fixture_path)
            .expect("import xls")
            .workbook
            .date_system;

        let tmp = tempfile::tempdir().expect("temp dir");
        let renamed_path = tmp.path().join("basic.xlsx");
        std::fs::copy(fixture_path, &renamed_path).expect("copy fixture");

        let workbook = read_xlsx_blocking(&renamed_path)
            .expect("read xls workbook with wrong .xlsx extension");
        assert!(workbook.origin_xlsx_bytes.is_none());
        assert_eq!(workbook.date_system, expected_date_system);
        let sheet1 = workbook
            .sheets
            .iter()
            .find(|s| s.name == "Sheet1")
            .expect("Sheet1 exists");
        assert_eq!(
            sheet1.get_cell(0, 0).computed_value,
            CellScalar::Text("Hello".to_string())
        );
    }

    #[test]
    fn reads_xlsb_fixture_with_xlsx_extension() {
        let fixture_path = Path::new(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../../crates/formula-xlsb/tests/fixtures/simple.xlsb"
        ));
        let tmp = tempfile::tempdir().expect("temp dir");
        let renamed_path = tmp.path().join("simple.xlsx");
        std::fs::copy(fixture_path, &renamed_path).expect("copy fixture");

        let workbook = read_xlsx_blocking(&renamed_path)
            .expect("read xlsb workbook with wrong .xlsx extension");
        let renamed_str = renamed_path.to_string_lossy().to_string();
        assert_eq!(
            workbook.origin_xlsb_path.as_deref(),
            Some(renamed_str.as_str())
        );
        assert_eq!(workbook.sheets.len(), 1);
        let sheet = &workbook.sheets[0];
        assert_eq!(sheet.name, "Sheet1");
        assert_eq!(
            sheet.get_cell(0, 0).computed_value,
            CellScalar::Text("Hello".to_string())
        );
    }

    #[test]
    fn reads_xlsb_date_system_1904_fixture() {
        let fixture_path = Path::new(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../../crates/formula-xlsb/tests/fixtures/date1904.xlsb"
        ));

        let workbook = read_xlsx_blocking(fixture_path).expect("read xlsb workbook");
        assert_eq!(workbook.date_system, WorkbookDateSystem::Excel1904);
    }

    #[test]
    fn reads_xls_date_system_1904_fixture() {
        let fixture_path = Path::new(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../../crates/formula-xls/tests/fixtures/date_system_1904.xls"
        ));
        let workbook = read_xlsx_blocking(fixture_path).expect("read xls workbook");
        assert_eq!(workbook.date_system, WorkbookDateSystem::Excel1904);
    }

    #[test]
    fn reads_xls_propagates_number_formats_into_cells() {
        let fixture_path = Path::new(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../../crates/formula-xls/tests/fixtures/dates.xls"
        ));
        let workbook = read_xlsx_blocking(fixture_path).expect("read xls workbook");

        let sheet = workbook
            .sheets
            .iter()
            .find(|s| crate::sheet_name::sheet_name_eq_case_insensitive(&s.name, "Dates"))
            .expect("Dates sheet exists");

        // `dates.xls` has a serial date value in A1 with Excel's default date format applied.
        let cell = sheet.get_cell(0, 0); // A1
        assert!(matches!(cell.computed_value, CellScalar::Number(_)));
        assert_eq!(cell.number_format.as_deref(), Some("m/d/yy"));
    }

    #[test]
    fn read_workbook_sniffs_xlsx_even_with_csv_extension() {
        let fixture_path = Path::new(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../../fixtures/xlsx/basic/basic.xlsx"
        ));
        let tmp = tempfile::tempdir().expect("temp dir");
        let misnamed_path = tmp.path().join("basic.csv");
        std::fs::copy(fixture_path, &misnamed_path).expect("copy fixture");

        let workbook = read_workbook_blocking(&misnamed_path).expect("read workbook");
        assert!(
            workbook.origin_xlsx_bytes.is_some(),
            "expected ZIP workbook to be parsed as XLSX"
        );
        assert_eq!(workbook.sheets.len(), 1);

        let sheet = &workbook.sheets[0];
        assert_eq!(sheet.get_cell(0, 0).computed_value, CellScalar::Number(1.0));
        assert_eq!(
            sheet.get_cell(0, 1).computed_value,
            CellScalar::Text("Hello".to_string())
        );
    }

    #[test]
    fn read_workbook_sniffs_csv_even_with_xlsx_extension() {
        let tmp = tempfile::tempdir().expect("temp dir");
        let misnamed_path = tmp.path().join("data.xlsx");
        std::fs::write(&misnamed_path, "col1,col2\n1,Hello\n").expect("write csv");

        let workbook = read_workbook_blocking(&misnamed_path).expect("read workbook");
        assert!(
            workbook.origin_xlsx_bytes.is_none(),
            "expected text workbook to not be treated as XLSX"
        );
        assert_eq!(workbook.sheets.len(), 1);

        let sheet = &workbook.sheets[0];
        assert!(
            sheet.columnar.is_some(),
            "expected CSV import to create a columnar-backed sheet"
        );
        assert_eq!(sheet.get_cell(0, 0).computed_value, CellScalar::Number(1.0));
        assert_eq!(
            sheet.get_cell(0, 1).computed_value,
            CellScalar::Text("Hello".to_string())
        );
    }

    #[cfg(not(feature = "parquet"))]
    #[test]
    fn read_workbook_errors_on_parquet_when_feature_disabled() {
        let fixture_path = Path::new(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../../packages/data-io/test/fixtures/simple.parquet"
        ));

        let err =
            read_workbook_blocking(fixture_path).expect_err("expected parquet import to fail");
        assert!(
            err.to_string()
                .contains("parquet support is not enabled in this build"),
            "unexpected error: {err}"
        );
    }

    #[test]
    fn reads_xlsb_populates_defined_names() {
        let fixture_path = Path::new(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../../crates/formula-xlsb/tests/fixtures_metadata/defined-names.xlsb"
        ));
        let workbook = read_xlsx_blocking(fixture_path).expect("read xlsb workbook");

        assert!(
            workbook.defined_names.iter().any(|n| n.name == "ZedName"),
            "expected defined name ZedName, got: {:?}",
            workbook
                .defined_names
                .iter()
                .map(|n| n.name.as_str())
                .collect::<Vec<_>>()
        );
        assert!(
            workbook.defined_names.iter().any(|n| n.name == "LocalName"),
            "expected defined name LocalName, got: {:?}",
            workbook
                .defined_names
                .iter()
                .map(|n| n.name.as_str())
                .collect::<Vec<_>>()
        );

        let zed = workbook
            .defined_names
            .iter()
            .find(|n| n.name == "ZedName")
            .expect("ZedName exists");
        assert_eq!(zed.refers_to, "Sheet1!$B$1");
        assert!(
            zed.sheet_id.is_none(),
            "expected ZedName to be workbook-scoped"
        );

        let local = workbook
            .defined_names
            .iter()
            .find(|n| n.name == "LocalName")
            .expect("LocalName exists");
        assert_eq!(local.refers_to, "Sheet1!$A$1");
        assert_eq!(local.sheet_id.as_deref(), Some("Sheet1"));

        let hidden = workbook
            .defined_names
            .iter()
            .find(|n| n.name == "HiddenName")
            .expect("HiddenName exists");
        assert!(hidden.hidden);
        assert_eq!(hidden.refers_to, "Sheet1!$A$1:$B$2");
    }

    #[test]
    fn xlsb_roundtrip_save_as_is_lossless() {
        let fixture_path = Path::new(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../../crates/formula-xlsb/tests/fixtures/simple.xlsb"
        ));
        let workbook = read_xlsx_blocking(fixture_path).expect("read xlsb workbook");

        let tmp = tempfile::tempdir().expect("temp dir");
        let out_path = tmp.path().join("roundtrip.xlsb");
        write_xlsx_blocking(&out_path, &workbook).expect("write xlsb workbook");

        let report = diff_workbooks(fixture_path, &out_path).expect("diff workbooks");
        assert!(
            report.is_empty(),
            "expected no diffs, got:\n{}",
            report
                .differences
                .iter()
                .map(|d| d.to_string())
                .collect::<Vec<_>>()
                .join("\n")
        );
    }

    #[test]
    fn xlsb_cell_edit_changes_only_expected_parts() {
        let fixture_path = Path::new(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../../crates/formula-xlsb/tests/fixtures/simple.xlsb"
        ));
        let fixture_archive = WorkbookArchive::open(fixture_path).expect("open fixture archive");
        let fixture_has_calc_chain = fixture_archive.get("xl/calcChain.bin").is_some();

        let mut workbook = read_xlsx_blocking(fixture_path).expect("read xlsb workbook");
        let sheet_id = workbook.sheets[0].id.clone();
        workbook.sheet_mut(&sheet_id).unwrap().set_cell(
            0,
            1,
            Cell::from_literal(Some(CellScalar::Number(123.0))),
        );

        let tmp = tempfile::tempdir().expect("temp dir");
        let out_path = tmp.path().join("edited.xlsb");
        write_xlsx_blocking(&out_path, &workbook).expect("write xlsb workbook");

        let out_archive = WorkbookArchive::open(&out_path).expect("open written archive");
        assert!(
            out_archive.get("xl/workbook.bin").is_some(),
            "expected output to contain xl/workbook.bin"
        );
        assert!(
            out_archive.get("xl/workbook.xml").is_none(),
            "expected output .xlsb to not contain xl/workbook.xml"
        );

        let report = diff_workbooks(fixture_path, &out_path).expect("diff workbooks");
        let report_text = report
            .differences
            .iter()
            .map(|d| d.to_string())
            .collect::<Vec<_>>()
            .join("\n");

        let extra_parts: Vec<_> = report
            .differences
            .iter()
            .filter(|d| d.kind == "extra_part")
            .map(|d| d.part.clone())
            .collect();
        assert!(
            extra_parts.is_empty(),
            "unexpected extra parts: {extra_parts:?}\n{report_text}"
        );

        assert!(
            report
                .differences
                .iter()
                .any(|d| d.part == "xl/worksheets/sheet1.bin"),
            "expected worksheet part to change, got:\n{report_text}",
        );

        let mut allowed_parts: BTreeSet<String> =
            BTreeSet::from(["xl/worksheets/sheet1.bin".to_string()]);
        if fixture_has_calc_chain {
            allowed_parts.extend([
                "xl/calcChain.bin".to_string(),
                "[Content_Types].xml".to_string(),
                "xl/_rels/workbook.bin.rels".to_string(),
            ]);
        } else {
            assert!(
                report
                    .differences
                    .iter()
                    .all(|d| !d.part.starts_with("xl/calcChain.")),
                "did not expect calcChain changes for fixture without calcChain.bin; got:\n{report_text}",
            );
            assert!(
                out_archive.get("xl/calcChain.bin").is_none(),
                "written workbook should not gain xl/calcChain.bin"
            );
        }

        let missing_parts: Vec<_> = report
            .differences
            .iter()
            .filter(|d| d.kind == "missing_part")
            .map(|d| d.part.clone())
            .collect();
        if fixture_has_calc_chain {
            assert!(
                missing_parts == vec!["xl/calcChain.bin".to_string()],
                "expected only calcChain.bin to be missing; got {missing_parts:?}\n{report_text}"
            );
        } else {
            assert!(
                missing_parts.is_empty(),
                "unexpected missing parts: {missing_parts:?}\n{report_text}"
            );
        }

        let diff_parts: BTreeSet<String> =
            report.differences.iter().map(|d| d.part.clone()).collect();
        let unexpected_parts: Vec<_> = diff_parts.difference(&allowed_parts).cloned().collect();
        assert!(
            unexpected_parts.is_empty(),
            "unexpected diff parts: {unexpected_parts:?}\n{report_text}"
        );

        let patched = XlsbWorkbook::open(&out_path).expect("re-open patched xlsb");
        let sheet = patched.read_sheet(0).expect("read patched sheet");
        let b1 = sheet
            .cells
            .iter()
            .find(|c| c.row == 0 && c.col == 1)
            .expect("B1 exists");
        assert_eq!(b1.value, XlsbCellValue::Number(123.0));
    }

    #[test]
    fn xlsb_text_edit_preserves_shared_string_record() {
        let fixture_path = Path::new(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../../crates/formula-xlsb/tests/fixtures/simple.xlsb"
        ));
        let fixture_archive = WorkbookArchive::open(fixture_path).expect("open fixture archive");
        let fixture_has_calc_chain = fixture_archive.get("xl/calcChain.bin").is_some();

        let mut workbook = read_xlsx_blocking(fixture_path).expect("read xlsb workbook");
        let sheet_id = workbook.sheets[0].id.clone();
        let new_text = "formula-desktop-tauri-shared-string-edit";
        workbook.sheet_mut(&sheet_id).unwrap().set_cell(
            0,
            0,
            Cell::from_literal(Some(CellScalar::Text(new_text.to_string()))),
        );

        let tmp = tempfile::tempdir().expect("temp dir");
        let out_path = tmp.path().join("edited-text.xlsb");
        write_xlsx_blocking(&out_path, &workbook).expect("write xlsb workbook");

        let report = diff_workbooks(fixture_path, &out_path).expect("diff workbooks");
        let report_text = report
            .differences
            .iter()
            .map(|d| d.to_string())
            .collect::<Vec<_>>()
            .join("\n");

        let extra_parts: Vec<_> = report
            .differences
            .iter()
            .filter(|d| d.kind == "extra_part")
            .map(|d| d.part.clone())
            .collect();
        assert!(
            extra_parts.is_empty(),
            "unexpected extra parts: {extra_parts:?}\n{report_text}"
        );

        for expected_part in ["xl/worksheets/sheet1.bin", "xl/sharedStrings.bin"] {
            assert!(
                report.differences.iter().any(|d| d.part == expected_part),
                "expected {expected_part} to change, got:\n{report_text}",
            );
        }

        let mut allowed_parts: BTreeSet<String> = BTreeSet::from([
            "xl/worksheets/sheet1.bin".to_string(),
            "xl/sharedStrings.bin".to_string(),
        ]);
        if fixture_has_calc_chain {
            allowed_parts.extend([
                "xl/calcChain.bin".to_string(),
                "[Content_Types].xml".to_string(),
                "xl/_rels/workbook.bin.rels".to_string(),
            ]);
        } else {
            assert!(
                report
                    .differences
                    .iter()
                    .all(|d| !d.part.starts_with("xl/calcChain.")),
                "did not expect calcChain changes for fixture without calcChain.bin; got:\n{report_text}",
            );
            let out_archive = WorkbookArchive::open(&out_path).expect("open written archive");
            assert!(
                out_archive.get("xl/calcChain.bin").is_none(),
                "written workbook should not gain xl/calcChain.bin"
            );
        }

        let missing_parts: Vec<_> = report
            .differences
            .iter()
            .filter(|d| d.kind == "missing_part")
            .map(|d| d.part.clone())
            .collect();
        if fixture_has_calc_chain {
            assert!(
                missing_parts == vec!["xl/calcChain.bin".to_string()],
                "expected only calcChain.bin to be missing; got {missing_parts:?}\n{report_text}"
            );
        } else {
            assert!(
                missing_parts.is_empty(),
                "unexpected missing parts: {missing_parts:?}\n{report_text}"
            );
        }

        let diff_parts: BTreeSet<String> =
            report.differences.iter().map(|d| d.part.clone()).collect();
        let unexpected_parts: Vec<_> = diff_parts.difference(&allowed_parts).cloned().collect();
        assert!(
            unexpected_parts.is_empty(),
            "unexpected diff parts: {unexpected_parts:?}\n{report_text}"
        );

        let patched = XlsbWorkbook::open(&out_path).expect("re-open patched xlsb");
        let sheet = patched.read_sheet(0).expect("read patched sheet");
        let a1 = sheet
            .cells
            .iter()
            .find(|c| c.row == 0 && c.col == 0)
            .expect("A1 exists");
        assert_eq!(a1.value, XlsbCellValue::Text(new_text.to_string()));

        let archive = WorkbookArchive::open(&out_path).expect("open written archive");
        let sheet_bin = archive
            .get("xl/worksheets/sheet1.bin")
            .expect("sheet1.bin exists");
        let (record_id, payload) =
            find_xlsb_cell_record(sheet_bin, 0, 0).expect("find A1 cell record");
        assert_eq!(
            record_id, 0x0007,
            "expected BrtCellIsst/STRING record id for A1"
        );
        assert!(
            payload.len() >= 12,
            "expected A1 payload to contain shared string index, got {} bytes",
            payload.len()
        );
        let isst = u32::from_le_bytes(payload[8..12].try_into().unwrap()) as usize;
        let shared_strings = patched.shared_strings();
        assert!(
            isst < shared_strings.len(),
            "A1 shared string index {isst} out of bounds ({} strings)",
            shared_strings.len()
        );
        assert_eq!(shared_strings[isst], new_text);
    }

    #[test]
    fn saves_xlsb_fixture_with_cleared_formula_cell() {
        let fixture_path = Path::new(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../../crates/formula-xlsb/tests/fixtures/simple.xlsb"
        ));

        let workbook = read_xlsx_blocking(fixture_path).expect("read xlsb workbook");

        let mut state = AppState::new();
        let info = state.load_workbook(workbook);
        let sheet_id = info.sheets[0].id.clone();
        state
            .set_cell(&sheet_id, 0, 2, None, None)
            .expect("clear formula cell");

        let tmp = tempfile::tempdir().expect("temp dir");
        let out_path = tmp.path().join("cleared-formula.xlsb");
        write_xlsx_blocking(&out_path, state.get_workbook().unwrap()).expect("write xlsb workbook");

        let patched_wb = XlsbWorkbook::open(&out_path).expect("open patched workbook");
        let patched_sheet = patched_wb.read_sheet(0).expect("read patched sheet");
        let c1 = patched_sheet
            .cells
            .iter()
            .find(|c| (c.row, c.col) == (0, 2))
            .expect("C1 exists");
        assert_eq!(c1.value, XlsbCellValue::Blank);
        assert!(c1.formula.is_none(), "expected formula to be cleared");
    }

    #[test]
    fn xlsb_edit_then_revert_does_not_change_workbook() {
        use serde_json::Value as JsonValue;

        let fixture_path = Path::new(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../../crates/formula-xlsb/tests/fixtures/simple.xlsb"
        ));
        let workbook = read_xlsx_blocking(fixture_path).expect("read xlsb workbook");

        let mut state = AppState::new();
        let info = state.load_workbook(workbook);
        let sheet_id = info.sheets[0].id.clone();

        // Pick a cell outside the used range so we can reliably "return to empty".
        state
            .set_cell(&sheet_id, 50, 50, Some(JsonValue::from(123)), None)
            .expect("edit cell");
        state
            .set_cell(&sheet_id, 50, 50, None, None)
            .expect("revert cell");

        let tmp = tempfile::tempdir().expect("temp dir");
        let out_path = tmp.path().join("reverted.xlsb");
        write_xlsx_blocking(&out_path, state.get_workbook().unwrap()).expect("write workbook");

        let report = diff_workbooks(fixture_path, &out_path).expect("diff workbooks");
        assert!(
            report.is_empty(),
            "expected no diffs, got:\n{}",
            report
                .differences
                .iter()
                .map(|d| d.to_string())
                .collect::<Vec<_>>()
                .join("\n")
        );
    }

    #[test]
    fn xlsb_formula_fallback_fills_missing_formula_text() {
        // Simulate a formula-xlsb cell that has a cached value but no decoded formula text,
        // and a Calamine-provided lookup table for formulas.
        let lookup = HashMap::from([((0, 2), "=B1*2".to_string())]);
        let mut sheet = Sheet::new("Sheet1".to_string(), "Sheet1".to_string());
        apply_xlsb_formula_fallback(
            &mut sheet,
            vec![(0, 2, CellScalar::Number(85.0), None)],
            &lookup,
        );

        let cell = sheet.get_cell(0, 2);
        assert_eq!(cell.formula.as_deref(), Some("=B1*2"));
        assert_eq!(cell.computed_value, CellScalar::Number(85.0));
    }

    #[test]
    fn reads_xlsb_date_formats_via_styles_bin() {
        let fixture_path = Path::new(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../../crates/formula-xlsb/tests/fixtures_styles/date.xlsb"
        ));

        let workbook = read_xlsx_blocking(fixture_path).expect("read xlsb workbook");
        assert_eq!(workbook.sheets.len(), 1);

        let mut state = AppState::new();
        let info = state.load_workbook(workbook);
        let sheet_id = info.sheets[0].id.clone();

        let expected = format_value(
            FormatValue::Number(44927.0),
            Some("m/d/yyyy"),
            &FormatOptions::default(),
        )
        .text;
        let cell = state.get_cell(&sheet_id, 0, 0).expect("get cell");
        assert_eq!(cell.value, CellScalar::Number(44927.0));
        assert_eq!(cell.display_value, expected);
    }

    #[test]
    fn reads_xlsx_propagates_number_formats_into_cells() {
        let fixture_path = Path::new(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../../fixtures/xlsx/styles/varied_styles.xlsx"
        ));

        let workbook = read_xlsx_blocking(fixture_path).expect("read xlsx workbook");
        assert_eq!(workbook.sheets.len(), 1);

        // `fixtures/xlsx/styles/varied_styles.xlsx` has a date-formatted serial in I1 (style XF
        // with built-in numFmtId=14).
        let sheet = &workbook.sheets[0];
        let cell = sheet.get_cell(0, 8); // I1
        assert_eq!(cell.computed_value, CellScalar::Number(44927.0));
        assert_eq!(cell.number_format.as_deref(), Some("m/d/yyyy"));
    }

    #[test]
    fn reads_rich_text_shared_strings_fixture_as_plain_text() {
        let fixture_path = Path::new(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../../fixtures/xlsx/styles/rich-text-shared-strings.xlsx"
        ));

        let workbook = read_xlsx_blocking(fixture_path).expect("read rich-text fixture");
        assert_eq!(workbook.sheets.len(), 1);

        let sheet = &workbook.sheets[0];
        assert_eq!(
            sheet.get_cell(0, 0).computed_value,
            CellScalar::Text("Hello Bold Italic".to_string())
        );
    }

    #[test]
    fn reads_multi_sheet_fixture_preserves_sheet_order() {
        let fixture_path = Path::new(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../../fixtures/xlsx/basic/multi-sheet.xlsx"
        ));

        let workbook = read_xlsx_blocking(fixture_path).expect("read multi-sheet workbook");
        let sheet_names: Vec<_> = workbook.sheets.iter().map(|s| s.name.as_str()).collect();
        assert_eq!(sheet_names, vec!["Sheet1", "Sheet2"]);
    }

    #[test]
    fn save_after_sheet_rename_uses_stable_worksheet_part() {
        let fixture_path = Path::new(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../../fixtures/xlsx/basic/multi-sheet.xlsx"
        ));
        let original_bytes = std::fs::read(fixture_path).expect("read fixture bytes");
        let mut workbook = read_xlsx_blocking(fixture_path).expect("read multi-sheet workbook");
        assert_eq!(
            workbook.sheets[1].xlsx_worksheet_part.as_deref(),
            Some("xl/worksheets/sheet2.xml"),
            "expected read_xlsx_blocking to record worksheet part names for xlsx inputs"
        );

        // Simulate in-app sheet rename without rewriting `workbook.xml`.
        workbook.sheets[1].name = "Renamed".to_string();
        workbook.sheets[1].set_cell(0, 0, Cell::from_literal(Some(CellScalar::Number(123.0))));

        let tmp = tempfile::tempdir().expect("temp dir");
        let out_path = tmp.path().join("renamed.xlsx");
        let written_bytes = write_xlsx_blocking(&out_path, &workbook).expect("write workbook");

        assert_non_worksheet_parts_preserved(&original_bytes, written_bytes.as_ref());

        let doc = formula_xlsx::load_from_bytes(written_bytes.as_ref())
            .expect("load saved workbook from bytes");
        let sheet = doc
            .workbook
            .sheet_by_name("Sheet2")
            .expect("original sheet name should still exist in workbook.xml");
        assert_eq!(
            sheet.value(formula_model::CellRef::new(0, 0)),
            ModelCellValue::Number(123.0)
        );
    }

    #[test]
    fn sheet_visibility_and_tab_color_are_patched_on_save() {
        // Create a small workbook that we can treat as an "origin XLSX" package.
        let mut model = formula_model::Workbook::new();
        model.add_sheet("Sheet1").expect("add Sheet1");
        model.add_sheet("Sheet2").expect("add Sheet2");
        model.add_sheet("Sheet3").expect("add Sheet3");
        let mut cursor = Cursor::new(Vec::new());
        formula_xlsx::write_workbook_to_writer(&model, &mut cursor).expect("write base workbook");
        let base_bytes = cursor.into_inner();

        let tmp = tempfile::tempdir().expect("temp dir");
        let base_path = tmp.path().join("base.xlsx");
        std::fs::write(&base_path, &base_bytes).expect("write base file");

        let mut workbook = read_xlsx_blocking(&base_path).expect("read base workbook");
        assert!(
            workbook.origin_xlsx_bytes.is_some(),
            "expected read_xlsx_blocking to retain origin bytes"
        );

        // Apply in-app edits that must be persisted via patching workbook.xml / worksheet sheetPr.
        workbook.sheets[0].tab_color = Some(TabColor::rgb("FFFF0000"));
        workbook.sheets[1].visibility = SheetVisibility::Hidden;
        workbook.sheets[2].visibility = SheetVisibility::VeryHidden;

        // Include at least one cell patch so this exercises "cell patches + metadata overrides"
        // rather than only the metadata fast-path.
        workbook.sheets[0].set_cell(0, 0, Cell::from_literal(Some(CellScalar::Number(123.0))));

        let out_path = tmp.path().join("patched.xlsx");
        let written_bytes =
            write_xlsx_blocking(&out_path, &workbook).expect("write patched workbook");

        let roundtrip =
            formula_xlsx::read_workbook_from_reader(Cursor::new(written_bytes.as_ref()))
                .expect("read patched workbook");

        let sheet1 = roundtrip.sheet_by_name("Sheet1").expect("Sheet1 exists");
        assert_eq!(
            sheet1.tab_color.as_ref().and_then(|c| c.rgb.as_deref()),
            Some("FFFF0000")
        );
        assert_eq!(
            sheet1.value(formula_model::CellRef::new(0, 0)),
            ModelCellValue::Number(123.0)
        );

        let sheet2 = roundtrip.sheet_by_name("Sheet2").expect("Sheet2 exists");
        assert_eq!(sheet2.visibility, SheetVisibility::Hidden);

        let sheet3 = roundtrip.sheet_by_name("Sheet3").expect("Sheet3 exists");
        assert_eq!(sheet3.visibility, SheetVisibility::VeryHidden);
    }

    #[test]
    fn patch_save_persists_column_widths_and_hidden_columns() {
        let fixture_path = Path::new(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../../fixtures/xlsx/basic/basic.xlsx"
        ));

        let mut workbook = read_xlsx_blocking(fixture_path).expect("read xlsx workbook");
        assert!(
            workbook.origin_xlsx_bytes.is_some(),
            "expected read_xlsx_blocking to retain origin bytes"
        );

        // Set a custom width for column A and hide column B.
        workbook.sheets[0].col_properties.insert(
            0,
            ModelColProperties {
                width: Some(20.0),
                hidden: false,
                style_id: None,
            },
        );
        workbook.sheets[0].col_properties.insert(
            1,
            ModelColProperties {
                width: None,
                hidden: true,
                style_id: None,
            },
        );

        let tmp = tempfile::tempdir().expect("temp dir");
        let out_path = tmp.path().join("cols.xlsx");
        let written_bytes = write_xlsx_blocking(&out_path, &workbook).expect("write workbook");

        let pkg = XlsxPackage::from_bytes(written_bytes.as_ref()).expect("parse output package");
        let sheet_xml = pkg
            .part("xl/worksheets/sheet1.xml")
            .expect("sheet1.xml exists");
        let sheet_xml_str = std::str::from_utf8(sheet_xml).expect("sheet xml is utf8");

        assert!(
            sheet_xml_str.contains("<cols"),
            "expected sheet to contain <cols>, got:\n{sheet_xml_str}"
        );
        assert!(
            sheet_xml_str.contains("min=\"1\"")
                && sheet_xml_str.contains("max=\"1\"")
                && sheet_xml_str.contains("width=\"20\"")
                && sheet_xml_str.contains("customWidth=\"1\""),
            "expected sheet xml to contain a width override for column A, got:\n{sheet_xml_str}"
        );
        assert!(
            sheet_xml_str.contains("min=\"2\"")
                && sheet_xml_str.contains("max=\"2\"")
                && sheet_xml_str.contains("hidden=\"1\""),
            "expected sheet xml to contain a hidden=1 override for column B, got:\n{sheet_xml_str}"
        );

        // Re-open via the normal XLSX read path and ensure the backend model picks up the column metadata.
        let roundtrip = read_xlsx_blocking(&out_path).expect("read patched workbook");
        let sheet = &roundtrip.sheets[0];
        assert_eq!(
            sheet.col_properties.get(&0).and_then(|p| p.width),
            Some(20.0)
        );
        assert_eq!(sheet.col_properties.get(&1).map(|p| p.hidden), Some(true));
    }

    #[test]
    fn reads_formula_fixture_with_equals_prefix() {
        let fixture_path = Path::new(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../../fixtures/xlsx/formulas/formulas.xlsx"
        ));

        let workbook = read_xlsx_blocking(fixture_path).expect("read formula workbook");
        assert_eq!(workbook.sheets.len(), 1);
        let sheet = &workbook.sheets[0];
        assert_eq!(sheet.get_cell(0, 2).formula.as_deref(), Some("=A1+B1"));
        assert_eq!(sheet.get_cell(0, 2).computed_value, CellScalar::Number(3.0));
    }

    #[test]
    fn reads_csv_into_columnar_backed_sheet() {
        let tmp = tempfile::tempdir().expect("temp dir");
        let path = tmp.path().join("data.csv");
        std::fs::write(&path, "id,name\n1,hello\n2,world\n").expect("write csv");

        let workbook = read_csv_blocking(&path).expect("read csv");
        assert_eq!(workbook.sheets.len(), 1);
        let sheet = &workbook.sheets[0];

        assert_eq!(sheet.get_cell(0, 0).computed_value, CellScalar::Number(1.0));
        assert_eq!(
            sheet.get_cell(1, 1).computed_value,
            CellScalar::Text("world".to_string())
        );
    }

    #[cfg(feature = "parquet")]
    #[test]
    fn reads_parquet_into_columnar_backed_sheet() {
        let fixture_path = Path::new(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../../packages/data-io/test/fixtures/simple.parquet"
        ));

        let workbook = read_parquet_blocking(fixture_path).expect("read parquet");
        assert_eq!(workbook.sheets.len(), 1);
        let sheet = &workbook.sheets[0];

        // Schema: id, name, active, score
        assert_eq!(sheet.get_cell(0, 0).computed_value, CellScalar::Number(1.0));
        assert_eq!(
            sheet.get_cell(1, 1).computed_value,
            CellScalar::Text("Bob".to_string())
        );
        assert_eq!(sheet.get_cell(2, 2).computed_value, CellScalar::Bool(true));
        assert_eq!(
            sheet.get_cell(2, 3).computed_value,
            CellScalar::Number(3.75)
        );
    }

    #[test]
    fn app_workbook_to_formula_model_preserves_columnar_backing_and_overlay_cells() {
        let tmp = tempfile::tempdir().expect("temp dir");
        let path = tmp.path().join("data.csv");
        std::fs::write(&path, "id,name\n1,hello\n2,world\n").expect("write csv");

        let mut workbook = read_csv_blocking(&path).expect("read csv");
        assert_eq!(workbook.sheets.len(), 1);

        // Override a single cell on top of the columnar backing.
        workbook.sheets[0].set_cell(
            0,
            1,
            Cell::from_literal(Some(CellScalar::Text("override".to_string()))),
        );

        let model = app_workbook_to_formula_model(&workbook).expect("convert workbook to model");
        let model_sheet = model
            .sheet_by_name(&workbook.sheets[0].name)
            .expect("sheet exists");

        assert!(
            model_sheet.columnar_table().is_some(),
            "expected model worksheet to preserve columnar backing"
        );

        // Non-overridden cells should not be materialized into the sparse cell map.
        let a2 = formula_model::CellRef::new(1, 0);
        assert!(
            model_sheet.cell(a2).is_none(),
            "expected A2 to be served from columnar backing, not stored as a cell record"
        );
        assert_eq!(model_sheet.value(a2), ModelCellValue::Number(2.0));

        // Overridden cells should appear in the sparse cell overlay and take precedence.
        let b1 = formula_model::CellRef::new(0, 1);
        assert!(
            model_sheet.cell(b1).is_some(),
            "expected B1 override to be stored as a sparse cell record"
        );
        assert_eq!(
            model_sheet.value(b1),
            ModelCellValue::String("override".to_string())
        );
    }

    #[test]
    fn app_workbook_to_formula_model_roundtrip_preserves_number_formats() {
        let mut workbook = Workbook::new_empty(None);
        workbook.add_sheet("Sheet1".to_string());
        let sheet_id = workbook.sheets[0].id.clone();

        let mut cell = Cell::from_literal(Some(CellScalar::Number(44927.0)));
        cell.number_format = Some("m/d/yyyy".to_string());
        workbook.sheet_mut(&sheet_id).unwrap().set_cell(0, 0, cell);

        let tmp = tempfile::tempdir().expect("temp dir");
        let out_path = tmp.path().join("roundtrip.xlsx");
        write_xlsx_blocking(&out_path, &workbook).expect("write workbook");

        let loaded = read_xlsx_blocking(&out_path).expect("read workbook");
        let sheet = &loaded.sheets[0];
        let loaded_cell = sheet.get_cell(0, 0);
        assert_eq!(loaded_cell.computed_value, CellScalar::Number(44927.0));
        assert_eq!(loaded_cell.number_format.as_deref(), Some("m/d/yyyy"));
    }

    #[test]
    fn app_workbook_to_formula_model_roundtrip_preserves_col_properties() {
        let mut workbook = Workbook::new_empty(None);
        workbook.add_sheet("Sheet1".to_string());

        workbook.sheets[0].col_properties.insert(
            0,
            formula_model::ColProperties {
                width: Some(20.0),
                hidden: false,
                style_id: None,
            },
        );
        workbook.sheets[0].col_properties.insert(
            1,
            formula_model::ColProperties {
                width: None,
                hidden: true,
                style_id: None,
            },
        );

        let tmp = tempfile::tempdir().expect("temp dir");
        let out_path = tmp.path().join("col_properties.xlsx");
        write_xlsx_blocking(&out_path, &workbook).expect("write workbook");

        let loaded = read_xlsx_blocking(&out_path).expect("read workbook");
        let sheet = &loaded.sheets[0];

        assert_eq!(
            sheet.col_properties.get(&0).and_then(|p| p.width),
            Some(20.0)
        );
        assert_eq!(sheet.col_properties.get(&1).map(|p| p.hidden), Some(true));
    }

    #[test]
    fn app_workbook_to_formula_model_roundtrip_preserves_sheet_default_col_width() {
        let mut workbook = Workbook::new_empty(None);
        workbook.add_sheet("Sheet1".to_string());
        workbook.sheets[0].default_col_width = Some(20.0);

        let tmp = tempfile::tempdir().expect("temp dir");
        let out_path = tmp.path().join("default_col_width.xlsx");
        write_xlsx_blocking(&out_path, &workbook).expect("write workbook");

        let loaded = read_xlsx_blocking(&out_path).expect("read workbook");
        let sheet = &loaded.sheets[0];
        assert_eq!(sheet.default_col_width, Some(20.0));
    }

    #[test]
    fn reads_csv_with_invalid_file_name_sanitizes_sheet_name_and_writes_xlsx() {
        let tmp = tempfile::tempdir().expect("temp dir");
        // Use characters that are invalid for Excel sheet names but valid on common filesystems.
        let path = tmp.path().join("bad[name]test.csv");
        std::fs::write(&path, "id,name\n1,hello\n").expect("write csv");

        let workbook = read_csv_blocking(&path).expect("read csv");
        assert_eq!(workbook.sheets.len(), 1);
        let sheet = &workbook.sheets[0];

        assert_eq!(sheet.name, "badnametest");
        assert!(
            !sheet.name.trim().is_empty(),
            "sheet name should be non-empty"
        );
        for ch in [':', '\\', '/', '?', '*', '[', ']'] {
            assert!(
                !sheet.name.contains(ch),
                "expected sanitized sheet name to not contain {ch}, got: {}",
                sheet.name
            );
        }

        let out_path = tmp.path().join("sanitized.xlsx");
        write_xlsx_blocking(&out_path, &workbook).expect("write xlsx");
    }

    #[test]
    fn reads_xltx_fixture_via_xlsx_path() {
        let src_path = Path::new(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../../fixtures/xlsx/basic/basic.xlsx"
        ));

        let tmp = tempfile::tempdir().expect("temp dir");
        let xltx_path = tmp.path().join("basic.xltx");
        std::fs::copy(src_path, &xltx_path).expect("copy fixture to .xltx");

        let workbook = read_xlsx_blocking(&xltx_path).expect("read xltx workbook");
        assert_eq!(workbook.sheets.len(), 1);
    }

    #[test]
    fn reads_xltm_and_xlam_capture_vba_project() {
        let src_path = Path::new(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../../fixtures/xlsx/macros/basic.xlsm"
        ));

        let tmp = tempfile::tempdir().expect("temp dir");
        for ext in ["xltm", "xlam"] {
            let dst = tmp.path().join(format!("basic.{ext}"));
            std::fs::copy(src_path, &dst).expect("copy macro fixture");

            let workbook = read_xlsx_blocking(&dst).expect("read macro workbook");
            assert!(
                workbook.vba_project_bin.is_some(),
                "expected vba_project_bin to be captured for {ext}"
            );
        }
    }

    #[test]
    fn xltm_save_preserves_vba_and_xltx_save_strips_vba() {
        let fixture_path = Path::new(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../../fixtures/xlsx/macros/basic.xlsm"
        ));

        let tmp = tempfile::tempdir().expect("temp dir");
        let xltm_path = tmp.path().join("basic.xltm");
        std::fs::copy(fixture_path, &xltm_path).expect("copy fixture to .xltm");

        let original_bytes = std::fs::read(&xltm_path).expect("read xltm bytes");
        let original_pkg = XlsxPackage::from_bytes(&original_bytes).expect("parse fixture package");
        let original_vba = original_pkg
            .vba_project_bin()
            .expect("fixture has vbaProject.bin")
            .to_vec();

        let workbook = read_xlsx_blocking(&xltm_path).expect("read xltm workbook");

        // Save back to `.xltm` should preserve VBA verbatim.
        let out_xltm = tmp.path().join("roundtrip.xltm");
        write_xlsx_blocking(&out_xltm, &workbook).expect("write xltm workbook");
        let written_bytes = std::fs::read(&out_xltm).expect("read written xltm");
        let written_pkg = XlsxPackage::from_bytes(&written_bytes).expect("parse written package");
        assert_eq!(
            written_pkg
                .vba_project_bin()
                .expect("written xltm should contain vbaProject.bin"),
            original_vba.as_slice()
        );

        // Save to `.xlam` should also preserve VBA verbatim.
        let out_xlam = tmp.path().join("roundtrip.xlam");
        write_xlsx_blocking(&out_xlam, &workbook).expect("write xlam workbook");
        let written_bytes = std::fs::read(&out_xlam).expect("read written xlam");
        let written_pkg = XlsxPackage::from_bytes(&written_bytes).expect("parse written package");
        assert_eq!(
            written_pkg
                .vba_project_bin()
                .expect("written xlam should contain vbaProject.bin"),
            original_vba.as_slice()
        );

        // Save to `.xltx` should strip VBA.
        let out_xltx = tmp.path().join("converted.xltx");
        write_xlsx_blocking(&out_xltx, &workbook).expect("write xltx workbook");
        let written_bytes = std::fs::read(&out_xltx).expect("read written xltx");
        let written_pkg = XlsxPackage::from_bytes(&written_bytes).expect("parse written package");
        assert!(
            written_pkg.vba_project_bin().is_none(),
            "expected vbaProject.bin to be removed when saving as .xltx"
        );
    }

    #[test]
    fn xlsx_save_strips_macrosheets_even_without_vba_project_bin() {
        // Build a macro-enabled package that contains XLM macrosheets + dialog sheets but no
        // `xl/vbaProject.bin`. The desktop save path should still run the macro-stripping pipeline
        // when saving as `.xlsx`/`.xltx`.
        fn build_macrosheet_only_fixture() -> Vec<u8> {
            use std::io::Write;

            let content_types = r#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.ms-excel.sheet.macroEnabled.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/macrosheets/sheet2.xml" ContentType="application/vnd.ms-excel.macrosheet+xml"/>
  <Override PartName="/xl/dialogsheets/sheet3.xml" ContentType="application/vnd.ms-excel.dialogsheet+xml"/>
</Types>"#;

            let root_rels = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
    Target="xl/workbook.xml"/>
</Relationships>"#;

            let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
    <sheet name="MacroSheet" sheetId="2" r:id="rId2"/>
    <sheet name="DialogSheet" sheetId="3" r:id="rId3"/>
  </sheets>
</workbook>"#;

            let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
    Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2"
    Type="http://schemas.microsoft.com/office/2006/relationships/xlMacrosheet"
    Target="macrosheets/sheet2.xml"/>
  <Relationship Id="rId3"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/dialogsheet"
    Target="dialogsheets/sheet3.xml"/>
</Relationships>"#;

            let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>"#;

            let macro_sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<macroSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>"#;

            let dialog_sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<dialogsheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>"#;

            let empty_rels = br#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>"#;

            let cursor = Cursor::new(Vec::new());
            let mut zip = zip::ZipWriter::new(cursor);
            let options = zip::write::FileOptions::<()>::default()
                .compression_method(zip::CompressionMethod::Deflated);

            fn add_file(
                zip: &mut zip::ZipWriter<Cursor<Vec<u8>>>,
                options: zip::write::FileOptions<()>,
                name: &str,
                bytes: &[u8],
            ) {
                zip.start_file(name, options).unwrap();
                zip.write_all(bytes).unwrap();
            }

            add_file(
                &mut zip,
                options,
                "[Content_Types].xml",
                content_types.as_bytes(),
            );
            add_file(&mut zip, options, "_rels/.rels", root_rels.as_bytes());
            add_file(
                &mut zip,
                options,
                "xl/workbook.xml",
                workbook_xml.as_bytes(),
            );
            add_file(
                &mut zip,
                options,
                "xl/_rels/workbook.xml.rels",
                workbook_rels.as_bytes(),
            );
            add_file(
                &mut zip,
                options,
                "xl/worksheets/sheet1.xml",
                worksheet_xml.as_bytes(),
            );
            add_file(
                &mut zip,
                options,
                "xl/macrosheets/sheet2.xml",
                macro_sheet_xml.as_bytes(),
            );
            add_file(
                &mut zip,
                options,
                "xl/dialogsheets/sheet3.xml",
                dialog_sheet_xml.as_bytes(),
            );
            add_file(
                &mut zip,
                options,
                "xl/macrosheets/_rels/sheet2.xml.rels",
                empty_rels,
            );
            add_file(
                &mut zip,
                options,
                "xl/dialogsheets/_rels/sheet3.xml.rels",
                empty_rels,
            );

            zip.finish().unwrap().into_inner()
        }

        let bytes = build_macrosheet_only_fixture();
        let mut workbook = Workbook::new_empty(Some("macrosheet.xlsm".to_string()));
        workbook.origin_xlsx_bytes = Some(Arc::<[u8]>::from(bytes));
        // Ensure we are exercising the "macros present but vba_project_bin is None" path.
        workbook.vba_project_bin = None;

        let tmp = tempfile::tempdir().expect("temp dir");
        let out_path = tmp.path().join("out.xlsx");
        let out_bytes = write_xlsx_blocking(&out_path, &workbook).expect("write workbook");
        let out_pkg = XlsxPackage::from_bytes(out_bytes.as_ref()).expect("parse output package");

        assert!(
            !out_pkg.macro_presence().any(),
            "expected save-as-xlsx to strip macro content"
        );
        assert!(out_pkg.part("xl/macrosheets/sheet2.xml").is_none());
        assert!(out_pkg.part("xl/dialogsheets/sheet3.xml").is_none());

        let content_types = std::str::from_utf8(out_pkg.part("[Content_Types].xml").unwrap())
            .expect("content types xml utf-8");
        assert!(
            !content_types.contains("macroEnabled.main+xml"),
            "expected workbook content type to be downgraded (got {content_types:?})"
        );
    }

    #[test]
    fn saves_xlsx_family_with_correct_workbook_main_content_type_and_vba_policy() {
        let fixture_path = Path::new(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../../fixtures/xlsx/macros/basic.xlsm"
        ));
        let workbook = read_xlsx_blocking(fixture_path).expect("read workbook");

        let tmp = tempfile::tempdir().expect("temp dir");

        let cases = [
            ("xlsx", WorkbookKind::Workbook, false),
            ("xlsm", WorkbookKind::MacroEnabledWorkbook, true),
            ("xltx", WorkbookKind::Template, false),
            ("xltm", WorkbookKind::MacroEnabledTemplate, true),
            ("xlam", WorkbookKind::MacroEnabledAddIn, true),
        ];

        let all_kinds = [
            WorkbookKind::Workbook,
            WorkbookKind::MacroEnabledWorkbook,
            WorkbookKind::Template,
            WorkbookKind::MacroEnabledTemplate,
            WorkbookKind::MacroEnabledAddIn,
        ];

        for (ext, expected_kind, expect_vba) in cases {
            let out_path = tmp.path().join(format!("out.{ext}"));
            write_xlsx_blocking(&out_path, &workbook).expect("write workbook");

            let bytes = std::fs::read(&out_path).expect("read written workbook");
            let pkg = XlsxPackage::from_bytes(&bytes).expect("parse written package");

            let content_types = std::str::from_utf8(
                pkg.part("[Content_Types].xml")
                    .expect("expected [Content_Types].xml part"),
            )
            .expect("content types should be utf-8");
            assert!(
                content_types.contains(expected_kind.workbook_content_type()),
                "expected `[Content_Types].xml` to advertise the correct workbook main content type for .{ext} ({expected_kind:?})"
            );
            for other in all_kinds {
                if other == expected_kind {
                    continue;
                }
                assert!(
                    !content_types.contains(other.workbook_content_type()),
                    "expected `.{}\" workbook main content type to be absent when saving as .{ext}",
                    other.workbook_content_type()
                );
            }

            assert_eq!(
                pkg.vba_project_bin().is_some(),
                expect_vba,
                "unexpected VBA preservation policy for .{ext}"
            );
        }
    }

    #[test]
    fn saving_template_keeps_print_settings_roundtrip() {
        let fixture_path = Path::new(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../../fixtures/xlsx/basic/print-settings.xlsx"
        ));
        let mut workbook = read_xlsx_blocking(fixture_path).expect("read print settings workbook");
        assert!(
            !workbook.print_settings.sheets.is_empty(),
            "expected print-settings fixture to contain at least one sheet print settings entry"
        );

        // Flip orientation so we can assert it round-trips.
        let existing = workbook.print_settings.sheets[0].page_setup.orientation;
        let updated = match existing {
            formula_xlsx::print::Orientation::Portrait => {
                formula_xlsx::print::Orientation::Landscape
            }
            formula_xlsx::print::Orientation::Landscape => {
                formula_xlsx::print::Orientation::Portrait
            }
        };
        workbook.print_settings.sheets[0].page_setup.orientation = updated;

        let tmp = tempfile::tempdir().expect("temp dir");
        let out_path = tmp.path().join("out.xltx");
        write_xlsx_blocking(&out_path, &workbook).expect("write workbook");

        let bytes = std::fs::read(&out_path).expect("read written workbook");
        let roundtrip = read_workbook_print_settings(&bytes).expect("read print settings");
        assert!(
            !roundtrip.sheets.is_empty(),
            "expected saved workbook to still contain sheet print settings"
        );
        assert_eq!(roundtrip.sheets[0].page_setup.orientation, updated);
    }

    #[test]
    fn saving_date_system_1904_workbook_as_template_preserves_date1904_flag() {
        let fixture_path = Path::new(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../../fixtures/xlsx/metadata/date-system-1904.xlsx"
        ));
        let workbook = read_xlsx_blocking(fixture_path).expect("read date system workbook");
        assert_eq!(workbook.date_system, WorkbookDateSystem::Excel1904);

        let tmp = tempfile::tempdir().expect("temp dir");
        let out_path = tmp.path().join("out.xltx");
        write_xlsx_blocking(&out_path, &workbook).expect("write workbook");

        let bytes = std::fs::read(&out_path).expect("read written workbook");
        let pkg = XlsxPackage::from_bytes(&bytes).expect("parse written package");
        let workbook_xml = std::str::from_utf8(
            pkg.part("xl/workbook.xml")
                .expect("expected xl/workbook.xml part"),
        )
        .expect("workbook.xml should be utf-8");
        assert!(
            workbook_xml.contains("date1904=\"1\""),
            "expected xl/workbook.xml to preserve date1904=\"1\" when saving as .xltx"
        );
    }

    #[test]
    fn preserves_vba_project_bin_when_saving_xlsm() {
        let fixture_path = Path::new(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../../fixtures/xlsx/macros/basic.xlsm"
        ));
        let original_bytes = std::fs::read(fixture_path).expect("read fixture xlsm");
        let original_pkg = XlsxPackage::from_bytes(&original_bytes).expect("parse fixture package");
        let original_vba = original_pkg
            .vba_project_bin()
            .expect("fixture has vbaProject.bin")
            .to_vec();

        let workbook = read_xlsx_blocking(fixture_path).expect("read workbook");
        assert!(
            workbook.vba_project_bin.is_some(),
            "read_xlsx_blocking should capture vbaProject.bin"
        );

        let tmp = tempfile::tempdir().expect("temp dir");
        let out_path = tmp.path().join("roundtrip.xlsm");
        let _ = write_xlsx_blocking(&out_path, &workbook).expect("write workbook");

        let written_bytes = std::fs::read(&out_path).expect("read written file");
        let written_pkg = XlsxPackage::from_bytes(&written_bytes).expect("parse written package");
        let written_vba = written_pkg
            .vba_project_bin()
            .expect("written workbook should contain vbaProject.bin");

        assert_eq!(original_vba, written_vba);
    }

    #[test]
    fn preserves_chart_parts_when_saving_xlsx() {
        let fixture_path = Path::new(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../../fixtures/xlsx/charts/basic-chart.xlsx"
        ));
        let original_bytes = std::fs::read(fixture_path).expect("read chart fixture bytes");
        let original_pkg =
            XlsxPackage::from_bytes(&original_bytes).expect("parse chart fixture package");

        let mut loaded = read_xlsx_blocking(fixture_path).expect("read workbook");
        assert!(
            loaded.preserved_drawing_parts.is_some(),
            "expected chart parts to be captured for preservation"
        );

        // Clear origin bytes so we exercise the regeneration path + `apply_preserved_drawing_parts`.
        loaded.origin_xlsx_bytes = None;

        let tmp = tempfile::tempdir().expect("temp dir");
        let dst_path = tmp.path().join("chart-dst.xlsx");
        let _ = write_xlsx_blocking(&dst_path, &loaded).expect("write workbook");

        let roundtrip_bytes = std::fs::read(&dst_path).expect("read written workbook");
        let dst_pkg = XlsxPackage::from_bytes(&roundtrip_bytes).expect("parse dst pkg");

        // Drawing + chart parts should match byte-for-byte.
        for (name, part_bytes) in original_pkg.parts() {
            if name.starts_with("xl/drawings/")
                || name.starts_with("xl/charts/")
                || name.starts_with("xl/media/")
            {
                assert_eq!(
                    Some(part_bytes),
                    dst_pkg.part(name),
                    "missing or mismatched preserved part {name}"
                );
            }
        }

        // Verify the chart is still discoverable in the saved workbook.
        let src_charts = original_pkg.extract_charts().expect("extract src charts");
        let dst_charts = dst_pkg.extract_charts().expect("extract dst charts");
        assert_eq!(src_charts.len(), 1);
        assert_eq!(dst_charts.len(), 1);
        assert_eq!(src_charts[0].rel_id, dst_charts[0].rel_id);
        assert_eq!(src_charts[0].chart_part, dst_charts[0].chart_part);
        assert_eq!(src_charts[0].drawing_part, dst_charts[0].drawing_part);
    }

    #[test]
    fn preserves_pivot_slicer_and_timeline_parts_when_saving_xlsx() {
        let fixture_path = Path::new(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../../crates/formula-xlsx/tests/fixtures/pivot_slicers_and_chart.xlsx"
        ));
        let original_bytes = std::fs::read(fixture_path).expect("read pivot fixture");
        let original_pkg = XlsxPackage::from_bytes(&original_bytes).expect("parse pivot fixture");

        let mut workbook = read_xlsx_blocking(fixture_path).expect("read workbook");
        assert!(
            workbook.preserved_pivot_parts.is_some(),
            "expected pivot slicer/timeline parts to be captured for preservation"
        );

        // `read_xlsx_blocking` stores the original XLSX bytes, which means `write_xlsx_blocking`
        // will typically patch the existing package in-place. Clear it here so we exercise the
        // regeneration path + `apply_preserved_pivot_parts`.
        workbook.origin_xlsx_bytes = None;

        let tmp = tempfile::tempdir().expect("temp dir");
        let dst_path = tmp.path().join("pivot-roundtrip.xlsx");
        write_xlsx_blocking(&dst_path, &workbook).expect("write workbook");

        let roundtrip_bytes = std::fs::read(&dst_path).expect("read written workbook");
        let roundtrip_pkg = XlsxPackage::from_bytes(&roundtrip_bytes).expect("parse roundtrip pkg");

        for part in [
            "xl/pivotTables/pivotTable1.xml",
            "xl/slicers/slicer1.xml",
            "xl/slicers/_rels/slicer1.xml.rels",
            "xl/slicerCaches/slicerCache1.xml",
            "xl/slicerCaches/_rels/slicerCache1.xml.rels",
            "xl/timelines/timeline1.xml",
            "xl/timelines/_rels/timeline1.xml.rels",
            "xl/timelineCaches/timelineCacheDefinition1.xml",
            "xl/timelineCaches/_rels/timelineCacheDefinition1.xml.rels",
        ] {
            let original_part = original_pkg
                .part(part)
                .unwrap_or_else(|| panic!("fixture missing required part {part}"));
            let roundtrip_part = roundtrip_pkg
                .part(part)
                .unwrap_or_else(|| panic!("roundtrip missing required part {part}"));
            assert_eq!(original_part, roundtrip_part, "part {part} differs");
        }
    }

    #[test]
    fn roundtrip_preserves_fixture_parts_when_unmodified() {
        let fixture_path = Path::new(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../../fixtures/xlsx/styles/styles.xlsx"
        ));
        let workbook = read_xlsx_blocking(fixture_path).expect("read fixture workbook");

        let tmp = tempfile::tempdir().expect("temp dir");
        let out_path = tmp.path().join("roundtrip.xlsx");
        let _ = write_xlsx_blocking(&out_path, &workbook).expect("write workbook");

        assert_no_critical_diffs(fixture_path, &out_path);
    }

    #[test]
    fn saving_unmodified_workbook_preserves_original_bytes() {
        let fixture_path = Path::new(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../../fixtures/xlsx/styles/styles.xlsx"
        ));
        let original_bytes = std::fs::read(fixture_path).expect("read fixture bytes");
        let workbook = read_xlsx_blocking(fixture_path).expect("read fixture workbook");

        let tmp = tempfile::tempdir().expect("temp dir");
        let out_path = tmp.path().join("roundtrip.xlsx");
        write_xlsx_blocking(&out_path, &workbook).expect("write workbook");

        let written_bytes = std::fs::read(&out_path).expect("read written workbook");
        assert_eq!(original_bytes, written_bytes);
    }

    #[test]
    fn saving_unmodified_xlsm_preserves_original_bytes() {
        let fixture_path = Path::new(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../../fixtures/xlsx/macros/basic.xlsm"
        ));
        let original_bytes = std::fs::read(fixture_path).expect("read fixture bytes");
        let workbook = read_xlsx_blocking(fixture_path).expect("read fixture workbook");

        let tmp = tempfile::tempdir().expect("temp dir");
        let out_path = tmp.path().join("roundtrip.xlsm");
        write_xlsx_blocking(&out_path, &workbook).expect("write workbook");

        let written_bytes = std::fs::read(&out_path).expect("read written workbook");
        assert_eq!(original_bytes, written_bytes);
    }

    #[test]
    fn power_query_part_roundtrip_is_preserved_and_can_be_updated() {
        // Create a minimal XLSX package in memory.
        let mut model = formula_model::Workbook::new();
        model.add_sheet("Sheet1").expect("add sheet");
        let mut cursor = Cursor::new(Vec::new());
        formula_xlsx::write_workbook_to_writer(&model, &mut cursor).expect("write workbook");
        let base_bytes = cursor.into_inner();

        // Save a version without the part so we can assert absent -> added transitions.
        let tmp = tempfile::tempdir().expect("temp dir");
        let no_pq_path = tmp.path().join("no-pq.xlsx");
        std::fs::write(&no_pq_path, &base_bytes).expect("write no-pq");

        let mut workbook = read_xlsx_blocking(&no_pq_path).expect("read workbook without pq");
        assert!(workbook.power_query_xml.is_none());
        assert!(workbook.original_power_query_xml.is_none());

        let added_xml = br#"<FormulaPowerQuery version="1"><![CDATA[{"queries":[{"id":"added"}]}]]></FormulaPowerQuery>"#.to_vec();
        workbook.power_query_xml = Some(added_xml.clone());
        let added_path = tmp.path().join("added.xlsx");
        let added_bytes = write_xlsx_blocking(&added_path, &workbook).expect("write added");
        let added_pkg = XlsxPackage::from_bytes(added_bytes.as_ref()).expect("parse added");
        assert_eq!(
            added_pkg.part(FORMULA_POWER_QUERY_PART),
            Some(added_xml.as_slice()),
            "expected save to add the new power-query.xml payload"
        );

        // Inject a Formula Power Query part.
        let initial_xml = br#"<FormulaPowerQuery version="1"><![CDATA[{"queries":[{"id":"q1"}]}]]></FormulaPowerQuery>"#.to_vec();
        let mut pkg = XlsxPackage::from_bytes(&base_bytes).expect("parse generated package");
        pkg.set_part(FORMULA_POWER_QUERY_PART, initial_xml.clone());
        let injected_bytes = pkg.write_to_bytes().expect("write injected package");

        let src_path = tmp.path().join("src.xlsx");
        std::fs::write(&src_path, &injected_bytes).expect("write src");

        let mut workbook = read_xlsx_blocking(&src_path).expect("read workbook");
        assert_eq!(
            workbook.power_query_xml.as_deref(),
            Some(initial_xml.as_slice())
        );
        assert_eq!(
            workbook.original_power_query_xml.as_deref(),
            Some(initial_xml.as_slice())
        );

        // Patch-based cell edits should preserve non-worksheet parts, including our PQ payload.
        let sheet_id = workbook.sheets[0].id.clone();
        workbook.sheet_mut(&sheet_id).unwrap().set_cell(
            0,
            0,
            Cell::from_literal(Some(CellScalar::Number(1.0))),
        );
        let patched_path = tmp.path().join("patched.xlsx");
        let patched_bytes = write_xlsx_blocking(&patched_path, &workbook).expect("write patched");
        let patched_pkg = XlsxPackage::from_bytes(patched_bytes.as_ref()).expect("parse patched");
        assert_eq!(
            patched_pkg.part(FORMULA_POWER_QUERY_PART),
            Some(initial_xml.as_slice()),
            "expected patch-based save to preserve power-query.xml verbatim"
        );

        // Changing the PQ payload should update the part while keeping the save streaming.
        let mut workbook = read_xlsx_blocking(&src_path).expect("read workbook for update");
        let updated_xml = br#"<FormulaPowerQuery version="1"><![CDATA[{"queries":[{"id":"q2"}]}]]></FormulaPowerQuery>"#.to_vec();
        workbook.power_query_xml = Some(updated_xml.clone());
        let sheet_id = workbook.sheets[0].id.clone();
        workbook.sheet_mut(&sheet_id).unwrap().set_cell(
            0,
            1,
            Cell::from_literal(Some(CellScalar::Number(2.0))),
        );

        let updated_path = tmp.path().join("updated.xlsx");
        let updated_bytes = write_xlsx_blocking(&updated_path, &workbook).expect("write updated");
        let updated_pkg = XlsxPackage::from_bytes(updated_bytes.as_ref()).expect("parse updated");
        assert_eq!(
            updated_pkg.part(FORMULA_POWER_QUERY_PART),
            Some(updated_xml.as_slice()),
            "expected updated save to write the new power-query.xml payload"
        );

        // Removing the PQ payload should delete the part from the package.
        let mut workbook = read_xlsx_blocking(&src_path).expect("read workbook for delete");
        workbook.power_query_xml = None;
        let deleted_path = tmp.path().join("deleted.xlsx");
        let deleted_bytes = write_xlsx_blocking(&deleted_path, &workbook).expect("write deleted");
        let deleted_pkg = XlsxPackage::from_bytes(deleted_bytes.as_ref()).expect("parse deleted");
        assert!(
            deleted_pkg.part(FORMULA_POWER_QUERY_PART).is_none(),
            "expected deleted save to remove power-query.xml"
        );
    }

    #[test]
    fn roundtrip_preserves_comments_when_unmodified() {
        let fixture_path = Path::new(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../../fixtures/xlsx/basic/comments.xlsx"
        ));
        let workbook = read_xlsx_blocking(fixture_path).expect("read comments fixture workbook");

        let tmp = tempfile::tempdir().expect("temp dir");
        let out_path = tmp.path().join("roundtrip.xlsx");
        write_xlsx_blocking(&out_path, &workbook).expect("write workbook");

        let report = xlsx_diff::diff_workbooks(fixture_path, &out_path).expect("diff workbooks");
        assert_eq!(
            report.count(Severity::Critical),
            0,
            "unexpected diffs: {report:?}"
        );
    }

    #[test]
    fn roundtrip_preserves_conditional_formatting_when_unmodified() {
        let fixture_path = Path::new(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../../fixtures/xlsx/conditional-formatting/conditional-formatting.xlsx"
        ));
        let workbook =
            read_xlsx_blocking(fixture_path).expect("read conditional formatting fixture workbook");

        let tmp = tempfile::tempdir().expect("temp dir");
        let out_path = tmp.path().join("roundtrip.xlsx");
        write_xlsx_blocking(&out_path, &workbook).expect("write workbook");

        let report = xlsx_diff::diff_workbooks(fixture_path, &out_path).expect("diff workbooks");
        assert_eq!(
            report.count(Severity::Critical),
            0,
            "unexpected diffs: {report:?}"
        );
    }

    #[test]
    fn cell_edit_preserves_comment_parts() {
        let fixture_path = Path::new(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../../fixtures/xlsx/basic/comments.xlsx"
        ));
        let original_bytes = std::fs::read(fixture_path).expect("read fixture bytes");
        let original_pkg =
            XlsxPackage::from_bytes(&original_bytes).expect("parse original package");
        let original_sheet_xml =
            std::str::from_utf8(original_pkg.part("xl/worksheets/sheet1.xml").unwrap())
                .expect("original sheet1.xml utf8");
        assert!(
            original_sheet_xml.contains("<legacyDrawing"),
            "expected fixture sheet1.xml to contain legacyDrawing for comments"
        );

        let mut workbook =
            read_xlsx_blocking(fixture_path).expect("read comments fixture workbook");

        let sheet_id = workbook.sheets[0].id.clone();
        workbook.sheet_mut(&sheet_id).unwrap().set_cell(
            0,
            0,
            Cell::from_literal(Some(CellScalar::Number(123.0))),
        );

        let tmp = tempfile::tempdir().expect("temp dir");
        let out_path = tmp.path().join("edited.xlsx");
        write_xlsx_blocking(&out_path, &workbook).expect("write workbook");

        let written_bytes = std::fs::read(&out_path).expect("read edited bytes");
        assert_non_worksheet_parts_preserved(&original_bytes, &written_bytes);

        let written_pkg = XlsxPackage::from_bytes(&written_bytes).expect("parse written package");
        let written_sheet_xml =
            std::str::from_utf8(written_pkg.part("xl/worksheets/sheet1.xml").unwrap())
                .expect("written sheet1.xml utf8");
        assert!(
            written_sheet_xml.contains("<legacyDrawing"),
            "expected patched sheet1.xml to retain legacyDrawing for comments"
        );

        let report = xlsx_diff::diff_workbooks(fixture_path, &out_path).expect("diff workbooks");
        assert!(
            !report.differences.iter().any(|d| d.kind == "missing_part"),
            "unexpected missing parts: {:?}",
            report
                .differences
                .iter()
                .filter(|d| d.kind == "missing_part")
                .collect::<Vec<_>>()
        );
        assert!(
            !report.differences.iter().any(|d| d.kind == "extra_part"),
            "unexpected extra parts: {:?}",
            report
                .differences
                .iter()
                .filter(|d| d.kind == "extra_part")
                .collect::<Vec<_>>()
        );

        let unexpected = report
            .differences
            .iter()
            .filter(|d| {
                d.severity != Severity::Info
                    && !d.part.starts_with("xl/worksheets/")
                    && d.part != "xl/sharedStrings.xml"
            })
            .collect::<Vec<_>>();
        assert!(unexpected.is_empty(), "unexpected diffs: {unexpected:?}");
    }

    #[test]
    fn roundtrip_preserves_pivot_parts_when_unmodified() {
        let fixture_path = Path::new(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../../fixtures/xlsx/pivots/pivot-fixture.xlsx"
        ));
        let workbook = read_xlsx_blocking(fixture_path).expect("read pivot fixture workbook");

        let tmp = tempfile::tempdir().expect("temp dir");
        let out_path = tmp.path().join("roundtrip.xlsx");
        write_xlsx_blocking(&out_path, &workbook).expect("write workbook");

        let report = xlsx_diff::diff_workbooks(fixture_path, &out_path).expect("diff workbooks");
        assert_eq!(
            report.count(Severity::Critical),
            0,
            "unexpected diffs: {report:?}"
        );
    }

    #[test]
    fn cell_edit_preserves_pivot_parts() {
        let fixture_path = Path::new(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../../fixtures/xlsx/pivots/pivot-fixture.xlsx"
        ));
        let original_bytes = std::fs::read(fixture_path).expect("read fixture bytes");
        let mut workbook = read_xlsx_blocking(fixture_path).expect("read pivot fixture workbook");

        let sheet_id = workbook.sheets[0].id.clone();
        workbook.sheet_mut(&sheet_id).unwrap().set_cell(
            0,
            0,
            Cell::from_literal(Some(CellScalar::Number(123.0))),
        );

        let tmp = tempfile::tempdir().expect("temp dir");
        let out_path = tmp.path().join("edited.xlsx");
        write_xlsx_blocking(&out_path, &workbook).expect("write workbook");

        let written_bytes = std::fs::read(&out_path).expect("read edited bytes");
        assert_non_worksheet_parts_preserved(&original_bytes, &written_bytes);

        let report = xlsx_diff::diff_workbooks(fixture_path, &out_path).expect("diff workbooks");
        assert!(
            !report.differences.iter().any(|d| d.kind == "missing_part"),
            "unexpected missing parts: {:?}",
            report
                .differences
                .iter()
                .filter(|d| d.kind == "missing_part")
                .collect::<Vec<_>>()
        );
        assert!(
            !report.differences.iter().any(|d| d.kind == "extra_part"),
            "unexpected extra parts: {:?}",
            report
                .differences
                .iter()
                .filter(|d| d.kind == "extra_part")
                .collect::<Vec<_>>()
        );

        let unexpected = report
            .differences
            .iter()
            .filter(|d| {
                d.severity != Severity::Info
                    && !d.part.starts_with("xl/worksheets/")
                    && d.part != "xl/sharedStrings.xml"
            })
            .collect::<Vec<_>>();
        assert!(unexpected.is_empty(), "unexpected diffs: {unexpected:?}");
    }

    #[test]
    fn cell_edit_preserves_image_parts() {
        let fixture_path = Path::new(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../../fixtures/xlsx/basic/image.xlsx"
        ));
        let original_bytes = std::fs::read(fixture_path).expect("read fixture bytes");
        let original_pkg =
            XlsxPackage::from_bytes(&original_bytes).expect("parse original package");
        let original_sheet_xml =
            std::str::from_utf8(original_pkg.part("xl/worksheets/sheet1.xml").unwrap())
                .expect("original sheet1.xml utf8");
        assert!(
            original_sheet_xml.contains("<drawing"),
            "expected fixture sheet1.xml to contain a drawing relationship"
        );

        let mut workbook = read_xlsx_blocking(fixture_path).expect("read image fixture workbook");

        let sheet_id = workbook.sheets[0].id.clone();
        workbook.sheet_mut(&sheet_id).unwrap().set_cell(
            1,
            1,
            Cell::from_literal(Some(CellScalar::Number(42.0))),
        );

        let tmp = tempfile::tempdir().expect("temp dir");
        let out_path = tmp.path().join("edited.xlsx");
        let _ = write_xlsx_blocking(&out_path, &workbook).expect("write workbook");

        let written_bytes = std::fs::read(&out_path).expect("read edited bytes");
        assert_ne!(
            original_bytes, written_bytes,
            "expected worksheet patching to produce a different file"
        );

        assert_non_worksheet_parts_preserved(&original_bytes, &written_bytes);

        let written_pkg = XlsxPackage::from_bytes(&written_bytes).expect("parse written package");
        let written_sheet_xml =
            std::str::from_utf8(written_pkg.part("xl/worksheets/sheet1.xml").unwrap())
                .expect("written sheet1.xml utf8");
        assert!(
            written_sheet_xml.contains("<drawing"),
            "expected patched sheet1.xml to retain drawing relationship"
        );

        let written = read_xlsx_blocking(&out_path).expect("read edited workbook");
        assert_eq!(
            written.sheets[0].get_cell(1, 1).computed_value,
            CellScalar::Number(42.0)
        );
    }

    #[test]
    fn cell_edit_preserves_hyperlink_relationships() {
        let fixture_path = Path::new(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../../fixtures/xlsx/hyperlinks/hyperlinks.xlsx"
        ));
        let original_bytes = std::fs::read(fixture_path).expect("read fixture bytes");
        let original_pkg =
            XlsxPackage::from_bytes(&original_bytes).expect("parse original package");
        let original_sheet_xml =
            std::str::from_utf8(original_pkg.part("xl/worksheets/sheet1.xml").unwrap())
                .expect("original sheet1.xml utf8");
        assert!(
            original_sheet_xml.contains("<hyperlinks>"),
            "expected fixture sheet1.xml to contain hyperlinks"
        );

        let mut workbook =
            read_xlsx_blocking(fixture_path).expect("read hyperlinks fixture workbook");

        let sheet_id = workbook.sheets[0].id.clone();
        workbook.sheet_mut(&sheet_id).unwrap().set_cell(
            0,
            1,
            Cell::from_literal(Some(CellScalar::Number(7.0))),
        );

        let tmp = tempfile::tempdir().expect("temp dir");
        let out_path = tmp.path().join("edited.xlsx");
        let _ = write_xlsx_blocking(&out_path, &workbook).expect("write workbook");

        let written_bytes = std::fs::read(&out_path).expect("read edited bytes");
        assert_ne!(
            original_bytes, written_bytes,
            "expected worksheet patching to produce a different file"
        );

        assert_non_worksheet_parts_preserved(&original_bytes, &written_bytes);

        let written_pkg = XlsxPackage::from_bytes(&written_bytes).expect("parse written package");
        let written_sheet_xml =
            std::str::from_utf8(written_pkg.part("xl/worksheets/sheet1.xml").unwrap())
                .expect("written sheet1.xml utf8");
        assert!(
            written_sheet_xml.contains("<hyperlinks>"),
            "expected patched sheet1.xml to retain hyperlinks"
        );
        assert!(
            written_sheet_xml.contains("ref=\"A1\""),
            "expected patched sheet1.xml to retain hyperlink refs"
        );

        let written = read_xlsx_blocking(&out_path).expect("read edited workbook");
        assert_eq!(
            written.sheets[0].get_cell(0, 1).computed_value,
            CellScalar::Number(7.0)
        );
    }

    #[test]
    fn read_xlsx_populates_defined_names() {
        let fixture_path = Path::new(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../../fixtures/xlsx/metadata/defined-names.xlsx"
        ));

        let workbook =
            read_xlsx_blocking(fixture_path).expect("read defined-names fixture workbook");
        assert!(
            workbook.defined_names.iter().any(|n| n.name == "ZedName"),
            "expected defined name ZedName, got: {:?}",
            workbook
                .defined_names
                .iter()
                .map(|n| n.name.as_str())
                .collect::<Vec<_>>()
        );
        assert!(
            workbook.defined_names.iter().any(|n| n.name == "MyRange"),
            "expected defined name MyRange, got: {:?}",
            workbook
                .defined_names
                .iter()
                .map(|n| n.name.as_str())
                .collect::<Vec<_>>()
        );

        let zed = workbook
            .defined_names
            .iter()
            .find(|n| n.name == "ZedName")
            .expect("ZedName exists");
        assert_eq!(zed.refers_to, "Sheet1!$B$1");
        assert!(
            zed.sheet_id.is_none(),
            "expected ZedName to be workbook-scoped"
        );
    }

    #[test]
    fn read_xlsx_populates_tables() {
        let mut model = formula_model::Workbook::new();
        let sheet_id = model.add_sheet("Sheet1").unwrap();

        let table = formula_model::Table {
            id: 1,
            name: "Table1".to_string(),
            display_name: "Table1".to_string(),
            range: formula_model::Range::from_a1("A1:B3").unwrap(),
            header_row_count: 1,
            totals_row_count: 0,
            columns: vec![
                formula_model::TableColumn {
                    id: 1,
                    name: "Amount".to_string(),
                    formula: None,
                    totals_formula: None,
                },
                formula_model::TableColumn {
                    id: 2,
                    name: "Category".to_string(),
                    formula: None,
                    totals_formula: None,
                },
            ],
            style: None,
            auto_filter: None,
            relationship_id: None,
            part_path: None,
        };

        model.add_table(sheet_id, table).unwrap();

        let tmp = tempfile::tempdir().expect("temp dir");
        let out_path = tmp.path().join("tables.xlsx");
        formula_xlsx::write_workbook(&model, &out_path).expect("write workbook with table");

        let workbook = read_xlsx_blocking(&out_path).expect("read workbook back");
        assert_eq!(workbook.tables.len(), 1);

        let t = &workbook.tables[0];
        assert_eq!(t.name, "Table1");
        assert_eq!(t.sheet_id, "Sheet1");
        assert_eq!(t.start_row, 0);
        assert_eq!(t.start_col, 0);
        assert_eq!(t.end_row, 2);
        assert_eq!(t.end_col, 1);
        assert_eq!(
            t.columns,
            vec!["Amount".to_string(), "Category".to_string()]
        );
    }

    #[test]
    fn regeneration_roundtrip_preserves_defined_names_and_tables() {
        let mut workbook = Workbook::new_empty(None);
        workbook.add_sheet("Sheet1".to_string());
        workbook.add_sheet("Sheet2".to_string());
        let sheet1_id = workbook.sheets[0].id.clone();
        let sheet2_id = workbook.sheets[1].id.clone();

        {
            let sheet = workbook.sheet_mut(&sheet1_id).expect("Sheet1 exists");
            sheet.set_cell(
                0,
                0,
                Cell::from_literal(Some(CellScalar::Text("Amount".to_string()))),
            );
            sheet.set_cell(
                0,
                1,
                Cell::from_literal(Some(CellScalar::Text("Category".to_string()))),
            );
            sheet.set_cell(1, 0, Cell::from_literal(Some(CellScalar::Number(10.0))));
            sheet.set_cell(
                1,
                1,
                Cell::from_literal(Some(CellScalar::Text("Food".to_string()))),
            );
            sheet.set_cell(2, 0, Cell::from_literal(Some(CellScalar::Number(5.0))));
            sheet.set_cell(
                2,
                1,
                Cell::from_literal(Some(CellScalar::Text("Other".to_string()))),
            );

            let mut formula_cell = Cell::from_formula("=SUM(A2:A3)".to_string());
            formula_cell.computed_value = CellScalar::Number(15.0);
            sheet.set_cell(0, 3, formula_cell);
        }

        workbook.defined_names.push(DefinedName {
            name: "MyRange".to_string(),
            refers_to: "Sheet1!$A$2:$A$3".to_string(),
            sheet_id: None,
            hidden: false,
        });
        workbook.defined_names.push(DefinedName {
            name: "LocalName".to_string(),
            refers_to: "Sheet2!$A$1".to_string(),
            sheet_id: Some(sheet2_id.clone()),
            hidden: false,
        });

        workbook.tables.push(Table {
            name: "Table1".to_string(),
            sheet_id: sheet1_id.clone(),
            start_row: 0,
            start_col: 0,
            end_row: 2,
            end_col: 1,
            columns: vec!["Amount".to_string(), "Category".to_string()],
        });

        assert!(
            workbook.origin_xlsx_bytes.is_none(),
            "expected regeneration path"
        );
        let tmp = tempfile::tempdir().expect("temp dir");
        let out_path = tmp.path().join("regen.xlsx");
        write_xlsx_blocking(&out_path, &workbook).expect("write workbook");

        let written = read_xlsx_blocking(&out_path).expect("read workbook back");

        let my_range = written
            .defined_names
            .iter()
            .find(|n| n.name == "MyRange")
            .expect("MyRange defined name exists");
        assert_eq!(my_range.refers_to, "Sheet1!$A$2:$A$3");
        assert!(my_range.sheet_id.is_none());

        let local_name = written
            .defined_names
            .iter()
            .find(|n| n.name == "LocalName")
            .expect("LocalName defined name exists");
        assert_eq!(local_name.refers_to, "Sheet2!$A$1");
        assert_eq!(local_name.sheet_id.as_deref(), Some("Sheet2"));

        assert_eq!(written.tables.len(), 1);
        let table = &written.tables[0];
        assert_eq!(table.name, "Table1");
        assert_eq!(table.sheet_id, "Sheet1");
        assert_eq!(table.start_row, 0);
        assert_eq!(table.start_col, 0);
        assert_eq!(table.end_row, 2);
        assert_eq!(table.end_col, 1);
        assert_eq!(
            table.columns,
            vec!["Amount".to_string(), "Category".to_string()]
        );

        assert_eq!(
            written.sheets[0].get_cell(0, 3).formula.as_deref(),
            Some("=SUM(A2:A3)")
        );
        assert_eq!(
            written.sheets[0].get_cell(0, 3).computed_value,
            CellScalar::Number(15.0)
        );
    }

    #[test]
    fn cell_edit_preserves_defined_names() {
        let fixture_path = Path::new(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../../fixtures/xlsx/metadata/defined-names.xlsx"
        ));
        let original_bytes = std::fs::read(fixture_path).expect("read fixture bytes");
        let mut workbook =
            read_xlsx_blocking(fixture_path).expect("read defined-names fixture workbook");

        let sheet_id = workbook.sheets[0].id.clone();
        workbook.sheet_mut(&sheet_id).unwrap().set_cell(
            0,
            2,
            Cell::from_literal(Some(CellScalar::Number(99.0))),
        );

        let tmp = tempfile::tempdir().expect("temp dir");
        let out_path = tmp.path().join("edited.xlsx");
        let _ = write_xlsx_blocking(&out_path, &workbook).expect("write workbook");

        let written_bytes = std::fs::read(&out_path).expect("read edited bytes");
        assert_ne!(
            original_bytes, written_bytes,
            "expected worksheet patching to produce a different file"
        );

        assert_non_worksheet_parts_preserved(&original_bytes, &written_bytes);

        let written = read_xlsx_blocking(&out_path).expect("read edited workbook");
        assert_eq!(
            written.sheets[0].get_cell(0, 2).computed_value,
            CellScalar::Number(99.0)
        );
    }

    #[test]
    fn cell_edit_preserves_external_link_parts() {
        let fixture_path = Path::new(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../../fixtures/xlsx/metadata/external-link.xlsx"
        ));
        let original_bytes = std::fs::read(fixture_path).expect("read fixture bytes");
        let mut workbook =
            read_xlsx_blocking(fixture_path).expect("read external-link fixture workbook");

        let sheet_id = workbook.sheets[0].id.clone();
        workbook.sheet_mut(&sheet_id).unwrap().set_cell(
            0,
            1,
            Cell::from_literal(Some(CellScalar::Number(5.0))),
        );

        let tmp = tempfile::tempdir().expect("temp dir");
        let out_path = tmp.path().join("edited.xlsx");
        let _ = write_xlsx_blocking(&out_path, &workbook).expect("write workbook");

        let written_bytes = std::fs::read(&out_path).expect("read edited bytes");
        assert_ne!(
            original_bytes, written_bytes,
            "expected worksheet patching to produce a different file"
        );

        assert_non_worksheet_parts_preserved(&original_bytes, &written_bytes);

        let written = read_xlsx_blocking(&out_path).expect("read edited workbook");
        assert_eq!(
            written.sheets[0].get_cell(0, 1).computed_value,
            CellScalar::Number(5.0)
        );
    }

    #[test]
    fn cell_edit_only_changes_worksheet_parts() {
        let fixture_path = Path::new(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../../fixtures/xlsx/styles/styles.xlsx"
        ));
        let mut workbook = read_xlsx_blocking(fixture_path).expect("read fixture workbook");

        let sheet_id = workbook.sheets[0].id.clone();
        workbook.sheet_mut(&sheet_id).unwrap().set_cell(
            0,
            0,
            Cell::from_literal(Some(CellScalar::Number(123.0))),
        );

        let tmp = tempfile::tempdir().expect("temp dir");
        let out_path = tmp.path().join("edited.xlsx");
        let _ = write_xlsx_blocking(&out_path, &workbook).expect("write workbook");

        let report = diff_workbooks(fixture_path, &out_path).expect("diff workbooks");
        assert!(
            !report.differences.iter().any(|d| d.kind == "missing_part"),
            "unexpected missing parts: {:?}",
            report
                .differences
                .iter()
                .filter(|d| d.kind == "missing_part")
                .collect::<Vec<_>>()
        );
        assert!(
            !report.differences.iter().any(|d| d.kind == "extra_part"),
            "unexpected extra parts: {:?}",
            report
                .differences
                .iter()
                .filter(|d| d.kind == "extra_part")
                .collect::<Vec<_>>()
        );

        let unexpected = report
            .differences
            .iter()
            .filter(|d| {
                d.severity != Severity::Info
                    && !d.part.starts_with("xl/worksheets/")
                    && d.part != "xl/sharedStrings.xml"
            })
            .collect::<Vec<_>>();
        assert!(unexpected.is_empty(), "unexpected diffs: {unexpected:?}");

        let written = read_xlsx_blocking(&out_path).expect("read edited workbook");
        assert_eq!(
            written.sheets[0].get_cell(0, 0).computed_value,
            CellScalar::Number(123.0)
        );
    }

    #[test]
    fn xlsm_cell_edit_preserves_vba_parts() {
        let fixture_path = Path::new(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../../fixtures/xlsx/macros/basic.xlsm"
        ));

        let original_bytes = std::fs::read(fixture_path).expect("read fixture xlsm");
        let original_pkg = XlsxPackage::from_bytes(&original_bytes).expect("parse fixture package");
        let original_vba = original_pkg
            .vba_project_bin()
            .expect("fixture has vbaProject.bin")
            .to_vec();

        let mut workbook = read_xlsx_blocking(fixture_path).expect("read fixture workbook");
        let sheet_id = workbook.sheets[0].id.clone();
        workbook.sheet_mut(&sheet_id).unwrap().set_cell(
            0,
            0,
            Cell::from_literal(Some(CellScalar::Number(7.0))),
        );

        let tmp = tempfile::tempdir().expect("temp dir");
        let out_path = tmp.path().join("edited.xlsm");
        let _ = write_xlsx_blocking(&out_path, &workbook).expect("write workbook");

        let written_bytes = std::fs::read(&out_path).expect("read written xlsm");
        let written_pkg = XlsxPackage::from_bytes(&written_bytes).expect("parse written package");
        let written_vba = written_pkg
            .vba_project_bin()
            .expect("written workbook should contain vbaProject.bin");
        assert_eq!(original_vba, written_vba);

        let report = diff_workbooks(fixture_path, &out_path).expect("diff workbooks");
        assert!(
            !report.differences.iter().any(|d| d.kind == "missing_part"),
            "unexpected missing parts"
        );
        let unexpected = report
            .differences
            .iter()
            .filter(|d| {
                d.severity != Severity::Info
                    && !d.part.starts_with("xl/worksheets/")
                    && d.part != "xl/sharedStrings.xml"
            })
            .collect::<Vec<_>>();
        assert!(unexpected.is_empty(), "unexpected diffs: {unexpected:?}");
    }

    #[test]
    fn saving_xlsm_as_xlsx_drops_vba_project() {
        let fixture_path = Path::new(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../../fixtures/xlsx/macros/basic.xlsm"
        ));
        let workbook = read_xlsx_blocking(fixture_path).expect("read fixture workbook");

        let tmp = tempfile::tempdir().expect("temp dir");
        let out_path = tmp.path().join("converted.xlsx");
        let _ = write_xlsx_blocking(&out_path, &workbook).expect("write workbook");

        let written_bytes = std::fs::read(&out_path).expect("read written xlsx");
        let written_pkg = XlsxPackage::from_bytes(&written_bytes).expect("parse written package");

        assert!(
            written_pkg.vba_project_bin().is_none(),
            "expected vbaProject.bin to be removed when saving as .xlsx"
        );

        let ct = std::str::from_utf8(written_pkg.part("[Content_Types].xml").unwrap()).unwrap();
        assert!(
            !ct.contains("vbaProject.bin"),
            "expected [Content_Types].xml to drop vbaProject.bin override"
        );
        assert!(
            !ct.contains("macroEnabled.main+xml"),
            "expected workbook content type to be converted back to .xlsx"
        );

        let rels =
            std::str::from_utf8(written_pkg.part("xl/_rels/workbook.xml.rels").unwrap()).unwrap();
        assert!(
            !rels.contains("relationships/vbaProject"),
            "expected workbook.xml.rels to drop the vbaProject relationship"
        );

        // Ensure macro stripping doesn't perturb unrelated parts.
        let mut ignore_parts = BTreeSet::new();
        ignore_parts.insert("xl/vbaProject.bin".to_string());
        ignore_parts.insert("[Content_Types].xml".to_string());
        ignore_parts.insert("xl/_rels/workbook.xml.rels".to_string());
        let options = DiffOptions {
            ignore_parts,
            ignore_globs: Vec::new(),
            ignore_paths: Vec::new(),
            strict_calc_chain: false,
        };
        let report =
            diff_workbooks_with_options(fixture_path, &out_path, &options).expect("diff workbooks");
        assert_eq!(
            report.count(Severity::Critical),
            0,
            "unexpected critical diffs after macro stripping: {report:?}"
        );
    }

    #[test]
    fn saving_xlsm_as_xlsx_drops_vba_project_and_applies_power_query_override() {
        let fixture_path = Path::new(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../../fixtures/xlsx/macros/basic.xlsm"
        ));
        let mut workbook = read_xlsx_blocking(fixture_path).expect("read fixture workbook");

        // Force the save path to apply a streaming part override in addition to macro stripping.
        let power_query_xml = br#"<FormulaPowerQuery version="1"><![CDATA[{"queries":[{"id":"q1"}]}]]></FormulaPowerQuery>"#.to_vec();
        workbook.power_query_xml = Some(power_query_xml.clone());

        let tmp = tempfile::tempdir().expect("temp dir");
        let out_path = tmp.path().join("converted-with-pq.xlsx");
        write_xlsx_blocking(&out_path, &workbook).expect("write workbook");

        let written_bytes = std::fs::read(&out_path).expect("read written xlsx");
        let written_pkg = XlsxPackage::from_bytes(&written_bytes).expect("parse written package");

        assert!(
            written_pkg.vba_project_bin().is_none(),
            "expected vbaProject.bin to be removed when saving as .xlsx"
        );
        assert_eq!(
            written_pkg.part(FORMULA_POWER_QUERY_PART),
            Some(power_query_xml.as_slice()),
            "expected power-query.xml override to be applied"
        );

        // Ensure macro stripping doesn't perturb unrelated parts.
        let mut ignore_parts = BTreeSet::new();
        ignore_parts.insert("xl/vbaProject.bin".to_string());
        ignore_parts.insert("[Content_Types].xml".to_string());
        ignore_parts.insert("xl/_rels/workbook.xml.rels".to_string());
        ignore_parts.insert(FORMULA_POWER_QUERY_PART.to_string());
        let options = DiffOptions {
            ignore_parts,
            ignore_globs: Vec::new(),
            ignore_paths: Vec::new(),
            strict_calc_chain: false,
        };
        let report =
            diff_workbooks_with_options(fixture_path, &out_path, &options).expect("diff workbooks");
        assert_eq!(
            report.count(Severity::Critical),
            0,
            "unexpected critical diffs after macro stripping + power query override: {report:?}"
        );
    }

    #[test]
    fn saving_xlsm_as_xltx_xltm_xlam_sets_expected_content_type_and_macro_parts() {
        let fixture_path = Path::new(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../../fixtures/xlsx/macros/basic.xlsm"
        ));
        let workbook = read_xlsx_blocking(fixture_path).expect("read fixture workbook");
        assert!(
            workbook.vba_project_bin.is_some(),
            "fixture should contain vbaProject.bin"
        );

        let tmp = tempfile::tempdir().expect("temp dir");

        for (ext, expects_vba, expected_ct) in [
            ("xltx", false, XLTX_WORKBOOK_CONTENT_TYPE),
            ("xltm", true, XLTM_WORKBOOK_CONTENT_TYPE),
            ("xlam", true, XLAM_WORKBOOK_CONTENT_TYPE),
        ] {
            let out_path = tmp.path().join(format!("converted.{ext}"));
            write_xlsx_blocking(&out_path, &workbook).expect("write workbook");

            let written_bytes = std::fs::read(&out_path).expect("read written bytes");
            let written_pkg =
                XlsxPackage::from_bytes(&written_bytes).expect("parse written package");

            assert_eq!(
                written_pkg.vba_project_bin().is_some(),
                expects_vba,
                "expected vbaProject.bin presence to be {expects_vba} for .{ext}"
            );

            let ct = std::str::from_utf8(written_pkg.part("[Content_Types].xml").unwrap())
                .expect("[Content_Types].xml should be utf8");
            assert!(
                workbook_override_matches_content_type(ct, expected_ct),
                "expected workbook override content type {expected_ct} for .{ext}, got: {ct}"
            );

            if ext == "xltx" {
                let rels =
                    std::str::from_utf8(written_pkg.part("xl/_rels/workbook.xml.rels").unwrap())
                        .expect("workbook.xml.rels should be utf8");
                assert!(
                    !rels.contains("relationships/vbaProject"),
                    "expected workbook.xml.rels to drop the vbaProject relationship"
                );
            }
        }
    }

    #[test]
    fn saving_xlsx_as_xltx_enforces_template_content_type() {
        let fixture_path = Path::new(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../../fixtures/xlsx/basic/basic.xlsx"
        ));
        let workbook = read_xlsx_blocking(fixture_path).expect("read fixture workbook");

        let tmp = tempfile::tempdir().expect("temp dir");
        let out_path = tmp.path().join("converted.xltx");
        write_xlsx_blocking(&out_path, &workbook).expect("write workbook");

        let written_bytes = std::fs::read(&out_path).expect("read written xltx");
        let written_pkg = XlsxPackage::from_bytes(&written_bytes).expect("parse written package");

        assert!(
            written_pkg.vba_project_bin().is_none(),
            "expected vbaProject.bin to remain absent when saving .xlsx as .xltx"
        );

        let ct = std::str::from_utf8(written_pkg.part("[Content_Types].xml").unwrap())
            .expect("[Content_Types].xml should be utf8");
        assert!(
            workbook_override_matches_content_type(ct, XLTX_WORKBOOK_CONTENT_TYPE),
            "expected workbook override content type {XLTX_WORKBOOK_CONTENT_TYPE} for .xltx, got: {ct}"
        );
    }

    #[test]
    fn storage_export_supports_xltx_xltm_xlam_macros_content_type_and_print_settings() {
        use formula_storage::ImportModelWorkbookOptions;

        let fixture_path = Path::new(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../../fixtures/xlsx/macros/basic.xlsm"
        ));
        let macro_wb = read_xlsx_blocking(fixture_path).expect("read xlsm fixture");
        let vba = macro_wb
            .vba_project_bin
            .clone()
            .expect("fixture has vbaProject.bin");

        // Create a minimal workbook in storage so we can exercise the storage export path.
        let mut model = formula_model::Workbook::new();
        model.add_sheet("Sheet1").expect("add sheet");
        let storage = formula_storage::Storage::open_in_memory().expect("open in-memory storage");
        let stored = storage
            .import_model_workbook(&model, ImportModelWorkbookOptions::new("test"))
            .expect("import model workbook");

        // App workbook metadata drives VBA + print settings in the export path.
        let mut workbook_meta = Workbook::new_empty(None);
        workbook_meta.add_sheet("Sheet1".to_string());
        workbook_meta.vba_project_bin = Some(vba);
        workbook_meta.print_settings = WorkbookPrintSettings {
            sheets: vec![formula_xlsx::print::SheetPrintSettings {
                sheet_name: "Sheet1".to_string(),
                print_area: None,
                print_titles: None,
                page_setup: formula_xlsx::print::PageSetup {
                    orientation: formula_xlsx::print::Orientation::Landscape,
                    ..Default::default()
                },
                manual_page_breaks: formula_xlsx::print::ManualPageBreaks::default(),
            }],
        };

        let tmp = tempfile::tempdir().expect("temp dir");

        for (ext, expects_vba, expected_ct) in [
            ("xltx", false, XLTX_WORKBOOK_CONTENT_TYPE),
            ("xltm", true, XLTM_WORKBOOK_CONTENT_TYPE),
            ("xlam", true, XLAM_WORKBOOK_CONTENT_TYPE),
        ] {
            let out_path = tmp.path().join(format!("exported.{ext}"));
            let bytes = crate::persistence::write_xlsx_from_storage(
                &storage,
                stored.id,
                &workbook_meta,
                &out_path,
            )
            .expect("write xlsx from storage");

            let pkg = XlsxPackage::from_bytes(bytes.as_ref()).expect("parse exported package");
            assert_eq!(
                pkg.vba_project_bin().is_some(),
                expects_vba,
                "expected vbaProject.bin presence to be {expects_vba} for .{ext}"
            );

            let ct = std::str::from_utf8(pkg.part("[Content_Types].xml").unwrap())
                .expect("[Content_Types].xml should be utf8");
            assert!(
                workbook_override_matches_content_type(ct, expected_ct),
                "expected workbook override content type {expected_ct} for .{ext}, got: {ct}"
            );

            if ext == "xltx" {
                // Assert print settings were applied for template output (storage export path).
                let settings = read_workbook_print_settings(bytes.as_ref())
                    .expect("read workbook print settings");
                let sheet = settings
                    .sheets
                    .iter()
                    .find(|s| s.sheet_name == "Sheet1")
                    .expect("Sheet1 print settings present");
                assert_eq!(
                    sheet.page_setup.orientation,
                    formula_xlsx::print::Orientation::Landscape
                );
            }
        }
    }

    #[test]
    fn open_save_xlsx_is_lossless_basic() {
        let fixture_path = Path::new(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../../fixtures/xlsx/basic/basic.xlsx"
        ));
        let workbook = read_xlsx_blocking(fixture_path).expect("read workbook");

        let tmp = tempfile::tempdir().expect("temp dir");
        let out_path = tmp.path().join("roundtrip.xlsx");
        let _ = write_xlsx_blocking(&out_path, &workbook).expect("write workbook");

        assert_no_critical_diffs(fixture_path, &out_path);
    }

    #[test]
    fn open_save_xlsx_is_lossless_metadata() {
        let fixture_path = Path::new(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../../fixtures/xlsx/metadata/row-col-properties.xlsx"
        ));
        let workbook = read_xlsx_blocking(fixture_path).expect("read workbook");

        let tmp = tempfile::tempdir().expect("temp dir");
        let out_path = tmp.path().join("roundtrip.xlsx");
        let _ = write_xlsx_blocking(&out_path, &workbook).expect("write workbook");

        assert_no_critical_diffs(fixture_path, &out_path);
    }

    #[test]
    fn open_save_xlsx_is_lossless_conditional_formatting() {
        let fixture_path = Path::new(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../../fixtures/xlsx/conditional-formatting/conditional-formatting.xlsx"
        ));
        let workbook = read_xlsx_blocking(fixture_path).expect("read workbook");

        let tmp = tempfile::tempdir().expect("temp dir");
        let out_path = tmp.path().join("roundtrip.xlsx");
        let _ = write_xlsx_blocking(&out_path, &workbook).expect("write workbook");

        assert_no_critical_diffs(fixture_path, &out_path);
    }

    #[test]
    fn open_save_xlsm_is_lossless_and_preserves_vba() {
        let fixture_path = Path::new(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../../fixtures/xlsx/macros/basic.xlsm"
        ));

        let original_bytes = std::fs::read(fixture_path).expect("read fixture xlsm");
        let original_pkg = XlsxPackage::from_bytes(&original_bytes).expect("parse fixture package");
        let original_vba = original_pkg
            .vba_project_bin()
            .expect("fixture has vbaProject.bin")
            .to_vec();

        let workbook = read_xlsx_blocking(fixture_path).expect("read workbook");
        let tmp = tempfile::tempdir().expect("temp dir");
        let out_path = tmp.path().join("roundtrip.xlsm");
        let _ = write_xlsx_blocking(&out_path, &workbook).expect("write workbook");

        assert_no_critical_diffs(fixture_path, &out_path);

        let written_bytes = std::fs::read(&out_path).expect("read written file");
        let written_pkg = XlsxPackage::from_bytes(&written_bytes).expect("parse written package");
        let written_vba = written_pkg
            .vba_project_bin()
            .expect("written workbook should contain vbaProject.bin")
            .to_vec();

        assert_eq!(original_vba, written_vba);
    }

    #[test]
    fn edited_cell_persists_after_save() {
        use serde_json::Value as JsonValue;

        let fixture_path = Path::new(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../../fixtures/xlsx/basic/basic.xlsx"
        ));
        let workbook = read_xlsx_blocking(fixture_path).expect("read workbook");

        let mut state = AppState::new();
        let info = state.load_workbook(workbook);
        let sheet_id = info.sheets[0].id.clone();

        state
            .set_cell(&sheet_id, 0, 0, Some(JsonValue::from(123)), None)
            .expect("edit cell");

        let tmp = tempfile::tempdir().expect("temp dir");
        let out_path = tmp.path().join("edited.xlsx");
        let written_bytes =
            write_xlsx_blocking(&out_path, state.get_workbook().unwrap()).expect("write workbook");

        let doc = formula_xlsx::load_from_bytes(written_bytes.as_ref())
            .expect("load saved workbook from bytes");
        let sheet = doc
            .workbook
            .sheet_by_name("Sheet1")
            .or_else(|| doc.workbook.sheets.first())
            .expect("sheet exists");

        assert_eq!(
            sheet.value(formula_model::CellRef::new(0, 0)),
            ModelCellValue::Number(123.0)
        );
    }

    #[test]
    fn edit_then_revert_does_not_change_workbook() {
        use serde_json::Value as JsonValue;

        let fixture_path = Path::new(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../../fixtures/xlsx/basic/basic.xlsx"
        ));
        let workbook = read_xlsx_blocking(fixture_path).expect("read workbook");

        let mut state = AppState::new();
        let info = state.load_workbook(workbook);
        let sheet_id = info.sheets[0].id.clone();

        // Pick a cell outside the used range so we can reliably "return to empty".
        state
            .set_cell(&sheet_id, 50, 50, Some(JsonValue::from(123)), None)
            .expect("edit cell");
        state
            .set_cell(&sheet_id, 50, 50, None, None)
            .expect("revert cell");

        let tmp = tempfile::tempdir().expect("temp dir");
        let out_path = tmp.path().join("reverted.xlsx");
        let _ =
            write_xlsx_blocking(&out_path, state.get_workbook().unwrap()).expect("write workbook");

        assert_no_critical_diffs(fixture_path, &out_path);
    }

    #[test]
    fn print_settings_update_is_applied() {
        let fixture_path = Path::new(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../../fixtures/xlsx/basic/print-settings.xlsx"
        ));
        let original_bytes = std::fs::read(fixture_path).expect("read fixture bytes");
        let mut workbook = read_xlsx_blocking(fixture_path).expect("read print-settings workbook");
        assert_eq!(workbook.print_settings.sheets.len(), 1);

        workbook.print_settings.sheets[0].page_setup.orientation =
            formula_xlsx::print::Orientation::Portrait;

        let tmp = tempfile::tempdir().expect("temp dir");
        let out_path = tmp.path().join("print-settings-updated.xlsx");
        write_xlsx_blocking(&out_path, &workbook).expect("write workbook");

        let written_bytes = std::fs::read(&out_path).expect("read written bytes");
        assert_ne!(
            original_bytes, written_bytes,
            "expected print settings change to rewrite the workbook"
        );

        let settings =
            read_workbook_print_settings(&written_bytes).expect("read workbook print settings");
        assert_eq!(settings.sheets.len(), 1);
        assert_eq!(
            settings.sheets[0].page_setup.orientation,
            formula_xlsx::print::Orientation::Portrait
        );

        let original_pkg =
            XlsxPackage::from_bytes(&original_bytes).expect("parse original package");
        let written_pkg = XlsxPackage::from_bytes(&written_bytes).expect("parse written package");
        for part in [
            "[Content_Types].xml",
            "_rels/.rels",
            "xl/_rels/workbook.xml.rels",
        ] {
            assert_eq!(
                original_pkg.part(part),
                written_pkg.part(part),
                "expected {part} to be preserved when updating print settings"
            );
        }

        let workbook_xml = std::str::from_utf8(written_pkg.part("xl/workbook.xml").unwrap())
            .expect("written workbook.xml utf8");
        assert!(
            workbook_xml.contains("_xlnm.Print_Area"),
            "expected print area defined name to remain present"
        );
        assert!(
            workbook_xml.contains("_xlnm.Print_Titles"),
            "expected print titles defined name to remain present"
        );
    }

    #[test]
    fn print_settings_edit_then_revert_does_not_change_workbook() {
        let fixture_path = Path::new(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../../fixtures/xlsx/basic/print-settings.xlsx"
        ));
        let original_bytes = std::fs::read(fixture_path).expect("read fixture bytes");
        let mut workbook = read_xlsx_blocking(fixture_path).expect("read print-settings workbook");

        workbook.print_settings.sheets[0].page_setup.orientation =
            formula_xlsx::print::Orientation::Portrait;
        // Restore baseline to ensure we don't churn the workbook when the user reverts changes.
        workbook.print_settings = workbook.original_print_settings.clone();

        let tmp = tempfile::tempdir().expect("temp dir");
        let out_path = tmp.path().join("print-settings-reverted.xlsx");
        write_xlsx_blocking(&out_path, &workbook).expect("write workbook");

        let written_bytes = std::fs::read(&out_path).expect("read written bytes");
        assert_eq!(original_bytes, written_bytes);
    }

    #[test]
    fn write_xlsx_blocking_replaces_existing_file() {
        let fixture_path = Path::new(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../../fixtures/xlsx/basic/basic.xlsx"
        ));
        let workbook = read_xlsx_blocking(fixture_path).expect("read fixture workbook");
        let expected = workbook
            .origin_xlsx_bytes
            .as_ref()
            .expect("origin bytes should be captured for xlsx inputs")
            .clone();

        let tmp = tempfile::tempdir().expect("temp dir");
        let out_path = tmp.path().join("existing.xlsx");
        std::fs::write(&out_path, b"old-bytes").expect("seed existing file");

        let written = write_xlsx_blocking(&out_path, &workbook).expect("write workbook");
        let file_bytes = std::fs::read(&out_path).expect("read written file");

        assert_eq!(file_bytes.as_slice(), written.as_ref());
        assert_eq!(file_bytes.as_slice(), expected.as_ref());
    }

    #[test]
    fn write_xlsx_blocking_creates_parent_dirs_for_patched_bytes() {
        let fixture_path = Path::new(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../../fixtures/xlsx/basic/basic.xlsx"
        ));
        let mut workbook = read_xlsx_blocking(fixture_path).expect("read fixture workbook");

        let sheet_id = workbook.sheets[0].id.clone();
        workbook
            .sheet_mut(&sheet_id)
            .expect("sheet exists")
            .set_cell(0, 0, Cell::from_literal(Some(CellScalar::Number(123.0))));

        let tmp = tempfile::tempdir().expect("temp dir");
        let out_path = tmp.path().join("nested/dir/patched.xlsx");

        let written = write_xlsx_blocking(&out_path, &workbook).expect("write patched workbook");
        let file_bytes = std::fs::read(&out_path).expect("read written file");

        assert_eq!(file_bytes.as_slice(), written.as_ref());
    }

    #[test]
    fn write_xlsx_blocking_creates_parent_dirs_for_generated_bytes() {
        let mut workbook = Workbook::new_empty(None);
        workbook.add_sheet("Sheet1".to_string());
        let sheet_id = workbook.sheets[0].id.clone();
        workbook
            .sheet_mut(&sheet_id)
            .expect("sheet exists")
            .set_cell(
                0,
                0,
                Cell::from_literal(Some(CellScalar::Text("Hello".to_string()))),
            );

        let tmp = tempfile::tempdir().expect("temp dir");
        let out_path = tmp.path().join("deep/nested/generated.xlsx");

        let written = write_xlsx_blocking(&out_path, &workbook).expect("write generated workbook");
        let file_bytes = std::fs::read(&out_path).expect("read written file");

        assert_eq!(file_bytes.as_slice(), written.as_ref());
    }

    #[test]
    fn xltx_save_writes_print_settings() {
        use formula_xlsx::print::{ManualPageBreaks, Orientation, PageSetup, SheetPrintSettings};

        let mut workbook = Workbook::new_empty(None);
        workbook.add_sheet("Sheet1".to_string());
        workbook.print_settings = WorkbookPrintSettings {
            sheets: vec![SheetPrintSettings {
                sheet_name: "Sheet1".to_string(),
                print_area: None,
                print_titles: None,
                page_setup: PageSetup {
                    orientation: Orientation::Landscape,
                    ..PageSetup::default()
                },
                manual_page_breaks: ManualPageBreaks::default(),
            }],
        };

        let tmp = tempfile::tempdir().expect("temp dir");
        let out_path = tmp.path().join("print-settings.xltx");
        let written_bytes = write_xlsx_blocking(&out_path, &workbook).expect("write workbook");

        let settings =
            read_workbook_print_settings(written_bytes.as_ref()).expect("read print settings");
        assert_eq!(settings, workbook.print_settings);
    }

    #[test]
    fn xltm_save_updates_workbook_date_system() {
        let mut workbook = Workbook::new_empty(None);
        workbook.add_sheet("Sheet1".to_string());
        workbook.date_system = WorkbookDateSystem::Excel1904;

        let tmp = tempfile::tempdir().expect("temp dir");
        let out_path = tmp.path().join("date-system-1904.xltm");
        let written_bytes = write_xlsx_blocking(&out_path, &workbook).expect("write workbook");

        let doc = formula_xlsx::load_from_bytes(written_bytes.as_ref())
            .expect("load workbook from bytes");
        assert_eq!(doc.workbook.date_system, WorkbookDateSystem::Excel1904);
    }
}
