use crate::state::{Cell, CellScalar};
use anyhow::Context;
use calamine::{open_workbook_auto, Data, Reader};
use formula_columnar::{ColumnType as ColumnarType, ColumnarTable, Value as ColumnarValue};
use formula_model::{
    import::{import_csv_to_columnar_table, CsvOptions, CsvTextEncoding},
    sanitize_sheet_name, CellValue as ModelCellValue, DateSystem as WorkbookDateSystem, WorksheetId,
};
use formula_xlsb::{
    CellEdit as XlsbCellEdit, CellValue as XlsbCellValue, OpenOptions as XlsbOpenOptions,
    XlsbWorkbook,
};
use formula_xlsx::drawingml::PreservedDrawingParts;
use formula_xlsx::print::{
    read_workbook_print_settings, write_workbook_print_settings, WorkbookPrintSettings,
};
use formula_xlsx::{
    patch_xlsx_streaming_workbook_cell_patches,
    patch_xlsx_streaming_workbook_cell_patches_with_part_overrides, strip_vba_project_streaming,
    CellPatch as XlsxCellPatch, PartOverride, PreservedPivotParts, WorkbookCellPatches, XlsxPackage,
};
use std::collections::{BTreeMap, HashMap, HashSet};
use std::io::{BufReader, Cursor};
use std::path::Path;
#[cfg(feature = "desktop")]
use std::path::PathBuf;
use std::sync::Arc;

use crate::macro_trust::compute_macro_fingerprint;

const FORMULA_POWER_QUERY_PART: &str = "xl/formula/power-query.xml";
const OLE_MAGIC: [u8; 8] = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];

#[derive(Clone, Debug)]
pub struct Sheet {
    pub id: String,
    pub name: String,
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

fn cfb_stream_exists<R: std::io::Read + std::io::Write + std::io::Seek>(
    ole: &mut cfb::CompoundFile<R>,
    name: &str,
) -> bool {
    if ole.open_stream(name).is_ok() {
        return true;
    }
    let with_leading_slash = format!("/{name}");
    ole.open_stream(&with_leading_slash).is_ok()
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

#[derive(Copy, Clone, Debug, Eq, PartialEq)]
enum SniffedWorkbookFormat {
    Xls,
    Xlsx,
    Xlsb,
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
        return Some(SniffedWorkbookFormat::Xls);
    }

    let is_zip = read >= 4
        && (header[..4] == ZIP_LOCAL_FILE_HEADER
            || header[..4] == ZIP_CENTRAL_DIRECTORY
            || header[..4] == ZIP_SPANNING_SIGNATURE);
    if !is_zip {
        return None;
    }

    let _ = file.seek(SeekFrom::Start(0));
    let mut archive = zip::ZipArchive::new(file).ok()?;
    if archive.by_name("xl/workbook.bin").is_ok() {
        return Some(SniffedWorkbookFormat::Xlsb);
    }
    if archive.by_name("xl/workbook.xml").is_ok() {
        return Some(SniffedWorkbookFormat::Xlsx);
    }

    None
}

fn read_xlsx_or_xlsm_blocking(path: &Path) -> anyhow::Result<Workbook> {
    let origin_xlsx_bytes = Arc::<[u8]>::from(
        std::fs::read(path).with_context(|| format!("read workbook bytes {:?}", path))?,
    );
    let workbook_model =
        formula_xlsx::read_workbook_from_reader(Cursor::new(origin_xlsx_bytes.as_ref()))
            .with_context(|| format!("parse xlsx {:?}", path))?;
    let print_settings = read_workbook_print_settings(origin_xlsx_bytes.as_ref())
        .ok()
        .unwrap_or_default();

    let mut out = Workbook {
        path: Some(path.to_string_lossy().to_string()),
        origin_path: Some(path.to_string_lossy().to_string()),
        origin_xlsx_bytes: Some(origin_xlsx_bytes.clone()),
        power_query_xml: None,
        origin_xlsb_path: None,
        vba_project_bin: None,
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
    out.vba_project_bin = formula_xlsx::read_part_from_reader(
        Cursor::new(origin_xlsx_bytes.as_ref()),
        "xl/vbaProject.bin",
    )
    .ok()
    .flatten();
    if let Ok(Some(power_query_xml)) = formula_xlsx::read_part_from_reader(
        Cursor::new(origin_xlsx_bytes.as_ref()),
        FORMULA_POWER_QUERY_PART,
    ) {
        out.power_query_xml = Some(power_query_xml.clone());
        out.original_power_query_xml = Some(power_query_xml);
    }
    if let (Some(origin), Some(vba)) = (out.origin_path.as_deref(), out.vba_project_bin.as_deref()) {
        out.macro_fingerprint = Some(compute_macro_fingerprint(origin, vba));
    }
    if let Ok(parts) = formula_xlsx::worksheet_parts_from_reader(Cursor::new(origin_xlsx_bytes.as_ref()))
    {
        for part in parts {
            worksheet_parts_by_name.insert(part.name, part.worksheet_part);
        }
    }
    if let Ok(preserved) = formula_xlsx::drawingml::preserve_drawing_parts_from_reader(Cursor::new(
        origin_xlsx_bytes.as_ref(),
    )) {
        if !preserved.is_empty() {
            out.preserved_drawing_parts = Some(preserved);
        }
    }
    if let Ok(preserved) =
        formula_xlsx::pivots::preserve_pivot_parts_from_reader(Cursor::new(origin_xlsx_bytes.as_ref()))
    {
        if !preserved.is_empty() {
            out.preserved_pivot_parts = Some(preserved);
        }
    }
    if let Ok(palette) = formula_xlsx::theme_palette_from_reader(Cursor::new(origin_xlsx_bytes.as_ref()))
    {
        out.theme_palette = palette;
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

fn read_xls_blocking(path: &Path) -> anyhow::Result<Workbook> {
    let imported = formula_xls::import_xls_path(path)
        .map_err(|e| anyhow::anyhow!(e))
        .with_context(|| format!("import xls {:?}", path))?;
    let workbook_model = imported.workbook;

    let mut out = Workbook {
        path: Some(path.to_string_lossy().to_string()),
        origin_path: Some(path.to_string_lossy().to_string()),
        origin_xlsx_bytes: None,
        power_query_xml: None,
        origin_xlsb_path: None,
        vba_project_bin: None,
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

pub fn read_xlsx_blocking(path: &Path) -> anyhow::Result<Workbook> {
    if let Ok(true) = is_encrypted_ooxml_workbook(path) {
        anyhow::bail!(
            "encrypted workbook not supported: workbook `{}` is password-protected/encrypted; remove password protection in Excel and try again",
            path.display()
        );
    }

    let extension = path
        .extension()
        .and_then(|s| s.to_str())
        .map(|s| s.to_ascii_lowercase());

    if matches!(extension.as_deref(), Some("xlsb")) {
        return read_xlsb_blocking(path);
    }

    if matches!(extension.as_deref(), Some("xlsx") | Some("xlsm")) {
        return read_xlsx_or_xlsm_blocking(path);
    }

    if matches!(extension.as_deref(), Some("xls")) {
        return read_xls_blocking(path);
    }

    // For unknown extensions (or missing extensions), sniff the file signature/ZIP contents
    // to route to the appropriate reader.
    if !matches!(extension.as_deref(), Some("csv")) {
        if let Some(format) = sniff_workbook_format(path) {
            match format {
                SniffedWorkbookFormat::Xls => return read_xls_blocking(path),
                SniffedWorkbookFormat::Xlsx => return read_xlsx_or_xlsm_blocking(path),
                SniffedWorkbookFormat::Xlsb => return read_xlsb_blocking(path),
            }
        }
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
        ModelCellValue::Array(arr) => CellScalar::Text(format!("{:?}", arr.data)),
        ModelCellValue::Spill(_) => CellScalar::Error("#SPILL!".to_string()),
    }
}

pub fn read_csv_blocking(path: &Path) -> anyhow::Result<Workbook> {
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

fn read_xlsb_blocking(path: &Path) -> anyhow::Result<Workbook> {
    let wb = XlsbWorkbook::open(path).with_context(|| format!("open xlsb workbook {:?}", path))?;

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
        origin_xlsb_path: Some(path.to_string_lossy().to_string()),
        vba_project_bin: None,
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

pub fn write_xlsx_blocking(path: &Path, workbook: &Workbook) -> anyhow::Result<Arc<[u8]>> {
    let extension = path
        .extension()
        .and_then(|s| s.to_str())
        .map(|s| s.to_ascii_lowercase());

    if matches!(extension.as_deref(), Some("xlsb")) {
        return write_xlsb_blocking(path, workbook);
    }

    let xlsx_date_system = match workbook.date_system {
        WorkbookDateSystem::Excel1900 => formula_xlsx::DateSystem::V1900,
        WorkbookDateSystem::Excel1904 => formula_xlsx::DateSystem::V1904,
    };

    if let Some(origin_bytes) = workbook.origin_xlsx_bytes.as_deref() {
        let print_settings_changed = workbook.print_settings != workbook.original_print_settings;
        let power_query_changed = workbook.power_query_xml != workbook.original_power_query_xml;

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

        let wants_drop_vba =
            matches!(extension.as_deref(), Some("xlsx")) && workbook.vba_project_bin.is_some();

        if patches.is_empty() && !print_settings_changed && !wants_drop_vba && !power_query_changed {
            std::fs::write(path, origin_bytes)
                .with_context(|| format!("write workbook {:?}", path))?;
            return Ok(workbook
                .origin_xlsx_bytes
                .as_ref()
                .expect("origin_xlsx_bytes should be Some when origin_bytes is Some")
                .clone());
        }

        if patches.is_empty() && print_settings_changed && !wants_drop_vba && !power_query_changed {
            let bytes = write_workbook_print_settings(origin_bytes, &workbook.print_settings)
                .map_err(|e| anyhow::anyhow!(e.to_string()))?;

            let bytes = Arc::<[u8]>::from(bytes);
            std::fs::write(path, bytes.as_ref())
                .with_context(|| format!("write workbook {:?}", path))?;
            return Ok(bytes);
        }

        if !wants_drop_vba && power_query_changed {
            let mut part_overrides: HashMap<String, PartOverride> = HashMap::new();
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
                    part_overrides.insert(FORMULA_POWER_QUERY_PART.to_string(), PartOverride::Remove);
                }
            }

            let mut cursor = Cursor::new(Vec::new());
            patch_xlsx_streaming_workbook_cell_patches_with_part_overrides(
                Cursor::new(origin_bytes),
                &mut cursor,
                &patches,
                &part_overrides,
            )
            .context("apply worksheet cell patches + part overrides (streaming)")?;

            let mut bytes = cursor.into_inner();
            if print_settings_changed {
                bytes = write_workbook_print_settings(&bytes, &workbook.print_settings)
                    .map_err(|e| anyhow::anyhow!(e.to_string()))?;
            }

            let bytes = Arc::<[u8]>::from(bytes);
            std::fs::write(path, bytes.as_ref())
                .with_context(|| format!("write workbook {:?}", path))?;
            return Ok(bytes);
        }

        if !patches.is_empty() && !wants_drop_vba && !power_query_changed {
            let mut cursor = Cursor::new(Vec::new());
            patch_xlsx_streaming_workbook_cell_patches(
                Cursor::new(origin_bytes),
                &mut cursor,
                &patches,
            )
            .context("apply worksheet cell patches (streaming)")?;
            let mut bytes = cursor.into_inner();
            if print_settings_changed {
                bytes = write_workbook_print_settings(&bytes, &workbook.print_settings)
                    .map_err(|e| anyhow::anyhow!(e.to_string()))?;
            }

            let bytes = Arc::<[u8]>::from(bytes);
            std::fs::write(path, bytes.as_ref())
                .with_context(|| format!("write workbook {:?}", path))?;
            return Ok(bytes);
        }

        if wants_drop_vba && !power_query_changed {
            let mut bytes = if patches.is_empty() {
                let mut cursor = Cursor::new(Vec::new());
                strip_vba_project_streaming(Cursor::new(origin_bytes), &mut cursor)
                    .context("strip VBA project (streaming)")?;
                cursor.into_inner()
            } else {
                let mut cursor = Cursor::new(Vec::new());
                patch_xlsx_streaming_workbook_cell_patches(
                    Cursor::new(origin_bytes),
                    &mut cursor,
                    &patches,
                )
                .context("apply worksheet cell patches (streaming)")?;
                let patched = cursor.into_inner();

                let mut stripped = Cursor::new(Vec::new());
                strip_vba_project_streaming(Cursor::new(patched), &mut stripped)
                    .context("strip VBA project (streaming)")?;
                stripped.into_inner()
            };

            if print_settings_changed {
                bytes = write_workbook_print_settings(&bytes, &workbook.print_settings)
                    .map_err(|e| anyhow::anyhow!(e.to_string()))?;
            }

            let bytes = Arc::<[u8]>::from(bytes);
            std::fs::write(path, bytes.as_ref())
                .with_context(|| format!("write workbook {:?}", path))?;
            return Ok(bytes);
        }

        let mut pkg =
            XlsxPackage::from_bytes(origin_bytes).context("parse original workbook package")?;
        if !patches.is_empty() {
            pkg.apply_cell_patches(&patches)
                .context("apply worksheet cell patches")?;
        }

        match workbook.power_query_xml.as_ref() {
            Some(bytes) => pkg.set_part(FORMULA_POWER_QUERY_PART, bytes.clone()),
            None => {
                pkg.parts_map_mut().remove(FORMULA_POWER_QUERY_PART);
            }
        }

        if wants_drop_vba {
            pkg.remove_vba_project()
                .context("remove VBA parts for .xlsx")?;
        }

        if matches!(extension.as_deref(), Some("xlsx") | Some("xlsm"))
            && matches!(workbook.date_system, WorkbookDateSystem::Excel1904)
        {
            pkg.set_workbook_date_system(xlsx_date_system)
                .context("set workbook date system")?;
        }

        let mut bytes = pkg
            .write_to_bytes()
            .context("write patched workbook package")?;

        if print_settings_changed {
            bytes = write_workbook_print_settings(&bytes, &workbook.print_settings)
                .map_err(|e| anyhow::anyhow!(e.to_string()))?;
        }

        let bytes = Arc::<[u8]>::from(bytes);
        std::fs::write(path, bytes.as_ref())
            .with_context(|| format!("write workbook {:?}", path))?;
        return Ok(bytes);
    }

    let model = app_workbook_to_formula_model(workbook).context("convert workbook to model")?;
    let mut cursor = Cursor::new(Vec::new());
    formula_xlsx::write_workbook_to_writer(&model, &mut cursor)
        .map_err(|e| anyhow::anyhow!(e.to_string()))
        .with_context(|| "serialize workbook to buffer")?;
    let mut bytes = cursor.into_inner();
    let wants_vba =
        workbook.vba_project_bin.is_some() && matches!(extension.as_deref(), Some("xlsm"));
    let wants_preserved_drawings = workbook.preserved_drawing_parts.is_some();
    let wants_preserved_pivots = workbook.preserved_pivot_parts.is_some();
    let needs_date_system_update = matches!(extension.as_deref(), Some("xlsx") | Some("xlsm"))
        && matches!(workbook.date_system, WorkbookDateSystem::Excel1904);
    let wants_power_query = workbook.power_query_xml.is_some();

    if wants_vba
        || wants_preserved_drawings
        || wants_preserved_pivots
        || wants_power_query
        || needs_date_system_update
    {
        let mut pkg = XlsxPackage::from_bytes(&bytes).context("parse generated workbook package")?;

        if wants_vba {
            pkg.set_part(
                "xl/vbaProject.bin",
                workbook.vba_project_bin.clone().expect("checked is_some"),
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

    if matches!(extension.as_deref(), Some("xlsx") | Some("xlsm")) {
        bytes = write_workbook_print_settings(&bytes, &workbook.print_settings)
            .map_err(|e| anyhow::anyhow!(e.to_string()))?;
    }

    let bytes = Arc::<[u8]>::from(bytes);
    std::fs::write(path, bytes.as_ref()).with_context(|| format!("write workbook {:?}", path))?;
    Ok(bytes)
}

fn write_xlsb_blocking(path: &Path, workbook: &Workbook) -> anyhow::Result<Arc<[u8]>> {
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

    let dest_is_origin = std::fs::canonicalize(origin_path)
        .ok()
        .zip(std::fs::canonicalize(path).ok())
        .is_some_and(|(origin, dest)| origin == dest);

    let mut temp_paths: Vec<std::path::PathBuf> = Vec::new();
    let res = (|| -> anyhow::Result<()> {
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

        // Avoid writing directly over the source workbook since `formula-xlsb` streams from
        // `origin_path` while writing the destination ZIP.
        let final_out_path = if dest_is_origin {
            let dir = path.parent().unwrap_or_else(|| Path::new("."));
            let pid = std::process::id();
            let nanos = std::time::SystemTime::now()
                .duration_since(std::time::UNIX_EPOCH)
                .unwrap_or_default()
                .as_nanos();
            let mut candidate = dir.join(format!(".formula-xlsb-save-{pid}-{nanos}-final.xlsb"));
            let mut bump = 0u32;
            while candidate.exists() {
                bump += 1;
                candidate = dir.join(format!(
                    ".formula-xlsb-save-{pid}-{nanos}-{bump}-final.xlsb"
                ));
            }
            temp_paths.push(candidate.clone());
            candidate
        } else {
            path.to_path_buf()
        };

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
                let (current_input, current_formula) = match sheet.cells.get(&(*row, *col)) {
                    Some(cell) => (cell.input_value.clone(), cell.formula.clone()),
                    None => (None, None),
                };

                if let Some((baseline_value, baseline_formula)) = workbook
                    .cell_input_baseline
                    .get(&(sheet.id.clone(), *row, *col))
                {
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
                        new_formula: None,
                        new_rgcb: None,
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
                xlsb.save_with_cell_edits_streaming_shared_strings(&final_out_path, sheet_index, edits)
                    .with_context(|| format!("save edited xlsb {:?}", final_out_path))?;
            } else {
                xlsb.save_with_cell_edits_streaming(&final_out_path, sheet_index, edits)
                    .with_context(|| format!("save edited xlsb {:?}", final_out_path))?;
            }
            return Ok(());
        }

        let has_text_edits = edits_by_sheet
            .values()
            .flatten()
            .any(|edit| {
                matches!(edit.new_value, XlsbCellValue::Text(_))
                    && edit.new_formula.is_none()
                    && edit.new_rgcb.is_none()
            });

        // Prefer a single-pass multi-sheet streaming save. Keep the older "patch through temp
        // workbooks" approach only as a fallback if the multi-sheet writer errors.
        let multi_res = if has_text_edits {
            xlsb.save_with_cell_edits_streaming_multi_shared_strings(&final_out_path, &edits_by_sheet)
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
                wb.save_with_cell_edits_streaming_shared_strings(&out_path, sheet_index, sheet_edits)
                    .with_context(|| format!("save edited xlsb {:?}", out_path))?;
            } else {
                wb.save_with_cell_edits_streaming(&out_path, sheet_index, sheet_edits)
                    .with_context(|| format!("save edited xlsb {:?}", out_path))?;
            }

            source_path = out_path;
        }

        Ok(())
    })();

    if let Err(err) = res {
        for tmp in &temp_paths {
            let _ = std::fs::remove_file(tmp);
        }
        return Err(err);
    }

    if dest_is_origin {
        // We've already written to a temp path in the same directory.
        let tmp_final = temp_paths
            .iter()
            .find(|p| {
                p.file_name()
                    .is_some_and(|n| n.to_string_lossy().contains("final"))
            })
            .cloned();
        if let Some(tmp_final) = tmp_final {
            #[cfg(windows)]
            std::fs::remove_file(path)
                .with_context(|| format!("remove original xlsb {:?}", path))?;
            std::fs::rename(&tmp_final, path)
                .with_context(|| format!("replace original xlsb {:?}", path))?;
        }
    }

    for tmp in &temp_paths {
        let _ = std::fs::remove_file(tmp);
    }

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

    let mut sheet_id_by_app_id: HashMap<String, WorksheetId> = HashMap::new();
    let mut sheet_id_by_name: HashMap<String, WorksheetId> = HashMap::new();
    for sheet in &workbook.sheets {
        let sheet_id = out
            .add_sheet(sheet.name.clone())
            .map_err(|e| anyhow::anyhow!(e.to_string()))
            .with_context(|| format!("add sheet {}", sheet.name))?;
        sheet_id_by_app_id.insert(sheet.id.clone(), sheet_id);
        sheet_id_by_name.insert(sheet.name.clone(), sheet_id);
    }

    for sheet in &workbook.sheets {
        let model_sheet_id = sheet_id_by_app_id
            .get(&sheet.id)
            .copied()
            .ok_or_else(|| anyhow::anyhow!("missing sheet id for {}", sheet.id))?;
        let model_sheet = out
            .sheet_mut(model_sheet_id)
            .ok_or_else(|| anyhow::anyhow!("sheet missing from model: {}", sheet.id))?;

        if let Some(columnar) = sheet.columnar.as_ref() {
            // Preserve columnar-backed worksheets without materializing the full dataset
            // into the sparse cell map. The XLSX writer can stream from the columnar
            // table, while `sheet.cells` acts as an overlay for edits/formulas.
            model_sheet.set_columnar_table(formula_model::CellRef::new(0, 0), columnar.clone());
        }

        for ((row, col), cell) in sheet.cells_iter() {
            let row_u32 = u32::try_from(row).map_err(|_| anyhow::anyhow!("row out of bounds: {row}"))?;
            let col_u32 = u32::try_from(col).map_err(|_| anyhow::anyhow!("col out of bounds: {col}"))?;
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
        use std::io::Read;
        use xlsx_diff::{diff_workbooks, diff_workbooks_with_options, DiffOptions, Severity, WorkbookArchive};

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
    fn read_xlsx_blocking_errors_on_encrypted_ooxml_container() {
        let tmp = tempfile::tempdir().expect("temp dir");

        let cursor = Cursor::new(Vec::new());
        let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
        ole.create_stream("EncryptionInfo")
            .expect("create EncryptionInfo stream");
        ole.create_stream("EncryptedPackage")
            .expect("create EncryptedPackage stream");
        let bytes = ole.into_inner().into_inner();

        for filename in ["encrypted.xlsx", "encrypted.xls"] {
            let path = tmp.path().join(filename);
            std::fs::write(&path, &bytes).expect("write encrypted fixture");

            let err = read_xlsx_blocking(&path).expect_err("expected encrypted workbook to error");
            let msg = err.to_string().to_lowercase();
            assert!(
                msg.contains("encrypted") || msg.contains("password"),
                "expected error message to mention encryption/password protection, got: {msg}"
            );
        }
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
        assert_eq!(workbook.origin_xlsb_path.as_deref(), Some(renamed_str.as_str()));

        assert_eq!(workbook.sheets.len(), 1);
        let sheet = &workbook.sheets[0];
        assert_eq!(sheet.name, "Sheet1");
        assert_eq!(
            sheet.get_cell(0, 0).computed_value,
            CellScalar::Text("Hello".to_string())
        );
        assert_eq!(sheet.get_cell(0, 1).computed_value, CellScalar::Number(42.5));
        assert_eq!(sheet.get_cell(0, 2).computed_value, CellScalar::Number(85.0));
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
        assert_eq!(sheet1.get_cell(1, 1).computed_value, CellScalar::Number(123.0));
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
        assert!(zed.sheet_id.is_none(), "expected ZedName to be workbook-scoped");

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

        let doc =
            formula_xlsx::load_from_bytes(written_bytes.as_ref()).expect("load saved workbook from bytes");
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
        assert_eq!(sheet.get_cell(2, 3).computed_value, CellScalar::Number(3.75));
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
    fn reads_csv_with_invalid_file_name_sanitizes_sheet_name_and_writes_xlsx() {
        let tmp = tempfile::tempdir().expect("temp dir");
        // Use characters that are invalid for Excel sheet names but valid on common filesystems.
        let path = tmp.path().join("bad[name]test.csv");
        std::fs::write(&path, "id,name\n1,hello\n").expect("write csv");

        let workbook = read_csv_blocking(&path).expect("read csv");
        assert_eq!(workbook.sheets.len(), 1);
        let sheet = &workbook.sheets[0];

        assert_eq!(sheet.name, "badnametest");
        assert!(!sheet.name.trim().is_empty(), "sheet name should be non-empty");
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
}
