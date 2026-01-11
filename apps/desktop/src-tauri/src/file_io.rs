use crate::state::{Cell, CellScalar};
use anyhow::Context;
use calamine::{open_workbook_auto, Data, Reader};
use formula_columnar::{ColumnType as ColumnarType, ColumnarTable, Value as ColumnarValue};
use formula_model::{
    import::{import_csv_to_columnar_table, CsvOptions},
    CellValue as ModelCellValue,
};
use formula_xlsb::{CellValue as XlsbCellValue, XlsbWorkbook};
use formula_xlsx::drawingml::PreservedDrawingParts;
use formula_xlsx::print::{
    read_workbook_print_settings, write_workbook_print_settings, WorkbookPrintSettings,
};
use formula_xlsx::{
    load_from_bytes, CellPatch as XlsxCellPatch, PreservedPivotParts, WorkbookCellPatches,
    XlsxPackage,
};
use rust_xlsxwriter::{Workbook as XlsxWorkbook, XlsxError};
use std::collections::{HashMap, HashSet};
use std::io::BufReader;
use std::path::Path;
#[cfg(feature = "desktop")]
use std::path::PathBuf;
use std::sync::Arc;

use crate::macro_trust::compute_macro_fingerprint;

#[derive(Clone, Debug)]
pub struct Sheet {
    pub id: String,
    pub name: String,
    pub(crate) cells: HashMap<(usize, usize), Cell>,
    pub(crate) dirty_cells: HashSet<(usize, usize)>,
    pub(crate) columnar: Option<Arc<ColumnarTable>>,
}

impl Sheet {
    pub fn new(id: String, name: String) -> Self {
        Self {
            id,
            name,
            cells: HashMap::new(),
            dirty_cells: HashSet::new(),
            columnar: None,
        }
    }

    pub fn get_cell(&self, row: usize, col: usize) -> Cell {
        if let Some(cell) = self.cells.get(&(row, col)) {
            return cell.clone();
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
                return match scalar {
                    CellScalar::Empty => Cell::empty(),
                    other => Cell::from_literal(Some(other)),
                };
            }
        }

        Cell::empty()
    }

    pub fn set_cell(&mut self, row: usize, col: usize, cell: Cell) {
        self.dirty_cells.insert((row, col));
        if cell.formula.is_none() && cell.input_value.is_none() {
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
    pub vba_project_bin: Option<Vec<u8>>,
    /// Stable identifier used for macro trust decisions (hash of workbook identity + `vbaProject.bin`).
    pub macro_fingerprint: Option<String>,
    pub preserved_drawing_parts: Option<PreservedDrawingParts>,
    pub preserved_pivot_parts: Option<PreservedPivotParts>,
    pub sheets: Vec<Sheet>,
    pub print_settings: WorkbookPrintSettings,
    pub(crate) original_print_settings: WorkbookPrintSettings,
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
            vba_project_bin: None,
            macro_fingerprint: None,
            preserved_drawing_parts: None,
            preserved_pivot_parts: None,
            sheets: Vec::new(),
            print_settings: WorkbookPrintSettings::default(),
            original_print_settings: WorkbookPrintSettings::default(),
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

pub fn read_xlsx_blocking(path: &Path) -> anyhow::Result<Workbook> {
    let extension = path
        .extension()
        .and_then(|s| s.to_str())
        .map(|s| s.to_ascii_lowercase());

    if matches!(extension.as_deref(), Some("xlsb")) {
        return read_xlsb_blocking(path);
    }

    if matches!(extension.as_deref(), Some("xlsx") | Some("xlsm")) {
        let origin_xlsx_bytes = Arc::<[u8]>::from(
            std::fs::read(path).with_context(|| format!("read workbook bytes {:?}", path))?,
        );
        let document = load_from_bytes(origin_xlsx_bytes.as_ref())
            .with_context(|| format!("parse xlsx {:?}", path))?;

        let print_settings = read_workbook_print_settings(origin_xlsx_bytes.as_ref())
            .ok()
            .unwrap_or_default();

        let mut out = Workbook {
            path: Some(path.to_string_lossy().to_string()),
            origin_path: Some(path.to_string_lossy().to_string()),
            origin_xlsx_bytes: Some(origin_xlsx_bytes.clone()),
            vba_project_bin: None,
            macro_fingerprint: None,
            preserved_drawing_parts: None,
            preserved_pivot_parts: None,
            sheets: Vec::new(),
            print_settings: print_settings.clone(),
            original_print_settings: print_settings,
            cell_input_baseline: HashMap::new(),
        };

        // Preserve macros: if the source file contains `xl/vbaProject.bin`, stash it so that
        // `write_xlsx_blocking` can re-inject it when saving as `.xlsm`.
        //
        // Note: formula-xlsx only understands XLSX/XLSM ZIP containers (not legacy XLS).
        if let Ok(pkg) = XlsxPackage::from_bytes(origin_xlsx_bytes.as_ref()) {
            out.vba_project_bin = pkg.vba_project_bin().map(|b| b.to_vec());
            if let (Some(origin), Some(vba)) =
                (out.origin_path.as_deref(), out.vba_project_bin.as_deref())
            {
                out.macro_fingerprint = Some(compute_macro_fingerprint(origin, vba));
            }
            if let Ok(preserved) = pkg.preserve_drawing_parts() {
                if !preserved.is_empty() {
                    out.preserved_drawing_parts = Some(preserved);
                }
            }
            if let Ok(preserved) = pkg.preserve_pivot_parts() {
                if !preserved.is_empty() {
                    out.preserved_pivot_parts = Some(preserved);
                }
            }
        }

        out.sheets = document
            .workbook
            .sheets
            .iter()
            .map(formula_model_sheet_to_app_sheet)
            .collect::<anyhow::Result<Vec<_>>>()?;

        out.ensure_sheet_ids();
        for sheet in &mut out.sheets {
            sheet.clear_dirty_cells();
        }
        return Ok(out);
    }

    let mut workbook =
        open_workbook_auto(path).with_context(|| format!("open workbook {:?}", path))?;
    let sheet_names = workbook.sheet_names().to_owned();

    let mut out = Workbook {
        path: Some(path.to_string_lossy().to_string()),
        origin_path: Some(path.to_string_lossy().to_string()),
        origin_xlsx_bytes: None,
        vba_project_bin: None,
        macro_fingerprint: None,
        preserved_drawing_parts: None,
        preserved_pivot_parts: None,
        sheets: Vec::new(),
        print_settings: WorkbookPrintSettings::default(),
        original_print_settings: WorkbookPrintSettings::default(),
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
                if formula.trim().is_empty() {
                    continue;
                }
                let normalized = if formula.starts_with('=') {
                    formula.to_string()
                } else {
                    format!("={formula}")
                };

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

fn formula_model_sheet_to_app_sheet(sheet: &formula_model::Worksheet) -> anyhow::Result<Sheet> {
    let mut out = Sheet::new(sheet.name.clone(), sheet.name.clone());

    for (cell_ref, cell) in sheet.iter_cells() {
        let row = cell_ref.row as usize;
        let col = cell_ref.col as usize;

        let cached_value = formula_model_value_to_scalar(&cell.value);
        if let Some(formula) = cell.formula.as_deref() {
            if formula.trim().is_empty() {
                continue;
            }
            let normalized = normalize_formula_text(formula);
            let mut c = Cell::from_formula(normalized);
            c.computed_value = cached_value;
            out.set_cell(row, col, c);
            continue;
        }

        if matches!(cached_value, CellScalar::Empty) {
            continue;
        }

        out.set_cell(row, col, Cell::from_literal(Some(cached_value)));
    }

    Ok(out)
}

fn normalize_formula_text(formula: &str) -> String {
    if formula.starts_with('=') {
        formula.to_string()
    } else {
        format!("={formula}")
    }
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
    let table = import_csv_to_columnar_table(reader, CsvOptions::default())
        .map_err(|e| anyhow::anyhow!(e.to_string()))
        .with_context(|| format!("import csv {:?}", path))?;

    let sheet_name = path
        .file_stem()
        .and_then(|s| s.to_str())
        .filter(|s| !s.trim().is_empty())
        .unwrap_or("Sheet1")
        .to_string();
    let mut sheet = Sheet::new(sheet_name.clone(), sheet_name);
    sheet.set_columnar_table(Arc::new(table));

    let mut out = Workbook {
        path: Some(path.to_string_lossy().to_string()),
        origin_path: Some(path.to_string_lossy().to_string()),
        origin_xlsx_bytes: None,
        vba_project_bin: None,
        macro_fingerprint: None,
        preserved_drawing_parts: None,
        preserved_pivot_parts: None,
        sheets: vec![sheet],
        print_settings: WorkbookPrintSettings::default(),
        original_print_settings: WorkbookPrintSettings::default(),
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

    let mut out = Workbook {
        path: Some(path.to_string_lossy().to_string()),
        origin_path: Some(path.to_string_lossy().to_string()),
        origin_xlsx_bytes: None,
        vba_project_bin: None,
        macro_fingerprint: None,
        preserved_drawing_parts: None,
        preserved_pivot_parts: None,
        sheets: Vec::new(),
        print_settings: WorkbookPrintSettings::default(),
        original_print_settings: WorkbookPrintSettings::default(),
        cell_input_baseline: HashMap::new(),
    };

    // If formula-xlsb can detect a formula record but can't decode its RGCE token stream yet,
    // it surfaces the formula without any text. In that case we keep the cached value, but also
    // track the cell so we can fall back to Calamine's formula text extraction.
    let mut missing_formulas: Vec<(usize, String, Vec<(usize, usize, CellScalar)>)> = Vec::new();

    for (idx, sheet_meta) in wb.sheet_metas().iter().enumerate() {
        let mut sheet = Sheet::new(sheet_meta.name.clone(), sheet_meta.name.clone());
        let mut undecoded_formula_cells: Vec<(usize, usize, CellScalar)> = Vec::new();

        wb.for_each_cell(idx, |cell| {
            let row = cell.row as usize;
            let col = cell.col as usize;

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
                        let normalized = if formula.starts_with('=') {
                            formula
                        } else {
                            format!("={formula}")
                        };
                        let mut c = Cell::from_formula(normalized);
                        c.computed_value = value;
                        sheet.set_cell(row, col, c);
                    }
                    None => {
                        // Preserve the cached value and fill the formula text later (best-effort).
                        undecoded_formula_cells.push((row, col, value));
                    }
                },
                None => {
                    if matches!(value, CellScalar::Empty) {
                        return;
                    }
                    sheet.set_cell(row, col, Cell::from_literal(Some(value)));
                }
            }
        })
        .with_context(|| format!("read xlsb sheet {}", sheet_meta.name))?;

        if !undecoded_formula_cells.is_empty() {
            missing_formulas.push((out.sheets.len(), sheet_meta.name.clone(), undecoded_formula_cells));
        }
        out.sheets.push(sheet);
    }

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
        if formula.trim().is_empty() {
            continue;
        }

        let normalized = if formula.starts_with('=') {
            formula.to_string()
        } else {
            format!("={formula}")
        };

        out.insert((row_offset + row, col_offset + col), normalized);
    }

    out
}

fn apply_xlsb_formula_fallback(
    sheet: &mut Sheet,
    missing_cells: Vec<(usize, usize, CellScalar)>,
    formula_lookup: &HashMap<(usize, usize), String>,
) {
    for (row, col, cached_value) in missing_cells {
        if let Some(formula) = formula_lookup.get(&(row, col)) {
            let mut cell = Cell::from_formula(formula.clone());
            cell.computed_value = cached_value;
            sheet.set_cell(row, col, cell);
        } else if !matches!(cached_value, CellScalar::Empty) {
            sheet.set_cell(row, col, Cell::from_literal(Some(cached_value)));
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

    if let Some(origin_bytes) = workbook.origin_xlsx_bytes.as_deref() {
        let print_settings_changed = workbook.print_settings != workbook.original_print_settings;

        let mut pkg =
            XlsxPackage::from_bytes(origin_bytes).context("parse original workbook package")?;
        let mut patches = WorkbookCellPatches::default();
        for sheet in &workbook.sheets {
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
                    patches.set_cell(sheet.name.clone(), cell_ref, XlsxCellPatch::clear());
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

                patches.set_cell(sheet.name.clone(), cell_ref, patch);
            }
        }

        let wants_drop_vba =
            matches!(extension.as_deref(), Some("xlsx")) && pkg.vba_project_bin().is_some();

        if patches.is_empty() && !print_settings_changed && !wants_drop_vba {
            std::fs::write(path, origin_bytes)
                .with_context(|| format!("write workbook {:?}", path))?;
            return Ok(workbook
                .origin_xlsx_bytes
                .as_ref()
                .expect("origin_xlsx_bytes should be Some when origin_bytes is Some")
                .clone());
        }
        if !patches.is_empty() {
            pkg.apply_cell_patches(&patches)
                .context("apply worksheet cell patches")?;
        }

        if wants_drop_vba {
            pkg.remove_vba_project().context("remove VBA parts for .xlsx")?;
        }

        let mut bytes = pkg
            .write_to_bytes()
            .context("write patched workbook package")?;

        if matches!(extension.as_deref(), Some("xlsx") | Some("xlsm")) && print_settings_changed {
            bytes = write_workbook_print_settings(&bytes, &workbook.print_settings)
                .map_err(|e| anyhow::anyhow!(e.to_string()))?;
        }

        let bytes = Arc::<[u8]>::from(bytes);
        std::fs::write(path, bytes.as_ref()).with_context(|| format!("write workbook {:?}", path))?;
        return Ok(bytes);
    }

    let mut out = XlsxWorkbook::new();

    for sheet in &workbook.sheets {
        let worksheet = out.add_worksheet();
        worksheet
            .set_name(&sheet.name)
            .map_err(|e| anyhow::anyhow!(format!("sheet name error: {e:?}")))?;

        for ((row, col), cell) in sheet.cells_iter() {
            if let Some(formula) = &cell.formula {
                worksheet
                    .write_formula(row as u32, col as u16, formula.as_str())
                    .map_err(|e| anyhow::anyhow!(xlsx_err(e)))?;
                continue;
            }

            match &cell.computed_value {
                CellScalar::Empty => {}
                CellScalar::Number(n) => {
                    worksheet
                        .write_number(row as u32, col as u16, *n)
                        .map_err(|e| anyhow::anyhow!(xlsx_err(e)))?;
                }
                CellScalar::Text(s) => {
                    worksheet
                        .write_string(row as u32, col as u16, s)
                        .map_err(|e| anyhow::anyhow!(xlsx_err(e)))?;
                }
                CellScalar::Bool(b) => {
                    worksheet
                        .write_boolean(row as u32, col as u16, *b)
                        .map_err(|e| anyhow::anyhow!(xlsx_err(e)))?;
                }
                CellScalar::Error(e) => {
                    worksheet
                        .write_string(row as u32, col as u16, e)
                        .map_err(|e| anyhow::anyhow!(xlsx_err(e)))?;
                }
            }
        }
    }

    let mut bytes = out
        .save_to_buffer()
        .map_err(|e| anyhow::anyhow!(xlsx_err(e)))
        .with_context(|| "serialize workbook to buffer")?;

    let wants_vba =
        workbook.vba_project_bin.is_some() && matches!(extension.as_deref(), Some("xlsm"));
    let wants_preserved_drawings = workbook.preserved_drawing_parts.is_some();
    let wants_preserved_pivots = workbook.preserved_pivot_parts.is_some();

    if wants_vba || wants_preserved_drawings || wants_preserved_pivots {
        let mut pkg =
            XlsxPackage::from_bytes(&bytes).context("parse generated workbook package")?;

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

fn xlsx_err(err: XlsxError) -> String {
    format!("{err:?}")
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::state::AppState;
    use rust_xlsxwriter::{Chart, ChartType};
    use xlsx_diff::{diff_workbooks, Severity};

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

        let original_names: HashSet<String> = original_pkg.part_names().map(str::to_owned).collect();
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
    fn xlsb_formula_fallback_fills_missing_formula_text() {
        // Simulate a formula-xlsb cell that has a cached value but no decoded formula text,
        // and a Calamine-provided lookup table for formulas.
        let lookup = HashMap::from([((0, 2), "=B1*2".to_string())]);
        let mut sheet = Sheet::new("Sheet1".to_string(), "Sheet1".to_string());
        apply_xlsb_formula_fallback(
            &mut sheet,
            vec![(0, 2, CellScalar::Number(85.0))],
            &lookup,
        );

        let cell = sheet.get_cell(0, 2);
        assert_eq!(cell.formula.as_deref(), Some("=B1*2"));
        assert_eq!(cell.computed_value, CellScalar::Number(85.0));
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
        let tmp = tempfile::tempdir().expect("temp dir");
        let src_path = tmp.path().join("chart-src.xlsx");
        let dst_path = tmp.path().join("chart-dst.xlsx");

        // Create a workbook with a simple chart.
        let mut workbook = XlsxWorkbook::new();
        let worksheet = workbook.add_worksheet();

        worksheet.write_string(0, 0, "Category").unwrap();
        worksheet.write_string(0, 1, "Value").unwrap();
        worksheet.write_string(1, 0, "A").unwrap();
        worksheet.write_number(1, 1, 2).unwrap();
        worksheet.write_string(2, 0, "B").unwrap();
        worksheet.write_number(2, 1, 4).unwrap();
        worksheet.write_string(3, 0, "C").unwrap();
        worksheet.write_number(3, 1, 3).unwrap();

        let mut chart = Chart::new(ChartType::Column);
        chart.title().set_name("Example Chart");

        let series = chart.add_series();
        series
            .set_categories("Sheet1!$A$2:$A$4")
            .set_values("Sheet1!$B$2:$B$4");

        worksheet.insert_chart(1, 3, &chart).unwrap();

        let bytes = workbook.save_to_buffer().expect("save workbook");
        std::fs::write(&src_path, &bytes).expect("write source workbook");

        // Load via the app IO path and save again.
        let loaded = read_xlsx_blocking(&src_path).expect("read workbook");
        assert!(
            loaded.preserved_drawing_parts.is_some(),
            "expected chart parts to be captured for preservation"
        );

        let _ = write_xlsx_blocking(&dst_path, &loaded).expect("write workbook");

        let roundtrip_bytes = std::fs::read(&dst_path).expect("read written workbook");
        let src_pkg = XlsxPackage::from_bytes(&bytes).expect("parse src pkg");
        let dst_pkg = XlsxPackage::from_bytes(&roundtrip_bytes).expect("parse dst pkg");

        // Drawing + chart parts should match byte-for-byte.
        for (name, part_bytes) in src_pkg.parts() {
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
        let src_charts = src_pkg.extract_charts().expect("extract src charts");
        let dst_charts = dst_pkg.extract_charts().expect("extract dst charts");
        assert_eq!(src_charts.len(), 1);
        assert_eq!(dst_charts.len(), 1);
        assert_eq!(src_charts[0].rel_id, dst_charts[0].rel_id);
        assert_eq!(src_charts[0].chart_part, dst_charts[0].chart_part);
        assert_eq!(src_charts[0].drawing_part, dst_charts[0].drawing_part);
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
        assert_eq!(report.count(Severity::Critical), 0, "unexpected diffs: {report:?}");
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
        assert_eq!(report.count(Severity::Critical), 0, "unexpected diffs: {report:?}");
    }

    #[test]
    fn cell_edit_preserves_comment_parts() {
        let fixture_path = Path::new(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../../fixtures/xlsx/basic/comments.xlsx"
        ));
        let mut workbook = read_xlsx_blocking(fixture_path).expect("read comments fixture workbook");

        let sheet_id = workbook.sheets[0].id.clone();
        workbook
            .sheet_mut(&sheet_id)
            .unwrap()
            .set_cell(0, 0, Cell::from_literal(Some(CellScalar::Number(123.0))));

        let tmp = tempfile::tempdir().expect("temp dir");
        let out_path = tmp.path().join("edited.xlsx");
        write_xlsx_blocking(&out_path, &workbook).expect("write workbook");

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
        assert_eq!(report.count(Severity::Critical), 0, "unexpected diffs: {report:?}");
    }

    #[test]
    fn cell_edit_preserves_pivot_parts() {
        let fixture_path = Path::new(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../../fixtures/xlsx/pivots/pivot-fixture.xlsx"
        ));
        let mut workbook = read_xlsx_blocking(fixture_path).expect("read pivot fixture workbook");

        let sheet_id = workbook.sheets[0].id.clone();
        workbook
            .sheet_mut(&sheet_id)
            .unwrap()
            .set_cell(0, 0, Cell::from_literal(Some(CellScalar::Number(123.0))));

        let tmp = tempfile::tempdir().expect("temp dir");
        let out_path = tmp.path().join("edited.xlsx");
        write_xlsx_blocking(&out_path, &workbook).expect("write workbook");

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
        let original_pkg = XlsxPackage::from_bytes(&original_bytes).expect("parse original package");
        let original_sheet_xml =
            std::str::from_utf8(original_pkg.part("xl/worksheets/sheet1.xml").unwrap())
                .expect("original sheet1.xml utf8");
        assert!(
            original_sheet_xml.contains("<drawing"),
            "expected fixture sheet1.xml to contain a drawing relationship"
        );

        let mut workbook = read_xlsx_blocking(fixture_path).expect("read image fixture workbook");

        let sheet_id = workbook.sheets[0].id.clone();
        workbook
            .sheet_mut(&sheet_id)
            .unwrap()
            .set_cell(1, 1, Cell::from_literal(Some(CellScalar::Number(42.0))));

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
        let original_pkg = XlsxPackage::from_bytes(&original_bytes).expect("parse original package");
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
        workbook
            .sheet_mut(&sheet_id)
            .unwrap()
            .set_cell(0, 1, Cell::from_literal(Some(CellScalar::Number(7.0))));

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
    fn cell_edit_preserves_defined_names() {
        let fixture_path = Path::new(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../../fixtures/xlsx/metadata/defined-names.xlsx"
        ));
        let original_bytes = std::fs::read(fixture_path).expect("read fixture bytes");
        let mut workbook =
            read_xlsx_blocking(fixture_path).expect("read defined-names fixture workbook");

        let sheet_id = workbook.sheets[0].id.clone();
        workbook
            .sheet_mut(&sheet_id)
            .unwrap()
            .set_cell(0, 2, Cell::from_literal(Some(CellScalar::Number(99.0))));

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
        workbook
            .sheet_mut(&sheet_id)
            .unwrap()
            .set_cell(0, 1, Cell::from_literal(Some(CellScalar::Number(5.0))));

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
        workbook
            .sheet_mut(&sheet_id)
            .unwrap()
            .set_cell(0, 0, Cell::from_literal(Some(CellScalar::Number(123.0))));

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
        workbook
            .sheet_mut(&sheet_id)
            .unwrap()
            .set_cell(0, 0, Cell::from_literal(Some(CellScalar::Number(7.0))));

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

        let rels = std::str::from_utf8(written_pkg.part("xl/_rels/workbook.xml.rels").unwrap())
            .unwrap();
        assert!(
            !rels.contains("relationships/vbaProject"),
            "expected workbook.xml.rels to drop the vbaProject relationship"
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
        let written_bytes = write_xlsx_blocking(&out_path, state.get_workbook().unwrap())
            .expect("write workbook");

        let doc = load_from_bytes(written_bytes.as_ref()).expect("load saved workbook from bytes");
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
        let _ = write_xlsx_blocking(&out_path, state.get_workbook().unwrap())
            .expect("write workbook");

        assert_no_critical_diffs(fixture_path, &out_path);
    }
}
