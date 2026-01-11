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
use formula_xlsx::{load_from_bytes, XlsxPackage};
use rust_xlsxwriter::{Workbook as XlsxWorkbook, XlsxError};
use std::collections::HashMap;
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
    pub(crate) columnar: Option<Arc<ColumnarTable>>,
}

impl Sheet {
    pub fn new(id: String, name: String) -> Self {
        Self {
            id,
            name,
            cells: HashMap::new(),
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
    pub vba_project_bin: Option<Vec<u8>>,
    /// Stable identifier used for macro trust decisions (hash of workbook identity + `vbaProject.bin`).
    pub macro_fingerprint: Option<String>,
    pub preserved_drawing_parts: Option<PreservedDrawingParts>,
    pub sheets: Vec<Sheet>,
    pub print_settings: WorkbookPrintSettings,
}

impl Workbook {
    pub fn new_empty(path: Option<String>) -> Self {
        Self {
            origin_path: path.clone(),
            path,
            vba_project_bin: None,
            macro_fingerprint: None,
            preserved_drawing_parts: None,
            sheets: Vec::new(),
            print_settings: WorkbookPrintSettings::default(),
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
        let bytes =
            std::fs::read(path).with_context(|| format!("read workbook bytes {:?}", path))?;
        let document = load_from_bytes(&bytes).with_context(|| format!("parse xlsx {:?}", path))?;

        let print_settings = read_workbook_print_settings(&bytes)
            .ok()
            .unwrap_or_default();

        let mut out = Workbook {
            path: Some(path.to_string_lossy().to_string()),
            origin_path: Some(path.to_string_lossy().to_string()),
            vba_project_bin: None,
            macro_fingerprint: None,
            preserved_drawing_parts: None,
            sheets: Vec::new(),
            print_settings,
        };

        // Preserve macros: if the source file contains `xl/vbaProject.bin`, stash it so that
        // `write_xlsx_blocking` can re-inject it when saving as `.xlsm`.
        //
        // Note: formula-xlsx only understands XLSX/XLSM ZIP containers (not legacy XLS).
        if let Ok(pkg) = XlsxPackage::from_bytes(&bytes) {
            out.vba_project_bin = pkg.vba_project_bin().map(|b| b.to_vec());
            if let (Some(origin), Some(vba)) = (out.origin_path.as_deref(), out.vba_project_bin.as_deref()) {
                out.macro_fingerprint = Some(compute_macro_fingerprint(origin, vba));
            }
            if let Ok(preserved) = pkg.preserve_drawing_parts() {
                if !preserved.is_empty() {
                    out.preserved_drawing_parts = Some(preserved);
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
        return Ok(out);
    }

    let mut workbook =
        open_workbook_auto(path).with_context(|| format!("open workbook {:?}", path))?;
    let sheet_names = workbook.sheet_names().to_owned();

    let mut out = Workbook {
        path: Some(path.to_string_lossy().to_string()),
        origin_path: Some(path.to_string_lossy().to_string()),
        vba_project_bin: None,
        macro_fingerprint: None,
        preserved_drawing_parts: None,
        sheets: Vec::new(),
        print_settings: WorkbookPrintSettings::default(),
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
        vba_project_bin: None,
        macro_fingerprint: None,
        preserved_drawing_parts: None,
        sheets: vec![sheet],
        print_settings: WorkbookPrintSettings::default(),
    };
    out.ensure_sheet_ids();
    Ok(out)
}

fn read_xlsb_blocking(path: &Path) -> anyhow::Result<Workbook> {
    let wb = XlsbWorkbook::open(path).with_context(|| format!("open xlsb workbook {:?}", path))?;

    let mut out = Workbook {
        path: Some(path.to_string_lossy().to_string()),
        origin_path: Some(path.to_string_lossy().to_string()),
        vba_project_bin: None,
        macro_fingerprint: None,
        preserved_drawing_parts: None,
        sheets: Vec::new(),
        print_settings: WorkbookPrintSettings::default(),
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
    // Codes match Excel's internal error ids used by XLSB. Keep the string form compatible with
    // what calamine uses for `Data::Error` so UI/engine layers behave consistently.
    match code {
        0x00 => "#NULL!".to_string(),
        0x07 => "#DIV/0!".to_string(),
        0x0F => "#VALUE!".to_string(),
        0x17 => "#REF!".to_string(),
        0x1D => "#NAME?".to_string(),
        0x24 => "#NUM!".to_string(),
        0x2A => "#N/A".to_string(),
        0x2B => "#GETTING_DATA".to_string(),
        other => format!("#ERR({other:#04x})"),
    }
}

#[cfg(feature = "desktop")]
pub async fn write_xlsx(
    path: impl Into<PathBuf> + Send + 'static,
    workbook: Workbook,
) -> anyhow::Result<()> {
    let path = path.into();
    tauri::async_runtime::spawn_blocking(move || write_xlsx_blocking(&path, &workbook))
        .await
        .map_err(|e| anyhow::anyhow!(e.to_string()))?
}

pub fn write_xlsx_blocking(path: &Path, workbook: &Workbook) -> anyhow::Result<()> {
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

    let extension = path
        .extension()
        .and_then(|s| s.to_str())
        .map(|s| s.to_ascii_lowercase());

    let wants_vba =
        workbook.vba_project_bin.is_some() && matches!(extension.as_deref(), Some("xlsm"));
    let wants_preserved_drawings = workbook.preserved_drawing_parts.is_some();

    if wants_vba || wants_preserved_drawings {
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

        bytes = pkg
            .write_to_bytes()
            .context("repack workbook package with preserved parts")?;
    }

    if matches!(extension.as_deref(), Some("xlsx") | Some("xlsm")) {
        bytes = write_workbook_print_settings(&bytes, &workbook.print_settings)
            .map_err(|e| anyhow::anyhow!(e.to_string()))?;
    }

    std::fs::write(path, bytes).with_context(|| format!("write workbook {:?}", path))?;
    Ok(())
}

fn xlsx_err(err: XlsxError) -> String {
    format!("{err:?}")
}

#[cfg(test)]
mod tests {
    use super::*;
    use rust_xlsxwriter::{Chart, ChartType};

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
        write_xlsx_blocking(&out_path, &workbook).expect("write workbook");

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

        write_xlsx_blocking(&dst_path, &loaded).expect("write workbook");

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
}
