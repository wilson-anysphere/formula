use crate::state::{Cell, CellScalar};
use anyhow::Context;
use calamine::{open_workbook_auto, Data, Reader};
use rust_xlsxwriter::{Workbook as XlsxWorkbook, XlsxError};
use std::collections::HashMap;
use std::path::Path;
#[cfg(feature = "desktop")]
use std::path::PathBuf;

#[derive(Clone, Debug)]
pub struct Sheet {
    pub id: String,
    pub name: String,
    pub(crate) cells: HashMap<(usize, usize), Cell>,
}

impl Sheet {
    pub fn new(id: String, name: String) -> Self {
        Self {
            id,
            name,
            cells: HashMap::new(),
        }
    }

    pub fn get_cell(&self, row: usize, col: usize) -> Cell {
        self.cells
            .get(&(row, col))
            .cloned()
            .unwrap_or_else(Cell::empty)
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
}

#[derive(Clone, Debug)]
pub struct Workbook {
    pub path: Option<String>,
    pub sheets: Vec<Sheet>,
}

impl Workbook {
    pub fn new_empty(path: Option<String>) -> Self {
        Self {
            path,
            sheets: Vec::new(),
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

pub fn read_xlsx_blocking(path: &Path) -> anyhow::Result<Workbook> {
    let mut workbook =
        open_workbook_auto(path).with_context(|| format!("open workbook {:?}", path))?;
    let sheet_names = workbook.sheet_names().to_owned();

    let mut out = Workbook {
        path: Some(path.to_string_lossy().to_string()),
        sheets: Vec::new(),
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

    out.save(path)
        .map_err(|e| anyhow::anyhow!(xlsx_err(e)))
        .with_context(|| format!("save workbook {:?}", path))?;
    Ok(())
}

fn xlsx_err(err: XlsxError) -> String {
    format!("{err:?}")
}
