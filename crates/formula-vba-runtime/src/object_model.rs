use std::cell::RefCell;
use std::collections::HashMap;
use std::rc::Rc;

use crate::runtime::VbaError;
use crate::value::VbaValue;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct VbaRangeRef {
    pub sheet: usize,
    pub start_row: u32,
    pub start_col: u32,
    pub end_row: u32,
    pub end_col: u32,
}

/// A minimal spreadsheet API that the VBA runtime can manipulate.
pub trait Spreadsheet {
    fn sheet_count(&self) -> usize;
    fn sheet_name(&self, sheet: usize) -> Option<&str>;
    fn sheet_index(&self, name: &str) -> Option<usize>;

    fn active_sheet(&self) -> usize;
    fn set_active_sheet(&mut self, sheet: usize) -> Result<(), VbaError>;

    fn active_cell(&self) -> (u32, u32);
    fn set_active_cell(&mut self, row: u32, col: u32) -> Result<(), VbaError>;

    fn get_cell_value(&self, sheet: usize, row: u32, col: u32) -> Result<VbaValue, VbaError>;
    fn set_cell_value(
        &mut self,
        sheet: usize,
        row: u32,
        col: u32,
        value: VbaValue,
    ) -> Result<(), VbaError>;

    fn get_cell_formula(
        &self,
        sheet: usize,
        row: u32,
        col: u32,
    ) -> Result<Option<String>, VbaError>;

    fn set_cell_formula(
        &mut self,
        sheet: usize,
        row: u32,
        col: u32,
        formula: String,
    ) -> Result<(), VbaError>;

    fn log(&mut self, message: String);
}

#[derive(Clone)]
pub struct VbaObjectRef(Rc<RefCell<VbaObject>>);

impl std::fmt::Debug for VbaObjectRef {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "VbaObjectRef(..)")
    }
}

impl VbaObjectRef {
    pub fn new(obj: VbaObject) -> Self {
        Self(Rc::new(RefCell::new(obj)))
    }

    pub fn borrow(&self) -> std::cell::Ref<'_, VbaObject> {
        self.0.borrow()
    }

    pub fn borrow_mut(&self) -> std::cell::RefMut<'_, VbaObject> {
        self.0.borrow_mut()
    }
}

#[derive(Debug, Clone)]
pub enum VbaObject {
    Application,
    Workbook,
    Worksheet { sheet: usize },
    Range(VbaRangeRef),
    Collection { items: Vec<VbaValue> },
}

pub fn a1_to_row_col(a1: &str) -> Result<(u32, u32), VbaError> {
    // A1 -> (row, col) 1-based
    let a1 = a1.trim();
    let mut col: u32 = 0;
    let mut row_str = String::new();
    for ch in a1.chars() {
        if ch.is_ascii_alphabetic() {
            col = col
                .checked_mul(26)
                .and_then(|v| v.checked_add((ch.to_ascii_uppercase() as u8 - b'A' + 1) as u32))
                .ok_or_else(|| VbaError::Runtime(format!("Invalid A1 reference: {a1}")))?;
        } else if ch.is_ascii_digit() {
            row_str.push(ch);
        } else if ch == '$' {
            // ignore absolute markers
        } else {
            return Err(VbaError::Runtime(format!("Invalid A1 reference: {a1}")));
        }
    }
    if col == 0 || row_str.is_empty() {
        return Err(VbaError::Runtime(format!("Invalid A1 reference: {a1}")));
    }
    let row: u32 = row_str
        .parse()
        .map_err(|_| VbaError::Runtime(format!("Invalid A1 reference: {a1}")))?;
    Ok((row, col))
}

pub fn row_col_to_a1(row: u32, col: u32) -> Result<String, VbaError> {
    if row == 0 || col == 0 {
        return Err(VbaError::Runtime("Row/col are 1-based".to_string()));
    }
    let mut c = col;
    let mut letters = Vec::new();
    while c > 0 {
        let rem = ((c - 1) % 26) as u8;
        letters.push((b'A' + rem) as char);
        c = (c - 1) / 26;
    }
    letters.reverse();
    Ok(format!(
        "{}{}",
        letters.into_iter().collect::<String>(),
        row
    ))
}

fn parse_range_a1(a1: &str) -> Result<(u32, u32, u32, u32), VbaError> {
    let parts: Vec<&str> = a1.split(':').collect();
    if parts.len() == 1 {
        let (r, c) = a1_to_row_col(parts[0])?;
        Ok((r, c, r, c))
    } else if parts.len() == 2 {
        let (r1, c1) = a1_to_row_col(parts[0])?;
        let (r2, c2) = a1_to_row_col(parts[1])?;
        Ok((r1.min(r2), c1.min(c2), r1.max(r2), c1.max(c2)))
    } else {
        Err(VbaError::Runtime(format!("Invalid A1 range: {a1}")))
    }
}

/// Parse an A1 reference into `(start_row, start_col, end_row, end_col)` (all 1-based).
pub fn a1_to_row_col_range(a1: &str) -> Result<(u32, u32, u32, u32), VbaError> {
    parse_range_a1(a1)
}

/// A very small in-memory workbook model used by tests and examples.
#[derive(Debug, Default)]
pub struct InMemoryWorkbook {
    sheets: Vec<Sheet>,
    active_sheet: usize,
    active_cell: (u32, u32),
    pub output: Vec<String>,
}

#[derive(Debug, Default)]
struct Sheet {
    name: String,
    cells: HashMap<(u32, u32), Cell>,
}

#[derive(Debug, Default)]
struct Cell {
    value: VbaValue,
    formula: Option<String>,
}

impl InMemoryWorkbook {
    pub fn new() -> Self {
        let mut wb = Self::default();
        wb.add_sheet("Sheet1");
        wb.active_sheet = 0;
        wb.active_cell = (1, 1);
        wb
    }

    pub fn add_sheet(&mut self, name: &str) -> usize {
        let idx = self.sheets.len();
        self.sheets.push(Sheet {
            name: name.to_string(),
            cells: HashMap::new(),
        });
        idx
    }

    pub fn get_value_a1(&self, sheet: &str, a1: &str) -> Result<VbaValue, VbaError> {
        let sheet_idx = self
            .sheet_index(sheet)
            .ok_or_else(|| VbaError::Runtime(format!("Unknown sheet: {sheet}")))?;
        let (row, col) = a1_to_row_col(a1)?;
        self.get_cell_value(sheet_idx, row, col)
    }

    pub fn get_formula_a1(&self, sheet: &str, a1: &str) -> Result<Option<String>, VbaError> {
        let sheet_idx = self
            .sheet_index(sheet)
            .ok_or_else(|| VbaError::Runtime(format!("Unknown sheet: {sheet}")))?;
        let (row, col) = a1_to_row_col(a1)?;
        self.get_cell_formula(sheet_idx, row, col)
    }

    pub fn set_value_a1(&mut self, sheet: &str, a1: &str, value: VbaValue) -> Result<(), VbaError> {
        let sheet_idx = self
            .sheet_index(sheet)
            .ok_or_else(|| VbaError::Runtime(format!("Unknown sheet: {sheet}")))?;
        let (row, col) = a1_to_row_col(a1)?;
        self.set_cell_value(sheet_idx, row, col, value)
    }

    pub fn set_formula_a1(&mut self, sheet: &str, a1: &str, formula: &str) -> Result<(), VbaError> {
        let sheet_idx = self
            .sheet_index(sheet)
            .ok_or_else(|| VbaError::Runtime(format!("Unknown sheet: {sheet}")))?;
        let (row, col) = a1_to_row_col(a1)?;
        self.set_cell_formula(sheet_idx, row, col, formula.to_string())
    }

    pub fn range_ref(&self, sheet: usize, a1: &str) -> Result<VbaRangeRef, VbaError> {
        let (r1, c1, r2, c2) = parse_range_a1(a1)?;
        Ok(VbaRangeRef {
            sheet,
            start_row: r1,
            start_col: c1,
            end_row: r2,
            end_col: c2,
        })
    }
}

impl Spreadsheet for InMemoryWorkbook {
    fn sheet_count(&self) -> usize {
        self.sheets.len()
    }

    fn sheet_name(&self, sheet: usize) -> Option<&str> {
        self.sheets.get(sheet).map(|s| s.name.as_str())
    }

    fn sheet_index(&self, name: &str) -> Option<usize> {
        self.sheets
            .iter()
            .position(|s| s.name.eq_ignore_ascii_case(name))
    }

    fn active_sheet(&self) -> usize {
        self.active_sheet
    }

    fn set_active_sheet(&mut self, sheet: usize) -> Result<(), VbaError> {
        if sheet >= self.sheets.len() {
            return Err(VbaError::Runtime(format!(
                "Sheet index out of range: {sheet}"
            )));
        }
        self.active_sheet = sheet;
        Ok(())
    }

    fn active_cell(&self) -> (u32, u32) {
        self.active_cell
    }

    fn set_active_cell(&mut self, row: u32, col: u32) -> Result<(), VbaError> {
        if row == 0 || col == 0 {
            return Err(VbaError::Runtime("ActiveCell is 1-based".to_string()));
        }
        self.active_cell = (row, col);
        Ok(())
    }

    fn get_cell_value(&self, sheet: usize, row: u32, col: u32) -> Result<VbaValue, VbaError> {
        let sh = self
            .sheets
            .get(sheet)
            .ok_or_else(|| VbaError::Runtime(format!("Unknown sheet index: {sheet}")))?;
        Ok(sh
            .cells
            .get(&(row, col))
            .map(|c| c.value.clone())
            .unwrap_or(VbaValue::Empty))
    }

    fn set_cell_value(
        &mut self,
        sheet: usize,
        row: u32,
        col: u32,
        value: VbaValue,
    ) -> Result<(), VbaError> {
        let sh = self
            .sheets
            .get_mut(sheet)
            .ok_or_else(|| VbaError::Runtime(format!("Unknown sheet index: {sheet}")))?;
        let cell = sh.cells.entry((row, col)).or_default();
        cell.value = value;
        Ok(())
    }

    fn get_cell_formula(
        &self,
        sheet: usize,
        row: u32,
        col: u32,
    ) -> Result<Option<String>, VbaError> {
        let sh = self
            .sheets
            .get(sheet)
            .ok_or_else(|| VbaError::Runtime(format!("Unknown sheet index: {sheet}")))?;
        Ok(sh.cells.get(&(row, col)).and_then(|c| c.formula.clone()))
    }

    fn set_cell_formula(
        &mut self,
        sheet: usize,
        row: u32,
        col: u32,
        formula: String,
    ) -> Result<(), VbaError> {
        let sh = self
            .sheets
            .get_mut(sheet)
            .ok_or_else(|| VbaError::Runtime(format!("Unknown sheet index: {sheet}")))?;
        let cell = sh.cells.entry((row, col)).or_default();
        cell.formula = Some(formula);
        Ok(())
    }

    fn log(&mut self, message: String) {
        self.output.push(message);
    }
}

/// Helper to create a `Range` object for a given A1 reference on the active sheet.
pub fn range_on_active_sheet(sheet: &dyn Spreadsheet, a1: &str) -> Result<VbaObjectRef, VbaError> {
    let (r1, c1, r2, c2) = parse_range_a1(a1)?;
    Ok(VbaObjectRef::new(VbaObject::Range(VbaRangeRef {
        sheet: sheet.active_sheet(),
        start_row: r1,
        start_col: c1,
        end_row: r2,
        end_col: c2,
    })))
}
