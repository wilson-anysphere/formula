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
    /// Set the literal value for a cell and clear any existing formula.
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

    /// Set a formula for a cell.
    fn set_cell_formula(
        &mut self,
        sheet: usize,
        row: u32,
        col: u32,
        formula: String,
    ) -> Result<(), VbaError>;

    /// Clear the value and formula for a cell (equivalent to Excel's `ClearContents`).
    fn clear_cell_contents(&mut self, sheet: usize, row: u32, col: u32) -> Result<(), VbaError>;

    fn log(&mut self, message: String);

    /// Best-effort "used cell" queries for optimizing operations like `Range.End` without
    /// scanning the full Excel sheet extents.
    ///
    /// Implementations should consider a cell "used" if it has a non-empty value or a formula.
    /// The runtime will fall back to cell-by-cell scanning when these return `None`.
    fn last_used_row_in_column(&self, _sheet: usize, _col: u32, _start_row: u32) -> Option<u32> {
        None
    }

    fn next_used_row_in_column(&self, _sheet: usize, _col: u32, _start_row: u32) -> Option<u32> {
        None
    }

    fn last_used_col_in_row(&self, _sheet: usize, _row: u32, _start_col: u32) -> Option<u32> {
        None
    }

    fn next_used_col_in_row(&self, _sheet: usize, _row: u32, _start_col: u32) -> Option<u32> {
        None
    }
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
    RangeRows { range: VbaRangeRef },
    RangeColumns { range: VbaRangeRef },
    Collection { items: Vec<VbaValue> },
    Dictionary { items: HashMap<String, VbaValue> },
    Err(VbaErrObject),
}

#[derive(Debug, Clone, Default)]
pub struct VbaErrObject {
    pub number: i32,
    pub description: String,
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
    const MAX_ROW: u32 = 1_048_576;
    const MAX_COL: u32 = 16_384;

    #[derive(Clone, Copy)]
    enum A1Ref {
        Cell { row: u32, col: u32 },
        Row { row: u32 },
        Col { col: u32 },
    }

    fn parse_ref(token: &str) -> Result<A1Ref, VbaError> {
        let token = token.trim();
        if token.is_empty() {
            return Err(VbaError::Runtime("Invalid A1 reference: empty".to_string()));
        }

        let mut letters = String::new();
        let mut digits = String::new();
        for ch in token.chars() {
            if ch == '$' {
                continue;
            }
            if ch.is_ascii_alphabetic() {
                letters.push(ch);
            } else if ch.is_ascii_digit() {
                digits.push(ch);
            } else {
                return Err(VbaError::Runtime(format!("Invalid A1 reference: {token}")));
            }
        }

        match (!letters.is_empty(), !digits.is_empty()) {
            (true, true) => {
                let (row, col) = a1_to_row_col(&format!("{letters}{digits}"))?;
                Ok(A1Ref::Cell { row, col })
            }
            (true, false) => {
                // Entire column reference like `A` or `AA`.
                let mut col: u32 = 0;
                for ch in letters.chars() {
                    col = col
                        .checked_mul(26)
                        .and_then(|v| v.checked_add((ch.to_ascii_uppercase() as u8 - b'A' + 1) as u32))
                        .ok_or_else(|| VbaError::Runtime(format!("Invalid A1 reference: {token}")))?;
                }
                if col == 0 {
                    return Err(VbaError::Runtime(format!("Invalid A1 reference: {token}")));
                }
                Ok(A1Ref::Col { col })
            }
            (false, true) => {
                // Entire row reference like `1`.
                let row: u32 = digits
                    .parse()
                    .map_err(|_| VbaError::Runtime(format!("Invalid A1 reference: {token}")))?;
                if row == 0 {
                    return Err(VbaError::Runtime(format!("Invalid A1 reference: {token}")));
                }
                Ok(A1Ref::Row { row })
            }
            (false, false) => Err(VbaError::Runtime(format!("Invalid A1 reference: {token}"))),
        }
    }

    let parts: Vec<&str> = a1.split(':').collect();
    let (start, end) = match parts.as_slice() {
        [single] => {
            let r = parse_ref(single)?;
            (r, r)
        }
        [a, b] => (parse_ref(a)?, parse_ref(b)?),
        _ => return Err(VbaError::Runtime(format!("Invalid A1 range: {a1}"))),
    };

    match (start, end) {
        (A1Ref::Cell { row: r1, col: c1 }, A1Ref::Cell { row: r2, col: c2 }) => {
            Ok((r1.min(r2), c1.min(c2), r1.max(r2), c1.max(c2)))
        }
        (A1Ref::Col { col: c1 }, A1Ref::Col { col: c2 }) => Ok((1, c1.min(c2), MAX_ROW, c1.max(c2))),
        (A1Ref::Row { row: r1 }, A1Ref::Row { row: r2 }) => Ok((r1.min(r2), 1, r1.max(r2), MAX_COL)),
        _ => Err(VbaError::Runtime(format!("Invalid A1 range: {a1}"))),
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
        cell.formula = None;
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
        cell.formula = if formula.trim().is_empty() {
            None
        } else {
            Some(formula)
        };
        Ok(())
    }

    fn clear_cell_contents(&mut self, sheet: usize, row: u32, col: u32) -> Result<(), VbaError> {
        let sh = self
            .sheets
            .get_mut(sheet)
            .ok_or_else(|| VbaError::Runtime(format!("Unknown sheet index: {sheet}")))?;
        sh.cells.remove(&(row, col));
        Ok(())
    }

    fn log(&mut self, message: String) {
        self.output.push(message);
    }

    fn last_used_row_in_column(&self, sheet: usize, col: u32, start_row: u32) -> Option<u32> {
        let sh = self.sheets.get(sheet)?;
        sh.cells
            .iter()
            .filter_map(|(&(row, c), cell)| {
                if c != col || row > start_row {
                    return None;
                }
                if !matches!(cell.value, VbaValue::Empty) || cell.formula.is_some() {
                    Some(row)
                } else {
                    None
                }
            })
            .max()
    }

    fn next_used_row_in_column(&self, sheet: usize, col: u32, start_row: u32) -> Option<u32> {
        let sh = self.sheets.get(sheet)?;
        sh.cells
            .iter()
            .filter_map(|(&(row, c), cell)| {
                if c != col || row < start_row {
                    return None;
                }
                if !matches!(cell.value, VbaValue::Empty) || cell.formula.is_some() {
                    Some(row)
                } else {
                    None
                }
            })
            .min()
    }

    fn last_used_col_in_row(&self, sheet: usize, row: u32, start_col: u32) -> Option<u32> {
        let sh = self.sheets.get(sheet)?;
        sh.cells
            .iter()
            .filter_map(|(&(r, col), cell)| {
                if r != row || col > start_col {
                    return None;
                }
                if !matches!(cell.value, VbaValue::Empty) || cell.formula.is_some() {
                    Some(col)
                } else {
                    None
                }
            })
            .max()
    }

    fn next_used_col_in_row(&self, sheet: usize, row: u32, start_col: u32) -> Option<u32> {
        let sh = self.sheets.get(sheet)?;
        sh.cells
            .iter()
            .filter_map(|(&(r, col), cell)| {
                if r != row || col < start_col {
                    return None;
                }
                if !matches!(cell.value, VbaValue::Empty) || cell.formula.is_some() {
                    Some(col)
                } else {
                    None
                }
            })
            .min()
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
