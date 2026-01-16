use std::collections::HashSet;

use thiserror::Error;

use crate::value::text_eq_case_insensitive;
use crate::{CellRef, Range};
use serde::{Deserialize, Serialize};

pub use crate::autofilter::{
    FilterColumn, SheetAutoFilter as AutoFilter, SortCondition, SortState,
};

/// Errors that can occur when creating or mutating a table.
#[derive(Debug, Clone, PartialEq, Eq, Error)]
pub enum TableError {
    #[error("table name cannot be empty")]
    EmptyName,
    #[error("table name exceeds Excel's 255 character limit")]
    NameTooLong,
    #[error("table name must start with an ASCII letter or '_'")]
    InvalidStartChar,
    #[error("table name contains invalid character '{ch}'")]
    InvalidChar { ch: char },
    #[error("table name conflicts with a cell or range reference")]
    ConflictsWithCellReference,
    #[error("table name is reserved")]
    ReservedName,
    #[error("table name already exists in workbook")]
    DuplicateName,
    #[error("worksheet not found")]
    SheetNotFound,
    #[error("table not found")]
    TableNotFound,
    #[error("table range is too small for header/totals row settings")]
    InvalidRange,
}

/// Identifier for a table within a worksheet.
///
/// Excel tables are workbook-scoped by name, but APIs may still need to refer to
/// tables by either their name or stable `id`.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(untagged)]
pub enum TableIdentifier {
    Name(String),
    Id(u32),
}

impl From<u32> for TableIdentifier {
    fn from(value: u32) -> Self {
        TableIdentifier::Id(value)
    }
}

impl From<String> for TableIdentifier {
    fn from(value: String) -> Self {
        TableIdentifier::Name(value)
    }
}

impl From<&str> for TableIdentifier {
    fn from(value: &str) -> Self {
        TableIdentifier::Name(value.to_string())
    }
}

/// Validate an Excel table name (ListObject name).
///
/// This mirrors Excel's rules approximately:
/// - Names are non-empty, <= 255 chars.
/// - First character must be an ASCII letter or `_`.
/// - Remaining characters may contain ASCII letters, digits, `_`, or `.`.
/// - Names may not look like A1 or R1C1 references (e.g. `A1`, `R1C1`).
/// - Names may not be reserved (`R`, `C`, `TRUE`, `FALSE`).
///
/// Workbook-wide uniqueness is enforced by [`crate::Workbook`] APIs.
pub fn validate_table_name(name: &str) -> Result<(), TableError> {
    let name = name.trim();
    if name.is_empty() {
        return Err(TableError::EmptyName);
    }
    if name.chars().count() > 255 {
        return Err(TableError::NameTooLong);
    }

    let mut chars = name.chars();
    let Some(first) = chars.next() else {
        return Err(TableError::EmptyName);
    };
    if !(first.is_ascii_alphabetic() || first == '_') {
        return Err(TableError::InvalidStartChar);
    }

    for ch in chars {
        if !(ch.is_ascii_alphanumeric() || ch == '_' || ch == '.') {
            return Err(TableError::InvalidChar { ch });
        }
    }

    // Names cannot look like cell/range references in A1 notation.
    if Range::from_a1(name).is_ok() {
        return Err(TableError::ConflictsWithCellReference);
    }

    // Names cannot look like R1C1 references.
    if name.eq_ignore_ascii_case("R")
        || name.eq_ignore_ascii_case("C")
        || name.eq_ignore_ascii_case("TRUE")
        || name.eq_ignore_ascii_case("FALSE")
    {
        return Err(TableError::ReservedName);
    }
    if looks_like_r1c1_reference(name) {
        return Err(TableError::ConflictsWithCellReference);
    }

    Ok(())
}

fn looks_like_r1c1_reference(name: &str) -> bool {
    // R1C1 or R1 or C1 style references (case-insensitive).
    // We treat all of these as invalid table names.
    let bytes = name.as_bytes();
    let Some((&first, rest)) = bytes.split_first() else {
        return false;
    };

    match first.to_ascii_uppercase() {
        b'R' => {
            if rest.is_empty() {
                return false;
            }

            // R<digits>
            let mut i = 0usize;
            while i < rest.len() && rest[i].is_ascii_digit() {
                i += 1;
            }
            if i == rest.len() {
                return true;
            }

            // R<digits>C<digits>
            if i == 0 || rest[i].to_ascii_uppercase() != b'C' {
                return false;
            }
            i += 1;
            if i >= rest.len() {
                return false;
            }
            let start = i;
            while i < rest.len() && rest[i].is_ascii_digit() {
                i += 1;
            }
            i > start && i == rest.len()
        }
        b'C' => !rest.is_empty() && rest.iter().all(|b| b.is_ascii_digit()),
        _ => false,
    }
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct TableStyleInfo {
    pub name: String,
    pub show_first_column: bool,
    pub show_last_column: bool,
    pub show_row_stripes: bool,
    pub show_column_stripes: bool,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct TableColumn {
    pub id: u32,
    pub name: String,
    /// Formula for calculated columns (stored without leading '=').
    pub formula: Option<String>,
    /// Totals row formula (stored without leading '=').
    pub totals_formula: Option<String>,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct Table {
    pub id: u32,
    pub name: String,
    pub display_name: String,
    pub range: Range,
    pub header_row_count: u32,
    pub totals_row_count: u32,
    pub columns: Vec<TableColumn>,
    pub style: Option<TableStyleInfo>,
    pub auto_filter: Option<AutoFilter>,
    /// Relationship ID (`r:id`) from the parent worksheet, if known.
    pub relationship_id: Option<String>,
    /// Path to the table part within the xlsx package, if known (e.g. `xl/tables/table1.xml`).
    pub part_path: Option<String>,
}

impl Table {
    pub(crate) fn rewrite_sheet_references(&mut self, old_name: &str, new_name: &str) {
        for column in &mut self.columns {
            if let Some(formula) = column.formula.as_mut() {
                *formula = crate::rewrite_sheet_names_in_formula(formula, old_name, new_name);
            }
            if let Some(formula) = column.totals_formula.as_mut() {
                *formula = crate::rewrite_sheet_names_in_formula(formula, old_name, new_name);
            }
        }
    }

    pub(crate) fn rewrite_sheet_references_internal_refs_only(
        &mut self,
        old_name: &str,
        new_name: &str,
    ) {
        for column in &mut self.columns {
            if let Some(formula) = column.formula.as_mut() {
                *formula =
                    crate::formula_rewrite::rewrite_sheet_names_in_formula_internal_refs_only(
                        formula, old_name, new_name,
                    );
            }
            if let Some(formula) = column.totals_formula.as_mut() {
                *formula =
                    crate::formula_rewrite::rewrite_sheet_names_in_formula_internal_refs_only(
                        formula, old_name, new_name,
                    );
            }
        }
    }

    pub(crate) fn rewrite_table_references(&mut self, renames: &[(String, String)]) {
        for column in &mut self.columns {
            if let Some(formula) = column.formula.as_mut() {
                *formula = crate::rewrite_table_names_in_formula(formula, renames);
            }
            if let Some(formula) = column.totals_formula.as_mut() {
                *formula = crate::rewrite_table_names_in_formula(formula, renames);
            }
        }
    }

    pub(crate) fn invalidate_deleted_sheet_references(
        &mut self,
        deleted_sheet: &str,
        sheet_order: &[String],
    ) {
        for column in &mut self.columns {
            if let Some(formula) = column.formula.as_mut() {
                *formula = crate::rewrite_deleted_sheet_references_in_formula(
                    formula,
                    deleted_sheet,
                    sheet_order,
                );
            }
            if let Some(formula) = column.totals_formula.as_mut() {
                *formula = crate::rewrite_deleted_sheet_references_in_formula(
                    formula,
                    deleted_sheet,
                    sheet_order,
                );
            }
        }
    }

    pub fn set_range(&mut self, new_range: Range) -> Result<(), TableError> {
        let required_rows = self.header_row_count + self.totals_row_count;
        if new_range.height() < required_rows {
            return Err(TableError::InvalidRange);
        }

        fn parse_default_column_number(name: &str) -> Option<u32> {
            let bytes = name.as_bytes();
            let prefix = b"column";
            if bytes.len() <= prefix.len() {
                return None;
            }
            if !bytes
                .get(..prefix.len())
                .is_some_and(|p| p.eq_ignore_ascii_case(prefix))
            {
                return None;
            }
            let digits = &name[prefix.len()..];
            if digits.is_empty() {
                return None;
            }
            // Only treat canonical `Column{n}` (no leading zeros) as a collision with our generated
            // default names. This matches the existing string-based collision behavior (e.g.
            // `Column01` does not collide with `Column1`).
            let digit_bytes = digits.as_bytes();
            if digit_bytes.len() > 1 && digit_bytes[0] == b'0' {
                return None;
            }
            if !digit_bytes.iter().all(|b| b.is_ascii_digit()) {
                return None;
            }
            digits.parse().ok().filter(|n: &u32| *n > 0)
        }

        let new_col_count = new_range.width() as usize;
        let current_col_count = self.columns.len();

        if new_col_count < current_col_count {
            self.columns.truncate(new_col_count);
        } else if new_col_count > current_col_count {
            let mut used_default_nums: HashSet<u32> = self
                .columns
                .iter()
                .filter_map(|c| parse_default_column_number(&c.name))
                .collect();
            let mut next_id = self.columns.iter().map(|c| c.id).max().unwrap_or(0) + 1;
            let mut next_default_num: u32 = 1;

            for _ in current_col_count..new_col_count {
                let name = loop {
                    let n = next_default_num;
                    next_default_num = next_default_num.saturating_add(1);
                    if used_default_nums.insert(n) {
                        break format!("Column{n}");
                    }
                };
                self.columns.push(TableColumn {
                    id: next_id,
                    name,
                    formula: None,
                    totals_formula: None,
                });
                next_id += 1;
            }
        }

        self.range = new_range;

        // Best-effort keep auto filter metadata consistent.
        if let Some(auto_filter) = &mut self.auto_filter {
            auto_filter.range = new_range;
            auto_filter
                .filter_columns
                .retain(|c| (c.col_id as usize) < new_col_count);
            if let Some(sort_state) = &mut auto_filter.sort_state {
                sort_state
                    .conditions
                    .retain(|cond| cond.range.intersects(&new_range));
            }
        }

        Ok(())
    }

    pub fn data_range(&self) -> Option<Range> {
        let r = self.range;
        let start_row = r.start.row.checked_add(self.header_row_count)?;
        let end_row = r.end.row.saturating_sub(self.totals_row_count);
        if start_row > end_row {
            return None;
        }
        Some(Range::new(
            CellRef::new(start_row, r.start.col),
            CellRef::new(end_row, r.end.col),
        ))
    }

    pub fn header_range(&self) -> Option<Range> {
        if self.header_row_count == 0 {
            return None;
        }
        let r = self.range;
        let header_end = self
            .header_row_count
            .checked_sub(1)
            .and_then(|delta| r.start.row.checked_add(delta))?;
        Some(Range::new(r.start, CellRef::new(header_end, r.end.col)))
    }

    pub fn totals_range(&self) -> Option<Range> {
        if self.totals_row_count == 0 {
            return None;
        }
        let r = self.range;
        let start_row = r
            .end
            .row
            .saturating_sub(self.totals_row_count.saturating_sub(1));
        Some(Range::new(CellRef::new(start_row, r.start.col), r.end))
    }

    pub fn column_index(&self, name: &str) -> Option<u32> {
        self.columns
            .iter()
            .position(|c| text_eq_case_insensitive(&c.name, name))
            .map(|idx| idx as u32)
    }

    pub fn column_range_in_area(&self, column_name: &str, area: TableArea) -> Option<Range> {
        let r = self.range;
        let col_offset = self.column_index(column_name)?;
        let col = r.start.col + col_offset;

        match area {
            TableArea::Headers => self.header_range().map(|hr| {
                Range::new(
                    CellRef::new(hr.start.row, col),
                    CellRef::new(hr.end.row, col),
                )
            }),
            TableArea::Totals => self.totals_range().map(|tr| {
                Range::new(
                    CellRef::new(tr.start.row, col),
                    CellRef::new(tr.end.row, col),
                )
            }),
            TableArea::Data => self.data_range().map(|dr| {
                Range::new(
                    CellRef::new(dr.start.row, col),
                    CellRef::new(dr.end.row, col),
                )
            }),
            TableArea::All => Some(Range::new(
                CellRef::new(r.start.row, col),
                CellRef::new(r.end.row, col),
            )),
        }
    }

    pub fn cell_for_this_row(&self, current_cell: CellRef, column_name: &str) -> Option<CellRef> {
        let r = self.range;
        let data_range = self.data_range()?;
        if !data_range.contains(current_cell) {
            return None;
        }
        let col_offset = self.column_index(column_name)?;
        Some(CellRef::new(current_cell.row, r.start.col + col_offset))
    }
}

#[derive(Debug, Copy, Clone, PartialEq, Eq)]
pub enum TableArea {
    Headers,
    Data,
    Totals,
    All,
}
