use crate::{CellRef, Range};
use serde::{Deserialize, Serialize};

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

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct FilterColumn {
    /// 0-based column offset within the table's autofilter range.
    pub col_id: u32,
    /// Allowed values for the filter (Excel "filters" element).
    pub values: Vec<String>,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct SortCondition {
    pub range: Range,
    pub descending: bool,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct SortState {
    pub conditions: Vec<SortCondition>,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct AutoFilter {
    pub range: Range,
    pub filter_columns: Vec<FilterColumn>,
    pub sort_state: Option<SortState>,
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

    pub fn data_range(&self) -> Option<Range> {
        let r = self.range;
        let start_row = r.start.row + self.header_row_count;
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
        Some(Range::new(
            r.start,
            CellRef::new(r.start.row + self.header_row_count - 1, r.end.col),
        ))
    }

    pub fn totals_range(&self) -> Option<Range> {
        if self.totals_row_count == 0 {
            return None;
        }
        let r = self.range;
        let start_row = r.end.row + 1 - self.totals_row_count;
        Some(Range::new(CellRef::new(start_row, r.start.col), r.end))
    }

    pub fn column_index(&self, name: &str) -> Option<u32> {
        self.columns
            .iter()
            .position(|c| c.name.eq_ignore_ascii_case(name))
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
