use super::{StructuredColumn, StructuredColumns, StructuredRef, StructuredRefItem};
use crate::eval::CellAddr;
use crate::value::{cmp_case_insensitive, ErrorKind};
use formula_model::table::{Table, TableArea};
use formula_model::{CellRef, Range};
use std::cmp::Ordering;

fn addr_to_model(addr: CellAddr) -> CellRef {
    CellRef::new(addr.row, addr.col)
}

fn column_index_ci(table: &Table, name: &str) -> Option<u32> {
    table
        .columns
        .iter()
        .position(|c| cmp_case_insensitive(&c.name, name) == Ordering::Equal)
        .map(|idx| idx as u32)
}

fn base_range_for_area(table: &Table, area: TableArea) -> Result<Range, ErrorKind> {
    Ok(match area {
        TableArea::Headers => table.header_range().ok_or(ErrorKind::Ref)?,
        TableArea::Totals => table.totals_range().ok_or(ErrorKind::Ref)?,
        TableArea::Data => table.data_range().ok_or(ErrorKind::Ref)?,
        TableArea::All => table.range,
    })
}

fn normalize_column_interval(start_idx: u32, end_idx: u32) -> (u32, u32) {
    if start_idx <= end_idx {
        (start_idx, end_idx)
    } else {
        (end_idx, start_idx)
    }
}

fn column_interval_ci(table: &Table, col: &StructuredColumn) -> Result<(u32, u32), ErrorKind> {
    match col {
        StructuredColumn::Single(name) => {
            let idx = column_index_ci(table, name).ok_or(ErrorKind::Ref)?;
            Ok((idx, idx))
        }
        StructuredColumn::Range { start, end } => {
            let start_idx = column_index_ci(table, start).ok_or(ErrorKind::Ref)?;
            let end_idx = column_index_ci(table, end).ok_or(ErrorKind::Ref)?;
            Ok(normalize_column_interval(start_idx, end_idx))
        }
    }
}

fn column_intervals_ci(
    table: &Table,
    columns: &StructuredColumns,
) -> Result<Vec<(u32, u32)>, ErrorKind> {
    match columns {
        StructuredColumns::All => Ok(Vec::new()),
        StructuredColumns::Single(name) => {
            let idx = column_index_ci(table, name).ok_or(ErrorKind::Ref)?;
            Ok(vec![(idx, idx)])
        }
        StructuredColumns::Range { start, end } => {
            let start_idx = column_index_ci(table, start).ok_or(ErrorKind::Ref)?;
            let end_idx = column_index_ci(table, end).ok_or(ErrorKind::Ref)?;
            Ok(vec![normalize_column_interval(start_idx, end_idx)])
        }
        StructuredColumns::Multi(parts) => {
            let mut out = Vec::with_capacity(parts.len());
            for part in parts {
                out.push(column_interval_ci(table, part)?);
            }
            Ok(out)
        }
    }
}

fn merge_intervals(mut intervals: Vec<(u32, u32)>) -> Vec<(u32, u32)> {
    if intervals.is_empty() {
        return intervals;
    }
    intervals.sort_by_key(|(start, end)| (*start, *end));
    let mut out: Vec<(u32, u32)> = Vec::new();
    let mut current = intervals[0];
    for (start, end) in intervals.into_iter().skip(1) {
        if start <= current.1.saturating_add(1) {
            current.1 = current.1.max(end);
        } else {
            out.push(current);
            current = (start, end);
        }
    }
    out.push(current);
    out
}

/// Resolve a structured reference into one or more concrete `(sheet_id, start, end)` ranges.
///
/// The caller provides `tables_by_sheet` indexed by `sheet_id`.
pub fn resolve_structured_ref(
    tables_by_sheet: &[Vec<Table>],
    origin_sheet: usize,
    origin_cell: CellAddr,
    sref: &StructuredRef,
) -> Result<Vec<(usize, CellAddr, CellAddr)>, ErrorKind> {
    let (sheet_id, table) = find_table(tables_by_sheet, origin_sheet, origin_cell, sref)?;
    // `@ThisRow` structured references are only valid on the table's own sheet. Once the engine has
    // stable sheet ids, callers can pass a different `origin_sheet`, so preserve the explicit
    // sheet-id check here.
    if sheet_id != origin_sheet
        && sref
            .items
            .iter()
            .any(|item| matches!(item, StructuredRefItem::ThisRow))
    {
        return Err(ErrorKind::Name);
    }

    let ranges = resolve_structured_ref_in_table(table, origin_cell, sref)?;
    Ok(ranges
        .into_iter()
        .map(|(start, end)| (sheet_id, start, end))
        .collect())
}

/// Resolve a structured reference against a specific table.
///
/// Returns one or more `(start, end)` ranges in the table's sheet coordinate space.
pub fn resolve_structured_ref_in_table(
    table: &Table,
    origin_cell: CellAddr,
    sref: &StructuredRef,
) -> Result<Vec<(CellAddr, CellAddr)>, ErrorKind> {
    // Resolve the column selection once. Column intervals are 0-based indices into the table's
    // column set (relative to `table.range.start.col`).
    let table_start = table.range.start;
    let table_width = table
        .range
        .end
        .col
        .saturating_sub(table_start.col)
        .saturating_add(1);
    if table_width == 0 {
        return Err(ErrorKind::Ref);
    }

    let col_intervals = if matches!(sref.columns, StructuredColumns::All) {
        vec![(0, table_width.saturating_sub(1))]
    } else {
        merge_intervals(column_intervals_ci(table, &sref.columns)?)
    };

    let mut row_intervals: Vec<(u32, u32)> = Vec::new();
    let mut push_item_rows = |item: StructuredRefItem| -> Result<(), ErrorKind> {
        match item {
            StructuredRefItem::ThisRow => {
                let data_range = table.data_range().ok_or(ErrorKind::Ref)?;
                if !data_range.contains(addr_to_model(origin_cell)) {
                    return Err(ErrorKind::Name);
                }
                row_intervals.push((origin_cell.row, origin_cell.row));
            }
            StructuredRefItem::Headers => {
                let base = base_range_for_area(table, TableArea::Headers)?;
                row_intervals.push((base.start.row, base.end.row));
            }
            StructuredRefItem::Totals => {
                let base = base_range_for_area(table, TableArea::Totals)?;
                row_intervals.push((base.start.row, base.end.row));
            }
            StructuredRefItem::All => {
                let base = base_range_for_area(table, TableArea::All)?;
                row_intervals.push((base.start.row, base.end.row));
            }
            StructuredRefItem::Data => {
                let base = base_range_for_area(table, TableArea::Data)?;
                row_intervals.push((base.start.row, base.end.row));
            }
        }
        Ok(())
    };

    // Excel defaults to `#Data` when no item specifier is present.
    if sref.items.is_empty() {
        push_item_rows(StructuredRefItem::Data)?;
    } else {
        for item in sref.items.iter().cloned() {
            push_item_rows(item)?;
        }
    }

    let row_intervals = merge_intervals(row_intervals);
    let mut out: Vec<(CellAddr, CellAddr)> =
        Vec::with_capacity(row_intervals.len().saturating_mul(col_intervals.len()));
    for (row_start, row_end) in row_intervals {
        for (left_idx, right_idx) in &col_intervals {
            out.push((
                CellAddr {
                    row: row_start,
                    col: table_start.col + *left_idx,
                },
                CellAddr {
                    row: row_end,
                    col: table_start.col + *right_idx,
                },
            ));
        }
    }

    // Stable ordering for deterministic union behavior.
    out.sort_by(|a, b| {
        a.0.row
            .cmp(&b.0.row)
            .then_with(|| a.0.col.cmp(&b.0.col))
            .then_with(|| a.1.row.cmp(&b.1.row))
            .then_with(|| a.1.col.cmp(&b.1.col))
    });

    Ok(out)
}

fn find_table<'a>(
    tables_by_sheet: &'a [Vec<Table>],
    origin_sheet: usize,
    origin_cell: CellAddr,
    sref: &StructuredRef,
) -> Result<(usize, &'a Table), ErrorKind> {
    if let Some(name) = &sref.table_name {
        for (sheet_id, tables) in tables_by_sheet.iter().enumerate() {
            if let Some(table) = tables.iter().find(|t| {
                cmp_case_insensitive(&t.name, name) == Ordering::Equal
                    || cmp_case_insensitive(&t.display_name, name) == Ordering::Equal
            }) {
                return Ok((sheet_id, table));
            }
        }
        return Err(ErrorKind::Name);
    }

    let tables = tables_by_sheet.get(origin_sheet).ok_or(ErrorKind::Ref)?;

    let origin_cell_model = addr_to_model(origin_cell);
    let table = tables
        .iter()
        .find(|t| t.range.contains(origin_cell_model))
        .ok_or(ErrorKind::Name)?;

    Ok((origin_sheet, table))
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::structured_refs::StructuredColumns;
    use formula_model::table::TableColumn;

    fn table_fixture() -> Table {
        Table {
            id: 1,
            name: "Table1".into(),
            display_name: "Table1".into(),
            range: Range::from_a1("A1:C3").unwrap(),
            header_row_count: 1,
            totals_row_count: 0,
            columns: vec![
                TableColumn {
                    id: 1,
                    name: "Col1".into(),
                    formula: None,
                    totals_formula: None,
                },
                TableColumn {
                    id: 2,
                    name: "Col2".into(),
                    formula: None,
                    totals_formula: None,
                },
                TableColumn {
                    id: 3,
                    name: "Col3".into(),
                    formula: None,
                    totals_formula: None,
                },
            ],
            style: None,
            auto_filter: None,
            relationship_id: None,
            part_path: None,
        }
    }

    #[test]
    fn resolves_table_column_data_range() {
        let tables = vec![vec![table_fixture()]];
        let sref = StructuredRef {
            table_name: Some("Table1".into()),
            items: Vec::new(),
            columns: StructuredColumns::Single("Col2".into()),
        };
        let ranges =
            resolve_structured_ref(&tables, 0, CellAddr { row: 0, col: 0 }, &sref).unwrap();
        let [(sheet_id, start, end)] = ranges.as_slice() else {
            panic!("expected a single resolved range");
        };
        assert_eq!(*sheet_id, 0);
        assert_eq!(*start, CellAddr { row: 1, col: 1 });
        assert_eq!(*end, CellAddr { row: 2, col: 1 });
    }
}
