use super::{StructuredColumn, StructuredColumns, StructuredRef, StructuredRefItem};
use crate::eval::CellAddr;
use formula_model::table::{Table, TableArea};
use formula_model::{CellRef, Range};

fn addr_to_model(addr: CellAddr) -> CellRef {
    CellRef::new(addr.row, addr.col)
}

fn model_to_addr(cell: CellRef) -> CellAddr {
    CellAddr {
        row: cell.row,
        col: cell.col,
    }
}

fn column_index_ci(table: &Table, name: &str) -> Option<u32> {
    table
        .columns
        .iter()
        .position(|c| c.name.eq_ignore_ascii_case(name))
        .map(|idx| idx as u32)
}

fn base_range_for_area(table: &Table, area: TableArea) -> Result<Range, String> {
    Ok(match area {
        TableArea::Headers => table
            .header_range()
            .ok_or_else(|| "table has no header row".to_string())?,
        TableArea::Totals => table
            .totals_range()
            .ok_or_else(|| "table has no totals row".to_string())?,
        TableArea::Data => table
            .data_range()
            .ok_or_else(|| "table has no data rows".to_string())?,
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

fn column_interval_ci(table: &Table, col: &StructuredColumn) -> Result<(u32, u32), String> {
    match col {
        StructuredColumn::Single(name) => {
            let idx =
                column_index_ci(table, name).ok_or_else(|| format!("unknown column '{name}'"))?;
            Ok((idx, idx))
        }
        StructuredColumn::Range { start, end } => {
            let start_idx =
                column_index_ci(table, start).ok_or_else(|| format!("unknown column '{start}'"))?;
            let end_idx =
                column_index_ci(table, end).ok_or_else(|| format!("unknown column '{end}'"))?;
            Ok(normalize_column_interval(start_idx, end_idx))
        }
    }
}

fn column_intervals_ci(
    table: &Table,
    columns: &StructuredColumns,
) -> Result<Vec<(u32, u32)>, String> {
    match columns {
        StructuredColumns::All => Ok(Vec::new()),
        StructuredColumns::Single(name) => {
            column_interval_ci(table, &StructuredColumn::Single(name.clone()))
                .map(|interval| vec![interval])
        }
        StructuredColumns::Range { start, end } => column_interval_ci(
            table,
            &StructuredColumn::Range {
                start: start.clone(),
                end: end.clone(),
            },
        )
        .map(|interval| vec![interval]),
        StructuredColumns::Multi(parts) => {
            let mut out = Vec::with_capacity(parts.len());
            for part in parts {
                out.push(column_interval_ci(table, part)?);
            }
            Ok(out)
        }
    }
}

fn merge_column_intervals(mut intervals: Vec<(u32, u32)>) -> Vec<(u32, u32)> {
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
) -> Result<Vec<(usize, CellAddr, CellAddr)>, String> {
    let (sheet_id, table) = find_table(tables_by_sheet, origin_sheet, origin_cell, sref)?;

    let item = sref.item.clone().unwrap_or(StructuredRefItem::Data);
    let ranges = match item {
        StructuredRefItem::ThisRow => resolve_this_row(table, origin_cell, &sref.columns)?,
        StructuredRefItem::Headers => resolve_area(table, TableArea::Headers, &sref.columns)?,
        StructuredRefItem::Totals => resolve_area(table, TableArea::Totals, &sref.columns)?,
        StructuredRefItem::All => resolve_area(table, TableArea::All, &sref.columns)?,
        StructuredRefItem::Data => resolve_area(table, TableArea::Data, &sref.columns)?,
    };

    Ok(ranges
        .into_iter()
        .map(|(start, end)| (sheet_id, start, end))
        .collect())
}

fn find_table<'a>(
    tables_by_sheet: &'a [Vec<Table>],
    origin_sheet: usize,
    origin_cell: CellAddr,
    sref: &StructuredRef,
) -> Result<(usize, &'a Table), String> {
    if let Some(name) = &sref.table_name {
        for (sheet_id, tables) in tables_by_sheet.iter().enumerate() {
            if let Some(table) = tables.iter().find(|t| {
                t.name.eq_ignore_ascii_case(name) || t.display_name.eq_ignore_ascii_case(name)
            }) {
                return Ok((sheet_id, table));
            }
        }
        return Err(format!("unknown table '{name}'"));
    }

    let tables = tables_by_sheet
        .get(origin_sheet)
        .ok_or_else(|| format!("sheet index {origin_sheet} out of bounds"))?;

    let origin_cell_model = addr_to_model(origin_cell);
    let table = tables
        .iter()
        .find(|t| t.range.contains(origin_cell_model))
        .ok_or_else(|| {
            "structured reference without table name used outside of a table".to_string()
        })?;

    Ok((origin_sheet, table))
}

fn resolve_area(
    table: &Table,
    area: TableArea,
    columns: &StructuredColumns,
) -> Result<Vec<(CellAddr, CellAddr)>, String> {
    let base = base_range_for_area(table, area)?;
    if matches!(columns, StructuredColumns::All) {
        return Ok(vec![(model_to_addr(base.start), model_to_addr(base.end))]);
    }

    let intervals = merge_column_intervals(column_intervals_ci(table, columns)?);
    let table_start = table.range.start;
    let mut out = Vec::with_capacity(intervals.len());
    for (left_idx, right_idx) in intervals {
        let range = Range::new(
            CellRef::new(base.start.row, table_start.col + left_idx),
            CellRef::new(base.end.row, table_start.col + right_idx),
        );
        out.push((model_to_addr(range.start), model_to_addr(range.end)));
    }
    Ok(out)
}

fn resolve_this_row(
    table: &Table,
    origin_cell: CellAddr,
    columns: &StructuredColumns,
) -> Result<Vec<(CellAddr, CellAddr)>, String> {
    let data_range = table
        .data_range()
        .ok_or_else(|| "table has no data rows".to_string())?;
    if !data_range.contains(addr_to_model(origin_cell)) {
        return Err("this-row structured reference used outside of table data row".to_string());
    }
    let row = origin_cell.row;

    match columns {
        StructuredColumns::All => Ok(vec![(
            CellAddr {
                row,
                col: table.range.start.col,
            },
            CellAddr {
                row,
                col: table.range.end.col,
            },
        )]),
        _ => {
            let intervals = merge_column_intervals(column_intervals_ci(table, columns)?);
            let mut out = Vec::with_capacity(intervals.len());
            for (left_idx, right_idx) in intervals {
                out.push((
                    CellAddr {
                        row,
                        col: table.range.start.col + left_idx,
                    },
                    CellAddr {
                        row,
                        col: table.range.start.col + right_idx,
                    },
                ));
            }
            Ok(out)
        }
    }
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
            item: None,
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
