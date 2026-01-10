use super::{StructuredColumns, StructuredRef, StructuredRefItem};
use crate::eval::CellAddr;
use formula_model::table::{Table, TableArea};
use formula_model::{CellRef, Range};

fn addr_to_model(addr: CellAddr) -> CellRef {
    CellRef::new(addr.row, addr.col)
}

fn model_to_addr(cell: CellRef) -> CellAddr {
    CellAddr { row: cell.row, col: cell.col }
}

fn column_index_ci(table: &Table, name: &str) -> Option<u32> {
    table
        .columns
        .iter()
        .position(|c| c.name.eq_ignore_ascii_case(name))
        .map(|idx| idx as u32)
}

fn column_range_in_area_ci(table: &Table, column_name: &str, area: TableArea) -> Option<Range> {
    let r = table.range;
    let col_offset = column_index_ci(table, column_name)?;
    let col = r.start.col + col_offset;

    match area {
        TableArea::Headers => table.header_range().map(|hr| {
            Range::new(
                CellRef::new(hr.start.row, col),
                CellRef::new(hr.end.row, col),
            )
        }),
        TableArea::Totals => table.totals_range().map(|tr| {
            Range::new(
                CellRef::new(tr.start.row, col),
                CellRef::new(tr.end.row, col),
            )
        }),
        TableArea::Data => table.data_range().map(|dr| {
            Range::new(
                CellRef::new(dr.start.row, col),
                CellRef::new(dr.end.row, col),
            )
        }),
        TableArea::All => Some(Range::new(CellRef::new(r.start.row, col), CellRef::new(r.end.row, col))),
    }
}

fn cell_for_this_row_ci(table: &Table, current_cell: CellRef, column_name: &str) -> Option<CellRef> {
    let r = table.range;
    let data_range = table.data_range()?;
    if !data_range.contains(current_cell) {
        return None;
    }
    let col_offset = column_index_ci(table, column_name)?;
    Some(CellRef::new(current_cell.row, r.start.col + col_offset))
}

/// Resolve a structured reference into a concrete `(sheet_id, start, end)` range.
///
/// The caller provides `tables_by_sheet` indexed by `sheet_id`.
pub fn resolve_structured_ref(
    tables_by_sheet: &[Vec<Table>],
    origin_sheet: usize,
    origin_cell: CellAddr,
    sref: &StructuredRef,
) -> Result<(usize, CellAddr, CellAddr), String> {
    let (sheet_id, table) = find_table(tables_by_sheet, origin_sheet, origin_cell, sref)?;

    let item = sref.item.clone().unwrap_or(StructuredRefItem::Data);
    let (start, end) = match item {
        StructuredRefItem::ThisRow => resolve_this_row(table, origin_cell, &sref.columns)?,
        StructuredRefItem::Headers => resolve_area(table, TableArea::Headers, &sref.columns)?,
        StructuredRefItem::Totals => resolve_area(table, TableArea::Totals, &sref.columns)?,
        StructuredRefItem::All => resolve_area(table, TableArea::All, &sref.columns)?,
        StructuredRefItem::Data => resolve_area(table, TableArea::Data, &sref.columns)?,
    };

    Ok((sheet_id, start, end))
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
        .ok_or_else(|| "structured reference without table name used outside of a table".to_string())?;

    Ok((origin_sheet, table))
}

fn resolve_area(table: &Table, area: TableArea, columns: &StructuredColumns) -> Result<(CellAddr, CellAddr), String> {
    match columns {
        StructuredColumns::All => {
            let range = match area {
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
            };
            Ok((model_to_addr(range.start), model_to_addr(range.end)))
        }
        StructuredColumns::Single(name) => {
            let range = column_range_in_area_ci(table, name, area)
                .ok_or_else(|| format!("unknown column '{name}'"))?;
            Ok((model_to_addr(range.start), model_to_addr(range.end)))
        }
        StructuredColumns::Range { start, end } => {
            let start_idx = column_index_ci(table, start).ok_or_else(|| format!("unknown column '{start}'"))?;
            let end_idx = column_index_ci(table, end).ok_or_else(|| format!("unknown column '{end}'"))?;
            let (left_idx, right_idx) = if start_idx <= end_idx {
                (start_idx, end_idx)
            } else {
                (end_idx, start_idx)
            };

            let base = match area {
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
            };

            let table_start = table.range.start;
            let range = Range::new(
                CellRef::new(base.start.row, table_start.col + left_idx),
                CellRef::new(base.end.row, table_start.col + right_idx),
            );
            Ok((model_to_addr(range.start), model_to_addr(range.end)))
        }
    }
}

fn resolve_this_row(
    table: &Table,
    origin_cell: CellAddr,
    columns: &StructuredColumns,
) -> Result<(CellAddr, CellAddr), String> {
    let row = origin_cell.row;
    let data_range = table
        .data_range()
        .ok_or_else(|| "table has no data rows".to_string())?;
    if row < data_range.start.row || row > data_range.end.row {
        return Err("this-row structured reference used outside of table data row".to_string());
    }

    match columns {
        StructuredColumns::All => Ok((
            CellAddr {
                row,
                col: table.range.start.col,
            },
            CellAddr { row, col: table.range.end.col },
        )),
        StructuredColumns::Single(name) => {
            let cell = cell_for_this_row_ci(table, addr_to_model(origin_cell), name)
                .ok_or_else(|| format!("unknown column '{name}'"))?;
            let addr = model_to_addr(cell);
            Ok((addr, addr))
        }
        StructuredColumns::Range { start, end } => {
            let start_idx = column_index_ci(table, start).ok_or_else(|| format!("unknown column '{start}'"))?;
            let end_idx = column_index_ci(table, end).ok_or_else(|| format!("unknown column '{end}'"))?;
            let (left_idx, right_idx) = if start_idx <= end_idx {
                (start_idx, end_idx)
            } else {
                (end_idx, start_idx)
            };
            Ok((
                CellAddr {
                    row,
                    col: table.range.start.col + left_idx,
                },
                CellAddr {
                    row,
                    col: table.range.start.col + right_idx,
                },
            ))
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
        let (_sheet, start, end) = resolve_structured_ref(&tables, 0, CellAddr { row: 0, col: 0 }, &sref).unwrap();
        assert_eq!(start, CellAddr { row: 1, col: 1 });
        assert_eq!(end, CellAddr { row: 2, col: 1 });
    }
}
