use formula_model::autofilter::{FilterJoin, SortCondition, SortState};
use formula_model::{Cell, CellRef, CellValue, FilterColumn, FilterCriterion, FilterValue, Range, SheetAutoFilter, Workbook};
use formula_xlsx::{read_workbook, write_workbook};
use tempfile::tempdir;
 
#[test]
fn worksheet_autofilter_round_trips_through_read_write_workbook() {
    let mut workbook = Workbook::new();
    let sheet_id = workbook.add_sheet("Sheet1").unwrap();
    let sheet = workbook.sheet_mut(sheet_id).unwrap();
 
    sheet.set_cell(CellRef::new(0, 0), Cell::new(CellValue::String("Name".into())));
    sheet.set_cell(CellRef::new(1, 0), Cell::new(CellValue::String("Alice".into())));
    sheet.set_cell(CellRef::new(2, 0), Cell::new(CellValue::String("Bob".into())));
 
    sheet.auto_filter = Some(SheetAutoFilter {
        range: Range::from_a1("A1:A3").unwrap(),
        filter_columns: vec![FilterColumn {
            col_id: 0,
            join: FilterJoin::Any,
            criteria: vec![FilterCriterion::Equals(FilterValue::Text("Alice".into()))],
            // Keep the legacy value list populated so a full model equality check holds
            // across read/write for simple `<filters>` payloads.
            values: vec!["Alice".into()],
            raw_xml: Vec::new(),
        }],
        sort_state: Some(SortState {
            conditions: vec![SortCondition {
                range: Range::from_a1("A1:A3").unwrap(),
                descending: true,
            }],
        }),
        raw_xml: Vec::new(),
    });
 
    let dir = tempdir().unwrap();
    let path = dir.path().join("autofilter.xlsx");
    write_workbook(&workbook, &path).unwrap();
 
    let loaded = read_workbook(&path).unwrap();
    let loaded_sheet = &loaded.sheets[0];
    assert_eq!(loaded_sheet.auto_filter, workbook.sheets[0].auto_filter);
}

