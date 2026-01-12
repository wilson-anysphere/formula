use formula_vba_runtime::{parse_program, InMemoryWorkbook, Spreadsheet, VbaRuntime, VbaValue};
use pretty_assertions::assert_eq;

#[test]
fn range_ref_entire_column_respects_sheet_dimensions() {
    let mut wb = InMemoryWorkbook::new();
    let sheet = wb.active_sheet();

    wb.set_sheet_dimensions(sheet, 2_000_000, 16_384).unwrap();

    let range = wb.range_ref(sheet, "A:A").unwrap();
    assert_eq!(range.start_row, 1);
    assert_eq!(range.start_col, 1);
    assert_eq!(range.end_row, 2_000_000);
    assert_eq!(range.end_col, 1);
}

#[test]
fn runtime_rows_count_and_range_end_respect_sheet_dimensions() {
    let code = r#"
Option Explicit

Sub Test()
    Range("A1") = Rows.Count
    Range("A2") = Columns.Count

    Cells(1999999, 1) = 1
    Cells(2000000, 1) = 2
    Range("B1") = Cells(1999999, 1).End(xlDown).Row
End Sub
"#;

    let program = parse_program(code).unwrap();
    let runtime = VbaRuntime::new(program);
    let mut wb = InMemoryWorkbook::new();
    let sheet = wb.active_sheet();
    wb.set_sheet_dimensions(sheet, 2_000_000, 16_384).unwrap();

    runtime.execute(&mut wb, "Test", &[]).unwrap();

    assert_eq!(
        wb.get_value_a1("Sheet1", "A1").unwrap(),
        VbaValue::Double(2_000_000.0)
    );
    assert_eq!(
        wb.get_value_a1("Sheet1", "A2").unwrap(),
        VbaValue::Double(16_384.0)
    );
    assert_eq!(
        wb.get_value_a1("Sheet1", "B1").unwrap(),
        VbaValue::Double(2_000_000.0)
    );
}

