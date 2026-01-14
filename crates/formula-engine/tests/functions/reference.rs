use formula_engine::{ErrorKind, Value};

use super::harness::{assert_number, TestSheet};
use formula_engine::locale::ValueLocaleConfig;

#[test]
fn row_and_column_without_args_use_current_cell() {
    let mut sheet = TestSheet::new();

    sheet.set_formula("B5", "=ROW()");
    sheet.set_formula("C5", "=COLUMN()");
    sheet.recalculate();

    assert_number(&sheet.get("B5"), 5.0);
    assert_number(&sheet.get("C5"), 3.0);
}

#[test]
fn row_and_column_with_reference() {
    let mut sheet = TestSheet::new();

    assert_number(&sheet.eval("=ROW(D10)"), 10.0);
    assert_number(&sheet.eval("=COLUMN(D10)"), 4.0);
}

#[test]
fn row_and_column_spill_for_rectangular_ranges() {
    let mut sheet = TestSheet::new();

    // Use a range that does not overlap the spill output to avoid circular dependencies.
    sheet.set_formula("A1", "=ROW(D4:F5)");
    sheet.set_formula("E1", "=COLUMN(D4:F5)");
    sheet.recalculate();

    // ROW(D4:F5) -> {4,4,4;5,5,5}
    assert_number(&sheet.get("A1"), 4.0);
    assert_number(&sheet.get("B1"), 4.0);
    assert_number(&sheet.get("C1"), 4.0);
    assert_number(&sheet.get("A2"), 5.0);
    assert_number(&sheet.get("B2"), 5.0);
    assert_number(&sheet.get("C2"), 5.0);

    // COLUMN(D4:F5) -> {4,5,6;4,5,6}
    assert_number(&sheet.get("E1"), 4.0);
    assert_number(&sheet.get("F1"), 5.0);
    assert_number(&sheet.get("G1"), 6.0);
    assert_number(&sheet.get("E2"), 4.0);
    assert_number(&sheet.get("F2"), 5.0);
    assert_number(&sheet.get("G2"), 6.0);
}

#[test]
fn row_and_column_handle_row_and_column_references() {
    let mut sheet = TestSheet::new();

    sheet.set_formula("J1", "=ROW(5:7)");
    sheet.set_formula("A10", "=COLUMN(D:F)");
    sheet.recalculate();

    // ROW(5:7) -> {5;6;7}
    assert_number(&sheet.get("J1"), 5.0);
    assert_number(&sheet.get("J2"), 6.0);
    assert_number(&sheet.get("J3"), 7.0);

    // COLUMN(D:F) -> {4,5,6}
    assert_number(&sheet.get("A10"), 4.0);
    assert_number(&sheet.get("B10"), 5.0);
    assert_number(&sheet.get("C10"), 6.0);
}

#[test]
fn rows_and_columns_return_dimensions() {
    let mut sheet = TestSheet::new();
    assert_number(&sheet.eval("=ROWS(B2:D3)"), 2.0);
    assert_number(&sheet.eval("=COLUMNS(B2:D3)"), 3.0);
    assert_number(&sheet.eval("=ROWS(5:7)"), 3.0);
    assert_number(&sheet.eval("=COLUMNS(A:C)"), 3.0);
}

#[test]
fn areas_counts_reference_unions() {
    let mut sheet = TestSheet::new();

    assert_number(&sheet.eval("=AREAS(A1:B2)"), 1.0);
    assert_number(&sheet.eval("=AREAS((A1:A2,C1:C2))"), 2.0);

    assert_eq!(sheet.eval("=AREAS(1)"), Value::Error(ErrorKind::Value));
    assert_eq!(sheet.eval("=AREAS(#REF!)"), Value::Error(ErrorKind::Ref));
}

#[test]
fn address_formats_a1_and_r1c1_styles() {
    let mut sheet = TestSheet::new();

    assert_eq!(sheet.eval("=ADDRESS(1,1)"), Value::Text("$A$1".to_string()));
    assert_eq!(sheet.eval("=ADDRESS(1,1,4)"), Value::Text("A1".to_string()));
    assert_eq!(
        sheet.eval("=ADDRESS(1,1,2)"),
        Value::Text("A$1".to_string())
    );
    assert_eq!(
        sheet.eval("=ADDRESS(1,1,3)"),
        Value::Text("$A1".to_string())
    );

    assert_eq!(
        sheet.eval("=ADDRESS(1,1,1,FALSE)"),
        Value::Text("R1C1".to_string())
    );
    assert_eq!(
        sheet.eval("=ADDRESS(1,1,4,FALSE)"),
        Value::Text("R[1]C[1]".to_string())
    );

    assert_eq!(
        sheet.eval("=ADDRESS(1,1,1,TRUE,\"Sheet2\")"),
        Value::Text("Sheet2!$A$1".to_string())
    );
    assert_eq!(
        sheet.eval("=ADDRESS(1,1,1,TRUE,\"My Sheet\")"),
        Value::Text("'My Sheet'!$A$1".to_string())
    );
    assert_eq!(
        sheet.eval("=ADDRESS(1,1,1,TRUE,\"Bob's Sheet\")"),
        Value::Text("'Bob''s Sheet'!$A$1".to_string())
    );
    assert_eq!(
        sheet.eval("=ADDRESS(1,1,1,TRUE,\"2024\")"),
        Value::Text("'2024'!$A$1".to_string())
    );
    assert_eq!(
        sheet.eval("=ADDRESS(1,1,1,TRUE,\"TRUE\")"),
        Value::Text("'TRUE'!$A$1".to_string())
    );
    assert_eq!(
        sheet.eval("=ADDRESS(1,1,1,TRUE,\"FALSE\")"),
        Value::Text("'FALSE'!$A$1".to_string())
    );
    assert_eq!(
        sheet.eval("=ADDRESS(1,1,1,TRUE,\"A1B\")"),
        Value::Text("A1B!$A$1".to_string())
    );
    assert_eq!(
        sheet.eval("=ADDRESS(1,1,1,TRUE,\"A1.B\")"),
        Value::Text("'A1.B'!$A$1".to_string())
    );

    // R1C1 sheet prefixes are quoted when they would otherwise be tokenized as R1C1 refs.
    assert_eq!(
        sheet.eval("=ADDRESS(1,1,1,FALSE,\"RC\")"),
        Value::Text("'RC'!R1C1".to_string())
    );

    // `sheet_text` is coerced to text using the workbook value locale.
    sheet.set_value_locale(ValueLocaleConfig::de_de());
    assert_eq!(
        sheet.eval("=ADDRESS(1,1,1,TRUE,1.5)"),
        Value::Text("'1,5'!$A$1".to_string())
    );

    assert_eq!(sheet.eval("=ADDRESS(0,1)"), Value::Error(ErrorKind::Value));
    assert_eq!(sheet.eval("=ADDRESS(1,0)"), Value::Error(ErrorKind::Value));
    assert_eq!(
        sheet.eval("=ADDRESS(1,1,0)"),
        Value::Error(ErrorKind::Value)
    );
}
