use formula_engine::{ErrorKind, Value};

use super::harness::{assert_number, TestSheet};

#[test]
fn cell_address_row_and_col() {
    let mut sheet = TestSheet::new();

    assert_eq!(
        sheet.eval("=CELL(\"address\",A1)"),
        Value::Text("$A$1".to_string())
    );
    assert_number(&sheet.eval("=CELL(\"row\",A10)"), 10.0);
    assert_number(&sheet.eval("=CELL(\"col\",C1)"), 3.0);
}

#[test]
fn cell_type_codes_match_excel() {
    let mut sheet = TestSheet::new();

    // Blank.
    sheet.set("A1", Value::Blank);
    assert_eq!(sheet.eval("=CELL(\"type\",A1)"), Value::Text("b".to_string()));

    // Number.
    sheet.set("A1", 1.0);
    assert_eq!(sheet.eval("=CELL(\"type\",A1)"), Value::Text("v".to_string()));

    // Text.
    sheet.set("A1", "x");
    assert_eq!(sheet.eval("=CELL(\"type\",A1)"), Value::Text("l".to_string()));
}

#[test]
fn cell_contents_returns_formula_text_or_value() {
    let mut sheet = TestSheet::new();

    sheet.set("A1", 5.0);
    assert_number(&sheet.eval("=CELL(\"contents\",A1)"), 5.0);

    sheet.set_formula("A1", "=1+1");
    assert_eq!(
        sheet.eval("=CELL(\"contents\",A1)"),
        Value::Text("=1+1".to_string())
    );
}

#[test]
fn info_recalc_and_unknown_keys() {
    let mut sheet = TestSheet::new();

    assert_eq!(
        sheet.eval("=INFO(\"recalc\")"),
        Value::Text("Automatic".to_string())
    );
    assert_eq!(sheet.eval("=INFO(\"no_such_key\")"), Value::Error(ErrorKind::Value));
}

#[test]
fn cell_errors_for_unknown_info_types() {
    let mut sheet = TestSheet::new();

    assert_eq!(
        sheet.eval("=CELL(\"no_such_info_type\",A1)"),
        Value::Error(ErrorKind::Value)
    );
}

