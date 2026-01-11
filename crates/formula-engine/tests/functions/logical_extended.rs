use formula_engine::{ErrorKind, Value};

use super::harness::{assert_number, TestSheet};

#[test]
fn ifs_short_circuits_in_scalar_mode_and_returns_na() {
    let mut sheet = TestSheet::new();
    assert_number(&sheet.eval("=IFS(TRUE, 1, TRUE, 1/0)"), 1.0);
    assert_eq!(
        sheet.eval("=IFS(FALSE, 1, FALSE, 2)"),
        Value::Error(ErrorKind::NA)
    );
    assert_eq!(
        sheet.eval("=IFS(TRUE, 1, FALSE)"),
        Value::Error(ErrorKind::Value)
    );
}

#[test]
fn ifs_spills_and_ignores_unselected_branch_errors_per_element() {
    let mut sheet = TestSheet::new();
    sheet.set_formula("A1", "=IFS({TRUE,FALSE}, 1, {FALSE,TRUE}, 1/0)");
    sheet.recalc();

    assert_number(&sheet.get("A1"), 1.0);
    assert_eq!(sheet.get("B1"), Value::Error(ErrorKind::Div0));
}

#[test]
fn ifs_returns_value_error_for_incompatible_branch_shapes() {
    let mut sheet = TestSheet::new();
    sheet.set_formula("A1", "=IFS({TRUE,FALSE}, {10,20}, {FALSE,TRUE}, {30;40})");
    sheet.recalc();

    assert_number(&sheet.get("A1"), 10.0);
    assert_eq!(sheet.get("B1"), Value::Error(ErrorKind::Value));
}

#[test]
fn switch_short_circuits_in_scalar_mode_and_supports_default() {
    let mut sheet = TestSheet::new();
    assert_number(&sheet.eval("=SWITCH(1, 1, 1, 2, 1/0)"), 1.0);
    assert_eq!(
        sheet.eval("=SWITCH(3, 1, 10, 2, 20)"),
        Value::Error(ErrorKind::NA)
    );
    assert_number(&sheet.eval("=SWITCH(3, 1, 10, 2, 20, 99)"), 99.0);
    assert_number(&sheet.eval("=SWITCH(\"a\", \"A\", 1, \"b\", 2)"), 1.0);
}

#[test]
fn switch_spills_and_returns_na_for_unmatched_elements() {
    let mut sheet = TestSheet::new();
    sheet.set_formula("A1", "=SWITCH({1,3}, 1, 10, 2, 20)");
    sheet.recalc();

    assert_number(&sheet.get("A1"), 10.0);
    assert_eq!(sheet.get("B1"), Value::Error(ErrorKind::NA));
}

#[test]
fn switch_spills_and_ignores_unselected_result_errors_per_element() {
    let mut sheet = TestSheet::new();
    sheet.set_formula("A1", "=SWITCH({1,2}, 1, 10, 2, 1/0)");
    sheet.recalc();

    assert_number(&sheet.get("A1"), 10.0);
    assert_eq!(sheet.get("B1"), Value::Error(ErrorKind::Div0));
}

#[test]
fn switch_returns_value_error_for_incompatible_result_shapes() {
    let mut sheet = TestSheet::new();
    sheet.set_formula("A1", "=SWITCH({1,2}, 1, {10,20}, 2, {30;40})");
    sheet.recalc();

    assert_number(&sheet.get("A1"), 10.0);
    assert_eq!(sheet.get("B1"), Value::Error(ErrorKind::Value));
}

#[test]
fn choose_spills_and_truncates_index() {
    let mut sheet = TestSheet::new();
    sheet.set_formula("A1", "=CHOOSE({1,2}, 10, 20)");
    sheet.recalc();

    assert_number(&sheet.get("A1"), 10.0);
    assert_number(&sheet.get("B1"), 20.0);

    assert_number(&sheet.eval("=CHOOSE(2.9, 10, 20, 30)"), 20.0);
    assert_eq!(
        sheet.eval("=CHOOSE(0, 10, 20)"),
        Value::Error(ErrorKind::Value)
    );
}

#[test]
fn choose_broadcasts_1x1_index_arrays_to_array_results() {
    let mut sheet = TestSheet::new();
    sheet.set_formula("A1", "=CHOOSE({1}, {10,20}, {30,40})");
    sheet.recalc();

    assert_number(&sheet.get("A1"), 10.0);
    assert_number(&sheet.get("B1"), 20.0);
}

#[test]
fn choose_selects_array_choices_elementwise() {
    let mut sheet = TestSheet::new();
    sheet.set_formula("A1", "=CHOOSE({1,2}, {10,20}, {30,40})");
    sheet.recalc();

    assert_number(&sheet.get("A1"), 10.0);
    assert_number(&sheet.get("B1"), 40.0);
}

#[test]
fn choose_ignores_unselected_choice_shape_mismatches() {
    let mut sheet = TestSheet::new();
    sheet.set_formula("A1", "=CHOOSE({1,1}, {10,20}, {30,40,50})");
    sheet.recalc();

    assert_number(&sheet.get("A1"), 10.0);
    assert_number(&sheet.get("B1"), 20.0);
}

#[test]
fn choose_returns_value_error_for_incompatible_choice_shapes() {
    let mut sheet = TestSheet::new();
    assert_eq!(
        sheet.eval("=CHOOSE({1,2}, {10,20}, {30,40,50})"),
        Value::Error(ErrorKind::Value)
    );
}

#[test]
fn choose_ignores_errors_in_unselected_branches_per_element() {
    let mut sheet = TestSheet::new();
    sheet.set_formula("A1", "=CHOOSE({1,2}, 1, 1/0)");
    sheet.recalc();

    assert_number(&sheet.get("A1"), 1.0);
    assert_eq!(sheet.get("B1"), Value::Error(ErrorKind::Div0));
}

#[test]
fn xor_handles_scalar_vs_range_semantics() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", Value::Blank);
    sheet.set("A2", Value::Text("x".to_string()));
    sheet.set("A3", true);
    sheet.set("A4", 0.0);
    sheet.set("A5", Value::Error(ErrorKind::Div0));

    assert_eq!(sheet.eval("=XOR(A1:A4)"), Value::Bool(true));
    assert_eq!(sheet.eval("=XOR(A2, TRUE)"), Value::Bool(true));
    assert_eq!(sheet.eval("=XOR(\"TRUE\", FALSE)"), Value::Bool(true));
    assert_eq!(sheet.eval("=XOR(\"x\", TRUE)"), Value::Error(ErrorKind::Value));
    assert_eq!(sheet.eval("=XOR(A5, TRUE)"), Value::Error(ErrorKind::Div0));
    assert_eq!(sheet.eval("=XOR({TRUE,FALSE,TRUE})"), Value::Bool(false));
}
