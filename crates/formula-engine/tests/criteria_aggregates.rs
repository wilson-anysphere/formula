use formula_engine::eval::CellAddr;
use formula_engine::locale::ValueLocaleConfig;
use formula_engine::value::{EntityValue, ErrorKind, RecordValue};
use formula_engine::{Engine, PrecedentNode, Value};

struct TestSheet {
    engine: Engine,
    sheet: &'static str,
    scratch_cell: &'static str,
}

impl TestSheet {
    fn new() -> Self {
        Self {
            engine: Engine::new(),
            sheet: "Sheet1",
            scratch_cell: "Z1",
        }
    }

    fn set(&mut self, addr: &str, value: impl Into<Value>) {
        self.engine
            .set_cell_value(self.sheet, addr, value)
            .expect("set cell value");
    }

    fn set_formula(&mut self, addr: &str, formula: &str) {
        self.engine
            .set_cell_formula(self.sheet, addr, formula)
            .expect("set cell formula");
    }

    fn eval(&mut self, formula: &str) -> Value {
        self.set_formula(self.scratch_cell, formula);
        self.engine.recalculate();
        self.engine.get_cell_value(self.sheet, self.scratch_cell)
    }
}

fn assert_number(value: &Value, expected: f64) {
    match value {
        Value::Number(n) => {
            assert!((*n - expected).abs() < 1e-9, "expected {expected}, got {n}");
        }
        other => panic!("expected number {expected}, got {other:?}"),
    }
}

#[test]
fn sumif_basic_and_optional_sum_range() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", 1);
    sheet.set("A2", 2);
    sheet.set("A3", 3);
    sheet.set("A4", 4);

    sheet.set("B1", 10);
    sheet.set("B2", 20);
    sheet.set("B3", 30);
    sheet.set("B4", 40);

    assert_number(&sheet.eval(r#"=SUMIF(A1:A4,">2",B1:B4)"#), 70.0);
    assert_number(&sheet.eval(r#"=SUMIF(A1:A4,">2")"#), 7.0);
    assert_number(&sheet.eval(r#"=SUMIF(A1:A4,">2",)"#), 7.0);
    assert_number(&sheet.eval(r#"=SUMIF(A1:A4,2,B1:B4)"#), 20.0);
}

#[test]
fn sumif_numeric_criteria_does_not_treat_text_as_zero() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", "x");
    sheet.set("A2", 0);
    // A3 left unset (blank) -> treated as 0 for numeric SUMIF criteria.

    sheet.set("B1", 5);
    sheet.set("B2", 10);
    sheet.set("B3", 20);

    sheet.set_formula(sheet.scratch_cell, "=SUMIF(A1:A3,0,B1:B3)");
    assert!(
        sheet.engine.bytecode_program_count() > 0,
        "expected SUMIF formula to compile to bytecode for this test"
    );
    sheet.engine.recalculate();
    assert_number(
        &sheet.engine.get_cell_value(sheet.sheet, sheet.scratch_cell),
        30.0,
    );
}

#[test]
fn sumifs_numeric_criteria_does_not_treat_text_as_zero() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", "x");
    sheet.set("A2", 0);
    // A3 left unset (blank) -> treated as 0 for numeric SUMIFS criteria.

    // Second criteria always matches so the result depends only on the numeric criteria.
    sheet.set("B1", 1);
    sheet.set("B2", 1);
    sheet.set("B3", 1);

    sheet.set("C1", 5);
    sheet.set("C2", 10);
    sheet.set("C3", 20);

    sheet.set_formula(sheet.scratch_cell, r#"=SUMIFS(C1:C3,A1:A3,0,B1:B3,">0")"#);
    assert!(
        sheet.engine.bytecode_program_count() > 0,
        "expected SUMIFS formula to compile to bytecode for this test"
    );
    sheet.engine.recalculate();
    assert_number(
        &sheet.engine.get_cell_value(sheet.sheet, sheet.scratch_cell),
        30.0,
    );
}

#[test]
fn averageif_numeric_criteria_does_not_treat_text_as_zero() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", "x");
    sheet.set("A2", 0);
    // A3 left unset (blank) -> treated as 0 for numeric AVERAGEIF criteria.

    sheet.set("B1", 5);
    sheet.set("B2", 10);
    sheet.set("B3", 20);

    sheet.set_formula(sheet.scratch_cell, "=AVERAGEIF(A1:A3,0,B1:B3)");
    assert!(
        sheet.engine.bytecode_program_count() > 0,
        "expected AVERAGEIF formula to compile to bytecode for this test"
    );
    sheet.engine.recalculate();
    assert_number(
        &sheet.engine.get_cell_value(sheet.sheet, sheet.scratch_cell),
        15.0,
    );
}

#[test]
fn averageifs_numeric_criteria_does_not_treat_text_as_zero() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", "x");
    sheet.set("A2", 0);
    // A3 left unset (blank) -> treated as 0 for numeric AVERAGEIFS criteria.

    // Second criteria always matches so the result depends only on the numeric criteria.
    sheet.set("B1", 1);
    sheet.set("B2", 1);
    sheet.set("B3", 1);

    sheet.set("C1", 5);
    sheet.set("C2", 10);
    sheet.set("C3", 20);

    sheet.set_formula(
        sheet.scratch_cell,
        r#"=AVERAGEIFS(C1:C3,A1:A3,0,B1:B3,">0")"#,
    );
    assert!(
        sheet.engine.bytecode_program_count() > 0,
        "expected AVERAGEIFS formula to compile to bytecode for this test"
    );
    sheet.engine.recalculate();
    assert_number(
        &sheet.engine.get_cell_value(sheet.sheet, sheet.scratch_cell),
        15.0,
    );
}

#[test]
fn averageif_treats_blank_average_range_as_omitted() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", 1);
    sheet.set("A2", 2);
    sheet.set("A3", 3);
    sheet.set("A4", 4);

    assert_number(&sheet.eval(r#"=AVERAGEIF(A1:A4,">2")"#), 3.5);
    assert_number(&sheet.eval(r#"=AVERAGEIF(A1:A4,">2",)"#), 3.5);
}

#[test]
fn sumif_supports_wildcards_and_blank_criteria() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", "apple");
    sheet.set("A2", "banana");
    sheet.set("A3", "apricot");

    sheet.set("B1", 1);
    sheet.set("B2", 2);
    sheet.set("B3", 3);

    assert_number(&sheet.eval(r#"=SUMIF(A1:A3,"ap*",B1:B3)"#), 4.0);

    // A4 is blank (unset); A5 is empty string.
    sheet.set("A5", "");
    sheet.set("B4", 4);
    sheet.set("B5", 5);
    sheet.set("A6", "x");
    sheet.set("B6", 6);

    assert_number(&sheet.eval(r#"=SUMIF(A4:A6,"",B4:B6)"#), 9.0);
    // Missing criteria argument is treated as blank criteria.
    assert_number(&sheet.eval(r#"=SUMIF(A4:A6,,B4:B6)"#), 9.0);
    assert_number(&sheet.eval(r#"=SUMIF(A4:A6,"<>",B4:B6)"#), 6.0);
}

#[test]
fn sumif_wildcards_match_entity_and_record_display_strings() {
    let mut sheet = TestSheet::new();
    // Force the AST evaluator so the criteria matcher sees Entity/Record values directly.
    sheet.engine.set_bytecode_enabled(false);

    sheet.set("A1", Value::Entity(EntityValue::new("Apple")));
    sheet.set("A2", Value::Record(RecordValue::new("Pineapple")));
    sheet.set("A3", "Banana");

    sheet.set("B1", 10);
    sheet.set("B2", 30);
    sheet.set("B3", 20);

    assert_number(&sheet.eval(r#"=SUMIF(A1:A3,"*pp*",B1:B3)"#), 40.0);
}

#[test]
fn sumifs_supports_blank_criteria_when_omitted() {
    let mut sheet = TestSheet::new();

    // A1 is blank (unset); A2 is empty string; A3 is non-blank.
    sheet.set("A2", "");
    sheet.set("A3", "x");

    sheet.set("B1", 4);
    sheet.set("B2", 5);
    sheet.set("B3", 6);

    assert_number(&sheet.eval(r#"=SUMIFS(B1:B3,A1:A3,)"#), 9.0);
}

#[test]
fn criteria_aggregates_reject_scalar_range_args() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", 1);

    assert_eq!(
        sheet.eval(r#"=SUMIF(1,">0")"#),
        Value::Error(ErrorKind::Value)
    );
    assert_eq!(
        sheet.eval(r#"=AVERAGEIF(1,">0")"#),
        Value::Error(ErrorKind::Value)
    );
    assert_eq!(
        sheet.eval(r#"=SUMIFS(1,A1:A1,">0")"#),
        Value::Error(ErrorKind::Value)
    );
    assert_eq!(
        sheet.eval(r#"=AVERAGEIFS(1,A1:A1,">0")"#),
        Value::Error(ErrorKind::Value)
    );
    assert_eq!(
        sheet.eval(r#"=MAXIFS(1,A1:A1,">0")"#),
        Value::Error(ErrorKind::Value)
    );
    assert_eq!(
        sheet.eval(r#"=MINIFS(1,A1:A1,">0")"#),
        Value::Error(ErrorKind::Value)
    );
}

#[test]
fn sumif_supports_boolean_and_error_criteria() {
    let mut sheet = TestSheet::new();

    sheet.set("A1", true);
    sheet.set("A2", false);
    sheet.set("A3", true);

    sheet.set("B1", 1);
    sheet.set("B2", 2);
    sheet.set("B3", 3);

    assert_number(&sheet.eval(r#"=SUMIF(A1:A3,TRUE,B1:B3)"#), 4.0);

    // Criteria strings can match errors without the criteria argument itself being an error.
    sheet.set("C1", Value::Error(ErrorKind::Div0));
    sheet.set("C2", 0);
    sheet.set("C3", Value::Error(ErrorKind::Div0));
    sheet.set("D1", 10);
    sheet.set("D2", 20);
    sheet.set("D3", 30);

    assert_number(&sheet.eval("=SUMIF(C1:C3,\"#DIV/0!\",D1:D3)"), 40.0);
    // If the criteria argument evaluates to an error value, it propagates.
    assert_eq!(
        sheet.eval(r#"=SUMIF(C1:C3,C1,D1:D3)"#),
        Value::Error(ErrorKind::Div0)
    );
}

#[test]
fn sumif_propagates_sum_range_errors_only_when_included() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", 1);
    sheet.set("A2", 2);

    sheet.set("B1", Value::Error(ErrorKind::Div0));
    sheet.set("B2", 5);

    // Error is not included (criteria doesn't match A1), so it is ignored.
    assert_number(&sheet.eval(r#"=SUMIF(A1:A2,2,B1:B2)"#), 5.0);

    // Error is included (criteria matches both), so it propagates.
    assert_eq!(
        sheet.eval(r#"=SUMIF(A1:A2,">0",B1:B2)"#),
        Value::Error(ErrorKind::Div0)
    );
}

#[test]
fn sumif_error_precedence_is_row_major_for_sparse_iteration() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", 1);
    sheet.set("A2", 1);

    // Both errors are included by the criteria. Excel returns the first error in range order.
    sheet.set("B1", Value::Error(ErrorKind::Ref));
    sheet.set("B2", Value::Error(ErrorKind::Div0));

    assert_eq!(
        sheet.eval(r#"=SUMIF(A1:A2,">0",B1:B2)"#),
        Value::Error(ErrorKind::Ref)
    );
}

#[test]
fn sumifs_multiple_criteria_and_shape_mismatch() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", "A");
    sheet.set("A2", "A");
    sheet.set("A3", "B");
    sheet.set("A4", "B");

    sheet.set("B1", 1);
    sheet.set("B2", 2);
    sheet.set("B3", 3);
    sheet.set("B4", 4);

    sheet.set("C1", 10);
    sheet.set("C2", 20);
    sheet.set("C3", 30);
    sheet.set("C4", 40);

    assert_number(&sheet.eval(r#"=SUMIFS(C1:C4,A1:A4,"A",B1:B4,">1")"#), 20.0);

    assert_eq!(
        sheet.eval(r#"=SUMIFS(C1:C4,A1:A3,"A")"#),
        Value::Error(ErrorKind::Value)
    );
}

#[test]
fn sumifs_propagates_sum_range_errors_only_when_included() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", 1);
    sheet.set("A2", 2);

    sheet.set("B1", Value::Error(ErrorKind::Div0));
    sheet.set("B2", 5);

    // Error is not included (criteria doesn't match A1), so it is ignored.
    assert_number(&sheet.eval(r#"=SUMIFS(B1:B2,A1:A2,2)"#), 5.0);

    // Error is included (criteria matches both), so it propagates.
    assert_eq!(
        sheet.eval(r#"=SUMIFS(B1:B2,A1:A2,">0")"#),
        Value::Error(ErrorKind::Div0)
    );
}

#[test]
fn averageif_and_averageifs_return_div0_when_no_numeric_cells_included() {
    let mut sheet = TestSheet::new();

    sheet.set("A1", 1);
    sheet.set("A2", 2);
    sheet.set("A3", 3);

    sheet.set("B1", "x");
    sheet.set("B2", "y");
    sheet.set("B3", "z");

    assert_eq!(
        sheet.eval(r#"=AVERAGEIF(A1:A3,">0",B1:B3)"#),
        Value::Error(ErrorKind::Div0)
    );

    sheet.set("C1", 10);
    sheet.set("C2", 20);
    sheet.set("C3", 30);
    sheet.set("C4", 40);

    sheet.set("D1", "A");
    sheet.set("D2", "B");
    sheet.set("D3", "C");
    sheet.set("D4", "D");

    assert_eq!(
        sheet.eval(r#"=AVERAGEIFS(C1:C4,D1:D4,"Z")"#),
        Value::Error(ErrorKind::Div0)
    );
}

#[test]
fn sumif_parses_date_criteria_strings() {
    let mut sheet = TestSheet::new();

    sheet.set_formula("A1", "=DATE(2019,12,31)");
    sheet.set_formula("A2", "=DATE(2020,1,1)");
    sheet.set_formula("A3", "=DATE(2020,1,2)");

    sheet.set("B1", 1);
    sheet.set("B2", 2);
    sheet.set("B3", 3);

    assert_number(&sheet.eval(r#"=SUMIF(A1:A3,">1/1/2020",B1:B3)"#), 3.0);
}

#[test]
fn sumif_indirect_records_dynamic_dependencies() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1).unwrap();
    engine.set_cell_value("Sheet1", "A2", 2).unwrap();
    engine.set_cell_value("Sheet1", "A3", 3).unwrap();
    engine
        .set_cell_formula("Sheet1", "Z1", r#"=SUMIF(INDIRECT("A1:A3"),">0")"#)
        .unwrap();
    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "Z1"), Value::Number(6.0));

    let precedents = engine.precedents("Sheet1", "Z1").unwrap();
    assert!(
        precedents.iter().any(|node| matches!(
            node,
            PrecedentNode::Range {
                sheet: 0,
                start: CellAddr { row: 0, col: 0 },
                end: CellAddr { row: 2, col: 0 },
            }
        )),
        "expected Z1 precedents to include Sheet1!A1:A3, got: {precedents:?}"
    );

    let dependents = engine.dependents("Sheet1", "A1").unwrap();
    assert!(
        dependents.iter().any(|node| matches!(
            node,
            PrecedentNode::Cell {
                sheet: 0,
                addr: CellAddr { row: 0, col: 25 },
            }
        )),
        "expected A1 dependents to include Sheet1!Z1, got: {dependents:?}"
    );
}

#[test]
fn countifs_indirect_records_dynamic_dependencies() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1).unwrap();
    engine.set_cell_value("Sheet1", "A2", 2).unwrap();
    engine.set_cell_value("Sheet1", "A3", 3).unwrap();
    engine
        .set_cell_formula("Sheet1", "Z1", r#"=COUNTIFS(INDIRECT("A1:A3"),">0")"#)
        .unwrap();
    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "Z1"), Value::Number(3.0));

    let precedents = engine.precedents("Sheet1", "Z1").unwrap();
    assert!(
        precedents.iter().any(|node| matches!(
            node,
            PrecedentNode::Range {
                sheet: 0,
                start: CellAddr { row: 0, col: 0 },
                end: CellAddr { row: 2, col: 0 },
            }
        )),
        "expected Z1 precedents to include Sheet1!A1:A3, got: {precedents:?}"
    );

    let dependents = engine.dependents("Sheet1", "A1").unwrap();
    assert!(
        dependents.iter().any(|node| matches!(
            node,
            PrecedentNode::Cell {
                sheet: 0,
                addr: CellAddr { row: 0, col: 25 },
            }
        )),
        "expected A1 dependents to include Sheet1!Z1, got: {dependents:?}"
    );
}

#[test]
fn maxifs_and_minifs_indirect_record_dynamic_dependencies() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1).unwrap();
    engine.set_cell_value("Sheet1", "A2", 0).unwrap();
    engine.set_cell_value("Sheet1", "A3", 1).unwrap();
    engine.set_cell_value("Sheet1", "B1", 10).unwrap();
    engine.set_cell_value("Sheet1", "B2", 20).unwrap();
    engine.set_cell_value("Sheet1", "B3", 30).unwrap();

    engine
        .set_cell_formula(
            "Sheet1",
            "Z1",
            r#"=MAXIFS(INDIRECT("B1:B3"),INDIRECT("A1:A3"),">0")"#,
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "Z2",
            r#"=MINIFS(INDIRECT("B1:B3"),INDIRECT("A1:A3"),">0")"#,
        )
        .unwrap();
    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "Z1"), Value::Number(30.0));
    assert_eq!(engine.get_cell_value("Sheet1", "Z2"), Value::Number(10.0));

    let precedents = engine.precedents("Sheet1", "Z1").unwrap();
    assert!(
        precedents.iter().any(|node| matches!(
            node,
            PrecedentNode::Range {
                sheet: 0,
                start: CellAddr { row: 0, col: 0 },
                end: CellAddr { row: 2, col: 0 },
            }
        )),
        "expected Z1 precedents to include Sheet1!A1:A3, got: {precedents:?}"
    );
    assert!(
        precedents.iter().any(|node| matches!(
            node,
            PrecedentNode::Range {
                sheet: 0,
                start: CellAddr { row: 0, col: 1 },
                end: CellAddr { row: 2, col: 1 },
            }
        )),
        "expected Z1 precedents to include Sheet1!B1:B3, got: {precedents:?}"
    );

    let dependents = engine.dependents("Sheet1", "A1").unwrap();
    assert!(
        dependents.iter().any(|node| matches!(
            node,
            PrecedentNode::Cell {
                sheet: 0,
                addr: CellAddr { row: 0, col: 25 },
            }
        )),
        "expected A1 dependents to include Sheet1!Z1, got: {dependents:?}"
    );
    let dependents = engine.dependents("Sheet1", "B1").unwrap();
    assert!(
        dependents.iter().any(|node| matches!(
            node,
            PrecedentNode::Cell {
                sheet: 0,
                addr: CellAddr { row: 0, col: 25 },
            }
        )),
        "expected B1 dependents to include Sheet1!Z1, got: {dependents:?}"
    );
}

#[test]
fn maxifs_and_minifs_require_matching_shapes() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", 1);
    sheet.set("A2", 2);
    sheet.set("A3", 3);
    sheet.set("A4", 4);
    sheet.set("B1", 10);
    sheet.set("B2", 20);
    sheet.set("B3", 30);
    sheet.set("B4", 40);

    // Same number of cells (4) but different shapes (4x1 vs 2x2).
    assert_eq!(
        sheet.eval(r#"=MAXIFS(B1:B4,A1:B2,">0")"#),
        Value::Error(ErrorKind::Value)
    );
    assert_eq!(
        sheet.eval(r#"=MINIFS(B1:B4,A1:B2,">0")"#),
        Value::Error(ErrorKind::Value)
    );
}

#[test]
fn maxifs_and_minifs_propagate_errors_only_when_included() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", 1);
    sheet.set("A2", 2);

    sheet.set("B1", Value::Error(ErrorKind::Div0));
    sheet.set("B2", 5);

    // Error is excluded by criteria, so it is ignored.
    assert_number(&sheet.eval(r#"=MAXIFS(B1:B2,A1:A2,2)"#), 5.0);
    assert_number(&sheet.eval(r#"=MINIFS(B1:B2,A1:A2,2)"#), 5.0);

    // Error is included, so it propagates.
    assert_eq!(
        sheet.eval(r#"=MAXIFS(B1:B2,A1:A2,">0")"#),
        Value::Error(ErrorKind::Div0)
    );
    assert_eq!(
        sheet.eval(r#"=MINIFS(B1:B2,A1:A2,">0")"#),
        Value::Error(ErrorKind::Div0)
    );
}

#[test]
fn maxifs_and_minifs_parse_date_criteria_strings() {
    let mut sheet = TestSheet::new();

    sheet.set_formula("A1", "=DATE(2019,12,31)");
    sheet.set_formula("A2", "=DATE(2020,1,1)");
    sheet.set_formula("A3", "=DATE(2020,1,2)");

    sheet.set("B1", 1);
    sheet.set("B2", 2);
    sheet.set("B3", 3);

    assert_number(&sheet.eval(r#"=MAXIFS(B1:B3,A1:A3,">12/31/2019")"#), 3.0);
    assert_number(&sheet.eval(r#"=MINIFS(B1:B3,A1:A3,">12/31/2019")"#), 2.0);
}

#[test]
fn maxifs_and_minifs_use_workbook_locale_for_numeric_criteria() {
    let mut engine = Engine::new();
    engine.set_value_locale(ValueLocaleConfig::de_de());

    engine.set_cell_value("Sheet1", "A1", 1).unwrap();
    engine.set_cell_value("Sheet1", "A2", 2).unwrap();
    engine.set_cell_value("Sheet1", "A3", 3).unwrap();

    engine.set_cell_value("Sheet1", "B1", 10).unwrap();
    engine.set_cell_value("Sheet1", "B2", 20).unwrap();
    engine.set_cell_value("Sheet1", "B3", 30).unwrap();

    engine
        .set_cell_formula("Sheet1", "Z1", r#"=MAXIFS(B1:B3,A1:A3,">1,5")"#)
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "Z2", r#"=MINIFS(B1:B3,A1:A3,">1,5")"#)
        .unwrap();

    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "Z1"), Value::Number(30.0));
    assert_eq!(engine.get_cell_value("Sheet1", "Z2"), Value::Number(20.0));
}

#[test]
fn maxifs_and_minifs_support_wildcards_blank_and_error_criteria() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", "apple");
    sheet.set("A2", "banana");
    sheet.set("A3", "apricot");

    sheet.set("B1", 1);
    sheet.set("B2", 2);
    sheet.set("B3", 3);

    assert_number(&sheet.eval(r#"=MAXIFS(B1:B3,A1:A3,"ap*")"#), 3.0);
    assert_number(&sheet.eval(r#"=MINIFS(B1:B3,A1:A3,"ap*")"#), 1.0);

    // A4 is implicit blank; A5 is explicit empty string.
    sheet.set("A5", "");
    sheet.set("A6", "x");
    sheet.set("B4", 4);
    sheet.set("B5", 5);
    sheet.set("B6", 6);

    assert_number(&sheet.eval(r#"=MAXIFS(B4:B6,A4:A6,"")"#), 5.0);
    assert_number(&sheet.eval(r#"=MINIFS(B4:B6,A4:A6,"")"#), 4.0);
    assert_number(&sheet.eval(r#"=MAXIFS(B4:B6,A4:A6,"<>")"#), 6.0);
    assert_number(&sheet.eval(r#"=MINIFS(B4:B6,A4:A6,"<>")"#), 6.0);

    // Criteria strings can match errors without the criteria argument itself being an error.
    sheet.set("C1", Value::Error(ErrorKind::Div0));
    sheet.set("C2", 0);
    sheet.set("C3", Value::Error(ErrorKind::Div0));
    sheet.set("D1", 10);
    sheet.set("D2", 20);
    sheet.set("D3", 30);

    assert_number(&sheet.eval(r##"=MAXIFS(D1:D3,C1:C3,"#DIV/0!")"##), 30.0);
    assert_number(&sheet.eval(r##"=MINIFS(D1:D3,C1:C3,"#DIV/0!")"##), 10.0);
}

#[test]
fn averageifs_propagates_average_range_errors_only_when_included() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", 1);
    sheet.set("A2", 2);

    sheet.set("B1", Value::Error(ErrorKind::Div0));
    sheet.set("B2", 5);

    // Error is excluded by criteria, so it is ignored.
    assert_number(&sheet.eval(r#"=AVERAGEIFS(B1:B2,A1:A2,2)"#), 5.0);

    // Error is included, so it propagates.
    assert_eq!(
        sheet.eval(r#"=AVERAGEIFS(B1:B2,A1:A2,">0")"#),
        Value::Error(ErrorKind::Div0)
    );
}
