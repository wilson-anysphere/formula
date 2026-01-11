use formula_engine::value::ErrorKind;
use formula_engine::{Engine, PrecedentNode, Value};
use formula_engine::eval::CellAddr;

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
    assert_number(&sheet.eval(r#"=SUMIF(A1:A4,2,B1:B4)"#), 20.0);
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
    assert_number(&sheet.eval(r#"=SUMIF(A4:A6,"<>",B4:B6)"#), 6.0);
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

    assert_number(
        &sheet.eval(r#"=SUMIFS(C1:C4,A1:A4,"A",B1:B4,">1")"#),
        20.0,
    );

    assert_eq!(
        sheet.eval(r#"=SUMIFS(C1:C4,A1:A3,"A")"#),
        Value::Error(ErrorKind::Value)
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
