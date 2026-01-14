use formula_engine::eval::parse_a1;
use formula_engine::{Engine, ErrorKind, Value};

use super::harness::{assert_number, TestSheet};

#[test]
fn mdeterm_handles_known_2x2_and_3x3_matrices() {
    let mut sheet = TestSheet::new();
    assert_number(&sheet.eval("=MDETERM({1,2;3,4})"), -2.0);
    assert_number(&sheet.eval("=MDETERM({1,2,3;0,1,4;5,6,0})"), 1.0);
}

#[test]
fn minverse_spills_and_matches_known_matrices() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=MINVERSE({1,2;3,5})")
        .unwrap();
    engine.recalculate_single_threaded();

    let (start, end) = engine.spill_range("Sheet1", "A1").expect("spill range");
    assert_eq!(start, parse_a1("A1").unwrap());
    assert_eq!(end, parse_a1("B2").unwrap());

    assert_number(&engine.get_cell_value("Sheet1", "A1"), -5.0);
    assert_number(&engine.get_cell_value("Sheet1", "B1"), 2.0);
    assert_number(&engine.get_cell_value("Sheet1", "A2"), 3.0);
    assert_number(&engine.get_cell_value("Sheet1", "B2"), -1.0);

    engine
        .set_cell_formula("Sheet1", "D1", "=MINVERSE({1,2,3;0,1,4;5,6,0})")
        .unwrap();
    engine.recalculate_single_threaded();

    let (start, end) = engine.spill_range("Sheet1", "D1").expect("spill range");
    assert_eq!(start, parse_a1("D1").unwrap());
    assert_eq!(end, parse_a1("F3").unwrap());

    // Inverse of {{1,2,3},{0,1,4},{5,6,0}} has determinant 1 and integer inverse:
    // {{-24,18,5},{20,-15,-4},{-5,4,1}}
    let expected = [(-24.0, 18.0, 5.0), (20.0, -15.0, -4.0), (-5.0, 4.0, 1.0)];
    for (r, (a, b, c)) in expected.into_iter().enumerate() {
        let row = r + 1;
        assert_number(&engine.get_cell_value("Sheet1", &format!("D{row}")), a);
        assert_number(&engine.get_cell_value("Sheet1", &format!("E{row}")), b);
        assert_number(&engine.get_cell_value("Sheet1", &format!("F{row}")), c);
    }
}

#[test]
fn mmult_returns_value_on_shape_mismatch() {
    let mut sheet = TestSheet::new();
    assert_eq!(
        sheet.eval("=MMULT({1,2;3,4},{1,2;3,4;5,6})"),
        Value::Error(ErrorKind::Value)
    );
}

#[test]
fn minverse_returns_num_on_singular_matrix() {
    let mut sheet = TestSheet::new();
    assert_eq!(
        sheet.eval("=MINVERSE({1,2;2,4})"),
        Value::Error(ErrorKind::Num)
    );
}

#[test]
fn munit_spills_identity_matrix() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "B2", "=MUNIT(3)")
        .unwrap();
    engine.recalculate_single_threaded();

    let (start, end) = engine.spill_range("Sheet1", "B2").expect("spill range");
    assert_eq!(start, parse_a1("B2").unwrap());
    assert_eq!(end, parse_a1("D4").unwrap());

    assert_eq!(engine.get_cell_value("Sheet1", "B2"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "C2"), Value::Number(0.0));
    assert_eq!(engine.get_cell_value("Sheet1", "D2"), Value::Number(0.0));

    assert_eq!(engine.get_cell_value("Sheet1", "B3"), Value::Number(0.0));
    assert_eq!(engine.get_cell_value("Sheet1", "C3"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "D3"), Value::Number(0.0));

    assert_eq!(engine.get_cell_value("Sheet1", "B4"), Value::Number(0.0));
    assert_eq!(engine.get_cell_value("Sheet1", "C4"), Value::Number(0.0));
    assert_eq!(engine.get_cell_value("Sheet1", "D4"), Value::Number(1.0));
}

#[test]
fn matrix_functions_coerce_matrix_elements_like_excel() {
    // Document coercion decisions for matrix elements:
    // - blanks -> 0
    // - booleans -> 1/0
    // - numeric text -> number
    // - non-numeric text -> #VALUE!
    // - errors propagate
    let mut sheet = TestSheet::new();

    // Range input with an implicit blank cell (B1 is missing/blank) should treat the blank as 0.
    sheet.set("A1", 1.0);
    sheet.set("A2", 3.0);
    sheet.set("B2", 4.0);
    sheet.set_formula("C1", "=MDETERM(A1:B2)");
    sheet.recalc();
    assert_number(&sheet.get("C1"), 4.0); // 1*4 - 3*0

    // Changing the blank to a number should update the determinant (dependency tracking).
    sheet.set("B1", 2.0);
    sheet.recalc();
    assert_number(&sheet.get("C1"), -2.0); // 1*4 - 3*2

    // Booleans inside the matrix are coerced like numbers.
    sheet.set("A1", true);
    sheet.recalc();
    assert_number(&sheet.get("C1"), -2.0); // TRUE=1

    // Numeric text is parsed.
    sheet.set("A1", 1.0);
    sheet.set("B1", Value::Text("2".to_string()));
    sheet.recalc();
    assert_number(&sheet.get("C1"), -2.0);

    // Non-numeric text yields #VALUE!.
    sheet.set("B1", Value::Text("x".to_string()));
    sheet.recalc();
    assert_eq!(sheet.get("C1"), Value::Error(ErrorKind::Value));

    // Errors inside the matrix propagate.
    sheet.set("B1", Value::Error(ErrorKind::Div0));
    sheet.recalc();
    assert_eq!(sheet.get("C1"), Value::Error(ErrorKind::Div0));
}
