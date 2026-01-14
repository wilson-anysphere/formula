use formula_engine::eval::parse_a1;
use formula_engine::{Engine, ErrorKind, Value};

use super::harness::{assert_number, TestSheet};

#[test]
fn index_reference_enables_range_operator_dynamic_ranges() {
    let mut sheet = TestSheet::new();
    for i in 1..=10 {
        sheet.set(&format!("A{i}"), i as f64);
    }

    assert_number(&sheet.eval("=SUM(INDEX(A1:A10,2):INDEX(A1:A10,5))"), 14.0);
}

#[test]
fn index_reference_supports_row_and_column_zero_slices() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", 1.0);
    sheet.set("B1", 2.0);
    sheet.set("C1", 3.0);

    sheet.set("A2", 10.0);
    sheet.set("B2", 20.0);
    sheet.set("C2", 30.0);

    sheet.set("A3", 100.0);
    sheet.set("B3", 200.0);
    sheet.set("C3", 300.0);

    assert_number(&sheet.eval("=SUM(INDEX(A1:C3,0,2))"), 222.0);
    // Blank row_num coerces to 0.
    assert_number(&sheet.eval("=SUM(INDEX(A1:C3,,2))"), 222.0);
    assert_number(&sheet.eval("=SUM(INDEX(A1:C3,2,0))"), 60.0);
    // Blank col_num coerces to 0.
    assert_number(&sheet.eval("=SUM(INDEX(A1:C3,2,))"), 60.0);
    assert_number(&sheet.eval("=SUM(INDEX(A1:C3,0,0))"), 666.0);
    assert_number(&sheet.eval("=SUM(INDEX(A1:C3,,))"), 666.0);

    // If column_num is omitted it defaults to 1 (first column slice).
    assert_number(&sheet.eval("=SUM(INDEX(A1:C3,0))"), 111.0);
}

#[test]
fn index_reference_supports_multi_area_union_and_area_num() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", 1.0);
    sheet.set("A2", 2.0);
    sheet.set("A3", 3.0);
    sheet.set("C1", 10.0);
    sheet.set("C2", 20.0);
    sheet.set("C3", 30.0);

    assert_eq!(
        sheet.eval("=INDEX((A1:A3,C1:C3),2,1,2)"),
        Value::Number(20.0)
    );
}

#[test]
fn index_reference_errors() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", 1.0);
    sheet.set("A2", 2.0);
    sheet.set("A3", 3.0);

    // column_num defaults to 1, so row_num=0 returns the first column slice.
    assert_number(&sheet.eval("=SUM(INDEX(A1:A3,0))"), 6.0);
    assert_eq!(sheet.eval("=INDEX(A1:A3,4)"), Value::Error(ErrorKind::Ref));
}

#[test]
fn index_truncates_fractional_indices_toward_zero() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", 10.0);
    sheet.set("A2", 20.0);
    sheet.set("A3", 30.0);
    sheet.set("B1", 1.0);
    sheet.set("C1", 2.0);

    // row_num/col_num are coerced via truncation toward 0.
    assert_eq!(sheet.eval("=INDEX(A1:A3,2.9)"), Value::Number(20.0));
    assert_eq!(sheet.eval("=INDEX(A1:C1,1,2.1)"), Value::Number(1.0));

    // Negative indices become #VALUE! after truncation.
    assert_eq!(
        sheet.eval("=INDEX(A1:A3,-1.1)"),
        Value::Error(ErrorKind::Value)
    );
    assert_eq!(
        sheet.eval("=INDEX(A1:C1,1,-2.9)"),
        Value::Error(ErrorKind::Value)
    );
}

#[test]
fn index_area_num_defaults_and_bounds() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", 1.0);
    sheet.set("A2", 2.0);
    sheet.set("A3", 3.0);
    sheet.set("C1", 10.0);
    sheet.set("C2", 20.0);
    sheet.set("C3", 30.0);

    // Default area_num is 1.
    assert_eq!(sheet.eval("=INDEX((A1:A3,C1:C3),2,1)"), Value::Number(2.0));

    // Out of bounds area_num -> #REF!
    assert_eq!(
        sheet.eval("=INDEX((A1:A3,C1:C3),2,1,0)"),
        Value::Error(ErrorKind::Ref)
    );
    assert_eq!(
        sheet.eval("=INDEX((A1:A3,C1:C3),2,1,3)"),
        Value::Error(ErrorKind::Ref)
    );

    // Passing area_num for a single-area reference behaves like Excel and errors unless it's 1.
    assert_eq!(
        sheet.eval("=INDEX(A1:A3,2,1,2)"),
        Value::Error(ErrorKind::Ref)
    );
}

#[test]
fn index_reference_slices_spill_when_used_as_formula_result() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "B1", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "C1", 3.0).unwrap();

    engine.set_cell_value("Sheet1", "A2", 10.0).unwrap();
    engine.set_cell_value("Sheet1", "B2", 20.0).unwrap();
    engine.set_cell_value("Sheet1", "C2", 30.0).unwrap();

    engine.set_cell_value("Sheet1", "A3", 100.0).unwrap();
    engine.set_cell_value("Sheet1", "B3", 200.0).unwrap();
    engine.set_cell_value("Sheet1", "C3", 300.0).unwrap();

    // Column slice spills vertically.
    engine
        .set_cell_formula("Sheet1", "E1", "=INDEX(A1:C3,0,2)")
        .unwrap();
    // Row slice spills horizontally.
    engine
        .set_cell_formula("Sheet1", "G1", "=INDEX(A1:C3,2,0)")
        .unwrap();
    // Full range spills as a 2D array.
    engine
        .set_cell_formula("Sheet1", "K1", "=INDEX(A1:C3,0,0)")
        .unwrap();

    engine.recalculate();

    // Column slice (B1:B3) -> {2;20;200}
    let (start, end) = engine.spill_range("Sheet1", "E1").expect("E1 spill range");
    assert_eq!(start, parse_a1("E1").unwrap());
    assert_eq!(end, parse_a1("E3").unwrap());
    assert_eq!(engine.get_cell_value("Sheet1", "E1"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "E2"), Value::Number(20.0));
    assert_eq!(engine.get_cell_value("Sheet1", "E3"), Value::Number(200.0));

    // Row slice (A2:C2) -> {10,20,30}
    let (start, end) = engine.spill_range("Sheet1", "G1").expect("G1 spill range");
    assert_eq!(start, parse_a1("G1").unwrap());
    assert_eq!(end, parse_a1("I1").unwrap());
    assert_eq!(engine.get_cell_value("Sheet1", "G1"), Value::Number(10.0));
    assert_eq!(engine.get_cell_value("Sheet1", "H1"), Value::Number(20.0));
    assert_eq!(engine.get_cell_value("Sheet1", "I1"), Value::Number(30.0));

    // Full range -> 3x3 spill.
    let (start, end) = engine.spill_range("Sheet1", "K1").expect("K1 spill range");
    assert_eq!(start, parse_a1("K1").unwrap());
    assert_eq!(end, parse_a1("M3").unwrap());
    assert_eq!(engine.get_cell_value("Sheet1", "K1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "L1"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "M1"), Value::Number(3.0));
    assert_eq!(engine.get_cell_value("Sheet1", "K2"), Value::Number(10.0));
    assert_eq!(engine.get_cell_value("Sheet1", "L2"), Value::Number(20.0));
    assert_eq!(engine.get_cell_value("Sheet1", "M2"), Value::Number(30.0));
    assert_eq!(engine.get_cell_value("Sheet1", "K3"), Value::Number(100.0));
    assert_eq!(engine.get_cell_value("Sheet1", "L3"), Value::Number(200.0));
    assert_eq!(engine.get_cell_value("Sheet1", "M3"), Value::Number(300.0));
}

#[test]
fn index_array_slices_spill_when_used_as_formula_result() {
    let mut engine = Engine::new();

    // Column slice spills vertically.
    engine
        .set_cell_formula("Sheet1", "A1", "=INDEX({1,2,3;4,5,6},0,2)")
        .unwrap();
    // Row slice spills horizontally.
    engine
        .set_cell_formula("Sheet1", "D1", "=INDEX({1,2,3;4,5,6},2,0)")
        .unwrap();
    // Default column_num is 1, so row_num=0 returns the first column.
    engine
        .set_cell_formula("Sheet1", "G1", "=INDEX({1,2,3;4,5,6},0)")
        .unwrap();

    engine.recalculate();

    // {1,2,3;4,5,6} column 2 -> {2;5}
    let (start, end) = engine.spill_range("Sheet1", "A1").expect("A1 spill range");
    assert_eq!(start, parse_a1("A1").unwrap());
    assert_eq!(end, parse_a1("A2").unwrap());
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "A2"), Value::Number(5.0));

    // {1,2,3;4,5,6} row 2 -> {4,5,6}
    let (start, end) = engine.spill_range("Sheet1", "D1").expect("D1 spill range");
    assert_eq!(start, parse_a1("D1").unwrap());
    assert_eq!(end, parse_a1("F1").unwrap());
    assert_eq!(engine.get_cell_value("Sheet1", "D1"), Value::Number(4.0));
    assert_eq!(engine.get_cell_value("Sheet1", "E1"), Value::Number(5.0));
    assert_eq!(engine.get_cell_value("Sheet1", "F1"), Value::Number(6.0));

    // {1,2,3;4,5,6} row_num=0 defaults to col 1 -> {1;4}
    let (start, end) = engine.spill_range("Sheet1", "G1").expect("G1 spill range");
    assert_eq!(start, parse_a1("G1").unwrap());
    assert_eq!(end, parse_a1("G2").unwrap());
    assert_eq!(engine.get_cell_value("Sheet1", "G1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "G2"), Value::Number(4.0));
}
