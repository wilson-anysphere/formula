use formula_engine::{parse_formula, Engine, ParseOptions, SerializeOptions, Value};
use pretty_assertions::assert_eq;

#[test]
fn parse_and_roundtrip_sheet_range_ref() {
    let ast = parse_formula("=SUM(Sheet1:Sheet3!A1)", ParseOptions::default()).unwrap();
    let roundtrip = ast.to_string(SerializeOptions::default()).unwrap();
    assert_eq!(roundtrip, "=SUM(Sheet1:Sheet3!A1)");
}

#[test]
fn parses_quoted_sheet_range_prefix() {
    let ast = parse_formula("=SUM('Sheet1:Sheet3'!A1)", ParseOptions::default()).unwrap();
    let roundtrip = ast.to_string(SerializeOptions::default()).unwrap();
    // Canonical serialization does not preserve the single-quoted span; it emits the equivalent
    // unquoted form when possible.
    assert_eq!(roundtrip, "=SUM(Sheet1:Sheet3!A1)");
}

#[test]
fn roundtrip_preserves_single_quoted_sheet_range_when_required() {
    let ast = parse_formula("=SUM('Sheet 1:Sheet 3'!A1)", ParseOptions::default()).unwrap();
    let roundtrip = ast.to_string(SerializeOptions::default()).unwrap();
    assert_eq!(roundtrip, "=SUM('Sheet 1:Sheet 3'!A1)");
}

#[test]
fn collapses_degenerate_sheet_range_refs_to_single_sheet() {
    let ast = parse_formula("=SUM(Sheet1:Sheet1!A1)", ParseOptions::default()).unwrap();
    let roundtrip = ast.to_string(SerializeOptions::default()).unwrap();
    assert_eq!(roundtrip, "=SUM(Sheet1!A1)");

    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 4.0).unwrap();
    engine
        .set_cell_formula("Summary", "A1", "=SUM(Sheet1:Sheet1!A1)")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Summary", "A1"), Value::Number(4.0));
}

#[test]
fn evaluates_sum_over_sheet_range_cell_ref() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet2", "A1", 2.0).unwrap();
    engine.set_cell_value("Sheet3", "A1", 3.0).unwrap();

    engine
        .set_cell_formula("Summary", "A1", "=SUM(Sheet1:Sheet3!A1)")
        .unwrap();
    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Summary", "A1"), Value::Number(6.0));
}

#[test]
fn evaluates_sum_over_quoted_sheet_range_with_spaces() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet 1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet 2", "A1", 2.0).unwrap();
    engine.set_cell_value("Sheet 3", "A1", 3.0).unwrap();

    engine
        .set_cell_formula("Summary", "A1", "=SUM('Sheet 1:Sheet 3'!A1)")
        .unwrap();
    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Summary", "A1"), Value::Number(6.0));
}

#[test]
fn evaluates_sum_over_sheet_range_area_ref() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 10.0).unwrap();
    engine.set_cell_value("Sheet2", "A1", 2.0).unwrap();
    engine.set_cell_value("Sheet2", "A2", 20.0).unwrap();
    engine.set_cell_value("Sheet3", "A1", 3.0).unwrap();
    engine.set_cell_value("Sheet3", "A2", 30.0).unwrap();

    engine
        .set_cell_formula("Summary", "A1", "=SUM(Sheet1:Sheet3!A1:A2)")
        .unwrap();
    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Summary", "A1"), Value::Number(66.0));
}

#[test]
fn evaluates_sum_over_sheet_range_column_range() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet2", "A2", 2.0).unwrap();
    engine.set_cell_value("Sheet3", "A3", 3.0).unwrap();
    engine.set_cell_value("Sheet2", "B1", 100.0).unwrap();

    engine
        .set_cell_formula("Summary", "A1", "=SUM(Sheet1:Sheet3!A:A)")
        .unwrap();
    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Summary", "A1"), Value::Number(6.0));
}

#[test]
fn evaluates_sum_over_sheet_range_row_range() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet2", "B1", 2.0).unwrap();
    engine.set_cell_value("Sheet3", "C1", 3.0).unwrap();
    engine.set_cell_value("Sheet3", "A2", 100.0).unwrap();

    engine
        .set_cell_formula("Summary", "A1", "=SUM(Sheet1:Sheet3!1:1)")
        .unwrap();
    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Summary", "A1"), Value::Number(6.0));
}

#[test]
fn recalculates_when_intermediate_sheet_changes() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet2", "A1", 2.0).unwrap();
    engine.set_cell_value("Sheet3", "A1", 3.0).unwrap();

    engine
        .set_cell_formula("Summary", "A1", "=SUM(Sheet1:Sheet3!A1)")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Summary", "A1"), Value::Number(6.0));

    engine.set_cell_value("Sheet2", "A1", 5.0).unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Summary", "A1"), Value::Number(9.0));
}

#[test]
fn evaluates_sum_over_reversed_sheet_range_ref() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet2", "A1", 2.0).unwrap();
    engine.set_cell_value("Sheet3", "A1", 3.0).unwrap();

    // Excel resolves 3D spans by workbook sheet order regardless of whether the
    // user writes them forward or reversed.
    engine
        .set_cell_formula("Summary", "A1", "=SUM(Sheet3:Sheet1!A1)")
        .unwrap();
    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Summary", "A1"), Value::Number(6.0));
}
