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
fn parses_double_quoted_sheet_range_prefix() {
    let ast = parse_formula("=SUM('Sheet 1':'Sheet 3'!A1)", ParseOptions::default()).unwrap();
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
fn sum_full_sheet_range_over_sheet_range_is_sparse_and_marks_dirty() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet2", "B2", 2.0).unwrap();
    engine.set_cell_value("Sheet3", "XFD1048576", 3.0).unwrap();

    engine
        // Place the formula on a different sheet so we don't create a circular reference: `A:XFD`
        // covers the entire sheet.
        .set_cell_formula("Summary", "C1", "=SUM(Sheet1:Sheet3!A:XFD)")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Summary", "C1"), Value::Number(6.0));

    // Updating a cell that was previously blank should still dirty the formula cell even though
    // the audit graph does not expand full-sheet ranges.
    engine.set_cell_value("Sheet2", "C3", 10.0).unwrap();
    assert!(engine.is_dirty("Summary", "C1"));
    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Summary", "C1"), Value::Number(16.0));
}

#[test]
fn union_over_sheet_range_refs_is_ref_error() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet2", "A1", 2.0).unwrap();
    engine.set_cell_value("Sheet3", "A1", 3.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 10.0).unwrap();

    // Union inside a function argument must be parenthesized to avoid being parsed as multiple
    // arguments. Excel's union/intersection reference algebra does not allow combining references
    // that resolve to different sheets, so 3D spans cannot participate (they expand to multiple
    // per-sheet areas).
    engine
        .set_cell_formula("Summary", "A1", "=SUM((Sheet1:Sheet3!A1,Sheet1!A2))")
        .unwrap();
    engine.recalculate_single_threaded();

    assert_eq!(
        engine.get_cell_value("Summary", "A1"),
        Value::Error(formula_engine::ErrorKind::Ref)
    );
}

#[test]
fn evaluates_sum_over_sheet_range_ref_and_additional_argument() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet2", "A1", 2.0).unwrap();
    engine.set_cell_value("Sheet3", "A1", 3.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 10.0).unwrap();

    // Use separate function arguments rather than the reference union operator.
    engine
        .set_cell_formula("Summary", "A1", "=SUM(Sheet1:Sheet3!A1,Sheet1!A2)")
        .unwrap();
    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Summary", "A1"), Value::Number(16.0));
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
