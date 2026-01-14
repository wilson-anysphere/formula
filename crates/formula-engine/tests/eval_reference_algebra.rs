use formula_engine::{Engine, ErrorKind, Value};
use pretty_assertions::assert_eq;

#[test]
fn sum_full_sheet_range_is_sparse_and_marks_dirty_via_calc_graph() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "B2", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "XFD1048576", 3.0).unwrap();

    engine
        // Place the formula on a different sheet to avoid a circular reference: `A:XFD`
        // covers the entire sheet.
        .set_cell_formula("Sheet2", "C1", "=SUM(Sheet1!A:XFD)")
        .unwrap();
    engine.recalculate();
    assert_eq!(engine.get_cell_value("Sheet2", "C1"), Value::Number(6.0));

    // Updating a cell that was previously blank should dirty the formula cell even though the
    // audit graph does not expand full-sheet ranges.
    engine.set_cell_value("Sheet1", "C3", 10.0).unwrap();
    assert!(engine.is_dirty("Sheet2", "C1"));
    engine.recalculate();
    assert_eq!(engine.get_cell_value("Sheet2", "C1"), Value::Number(16.0));
}

#[test]
fn empty_intersection_yields_null_error() {
    let mut engine = Engine::new();
    // Keep the formula outside of the referenced columns so this doesn't become a circular
    // reference (Excel surfaces circular refs as 0 with a warning).
    engine.set_cell_formula("Sheet1", "C1", "=A:A B:B").unwrap();
    engine.recalculate();
    assert_eq!(
        engine.get_cell_value("Sheet1", "C1"),
        Value::Error(ErrorKind::Null)
    );
}

#[test]
fn union_references_can_be_consumed_by_sum_and_counta() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "B1", 10.0).unwrap();
    engine.set_cell_value("Sheet1", "B2", "x").unwrap();

    engine
        .set_cell_formula("Sheet1", "C1", "=SUM((A1:A2,B1:B2))")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "C2", "=COUNTA((A1:A2,B1:B2))")
        .unwrap();
    engine.recalculate();

    // SUM ignores text/logicals in reference arguments.
    assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::Number(13.0));
    assert_eq!(engine.get_cell_value("Sheet1", "C2"), Value::Number(4.0));
}

#[test]
fn row_ranges_work() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "C1", 2.0).unwrap();
    engine
        .set_cell_formula("Sheet1", "B2", "=SUM(1:1)")
        .unwrap();
    engine.recalculate();
    assert_eq!(engine.get_cell_value("Sheet1", "B2"), Value::Number(3.0));

    // Row ranges cover thousands of cells; the audit graph won't expand them. Ensure calc-graph
    // tracking keeps the dependent dirty even when the edited cell didn't previously exist.
    engine.set_cell_value("Sheet1", "B1", 10.0).unwrap();
    assert!(engine.is_dirty("Sheet1", "B2"));
    engine.recalculate();
    assert_eq!(engine.get_cell_value("Sheet1", "B2"), Value::Number(13.0));
}

#[test]
fn deref_full_sheet_range_as_formula_result_returns_spill_error() {
    let mut engine = Engine::new();
    // A bare reference result is treated as a dynamic array. Materializing a full-sheet reference
    // would be prohibitively large, so the evaluator should surface `#SPILL!` instead of attempting
    // to allocate an array of billions of cells.
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine
        // Put the formula on another sheet to avoid a circular reference (the referenced range
        // spans the entire source sheet).
        .set_cell_formula("Sheet2", "A1", "=Sheet1!A:XFD")
        .unwrap();
    engine.recalculate();
    assert_eq!(
        engine.get_cell_value("Sheet2", "A1"),
        Value::Error(ErrorKind::Spill)
    );
}
