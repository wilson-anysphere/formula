use formula_engine::{Engine, ErrorKind, Value};
use pretty_assertions::assert_eq;

#[test]
fn sheet_span_ordering_is_correct_after_reorder_and_delete_many_sheets() {
    let mut engine = Engine::new();

    // Use >100 sheets to ensure we exercise the Snapshot tab-index cache on a large workbook.
    const SHEET_COUNT: usize = 128;

    for idx in 1..=SHEET_COUNT {
        let name = format!("Sheet{idx:03}");
        engine.set_cell_value(&name, "A1", idx as f64).unwrap();
    }

    engine
        .set_cell_formula("Summary", "A1", "=INDEX(Sheet001:Sheet128!A1,1,1,50)")
        .unwrap();
    engine
        .set_cell_formula("Summary", "A2", "=INDEX(Sheet001:Sheet128!A1,1,1,64)")
        .unwrap();
    engine
        .set_cell_formula("Summary", "A3", "=INDEX(Sheet001:Sheet128!A1,1,1,128)")
        .unwrap();

    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Summary", "A1"), Value::Number(50.0));
    assert_eq!(engine.get_cell_value("Summary", "A2"), Value::Number(64.0));
    assert_eq!(engine.get_cell_value("Summary", "A3"), Value::Number(128.0));

    // Reordering sheets changes workbook tab order; 3D spans and multi-area ordering must follow it.
    assert!(engine.reorder_sheet("Sheet064", 0));
    engine.recalculate_single_threaded();

    // `Sheet064` is now left of `Sheet001`, so it falls outside the `Sheet001:Sheet128` span.
    assert_eq!(engine.get_cell_value("Summary", "A1"), Value::Number(50.0));
    assert_eq!(engine.get_cell_value("Summary", "A2"), Value::Number(65.0));
    assert_eq!(
        engine.get_cell_value("Summary", "A3"),
        Value::Error(ErrorKind::Ref)
    );

    // Deleting an intermediate sheet should shrink the span without changing its boundaries.
    engine.delete_sheet("Sheet050").unwrap();
    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Summary", "A1"), Value::Number(51.0));
    assert_eq!(engine.get_cell_value("Summary", "A2"), Value::Number(66.0));
    assert_eq!(
        engine.get_cell_value("Summary", "A3"),
        Value::Error(ErrorKind::Ref)
    );
}
