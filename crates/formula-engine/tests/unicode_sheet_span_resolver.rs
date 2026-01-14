use formula_engine::editing::EditOp;
use formula_engine::Engine;
use pretty_assertions::assert_eq;

#[test]
fn engine_rewrites_sheet_spans_using_unicode_aware_sheet_lookup() {
    // `SS` should resolve to a sheet named `ß` (sharp s uppercases to `SS`) when determining
    // whether a sheet span contains the edited sheet.
    let mut engine = Engine::new();
    engine.ensure_sheet("ß");
    engine.ensure_sheet("Middle");
    engine.ensure_sheet("Sheet3");
    engine.ensure_sheet("Summary");

    engine
        .set_cell_formula("Summary", "A1", "=SUM(SS:Sheet3!A1)")
        .unwrap();

    engine
        .apply_operation(EditOp::InsertRows {
            sheet: "Middle".to_string(),
            row: 0,
            count: 1,
        })
        .unwrap();

    assert_eq!(
        engine.get_cell_formula("Summary", "A1"),
        Some("=SUM(SS:Sheet3!A2)")
    );
}

#[test]
fn engine_rewrites_sheet_spans_using_nfkc_case_insensitive_sheet_lookup() {
    // NFKC equivalence should be considered when resolving sheet names for 3D spans.
    //
    // U+212A KELVIN SIGN (K) is NFKC-equivalent to ASCII 'K', so a workbook sheet named `Kelvin`
    // should be addressable as `KELVIN` in a 3D span boundary.
    let mut engine = Engine::new();
    engine.ensure_sheet("Kelvin");
    engine.ensure_sheet("Middle");
    engine.ensure_sheet("Sheet3");
    engine.ensure_sheet("Summary");

    engine
        .set_cell_formula("Summary", "A1", "=SUM(KELVIN:Sheet3!A1)")
        .unwrap();

    engine
        .apply_operation(EditOp::InsertRows {
            sheet: "Middle".to_string(),
            row: 0,
            count: 1,
        })
        .unwrap();

    assert_eq!(
        engine.get_cell_formula("Summary", "A1"),
        Some("=SUM(KELVIN:Sheet3!A2)")
    );
}
