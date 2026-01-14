use formula_engine::{EditOp, Engine, NameDefinition, NameScope, Value};
use formula_model::{CellRef, Range, EXCEL_MAX_COLS};
use pretty_assertions::assert_eq;

fn cell(a1: &str) -> CellRef {
    CellRef::from_a1(a1).unwrap()
}

fn range(a1: &str) -> Range {
    Range::from_a1(a1).unwrap()
}

#[test]
fn insert_row_above_updates_references_and_moves_cells() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_formula("Sheet1", "B1", "=A1").unwrap();

    engine
        .apply_operation(EditOp::InsertRows {
            sheet: "Sheet1".to_string(),
            row: 0,
            count: 1,
        })
        .unwrap();

    assert_eq!(engine.get_cell_formula("Sheet1", "B2"), Some("=A2"));
    assert_eq!(engine.get_cell_value("Sheet1", "A2"), Value::Number(1.0));
}

#[test]
fn insert_row_above_updates_references_when_sheet_display_name_differs_from_key() {
    let mut engine = Engine::new();
    engine.ensure_sheet("sheet1_key");
    engine.set_sheet_display_name("sheet1_key", "Sheet1");

    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_formula("Sheet1", "B1", "=A1").unwrap();

    // Apply the edit using the *display name*.
    engine
        .apply_operation(EditOp::InsertRows {
            sheet: "Sheet1".to_string(),
            row: 0,
            count: 1,
        })
        .unwrap();

    assert_eq!(engine.get_cell_formula("Sheet1", "B2"), Some("=A2"));
    assert_eq!(engine.get_cell_value("Sheet1", "A2"), Value::Number(1.0));
}

#[test]
fn insert_row_above_rewrites_cross_sheet_refs_when_op_uses_sheet_key() {
    let mut engine = Engine::new();
    engine.ensure_sheet("sheet1_key");
    engine.ensure_sheet("sheet2_key");
    engine.set_sheet_display_name("sheet1_key", "Sheet1");
    engine.set_sheet_display_name("sheet2_key", "Sheet2");

    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine
        .set_cell_formula("Sheet2", "A1", "=Sheet1!A1")
        .unwrap();

    // Apply the edit using the *stable sheet key*.
    engine
        .apply_operation(EditOp::InsertRows {
            sheet: "sheet1_key".to_string(),
            row: 0,
            count: 1,
        })
        .unwrap();

    assert_eq!(engine.get_cell_formula("Sheet2", "A1"), Some("=Sheet1!A2"));
}

#[test]
fn insert_row_above_matches_sheet_names_case_insensitively_across_unicode() {
    let mut engine = Engine::new();

    // Use a Unicode sheet name that requires Unicode-aware case folding (ß -> SS).
    engine.set_cell_value("Straße", "A1", 1.0).unwrap();
    engine.set_cell_formula("Straße", "B1", "=A1").unwrap();
    engine
        .set_cell_formula("Other", "A1", "='Straße'!A1")
        .unwrap();

    // Apply the edit using a different (Unicode-folded) spelling of the sheet name.
    engine
        .apply_operation(EditOp::InsertRows {
            sheet: "STRASSE".to_string(),
            row: 0,
            count: 1,
        })
        .unwrap();

    assert_eq!(engine.get_cell_formula("Straße", "B2"), Some("=A2"));
    assert_eq!(engine.get_cell_formula("Other", "A1"), Some("='Straße'!A2"));
    assert_eq!(engine.get_cell_value("Straße", "A2"), Value::Number(1.0));
}

#[test]
fn insert_row_above_matches_sheet_names_nfkc_case_insensitively() {
    let mut engine = Engine::new();

    // Use a sheet name that requires NFKC normalization to match (K == K).
    engine.set_cell_value("Kelvin", "A1", 1.0).unwrap();
    engine.set_cell_formula("Kelvin", "B1", "=A1").unwrap();
    engine
        .set_cell_formula("Other", "A1", "='Kelvin'!A1")
        .unwrap();

    engine
        .apply_operation(EditOp::InsertRows {
            sheet: "KELVIN".to_string(),
            row: 0,
            count: 1,
        })
        .unwrap();

    assert_eq!(engine.get_cell_formula("Kelvin", "B2"), Some("=A2"));
    assert_eq!(engine.get_cell_formula("Other", "A1"), Some("='Kelvin'!A2"));
}

#[test]
fn delete_column_updates_references_and_creates_ref_errors() {
    let mut engine = Engine::new();
    engine.set_cell_formula("Sheet1", "C1", "=A1+B1").unwrap();

    engine
        .apply_operation(EditOp::DeleteCols {
            sheet: "Sheet1".to_string(),
            col: 0,
            count: 1,
        })
        .unwrap();

    // C1 shifted left to B1 and its A1 reference was deleted.
    assert_eq!(engine.get_cell_formula("Sheet1", "B1"), Some("=#REF!+A1"));
}

#[test]
fn move_range_rewrites_formulas_to_follow_moved_cells() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 42.0).unwrap();
    engine.set_cell_formula("Sheet1", "B1", "=A1").unwrap();
    engine.set_cell_formula("Sheet1", "C1", "=A1").unwrap();

    engine
        .apply_operation(EditOp::MoveRange {
            sheet: "Sheet1".to_string(),
            src: range("A1:B1"),
            dst_top_left: cell("A2"),
        })
        .unwrap();

    assert_eq!(engine.get_cell_formula("Sheet1", "B2"), Some("=A2"));
    assert_eq!(engine.get_cell_formula("Sheet1", "C1"), Some("=A2"));
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Blank);
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Blank);
}

#[test]
fn copy_range_adjusts_relative_references() {
    let mut engine = Engine::new();
    engine.set_cell_formula("Sheet1", "B1", "=A1").unwrap();

    engine
        .apply_operation(EditOp::CopyRange {
            sheet: "Sheet1".to_string(),
            src: range("B1"),
            dst_top_left: cell("B2"),
        })
        .unwrap();

    assert_eq!(engine.get_cell_formula("Sheet1", "B1"), Some("=A1"));
    assert_eq!(engine.get_cell_formula("Sheet1", "B2"), Some("=A2"));
}

#[test]
fn copy_range_adjusts_external_workbook_cell_references() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "B1", "=[Book.xlsx]Sheet1!A1")
        .unwrap();

    engine
        .apply_operation(EditOp::CopyRange {
            sheet: "Sheet1".to_string(),
            src: range("B1"),
            dst_top_left: cell("B2"),
        })
        .unwrap();

    assert_eq!(
        engine.get_cell_formula("Sheet1", "B1"),
        Some("=[Book.xlsx]Sheet1!A1")
    );
    assert_eq!(
        engine.get_cell_formula("Sheet1", "B2"),
        Some("=[Book.xlsx]Sheet1!A2")
    );
}

#[test]
fn fill_adjusts_external_workbook_range_references() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "C1", "=SUM([Book.xlsx]Sheet1!A1:B2)")
        .unwrap();

    engine
        .apply_operation(EditOp::Fill {
            sheet: "Sheet1".to_string(),
            src: range("C1"),
            dst: range("C1:C2"),
        })
        .unwrap();

    assert_eq!(
        engine.get_cell_formula("Sheet1", "C1"),
        Some("=SUM([Book.xlsx]Sheet1!A1:B2)")
    );
    assert_eq!(
        engine.get_cell_formula("Sheet1", "C2"),
        Some("=SUM([Book.xlsx]Sheet1!A2:B3)")
    );
}

#[test]
fn copy_range_does_not_adjust_external_workbook_absolute_references() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "B1", "=[Book.xlsx]Sheet1!$A$1")
        .unwrap();

    engine
        .apply_operation(EditOp::CopyRange {
            sheet: "Sheet1".to_string(),
            src: range("B1"),
            dst_top_left: cell("B2"),
        })
        .unwrap();

    assert_eq!(
        engine.get_cell_formula("Sheet1", "B1"),
        Some("=[Book.xlsx]Sheet1!$A$1")
    );
    assert_eq!(
        engine.get_cell_formula("Sheet1", "B2"),
        Some("=[Book.xlsx]Sheet1!$A$1")
    );
}

#[test]
fn copy_range_to_far_row_grows_sheet_dimensions() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 42.0).unwrap();

    engine
        .apply_operation(EditOp::CopyRange {
            sheet: "Sheet1".to_string(),
            src: range("A1"),
            dst_top_left: cell("A2000000"),
        })
        .unwrap();

    // Copying a literal value to a row beyond Excel's default should automatically grow the sheet
    // dimensions so the new cell remains addressable.
    assert_eq!(
        engine.get_cell_value("Sheet1", "A2000000"),
        Value::Number(42.0)
    );
    assert_eq!(
        engine.sheet_dimensions("Sheet1"),
        Some((2_000_000, EXCEL_MAX_COLS))
    );
}

#[test]
fn fill_repeats_formulas_and_updates_relative_references() {
    let mut engine = Engine::new();
    engine.set_cell_formula("Sheet1", "C1", "=A1+B1").unwrap();

    engine
        .apply_operation(EditOp::Fill {
            sheet: "Sheet1".to_string(),
            src: range("C1"),
            dst: range("C1:C3"),
        })
        .unwrap();

    assert_eq!(engine.get_cell_formula("Sheet1", "C1"), Some("=A1+B1"));
    assert_eq!(engine.get_cell_formula("Sheet1", "C2"), Some("=A2+B2"));
    assert_eq!(engine.get_cell_formula("Sheet1", "C3"), Some("=A3+B3"));
}

#[test]
fn structural_edits_update_sheet_qualified_references() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 0.0).unwrap();
    engine
        .set_cell_formula("Other", "A1", "=Sheet1!A1")
        .unwrap();

    engine
        .apply_operation(EditOp::InsertRows {
            sheet: "Sheet1".to_string(),
            row: 0,
            count: 1,
        })
        .unwrap();

    assert_eq!(engine.get_cell_formula("Other", "A1"), Some("=Sheet1!A2"));
}

#[test]
fn structural_edits_update_sheet_range_references() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet2", "A1", 2.0).unwrap();
    engine.set_cell_value("Sheet3", "A1", 3.0).unwrap();
    engine
        .set_cell_formula("Summary", "A1", "=SUM(Sheet1:Sheet3!A1)")
        .unwrap();

    engine
        .apply_operation(EditOp::InsertRows {
            sheet: "Sheet2".to_string(),
            row: 0,
            count: 1,
        })
        .unwrap();

    assert_eq!(
        engine.get_cell_formula("Summary", "A1"),
        Some("=SUM(Sheet1:Sheet3!A2)")
    );
}

#[test]
fn structural_edits_update_unicode_sheet_range_references() {
    // Regression test for Unicode sheet name matching in structural edit rewrites.
    //
    // The engine treats sheet names case-insensitively across Unicode (see `Workbook::sheet_id`),
    // so structural edits must resolve the same sheets when rewriting 3D spans.
    let mut engine = Engine::new();
    engine.set_cell_value("Café", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet2", "A1", 2.0).unwrap();
    engine.set_cell_value("Sheet3", "A1", 3.0).unwrap();
    // `CAFÉ` uses a different Unicode case for `É` than the sheet's display name.
    engine
        .set_cell_formula("Summary", "A1", "=SUM('CAFÉ:Sheet3'!A1)")
        .unwrap();

    engine
        .apply_operation(EditOp::InsertRows {
            sheet: "Sheet2".to_string(),
            row: 0,
            count: 1,
        })
        .unwrap();

    assert_eq!(
        engine.get_cell_formula("Summary", "A1"),
        Some("=SUM('CAFÉ:Sheet3'!A2)")
    );
}

#[test]
fn range_map_edits_update_unicode_sheet_range_references() {
    // Regression test for Unicode sheet name matching in range-map rewrites (insert/delete cells,
    // move range, etc.).
    //
    // Like structural edits, 3D spans are defined by sheet tab order and sheet name comparisons
    // must be Unicode-aware.
    let mut engine = Engine::new();
    engine.set_cell_value("Café", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet2", "A1", 2.0).unwrap();
    engine.set_cell_value("Sheet3", "A1", 3.0).unwrap();
    engine
        .set_cell_formula("Summary", "A1", "=SUM('CAFÉ:Sheet3'!A1)")
        .unwrap();

    engine
        .apply_operation(EditOp::InsertCellsShiftDown {
            sheet: "Sheet2".to_string(),
            range: range("A1"),
        })
        .unwrap();

    assert_eq!(
        engine.get_cell_formula("Summary", "A1"),
        Some("=SUM('CAFÉ:Sheet3'!A2)")
    );
}

#[test]
fn structural_edits_resolve_sheet_names_using_nfkc_normalization() {
    // Regression test for NFKC sheet name equivalence in structural edit rewrites.
    //
    // Excel compares sheet names case-insensitively across Unicode and applies compatibility
    // normalization (NFKC). For example, Å (ANGSTROM SIGN) is compatibility-equivalent to Å
    // (LATIN CAPITAL LETTER A WITH RING ABOVE). The engine's sheet lookup and structural formula
    // rewrite resolvers must treat them as the same sheet name.
    let mut engine = Engine::new();
    engine.set_cell_value("Å", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet2", "A1", 2.0).unwrap();
    engine.set_cell_value("Sheet3", "A1", 3.0).unwrap();
    engine
        // The 3D span uses Å, while the actual sheet name is Å.
        .set_cell_formula("Summary", "A1", "=SUM('Å:Sheet3'!A1)")
        .unwrap();

    engine
        .apply_operation(EditOp::InsertRows {
            sheet: "Sheet2".to_string(),
            row: 0,
            count: 1,
        })
        .unwrap();

    assert_eq!(
        engine.get_cell_formula("Summary", "A1"),
        Some("=SUM('Å:Sheet3'!A2)")
    );
}

#[test]
fn insert_row_updates_mixed_absolute_and_relative_references() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "B1", "=$A$1+$A1+A$1")
        .unwrap();

    engine
        .apply_operation(EditOp::InsertRows {
            sheet: "Sheet1".to_string(),
            row: 0,
            count: 1,
        })
        .unwrap();

    // Formula moved from B1 -> B2; all references should track the moved cells.
    assert_eq!(
        engine.get_cell_formula("Sheet1", "B2"),
        Some("=$A$2+$A2+A$2")
    );
}

#[test]
fn delete_col_updates_mixed_absolute_and_relative_references() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "C1", "=$A$1+$B$1")
        .unwrap();

    engine
        .apply_operation(EditOp::DeleteCols {
            sheet: "Sheet1".to_string(),
            col: 0,
            count: 1,
        })
        .unwrap();

    // Formula moved from C1 -> B1. $A$1 is deleted, and $B$1 shifts to $A$1.
    assert_eq!(engine.get_cell_formula("Sheet1", "B1"), Some("=#REF!+$A$1"));
}

#[test]
fn copy_range_adjusts_range_references() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "C1", "=SUM(A1:B2)")
        .unwrap();

    engine
        .apply_operation(EditOp::CopyRange {
            sheet: "Sheet1".to_string(),
            src: range("C1"),
            dst_top_left: cell("C2"),
        })
        .unwrap();

    assert_eq!(engine.get_cell_formula("Sheet1", "C2"), Some("=SUM(A2:B3)"));
}

#[test]
fn structural_edits_rewrite_quoted_sheet_names() {
    let mut engine = Engine::new();
    engine.set_cell_value("My Sheet", "A1", 0.0).unwrap();
    engine
        .set_cell_formula("Other", "A1", "='My Sheet'!A1")
        .unwrap();

    engine
        .apply_operation(EditOp::InsertRows {
            sheet: "My Sheet".to_string(),
            row: 0,
            count: 1,
        })
        .unwrap();

    assert_eq!(
        engine.get_cell_formula("Other", "A1"),
        Some("='My Sheet'!A2")
    );
}

#[test]
fn copy_range_adjusts_row_and_column_ranges() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "C1", "=SUM(1:1)")
        .unwrap();
    engine
        .apply_operation(EditOp::CopyRange {
            sheet: "Sheet1".to_string(),
            src: range("C1"),
            dst_top_left: cell("C2"),
        })
        .unwrap();
    assert_eq!(engine.get_cell_formula("Sheet1", "C2"), Some("=SUM(2:2)"));

    engine
        .set_cell_formula("Sheet1", "C3", "=SUM(A:A)")
        .unwrap();
    engine
        .apply_operation(EditOp::CopyRange {
            sheet: "Sheet1".to_string(),
            src: range("C3"),
            dst_top_left: cell("D3"),
        })
        .unwrap();
    assert_eq!(engine.get_cell_formula("Sheet1", "D3"), Some("=SUM(B:B)"));
}

#[test]
fn insert_cells_shift_right_moves_cells_and_rewrites_references() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "C1", 3.0).unwrap();
    engine.set_cell_formula("Sheet1", "D1", "=A1+C1").unwrap();

    engine
        .apply_operation(EditOp::InsertCellsShiftRight {
            sheet: "Sheet1".to_string(),
            range: range("A1:B1"),
        })
        .unwrap();

    // A1 moved to C1, and C1 moved to E1.
    assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "E1"), Value::Number(3.0));
    // Formula moved from D1 -> F1 and should track the moved cells.
    assert_eq!(engine.get_cell_formula("Sheet1", "F1"), Some("=C1+E1"));
}

#[test]
fn delete_cells_shift_left_creates_ref_errors_and_updates_shifted_references() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "B1", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "C1", 3.0).unwrap();
    engine.set_cell_value("Sheet1", "D1", 4.0).unwrap();
    engine.set_cell_formula("Sheet1", "E1", "=A1+D1").unwrap();
    // This reference points into the deleted region and should become #REF!
    engine.set_cell_formula("Sheet1", "A2", "=B1").unwrap();

    engine
        .apply_operation(EditOp::DeleteCellsShiftLeft {
            sheet: "Sheet1".to_string(),
            range: range("B1:C1"),
        })
        .unwrap();

    // D1 moved into B1.
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(4.0));
    // Formula moved from E1 -> C1 and should track the moved cell (D1 -> B1).
    assert_eq!(engine.get_cell_formula("Sheet1", "C1"), Some("=A1+B1"));
    // Reference into deleted region becomes #REF!, even though another cell moved into B1.
    assert_eq!(engine.get_cell_formula("Sheet1", "A2"), Some("=#REF!"));
}

#[test]
fn insert_cells_shift_down_rewrites_references_into_shifted_region() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 42.0).unwrap();
    engine.set_cell_formula("Sheet1", "B1", "=A1").unwrap();

    engine
        .apply_operation(EditOp::InsertCellsShiftDown {
            sheet: "Sheet1".to_string(),
            range: range("A1"),
        })
        .unwrap();

    // A1 moved down to A2; formula should follow it.
    assert_eq!(engine.get_cell_value("Sheet1", "A2"), Value::Number(42.0));
    assert_eq!(engine.get_cell_formula("Sheet1", "B1"), Some("=A2"));
}

#[test]
fn delete_cells_shift_up_rewrites_moved_references_and_invalidates_deleted_targets() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A3", 3.0).unwrap();
    engine.set_cell_formula("Sheet1", "B1", "=A3").unwrap();
    engine.set_cell_formula("Sheet1", "B2", "=A2").unwrap();

    engine
        .apply_operation(EditOp::DeleteCellsShiftUp {
            sheet: "Sheet1".to_string(),
            range: range("A1:A2"),
        })
        .unwrap();

    // A3 moved up to A1; B1 should follow that move.
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(3.0));
    assert_eq!(engine.get_cell_formula("Sheet1", "B1"), Some("=A1"));

    // Reference directly into deleted region becomes #REF!
    assert_eq!(engine.get_cell_formula("Sheet1", "B2"), Some("=#REF!"));
}

#[test]
fn insert_row_updates_spill_operator_references() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "D1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "E1", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "D2", 3.0).unwrap();
    engine.set_cell_value("Sheet1", "E2", 4.0).unwrap();
    engine.set_cell_formula("Sheet1", "A1", "=D1:E2").unwrap();
    engine.set_cell_formula("Sheet1", "G1", "=A1#").unwrap();

    engine
        .apply_operation(EditOp::InsertRows {
            sheet: "Sheet1".to_string(),
            row: 0,
            count: 1,
        })
        .unwrap();

    // Both the spill origin and the referencing formula should move down and update.
    assert_eq!(engine.get_cell_formula("Sheet1", "A2"), Some("=D2:E3"));
    assert_eq!(engine.get_cell_formula("Sheet1", "G2"), Some("=A2#"));

    engine.recalculate();
    assert_eq!(engine.get_cell_value("Sheet1", "G2"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "H3"), Value::Number(4.0));
}

#[test]
fn delete_column_converts_spill_operator_reference_to_ref_error() {
    let mut engine = Engine::new();
    engine.set_cell_formula("Sheet1", "B1", "=A1#").unwrap();

    engine
        .apply_operation(EditOp::DeleteCols {
            sheet: "Sheet1".to_string(),
            col: 0,
            count: 1,
        })
        .unwrap();

    // B1 shifted left to A1, and the referenced A column was deleted.
    assert_eq!(engine.get_cell_formula("Sheet1", "A1"), Some("=#REF!"));
}

#[test]
fn structural_edits_update_named_range_definitions() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine
        .define_name(
            "MyX",
            NameScope::Workbook,
            NameDefinition::Reference("Sheet1!A1".to_string()),
        )
        .unwrap();
    engine.set_cell_formula("Sheet1", "B1", "=MyX").unwrap();

    engine
        .apply_operation(EditOp::InsertRows {
            sheet: "Sheet1".to_string(),
            row: 0,
            count: 1,
        })
        .unwrap();

    assert_eq!(
        engine.get_name("MyX", NameScope::Workbook).cloned(),
        Some(NameDefinition::Reference("Sheet1!A2".to_string()))
    );

    // The formula cell moved to B2 and should still evaluate to the moved value once recalculated.
    assert_eq!(engine.get_cell_formula("Sheet1", "B2"), Some("=MyX"));
    engine.recalculate();
    assert_eq!(engine.get_cell_value("Sheet1", "B2"), Value::Number(1.0));
}

#[test]
fn move_range_updates_named_range_definitions() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 42.0).unwrap();
    engine
        .define_name(
            "MyX",
            NameScope::Workbook,
            NameDefinition::Reference("Sheet1!A1".to_string()),
        )
        .unwrap();
    engine.set_cell_formula("Sheet1", "C1", "=MyX").unwrap();

    engine
        .apply_operation(EditOp::MoveRange {
            sheet: "Sheet1".to_string(),
            src: range("A1:A1"),
            dst_top_left: cell("A2"),
        })
        .unwrap();

    assert_eq!(
        engine.get_name("MyX", NameScope::Workbook).cloned(),
        Some(NameDefinition::Reference("Sheet1!A2".to_string()))
    );
    engine.recalculate();
    assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::Number(42.0));
}
