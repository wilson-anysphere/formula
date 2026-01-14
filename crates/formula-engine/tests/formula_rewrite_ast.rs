use formula_engine::editing::rewrite::{
    rewrite_formula_for_copy_delta, rewrite_formula_for_range_map,
    rewrite_formula_for_range_map_with_resolver,
    rewrite_formula_for_sheet_delete,
    rewrite_formula_for_structural_edit, rewrite_formula_for_structural_edit_with_sheet_order_resolver,
    GridRange, RangeMapEdit, StructuralEdit,
};
use formula_engine::CellAddr;
use pretty_assertions::assert_eq;

#[test]
fn insert_row_updates_absolute_and_relative_a1_refs() {
    let edit = StructuralEdit::InsertRows {
        sheet: "Sheet1".to_string(),
        row: 0,
        count: 1,
    };
    let origin = CellAddr::new(2, 1);

    let (out, changed) = rewrite_formula_for_structural_edit("=$A$1+A1", "Sheet1", origin, &edit);

    assert!(changed);
    assert_eq!(out, "=$A$2+A2");
}

#[test]
fn copy_range_delta_updates_relative_but_not_absolute_refs() {
    let origin = CellAddr::new(2, 2);
    let (out, changed) = rewrite_formula_for_copy_delta("=$A$1+A1", "Sheet1", origin, 1, 1);

    assert!(changed);
    assert_eq!(out, "=$A$1+B2");
}

#[test]
fn copy_delta_shifts_external_cell_ref() {
    let origin = CellAddr::new(2, 2);
    let (out, changed) =
        rewrite_formula_for_copy_delta("=[Book.xlsx]Sheet1!A1", "Sheet1", origin, 1, 1);

    assert!(changed);
    assert_eq!(out, "=[Book.xlsx]Sheet1!B2");
}

#[test]
fn copy_delta_preserves_absolute_external_cell_ref() {
    let origin = CellAddr::new(2, 2);
    let (out, changed) =
        rewrite_formula_for_copy_delta("=[Book.xlsx]Sheet1!$A$1", "Sheet1", origin, 5, 5);

    assert!(!changed);
    assert_eq!(out, "=[Book.xlsx]Sheet1!$A$1");
}

#[test]
fn copy_delta_shifts_external_rectangular_range() {
    let origin = CellAddr::new(2, 2);
    let (out, changed) = rewrite_formula_for_copy_delta(
        "=SUM([Book.xlsx]Sheet1!A1:B2)",
        "Sheet1",
        origin,
        1,
        1,
    );

    assert!(changed);
    assert_eq!(out, "=SUM([Book.xlsx]Sheet1!B2:C3)");
}

#[test]
fn copy_delta_shifts_external_row_and_col_ranges() {
    let origin = CellAddr::new(0, 0);

    let (out, changed) =
        rewrite_formula_for_copy_delta("=SUM([Book.xlsx]Sheet1!A:A)", "Sheet1", origin, 0, 1);
    assert!(changed);
    assert_eq!(out, "=SUM([Book.xlsx]Sheet1!B:B)");

    let (out, changed) =
        rewrite_formula_for_copy_delta("=SUM([Book.xlsx]Sheet1!1:1)", "Sheet1", origin, 1, 0);
    assert!(changed);
    assert_eq!(out, "=SUM([Book.xlsx]Sheet1!2:2)");
}

#[test]
fn sheet_qualified_refs_only_rewrite_for_the_target_sheet() {
    let edit = StructuralEdit::InsertRows {
        sheet: "My Sheet".to_string(),
        row: 0,
        count: 1,
    };
    let origin = CellAddr::new(0, 0);

    let (out, changed) =
        rewrite_formula_for_structural_edit("=A1+'My Sheet'!A1", "Other", origin, &edit);

    assert!(changed);
    assert_eq!(out, "=A1+'My Sheet'!A2");
}

#[test]
fn insert_cols_updates_column_ranges() {
    let edit = StructuralEdit::InsertCols {
        sheet: "Sheet1".to_string(),
        col: 0,
        count: 1,
    };
    let origin = CellAddr::new(0, 1);

    let (out, changed) = rewrite_formula_for_structural_edit("=SUM(A:A)", "Sheet1", origin, &edit);

    assert!(changed);
    assert_eq!(out, "=SUM(B:B)");
}

#[test]
fn insert_rows_updates_row_ranges() {
    let edit = StructuralEdit::InsertRows {
        sheet: "Sheet1".to_string(),
        row: 0,
        count: 1,
    };
    let origin = CellAddr::new(1, 0);

    let (out, changed) = rewrite_formula_for_structural_edit("=SUM(1:1)", "Sheet1", origin, &edit);

    assert!(changed);
    assert_eq!(out, "=SUM(2:2)");
}

#[test]
fn insert_rows_updates_cell_ranges() {
    let edit = StructuralEdit::InsertRows {
        sheet: "Sheet1".to_string(),
        row: 0,
        count: 1,
    };
    let origin = CellAddr::new(3, 2);

    let (out, changed) =
        rewrite_formula_for_structural_edit("=SUM(A1:B2)", "Sheet1", origin, &edit);

    assert!(changed);
    assert_eq!(out, "=SUM(A2:B3)");
}

#[test]
fn delete_cols_creates_ref_errors_for_deleted_column_refs() {
    let edit = StructuralEdit::DeleteCols {
        sheet: "Sheet1".to_string(),
        col: 0,
        count: 1,
    };
    let origin = CellAddr::new(0, 1);

    let (out, changed) =
        rewrite_formula_for_structural_edit("=$A$1+A1+$A1+A$1", "Sheet1", origin, &edit);

    assert!(changed);
    assert_eq!(out, "=#REF!+#REF!+#REF!+#REF!");
}

#[test]
fn string_literals_that_look_like_refs_are_not_rewritten() {
    let edit = StructuralEdit::InsertRows {
        sheet: "Sheet1".to_string(),
        row: 0,
        count: 1,
    };
    let origin = CellAddr::new(1, 1);

    let (out, changed) = rewrite_formula_for_structural_edit("=\"A1\"&A1", "Sheet1", origin, &edit);

    assert!(changed);
    assert_eq!(out, "=\"A1\"&A2");
}

#[test]
fn structural_edits_rewrite_sheet_range_refs_when_edit_sheet_in_span() {
    let edit = StructuralEdit::InsertRows {
        sheet: "Sheet2".to_string(),
        row: 0,
        count: 1,
    };
    let origin = CellAddr::new(0, 0);

    let (out, changed) = rewrite_formula_for_structural_edit_with_sheet_order_resolver(
        "=SUM(Sheet1:Sheet3!A1)",
        "Summary",
        origin,
        &edit,
        |name| match name {
            "Sheet1" => Some(0),
            "Sheet2" => Some(1),
            "Sheet3" => Some(2),
            "Summary" => Some(3),
            _ => None,
        },
    );

    assert!(changed);
    assert_eq!(out, "=SUM(Sheet1:Sheet3!A2)");
}

#[test]
fn structural_edits_rewrite_reversed_sheet_range_refs_when_edit_sheet_in_span() {
    let edit = StructuralEdit::InsertRows {
        sheet: "Sheet2".to_string(),
        row: 0,
        count: 1,
    };
    let origin = CellAddr::new(0, 0);

    let (out, changed) = rewrite_formula_for_structural_edit_with_sheet_order_resolver(
        "=SUM(Sheet3:Sheet1!A1)",
        "Summary",
        origin,
        &edit,
        |name| match name {
            "Sheet1" => Some(0),
            "Sheet2" => Some(1),
            "Sheet3" => Some(2),
            "Summary" => Some(3),
            _ => None,
        },
    );

    assert!(changed);
    assert_eq!(out, "=SUM(Sheet3:Sheet1!A2)");
}

#[test]
fn structural_edits_use_tab_order_for_sheet_range_membership() {
    // Simulate a workbook where sheet ids are stable but not the same as tab order positions.
    //
    // Tab order is: Sheet2, Sheet3, Sheet1, Summary
    // So the 3D span `Sheet1:Sheet3` includes Sheet3 and Sheet1, but *not* Sheet2.
    let edit = StructuralEdit::InsertRows {
        sheet: "Sheet2".to_string(),
        row: 0,
        count: 1,
    };
    let origin = CellAddr::new(0, 0);

    let (out, changed) = rewrite_formula_for_structural_edit_with_sheet_order_resolver(
        "=SUM(Sheet1:Sheet3!A1)",
        "Summary",
        origin,
        &edit,
        |name| match name {
            "Sheet2" => Some(0),
            "Sheet3" => Some(1),
            "Sheet1" => Some(2),
            "Summary" => Some(3),
            _ => None,
        },
    );

    assert!(!changed);
    assert_eq!(out, "=SUM(Sheet1:Sheet3!A1)");
}

#[test]
fn range_map_edits_rewrite_sheet_range_refs_when_edit_sheet_in_span() {
    let edit = RangeMapEdit {
        sheet: "Sheet2".to_string(),
        moved_region: GridRange::new(0, 0, 0, 0),
        delta_row: 1,
        delta_col: 0,
        deleted_region: None,
    };
    let origin = CellAddr::new(0, 0);

    let (out, changed) = rewrite_formula_for_range_map_with_resolver(
        "=SUM(Sheet1:Sheet3!A1)",
        "Summary",
        origin,
        &edit,
        |name| match name {
            "Sheet1" => Some(0),
            "Sheet2" => Some(1),
            "Sheet3" => Some(2),
            "Summary" => Some(3),
            _ => None,
        },
    );

    assert!(changed);
    assert_eq!(out, "=SUM(Sheet1:Sheet3!A2)");
}

#[test]
fn range_map_edits_use_tab_order_for_sheet_range_membership() {
    // Mirror `structural_edits_use_tab_order_for_sheet_range_membership`, but for move/range-map
    // edits.
    //
    // Tab order is: Sheet2, Sheet3, Sheet1, Summary
    // So the 3D span `Sheet1:Sheet3` includes Sheet3 and Sheet1, but *not* Sheet2.
    let edit = RangeMapEdit {
        sheet: "Sheet2".to_string(),
        moved_region: GridRange::new(0, 0, 0, 0),
        delta_row: 1,
        delta_col: 0,
        deleted_region: None,
    };
    let origin = CellAddr::new(0, 0);

    let (out, changed) = rewrite_formula_for_range_map_with_resolver(
        "=SUM(Sheet1:Sheet3!A1)",
        "Summary",
        origin,
        &edit,
        |name| match name {
            "Sheet2" => Some(0),
            "Sheet3" => Some(1),
            "Sheet1" => Some(2),
            "Summary" => Some(3),
            _ => None,
        },
    );

    assert!(!changed);
    assert_eq!(out, "=SUM(Sheet1:Sheet3!A1)");
}

#[test]
fn structural_edits_match_sheet_names_case_insensitively_across_unicode() {
    // Excel compares sheet names case-insensitively across Unicode (with NFKC normalization).
    // The sharp s (`ß`) uppercases to `SS`, which should be treated as the same sheet name.
    let edit = StructuralEdit::InsertRows {
        sheet: "SS".to_string(),
        row: 0,
        count: 1,
    };
    let origin = CellAddr::new(0, 0);

    let (out, changed) = rewrite_formula_for_structural_edit("=A1+'ß'!A1", "Other", origin, &edit);

    assert!(changed);
    assert_eq!(out, "=A1+'ß'!A2");
}

#[test]
fn structural_edits_match_sheet_names_nfkc_case_insensitively() {
    // Excel applies compatibility normalization (NFKC) when comparing sheet names.
    // U+212A KELVIN SIGN (K) is NFKC-equivalent to ASCII 'K'.
    let edit = StructuralEdit::InsertRows {
        sheet: "KELVIN".to_string(),
        row: 0,
        count: 1,
    };
    let origin = CellAddr::new(0, 0);

    let (out, changed) =
        rewrite_formula_for_structural_edit("='Kelvin'!A1", "Other", origin, &edit);

    assert!(changed);
    assert_eq!(out, "='Kelvin'!A2");
}

#[test]
fn range_map_edits_match_sheet_names_case_insensitively_across_unicode() {
    // Excel compares sheet names case-insensitively across Unicode (with NFKC normalization).
    // The sharp s (`ß`) uppercases to `SS`, which should be treated as the same sheet name.
    let edit = RangeMapEdit {
        sheet: "SS".to_string(),
        moved_region: GridRange::new(0, 0, 0, 0),
        delta_row: 1,
        delta_col: 0,
        deleted_region: None,
    };
    let origin = CellAddr::new(0, 0);

    let (out, changed) = rewrite_formula_for_range_map("='ß'!A1", "Other", origin, &edit);

    assert!(changed);
    assert_eq!(out, "='ß'!A2");
}

#[test]
fn range_map_edits_match_sheet_names_nfkc_case_insensitively() {
    // Excel applies compatibility normalization (NFKC) when comparing sheet names.
    // U+212A KELVIN SIGN (K) is NFKC-equivalent to ASCII 'K'.
    let edit = RangeMapEdit {
        sheet: "KELVIN".to_string(),
        moved_region: GridRange::new(0, 0, 0, 0),
        delta_row: 1,
        delta_col: 0,
        deleted_region: None,
    };
    let origin = CellAddr::new(0, 0);

    let (out, changed) = rewrite_formula_for_range_map("='Kelvin'!A1", "Other", origin, &edit);

    assert!(changed);
    assert_eq!(out, "='Kelvin'!A2");
}

#[test]
fn sheet_delete_matches_sheet_names_case_insensitively_across_unicode() {
    // Deleting a sheet should invalidate references even when the delete request uses a Unicode
    // case-insensitive equivalent name (`ß` uppercases to `SS`).
    let origin = CellAddr::new(0, 0);
    let sheet_order = vec!["ß".to_string(), "Other".to_string()];

    let (out, changed) = rewrite_formula_for_sheet_delete("='ß'!A1", origin, "SS", &sheet_order);

    assert!(changed);
    assert_eq!(out, "=#REF!");
}

#[test]
fn sheet_delete_matches_sheet_names_nfkc_case_insensitively() {
    // Deleting a sheet should respect NFKC equivalence (K == K).
    let origin = CellAddr::new(0, 0);
    let sheet_order = vec!["Kelvin".to_string(), "Other".to_string()];

    let (out, changed) =
        rewrite_formula_for_sheet_delete("='Kelvin'!A1", origin, "KELVIN", &sheet_order);

    assert!(changed);
    assert_eq!(out, "=#REF!");
}

#[test]
fn sheet_delete_shifts_sheet_span_boundaries_using_nfkc_case_insensitive_matching() {
    // When deleting a 3D span boundary, Excel shifts it one sheet inward based on tab order.
    // Boundary matching should respect NFKC equivalence (K == K).
    let origin = CellAddr::new(0, 0);
    let sheet_order = vec![
        "Kelvin".to_string(),
        "Middle".to_string(),
        "Sheet3".to_string(),
        "Summary".to_string(),
    ];

    let (out, changed) = rewrite_formula_for_sheet_delete(
        "=SUM('Kelvin':Sheet3!A1)",
        origin,
        "KELVIN",
        &sheet_order,
    );

    assert!(changed);
    assert_eq!(out, "=SUM(Middle:Sheet3!A1)");
}
