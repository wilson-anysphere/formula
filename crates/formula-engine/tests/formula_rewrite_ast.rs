use formula_engine::editing::rewrite::{
    rewrite_formula_for_copy_delta, rewrite_formula_for_structural_edit, StructuralEdit,
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

    let (out, changed) =
        rewrite_formula_for_structural_edit("=$A$1+A1", "Sheet1", origin, &edit);

    assert!(changed);
    assert_eq!(out, "=$A$2+A2");
}

#[test]
fn copy_range_delta_updates_relative_but_not_absolute_refs() {
    let origin = CellAddr::new(2, 2);
    let (out, changed) =
        rewrite_formula_for_copy_delta("=$A$1+A1", "Sheet1", origin, 1, 1);

    assert!(changed);
    assert_eq!(out, "=$A$1+B2");
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

    let (out, changed) = rewrite_formula_for_structural_edit("=$A$1+A1+$A1+A$1", "Sheet1", origin, &edit);

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
