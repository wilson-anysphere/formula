use formula_engine::{EditOp, Engine, Value};
use formula_model::{CellRef, Range};
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
    assert_eq!(engine.get_cell_formula("Sheet1", "B2"), Some("=$A$2+$A2+A$2"));
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

    assert_eq!(engine.get_cell_formula("Other", "A1"), Some("='My Sheet'!A2"));
}
