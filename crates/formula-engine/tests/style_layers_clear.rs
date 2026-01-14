use formula_engine::value::Value;
use formula_engine::{EditOp, Engine};
use formula_model::{CellRef, Protection, Range, Style};

fn unlocked_style_id(engine: &mut Engine) -> u32 {
    engine.intern_style(Style {
        protection: Some(Protection {
            locked: false,
            hidden: false,
        }),
        ..Style::default()
    })
}

fn protect_value(engine: &mut Engine) -> Value {
    engine.recalculate();
    engine.get_cell_value("Sheet1", "B1")
}

#[test]
fn sheet_row_col_style_layers_clear_with_zero() {
    let mut engine = Engine::new();
    let unlocked = unlocked_style_id(&mut engine);

    engine
        .set_cell_formula("Sheet1", "B1", r#"=CELL("protect",A1)"#)
        .expect("set formula");

    // Sheet default style layer.
    engine.set_sheet_default_style_id("Sheet1", Some(unlocked));
    assert_eq!(protect_value(&mut engine), Value::Number(0.0));
    engine.set_sheet_default_style_id("Sheet1", Some(0));
    assert_eq!(protect_value(&mut engine), Value::Number(1.0));

    // Row style layer.
    engine.set_row_style_id("Sheet1", 0, Some(unlocked));
    assert_eq!(protect_value(&mut engine), Value::Number(0.0));
    engine.set_row_style_id("Sheet1", 0, Some(0));
    assert_eq!(protect_value(&mut engine), Value::Number(1.0));

    // Column style layer.
    engine.set_col_style_id("Sheet1", 0, Some(unlocked));
    assert_eq!(protect_value(&mut engine), Value::Number(0.0));
    engine.set_col_style_id("Sheet1", 0, Some(0));
    assert_eq!(protect_value(&mut engine), Value::Number(1.0));
}

#[test]
fn sheet_default_style_persists_across_structural_edits() {
    let mut engine = Engine::new();
    let unlocked = unlocked_style_id(&mut engine);

    // Use an implicit reference so the formula text does not need to be rewritten when cells move.
    engine
        .set_cell_formula("Sheet1", "A1", r#"=CELL("protect")"#)
        .expect("set formula");

    engine.set_sheet_default_style_id("Sheet1", Some(unlocked));
    engine.recalculate();
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(0.0));

    // Structural edits should not clear the sheet default style layer.
    engine
        .apply_operation(EditOp::InsertCols {
            sheet: "Sheet1".to_string(),
            col: 0,
            count: 1,
        })
        .expect("insert cols");
    engine.recalculate();
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(0.0));

    engine
        .apply_operation(EditOp::DeleteCols {
            sheet: "Sheet1".to_string(),
            col: 0,
            count: 1,
        })
        .expect("delete cols");
    engine.recalculate();
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(0.0));

    engine
        .apply_operation(EditOp::InsertRows {
            sheet: "Sheet1".to_string(),
            row: 0,
            count: 1,
        })
        .expect("insert rows");
    engine.recalculate();
    assert_eq!(engine.get_cell_value("Sheet1", "A2"), Value::Number(0.0));

    engine
        .apply_operation(EditOp::DeleteRows {
            sheet: "Sheet1".to_string(),
            row: 0,
            count: 1,
        })
        .expect("delete rows");
    engine.recalculate();
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(0.0));

    // Range-map edits (insert/delete cells) should also preserve the sheet default style layer.
    let a1 = Range::new(CellRef::new(0, 0), CellRef::new(0, 0));

    engine
        .apply_operation(EditOp::InsertCellsShiftRight {
            sheet: "Sheet1".to_string(),
            range: a1,
        })
        .expect("insert cells shift right");
    engine.recalculate();
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(0.0));

    engine
        .apply_operation(EditOp::DeleteCellsShiftLeft {
            sheet: "Sheet1".to_string(),
            range: a1,
        })
        .expect("delete cells shift left");
    engine.recalculate();
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(0.0));

    engine
        .apply_operation(EditOp::InsertCellsShiftDown {
            sheet: "Sheet1".to_string(),
            range: a1,
        })
        .expect("insert cells shift down");
    engine.recalculate();
    assert_eq!(engine.get_cell_value("Sheet1", "A2"), Value::Number(0.0));

    engine
        .apply_operation(EditOp::DeleteCellsShiftUp {
            sheet: "Sheet1".to_string(),
            range: a1,
        })
        .expect("delete cells shift up");
    engine.recalculate();
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(0.0));
}
