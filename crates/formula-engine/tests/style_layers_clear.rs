use formula_engine::Engine;
use formula_engine::value::Value;
use formula_model::{Protection, Style};

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
