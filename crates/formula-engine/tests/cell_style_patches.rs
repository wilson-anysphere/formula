use formula_engine::style_patch::{AlignmentPatch, ProtectionPatch, StylePatch};
use formula_engine::{Engine, Value};
use formula_model::HorizontalAlignment;

#[test]
fn cell_prefix_respects_explicit_alignment_null_clear() {
    let mut engine = Engine::new();

    // Column A: alignment center.
    engine.set_style_patch(
        1,
        StylePatch {
            alignment: Some(AlignmentPatch {
                horizontal: Some(Some(HorizontalAlignment::Center)),
            }),
            ..StylePatch::default()
        },
    );
    engine.set_col_patch_style_id("Sheet1", 0, 1);

    // A1: explicitly clear `alignment.horizontal` to remove inherited column alignment.
    engine.set_style_patch(
        2,
        StylePatch {
            alignment: Some(AlignmentPatch {
                horizontal: Some(None),
            }),
            ..StylePatch::default()
        },
    );
    engine.set_cell_patch_style_id("Sheet1", "A1", 2).unwrap();

    engine
        .set_cell_formula("Sheet1", "B1", r#"=CELL("prefix",A1)"#)
        .unwrap();
    engine.recalculate_single_threaded();

    assert_eq!(
        engine.get_cell_value("Sheet1", "B1"),
        Value::Text(String::new())
    );
}

#[test]
fn effective_number_format_respects_explicit_null_clear() {
    let mut engine = Engine::new();

    // Column A: apply a number format.
    engine.set_style_patch(
        1,
        StylePatch {
            number_format: Some(Some("0.00".to_string())),
            ..StylePatch::default()
        },
    );
    engine.set_col_patch_style_id("Sheet1", 0, 1);

    // A1: explicitly clear numberFormat.
    engine.set_style_patch(
        2,
        StylePatch {
            number_format: Some(None),
            ..StylePatch::default()
        },
    );
    engine.set_cell_patch_style_id("Sheet1", "A1", 2).unwrap();

    let style = engine.effective_cell_style("Sheet1", "A1").unwrap();
    assert!(style.number_format.is_none());
}

#[test]
fn cell_protect_respects_layered_locked_overrides() {
    let mut engine = Engine::new();

    // Row 1: unlocked.
    engine.set_style_patch(
        1,
        StylePatch {
            protection: Some(ProtectionPatch {
                locked: Some(Some(false)),
            }),
            ..StylePatch::default()
        },
    );
    engine.set_row_patch_style_id("Sheet1", 0, 1);

    // A1: locked (overrides row-level unlocked).
    engine.set_style_patch(
        2,
        StylePatch {
            protection: Some(ProtectionPatch {
                locked: Some(Some(true)),
            }),
            ..StylePatch::default()
        },
    );
    engine.set_cell_patch_style_id("Sheet1", "A1", 2).unwrap();

    engine
        .set_cell_formula("Sheet1", "B1", r#"=CELL("protect",A1)"#)
        .unwrap();
    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(1.0));
}
