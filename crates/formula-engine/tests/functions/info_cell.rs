use formula_engine::{ErrorKind, Value};
use formula_model::{Protection, Style};

use super::harness::{assert_number, TestSheet};

use formula_engine::eval::CompiledExpr;
use formula_engine::functions::{
    ArraySupport, FunctionContext, FunctionSpec, ThreadSafety, ValueType, Volatility,
};

fn recalc_tick_test(ctx: &dyn FunctionContext, _args: &[CompiledExpr]) -> Value {
    // Use only 53 bits so the f64 conversion is exact and comparisons remain deterministic.
    //
    // Prefer the low bits via a mask (instead of `>> 11`): it preserves exactness while making
    // collisions across consecutive recalc ticks astronomically unlikely (important because the
    // regression tests compare values for equality/inequality across ticks).
    let bits = ctx.volatile_rand_u64() & ((1u64 << 53) - 1);
    Value::Number(bits as f64)
}

inventory::submit! {
    FunctionSpec {
        name: "RECALC_TICK_TEST",
        min_args: 0,
        max_args: 0,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[],
        implementation: recalc_tick_test,
    }
}

#[test]
fn cell_address_row_and_col() {
    let mut sheet = TestSheet::new();

    assert_eq!(
        sheet.eval("=CELL(\"address\",A1)"),
        Value::Text("$A$1".to_string())
    );
    assert_number(&sheet.eval("=CELL(\"row\",A10)"), 10.0);
    assert_number(&sheet.eval("=CELL(\"col\",C1)"), 3.0);
}

#[test]
fn cell_format_reads_number_format_from_style_table() {
    use formula_engine::Engine;
    use formula_model::Style;

    let mut engine = Engine::new();
    let style_id = engine.intern_style(Style {
        number_format: Some("0.00".to_string()),
        ..Style::default()
    });
    engine
        .set_cell_style_id("Sheet1", "A1", style_id)
        .expect("set style id");
    engine
        .set_cell_formula("Sheet1", "B1", "=CELL(\"format\",A1)")
        .expect("set formula");
    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Sheet1", "B1"),
        Value::Text("F2".to_string())
    );
}

#[test]
fn cell_type_codes_match_excel() {
    let mut sheet = TestSheet::new();

    // Blank.
    sheet.set("A1", Value::Blank);
    assert_eq!(
        sheet.eval("=CELL(\"type\",A1)"),
        Value::Text("b".to_string())
    );

    // Number.
    sheet.set("A1", 1.0);
    assert_eq!(
        sheet.eval("=CELL(\"type\",A1)"),
        Value::Text("v".to_string())
    );

    // Text.
    sheet.set("A1", "x");
    assert_eq!(
        sheet.eval("=CELL(\"type\",A1)"),
        Value::Text("l".to_string())
    );
}

#[test]
fn cell_contents_returns_formula_text_or_value() {
    let mut sheet = TestSheet::new();

    sheet.set("A1", 5.0);
    assert_number(&sheet.eval("=CELL(\"contents\",A1)"), 5.0);

    sheet.set_formula("A1", "=1+1");
    assert_eq!(
        sheet.eval("=CELL(\"contents\",A1)"),
        Value::Text("=1+1".to_string())
    );
}

#[test]
fn cell_format_color_and_parentheses_default_style() {
    let mut sheet = TestSheet::new();

    assert_eq!(
        sheet.eval("=CELL(\"format\",A1)"),
        Value::Text("G".to_string())
    );
    assert_number(&sheet.eval("=CELL(\"color\",A1)"), 0.0);
    assert_number(&sheet.eval("=CELL(\"parentheses\",A1)"), 0.0);
}

#[test]
fn cell_format_reflects_explicit_cell_style() {
    use formula_engine::Engine;

    let mut engine = Engine::new();
    let style_id = engine.intern_style(Style {
        number_format: Some("0.00".into()),
        ..Default::default()
    });
    engine.set_cell_style_id("Sheet1", "A1", style_id).unwrap();
    engine
        .set_cell_formula("Sheet1", "B1", "=CELL(\"format\",A1)")
        .unwrap();
    engine.recalculate_single_threaded();

    assert_eq!(
        engine.get_cell_value("Sheet1", "B1"),
        Value::Text("F2".to_string())
    );
}

#[test]
fn cell_color_and_parentheses_reflect_explicit_cell_style() {
    use formula_engine::Engine;

    let mut engine = Engine::new();
    let style_id = engine.intern_style(Style {
        number_format: Some("0;[Red](0)".into()),
        ..Default::default()
    });
    engine.set_cell_style_id("Sheet1", "A1", style_id).unwrap();
    engine
        .set_cell_formula("Sheet1", "B1", "=CELL(\"color\",A1)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "C1", "=CELL(\"parentheses\",A1)")
        .unwrap();
    engine.recalculate_single_threaded();

    assert_number(&engine.get_cell_value("Sheet1", "B1"), 1.0);
    assert_number(&engine.get_cell_value("Sheet1", "C1"), 1.0);
}

#[test]
fn cell_format_and_flags_handle_builtin_placeholder_styles() {
    use formula_engine::Engine;

    let mut engine = Engine::new();
    let style_id = engine.intern_style(Style {
        number_format: Some("__builtin_numFmtId:6".into()),
        ..Default::default()
    });
    engine.set_cell_style_id("Sheet1", "A1", style_id).unwrap();
    engine
        .set_cell_formula("Sheet1", "B1", "=CELL(\"format\",A1)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "C1", "=CELL(\"color\",A1)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "D1", "=CELL(\"parentheses\",A1)")
        .unwrap();
    engine.recalculate_single_threaded();

    let fmt = match engine.get_cell_value("Sheet1", "B1") {
        Value::Text(s) => s,
        other => panic!("expected text for CELL(\"format\"), got {other:?}"),
    };
    assert!(
        fmt.starts_with('C'),
        "expected currency CELL(\"format\") code starting with 'C', got {fmt:?}"
    );
    assert_number(&engine.get_cell_value("Sheet1", "C1"), 1.0);
    assert_number(&engine.get_cell_value("Sheet1", "D1"), 1.0);
}

#[test]
fn cell_format_uses_cell_row_col_style_precedence() {
    use formula_engine::Engine;

    let mut engine = Engine::new();
    let col_style = engine.intern_style(Style {
        number_format: Some("0.00".into()),
        ..Default::default()
    });
    let row_style = engine.intern_style(Style {
        number_format: Some("0.000".into()),
        ..Default::default()
    });
    let cell_style = engine.intern_style(Style {
        number_format: Some("0.0".into()),
        ..Default::default()
    });

    // With no explicit cell or row style, fall back to the column style.
    engine.set_col_style_id("Sheet1", 0, Some(col_style));
    engine
        .set_cell_formula("Sheet1", "B1", "=CELL(\"format\",A1)")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Sheet1", "B1"),
        Value::Text("F2".to_string())
    );

    // Row style overrides column style.
    engine.set_row_style_id("Sheet1", 0, Some(row_style));
    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Sheet1", "B1"),
        Value::Text("F3".to_string())
    );

    // Explicit non-zero cell style overrides row/column styles.
    engine
        .set_cell_style_id("Sheet1", "A1", cell_style)
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Sheet1", "B1"),
        Value::Text("F1".to_string())
    );
}

#[test]
fn cell_format_uses_cell_row_col_sheet_style_precedence() {
    use formula_engine::Engine;

    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();

    // Define one style per layer so we can validate precedence:
    // sheet < col < row < cell.
    let sheet_style = engine.intern_style(Style {
        number_format: Some("0.0000".into()),
        ..Default::default()
    });
    let col_style = engine.intern_style(Style {
        number_format: Some("0.00".into()),
        ..Default::default()
    });
    let row_style = engine.intern_style(Style {
        number_format: Some("0.000".into()),
        ..Default::default()
    });
    let cell_style = engine.intern_style(Style {
        number_format: Some("0.0".into()),
        ..Default::default()
    });

    engine
        .set_cell_formula("Sheet1", "B1", "=CELL(\"format\",A1)")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Sheet1", "B1"),
        Value::Text("G".to_string())
    );

    // Sheet default is the fallback when there are no row/col/cell overrides.
    engine.set_sheet_default_style_id("Sheet1", Some(sheet_style));
    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Sheet1", "B1"),
        Value::Text("F4".to_string())
    );

    // Column overrides sheet default.
    engine.set_col_style_id("Sheet1", 0, Some(col_style));
    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Sheet1", "B1"),
        Value::Text("F2".to_string())
    );

    // Row overrides column (and sheet).
    engine.set_row_style_id("Sheet1", 0, Some(row_style));
    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Sheet1", "B1"),
        Value::Text("F3".to_string())
    );

    // Explicit cell style overrides row/column/sheet.
    engine
        .set_cell_style_id("Sheet1", "A1", cell_style)
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Sheet1", "B1"),
        Value::Text("F1".to_string())
    );
}

#[test]
fn cell_format_classifies_thousands_separated_numbers_as_n() {
    use formula_engine::Engine;

    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine
        .set_cell_formula("Sheet1", "B1", "=CELL(\"format\",A1)")
        .unwrap();

    // Explicit grouping format code (no decimals).
    let grouped0 = engine.intern_style(Style {
        number_format: Some("#,##0".to_string()),
        ..Style::default()
    });
    engine.set_cell_style_id("Sheet1", "A1", grouped0).unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Sheet1", "B1"),
        Value::Text("N0".to_string())
    );

    // Built-in placeholder variant (id 3 = `#,##0`).
    let grouped0_builtin = engine.intern_style(Style {
        number_format: Some("__builtin_numFmtId:3".to_string()),
        ..Style::default()
    });
    engine
        .set_cell_style_id("Sheet1", "A1", grouped0_builtin)
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Sheet1", "B1"),
        Value::Text("N0".to_string())
    );

    // Explicit grouping format code.
    let grouped = engine.intern_style(Style {
        number_format: Some("#,##0.00".to_string()),
        ..Style::default()
    });
    engine.set_cell_style_id("Sheet1", "A1", grouped).unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Sheet1", "B1"),
        Value::Text("N2".to_string())
    );

    // Built-in placeholder variant (id 4 = `#,##0.00`).
    let grouped_builtin = engine.intern_style(Style {
        number_format: Some("__builtin_numFmtId:4".to_string()),
        ..Style::default()
    });
    engine
        .set_cell_style_id("Sheet1", "A1", grouped_builtin)
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Sheet1", "B1"),
        Value::Text("N2".to_string())
    );
}

#[test]
fn cell_width_reflects_column_width_metadata() {
    use formula_engine::Engine;

    let mut engine = Engine::new();
    engine
        // Put the formula in a different cell to avoid creating a self-reference cycle.
        .set_cell_formula("Sheet1", "B1", "=CELL(\"width\",A1)")
        .unwrap();
    engine.recalculate_single_threaded();
    // Default Excel column width is 8.43 characters, but CELL("width") returns the rounded
    // integer component (plus 0.1 when the width is custom).
    assert_number(&engine.get_cell_value("Sheet1", "B1"), 8.0);

    // Update the column width metadata and ensure the formula result updates on recalc.
    engine.set_col_width("Sheet1", 0, Some(16.42578125));
    engine.recalculate_single_threaded();
    assert_number(&engine.get_cell_value("Sheet1", "B1"), 16.1);
}

#[test]
fn cell_width_self_reference_is_not_a_circular_dependency() {
    // `CELL("width", A1)` only depends on column metadata; the `A1` reference is used for its
    // address (column) only. Excel evaluates this successfully even when the formula is entered
    // into the referenced cell.
    let mut sheet = TestSheet::new();
    sheet.set_formula("A1", "=CELL(\"width\",A1)");
    sheet.recalculate();
    assert_number(&sheet.get("A1"), 8.0);
    assert_eq!(sheet.circular_reference_count(), 0);
}

#[test]
fn cell_width_offset_self_reference_is_not_a_circular_dependency() {
    // Similar to `CELL("width", A1)`, but with a reference-returning function. The width depends on
    // the *address* of the returned reference, not the referenced cell's value, so it should not be
    // treated as a calc-graph self-edge.
    let mut sheet = TestSheet::new();
    sheet.set_formula("A1", "=CELL(\"width\", OFFSET(A1, 0, 0))");
    sheet.recalculate();
    assert_number(&sheet.get("A1"), 8.0);
    assert_eq!(sheet.circular_reference_count(), 0);
}

#[test]
fn cell_width_omitted_reference_uses_current_cell_and_is_not_circular() {
    // Excel allows `CELL("width")` without the optional reference argument; it should use the
    // current cell as the implicit reference without introducing a self-edge.
    let mut sheet = TestSheet::new();
    sheet.set_formula("A1", "=CELL(\"width\")");
    sheet.recalculate();

    assert_number(&sheet.get("A1"), 8.0);
    assert_eq!(sheet.circular_reference_count(), 0);

    // Column metadata edits should affect the result on the next recalculation.
    sheet.set_col_width(0, Some(25.0));
    sheet.recalculate();
    assert_number(&sheet.get("A1"), 25.1);
    assert_eq!(sheet.circular_reference_count(), 0);
}

#[test]
fn cell_metadata_keys_return_ref_for_out_of_bounds_reference() {
    use formula_engine::Engine;

    // Restrict the sheet to only column A; reference column B should be out-of-bounds.
    let mut engine = Engine::new();
    engine.set_sheet_dimensions("Sheet1", 6, 1).unwrap(); // rows 1..=6, cols A only

    engine
        .set_cell_formula("Sheet1", "A1", "=CELL(\"protect\",B1)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A2", "=CELL(\"prefix\",B1)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A3", "=CELL(\"width\",B1)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A4", "=CELL(\"format\",B1)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A5", "=CELL(\"color\",B1)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A6", "=CELL(\"parentheses\",B1)")
        .unwrap();

    engine.recalculate_single_threaded();

    assert_eq!(
        engine.get_cell_value("Sheet1", "A1"),
        Value::Error(ErrorKind::Ref)
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "A2"),
        Value::Error(ErrorKind::Ref)
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "A3"),
        Value::Error(ErrorKind::Ref)
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "A4"),
        Value::Error(ErrorKind::Ref)
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "A5"),
        Value::Error(ErrorKind::Ref)
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "A6"),
        Value::Error(ErrorKind::Ref)
    );
}

#[test]
fn cell_width_name_ref_self_reference_is_not_a_circular_dependency() {
    use formula_engine::{Engine, NameDefinition, NameScope};

    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=CELL(\"width\",X)")
        .unwrap();
    engine
        .define_name(
            "X",
            NameScope::Workbook,
            NameDefinition::Reference("Sheet1!A1".to_string()),
        )
        .unwrap();
    engine.recalculate_single_threaded();
    assert_number(&engine.get_cell_value("Sheet1", "A1"), 8.0);
    assert_eq!(engine.circular_reference_count(), 0);
}

#[test]
fn cell_sheet_default_style_affects_format_prefix_and_protect() {
    use formula_engine::Engine;
    use formula_model::{Alignment, HorizontalAlignment, Protection};

    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();

    let sheet_style = engine.intern_style(Style {
        number_format: Some("0.00".to_string()),
        alignment: Some(Alignment {
            horizontal: Some(HorizontalAlignment::Left),
            ..Default::default()
        }),
        protection: Some(Protection {
            locked: false,
            hidden: false,
        }),
        ..Default::default()
    });
    engine.set_sheet_default_style_id("Sheet1", Some(sheet_style));

    engine
        .set_cell_formula("Sheet1", "B1", "=CELL(\"format\",A1)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B2", "=CELL(\"prefix\",A1)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B3", "=CELL(\"protect\",A1)")
        .unwrap();

    engine.recalculate_single_threaded();

    assert_eq!(
        engine.get_cell_value("Sheet1", "B1"),
        Value::Text("F2".to_string())
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "B2"),
        Value::Text("'".to_string())
    );
    assert_number(&engine.get_cell_value("Sheet1", "B3"), 0.0);
}

#[test]
fn cell_width_defaults_to_excel_standard_width() {
    let mut sheet = TestSheet::new();
    // Default Excel column width is 8.43 characters, but CELL("width") reports the rounded
    // integer width (plus a custom-width marker in the first decimal place).
    assert_number(&sheet.eval("=CELL(\"width\",A1)"), 8.0);
}

#[test]
fn cell_width_uses_sheet_default_width_when_present() {
    let mut sheet = TestSheet::new();
    sheet.set_default_col_width(Some(20.0));
    assert_number(&sheet.eval("=CELL(\"width\",A1)"), 20.0);
}

#[test]
fn cell_width_updates_on_sheet_default_width_change_in_automatic_mode() {
    use formula_engine::calc_settings::{CalcSettings, CalculationMode};
    use formula_engine::Engine;

    let mut engine = Engine::new();
    engine.set_calc_settings(CalcSettings {
        calculation_mode: CalculationMode::Automatic,
        ..CalcSettings::default()
    });

    engine
        .set_cell_formula("Sheet1", "B1", "=CELL(\"width\",A1)")
        .unwrap();
    assert_number(&engine.get_cell_value("Sheet1", "B1"), 8.0);

    // Regression guard: changing column width metadata should not force unrelated non-volatile
    // formulas to recalculate.
    engine
        .set_cell_formula("Sheet1", "B2", "=RECALC_TICK_TEST()")
        .unwrap();
    let before_tick = engine.get_cell_value("Sheet1", "B2");

    engine.set_sheet_default_col_width("Sheet1", Some(20.0));
    assert_number(&engine.get_cell_value("Sheet1", "B1"), 20.0);
    assert_eq!(engine.get_cell_value("Sheet1", "B2"), before_tick);

    // Clearing the sheet default should revert to Excel's standard width (8.43 -> 8.0 encoding)
    // without forcing unrelated non-volatile formulas to recalculate.
    engine.set_sheet_default_col_width("Sheet1", None);
    assert_number(&engine.get_cell_value("Sheet1", "B1"), 8.0);
    assert_eq!(engine.get_cell_value("Sheet1", "B2"), before_tick);
}

#[test]
fn cell_width_prefers_per_column_override_and_sets_custom_flag() {
    let mut sheet = TestSheet::new();
    sheet.set_default_col_width(Some(20.0));
    sheet.set_col_width(0, Some(25.0));

    assert_number(&sheet.eval("=CELL(\"width\",A1)"), 25.1);
    assert_number(&sheet.eval("=CELL(\"width\",B1)"), 20.0);
}

#[test]
fn cell_width_sets_custom_flag_even_when_width_equals_sheet_default() {
    // Excel's fractional marker (`.1`) indicates whether the width is an explicit per-column
    // override, *not* whether it differs from the sheet default.
    let mut sheet = TestSheet::new();
    sheet.set_default_col_width(Some(20.0));

    // Explicit override equal to the default should still set the custom-width flag.
    sheet.set_col_width(0, Some(20.0));
    assert_number(&sheet.eval("=CELL(\"width\",A1)"), 20.1);

    // Clearing the override should revert to the sheet default + `.0`.
    sheet.set_col_width(0, None);
    assert_number(&sheet.eval("=CELL(\"width\",A1)"), 20.0);
}

#[test]
fn cell_protect_defaults_to_locked() {
    let mut sheet = TestSheet::new();
    // Excel default: all cells are locked.
    assert_number(&sheet.eval("=CELL(\"protect\",A1)"), 1.0);
}

#[test]
fn cell_protect_inherits_column_locked_false() {
    let mut sheet = TestSheet::new();
    // Unlock the entire column A; A1 inherits it.
    let unlocked = sheet.intern_style(Style {
        protection: Some(Protection {
            locked: false,
            ..Protection::default()
        }),
        ..Style::default()
    });
    sheet.set_col_style_id(0, Some(unlocked));
    assert_number(&sheet.eval("=CELL(\"protect\",A1)"), 0.0);
}

#[test]
fn cell_protect_cell_override_beats_column() {
    let mut sheet = TestSheet::new();
    let unlocked = sheet.intern_style(Style {
        protection: Some(Protection {
            locked: false,
            ..Protection::default()
        }),
        ..Style::default()
    });
    sheet.set_col_style_id(0, Some(unlocked));
    // Explicit cell-level override should win over the unlocked column.
    let locked = sheet.intern_style(Style {
        protection: Some(Protection {
            locked: true,
            ..Protection::default()
        }),
        ..Style::default()
    });
    sheet.set_cell_style_id("A1", locked);
    assert_number(&sheet.eval("=CELL(\"protect\",A1)"), 1.0);
}

#[test]
fn cell_protect_ignores_sheet_protection_enabled() {
    let mut sheet = TestSheet::new();
    let unlocked = sheet.intern_style(Style {
        protection: Some(Protection {
            locked: false,
            ..Protection::default()
        }),
        ..Style::default()
    });
    sheet.set_col_style_id(0, Some(unlocked));
    assert_number(&sheet.eval("=CELL(\"protect\",A1)"), 0.0);

    // Enabling/disabling sheet protection does not change CELL("protect"): it reports formatting
    // state, not whether the protection is enforced.
    sheet.set_sheet_protection_enabled(true);
    assert_number(&sheet.eval("=CELL(\"protect\",A1)"), 0.0);
    sheet.set_sheet_protection_enabled(false);
    assert_number(&sheet.eval("=CELL(\"protect\",A1)"), 0.0);
}

#[test]
fn cell_prefix_alignment_codes_match_excel() {
    use formula_engine::Engine;
    use formula_model::{Alignment, HorizontalAlignment};

    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", "x").unwrap();
    engine
        .set_cell_formula("Sheet1", "B1", "=CELL(\"prefix\",A1)")
        .unwrap();

    // Default alignment ("General") returns the empty string.
    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Sheet1", "B1"),
        Value::Text(String::new())
    );

    for (alignment, expected) in [
        (HorizontalAlignment::Left, "'"),
        (HorizontalAlignment::Right, "\""),
        (HorizontalAlignment::Center, "^"),
        (HorizontalAlignment::Fill, "\\"),
        // Excel returns "" for other alignments, including Justify.
        (HorizontalAlignment::Justify, ""),
    ] {
        let style_id = engine.intern_style(Style {
            alignment: Some(Alignment {
                horizontal: Some(alignment),
                ..Default::default()
            }),
            ..Style::default()
        });
        engine.set_cell_style_id("Sheet1", "A1", style_id).unwrap();
        engine.recalculate_single_threaded();
        assert_eq!(
            engine.get_cell_value("Sheet1", "B1"),
            Value::Text(expected.to_string()),
            "alignment={alignment:?}"
        );
    }
}

#[test]
fn cell_prefix_general_is_empty_for_text_and_number() {
    use formula_engine::Engine;

    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine
        .set_cell_formula("Sheet1", "B1", "=CELL(\"prefix\",A1)")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Sheet1", "B1"),
        Value::Text(String::new())
    );

    engine.set_cell_value("Sheet1", "A1", "x").unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Sheet1", "B1"),
        Value::Text(String::new())
    );
}

#[test]
fn cell_prefix_respects_layered_alignment_and_explicit_clears() {
    use formula_engine::Engine;
    use formula_model::{Alignment, HorizontalAlignment};

    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", "x").unwrap();
    engine
        .set_cell_formula("Sheet1", "B1", "=CELL(\"prefix\",A1)")
        .unwrap();

    let style_left = engine.intern_style(Style {
        alignment: Some(Alignment {
            horizontal: Some(HorizontalAlignment::Left),
            ..Default::default()
        }),
        ..Style::default()
    });
    let style_right = engine.intern_style(Style {
        alignment: Some(Alignment {
            horizontal: Some(HorizontalAlignment::Right),
            ..Default::default()
        }),
        ..Style::default()
    });
    let style_center = engine.intern_style(Style {
        alignment: Some(Alignment {
            horizontal: Some(HorizontalAlignment::Center),
            ..Default::default()
        }),
        ..Style::default()
    });
    let style_general = engine.intern_style(Style {
        alignment: Some(Alignment {
            horizontal: Some(HorizontalAlignment::General),
            ..Default::default()
        }),
        ..Style::default()
    });

    // sheet < col < row < range-run < cell precedence.
    engine.set_col_style_id("Sheet1", 0, Some(style_left)); // col A
    engine.set_row_style_id("Sheet1", 0, Some(style_right)); // row 1
    engine.recalculate_single_threaded();
    // Row overrides col.
    assert_eq!(
        engine.get_cell_value("Sheet1", "B1"),
        Value::Text("\"".to_string())
    );

    // Cell override wins.
    engine
        .set_cell_style_id("Sheet1", "A1", style_center)
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Sheet1", "B1"),
        Value::Text("^".to_string())
    );

    // Explicit clear should override inherited formatting and revert to General (empty prefix).
    engine
        .set_cell_style_id("Sheet1", "A1", style_general)
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Sheet1", "B1"),
        Value::Text(String::new())
    );
}

#[test]
fn cell_prefix_respects_range_run_precedence() {
    use formula_engine::metadata::FormatRun;
    use formula_engine::Engine;
    use formula_model::{Alignment, HorizontalAlignment};

    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", "x").unwrap();
    engine
        .set_cell_formula("Sheet1", "B1", "=CELL(\"prefix\",A1)")
        .unwrap();

    let style_right = engine.intern_style(Style {
        alignment: Some(Alignment {
            horizontal: Some(HorizontalAlignment::Right),
            ..Default::default()
        }),
        ..Style::default()
    });
    let style_center = engine.intern_style(Style {
        alignment: Some(Alignment {
            horizontal: Some(HorizontalAlignment::Center),
            ..Default::default()
        }),
        ..Style::default()
    });
    let style_fill = engine.intern_style(Style {
        alignment: Some(Alignment {
            horizontal: Some(HorizontalAlignment::Fill),
            ..Default::default()
        }),
        ..Style::default()
    });
    let style_general = engine.intern_style(Style {
        alignment: Some(Alignment {
            horizontal: Some(HorizontalAlignment::General),
            ..Default::default()
        }),
        ..Style::default()
    });

    engine.set_row_style_id("Sheet1", 0, Some(style_right));
    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Sheet1", "B1"),
        Value::Text("\"".to_string())
    );

    // Range-run overrides row/col.
    engine
        .set_format_runs_by_col(
            "Sheet1",
            0, // col A
            vec![FormatRun {
                start_row: 0,
                end_row_exclusive: 10,
                style_id: style_center,
            }],
        )
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Sheet1", "B1"),
        Value::Text("^".to_string())
    );

    // Cell override wins over range-run.
    engine
        .set_cell_style_id("Sheet1", "A1", style_fill)
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Sheet1", "B1"),
        Value::Text("\\".to_string())
    );

    // Explicitly clearing the cell alignment should not fall back to the range run.
    engine
        .set_cell_style_id("Sheet1", "A1", style_general)
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Sheet1", "B1"),
        Value::Text(String::new())
    );
}

#[test]
fn cell_protect_and_format_respect_range_run_layer() {
    use formula_engine::metadata::FormatRun;
    use formula_engine::Engine;
    use formula_model::{Protection, Style};

    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", "x").unwrap();

    engine
        .set_cell_formula("Sheet1", "B1", "=CELL(\"protect\",A1)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B2", "=CELL(\"format\",A1)")
        .unwrap();
    engine.recalculate_single_threaded();

    // Excel defaults to locked cells and general number format.
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(1.0));

    let run_style_id = engine.intern_style(Style {
        protection: Some(Protection {
            locked: false,
            hidden: false,
        }),
        number_format: Some("__builtin_numFmtId:12".to_string()),
        ..Style::default()
    });

    // Apply range-run formatting (unlocked + number format) for A1.
    engine
        .set_format_runs_by_col(
            "Sheet1",
            0, // col A
            vec![FormatRun {
                start_row: 0,
                end_row_exclusive: 10,
                style_id: run_style_id,
            }],
        )
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(0.0));
    assert_eq!(
        engine.get_cell_value("Sheet1", "B2"),
        Value::Text("N".to_string())
    );

    // Cell style overrides range-run.
    let locked_style_id = engine.intern_style(Style {
        protection: Some(Protection {
            locked: true,
            hidden: false,
        }),
        ..Style::default()
    });
    engine
        .set_cell_style_id("Sheet1", "A1", locked_style_id)
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(1.0));
}

#[test]
fn info_recalc_defaults_to_manual_and_unknown_keys() {
    let mut sheet = TestSheet::new();

    assert_eq!(
        sheet.eval("=INFO(\"recalc\")"),
        // The engine defaults to manual calculation mode; callers can opt into Excel-like
        // automatic calculation via `Engine::set_calc_settings` / `CalcSettings.calculation_mode`.
        Value::Text("Manual".to_string())
    );
    assert_eq!(
        sheet.eval("=INFO(\"no_such_key\")"),
        Value::Error(ErrorKind::Value)
    );
}

#[test]
fn info_recalc_reflects_calc_settings() {
    use formula_engine::calc_settings::{CalcSettings, CalculationMode};
    use formula_engine::Engine;

    let mut engine = Engine::new();
    engine.set_calc_settings(CalcSettings {
        calculation_mode: CalculationMode::Automatic,
        ..CalcSettings::default()
    });
    engine
        .set_cell_formula("Sheet1", "A1", "=INFO(\"recalc\")")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Sheet1", "A1"),
        Value::Text("Automatic".to_string())
    );

    let mut engine = Engine::new();
    engine.set_calc_settings(CalcSettings {
        calculation_mode: CalculationMode::AutomaticNoTable,
        ..CalcSettings::default()
    });
    engine
        .set_cell_formula("Sheet1", "A1", "=INFO(\"recalc\")")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Sheet1", "A1"),
        Value::Text("Automatic except for tables".to_string())
    );

    let mut engine = Engine::new();
    engine.set_calc_settings(CalcSettings {
        calculation_mode: CalculationMode::Manual,
        ..CalcSettings::default()
    });
    engine
        .set_cell_formula("Sheet1", "A1", "=INFO(\"recalc\")")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Sheet1", "A1"),
        Value::Text("Manual".to_string())
    );
}

#[test]
fn info_recalc_refreshes_after_calc_settings_change() {
    use formula_engine::calc_settings::{CalcSettings, CalculationMode};
    use formula_engine::Engine;

    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=INFO(\"recalc\")")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Sheet1", "A1"),
        Value::Text("Manual".to_string())
    );

    // Changing calculation mode should refresh INFO("recalc") on the next recalculation tick,
    // even though no other cell values changed.
    engine.set_calc_settings(CalcSettings {
        calculation_mode: CalculationMode::Automatic,
        ..engine.calc_settings().clone()
    });
    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Sheet1", "A1"),
        Value::Text("Automatic".to_string())
    );
}

#[test]
fn info_recalc_refreshes_after_calc_settings_change_with_dynamic_key() {
    use formula_engine::calc_settings::{CalcSettings, CalculationMode};
    use formula_engine::Engine;

    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", "recalc").unwrap();
    engine
        .set_cell_formula("Sheet1", "B1", "=INFO(A1)")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Sheet1", "B1"),
        Value::Text("Manual".to_string())
    );

    engine.set_calc_settings(CalcSettings {
        calculation_mode: CalculationMode::AutomaticNoTable,
        ..engine.calc_settings().clone()
    });
    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Sheet1", "B1"),
        Value::Text("Automatic except for tables".to_string())
    );
}

#[test]
fn info_and_cell_keys_are_trimmed_and_case_insensitive() {
    let mut sheet = TestSheet::new();

    assert_eq!(
        sheet.eval("=INFO(\" ReCaLc \")"),
        Value::Text("Manual".to_string())
    );
    assert_number(&sheet.eval("=CELL(\" rOw \",A10)"), 10.0);
    assert_number(&sheet.eval("=CELL(\" cOl \",C1)"), 3.0);

    assert_eq!(sheet.eval("=INFO(\"\")"), Value::Error(ErrorKind::Value));
    assert_eq!(
        sheet.eval("=CELL(\" \",A1)"),
        Value::Error(ErrorKind::Value)
    );
}

#[test]
fn info_numfile_counts_sheets() {
    use formula_engine::Engine;

    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet2", "A1", 1.0).unwrap();
    engine
        .set_cell_formula("Sheet1", "B1", "=INFO(\"numfile\")")
        .unwrap();
    engine.recalculate_single_threaded();

    assert_number(&engine.get_cell_value("Sheet1", "B1"), 2.0);
}

#[test]
fn info_exposes_host_provided_metadata() {
    use formula_engine::{Engine, EngineInfo};

    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=INFO(\"system\")")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A2", "=INFO(\"directory\")")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A3", "=INFO(\"osversion\")")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A4", "=INFO(\"release\")")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A5", "=INFO(\"version\")")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A6", "=INFO(\"memavail\")")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A7", "=INFO(\"totmem\")")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A8", "=INFO(\"origin\")")
        .unwrap();
    engine
        .set_cell_formula("Sheet2", "A8", "=INFO(\"origin\")")
        .unwrap();

    // Unset metadata returns Excel `#N/A` for supported-but-unknown keys.
    //
    // `INFO("origin")` is always available and defaults to the top-left cell (`$A$1`).
    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Sheet1", "A1"),
        Value::Text("pcdos".to_string())
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "A2"),
        Value::Error(ErrorKind::NA)
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "A3"),
        Value::Error(ErrorKind::NA)
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "A4"),
        Value::Error(ErrorKind::NA)
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "A5"),
        Value::Error(ErrorKind::NA)
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "A6"),
        Value::Error(ErrorKind::NA)
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "A7"),
        Value::Error(ErrorKind::NA)
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "A8"),
        // Excel defaults to the top-left visible cell ($A$1) when no view origin is provided.
        Value::Text("$A$1".to_string())
    );
    assert_eq!(
        engine.get_cell_value("Sheet2", "A8"),
        Value::Text("$A$1".to_string())
    );

    engine.set_engine_info(EngineInfo {
        system: Some("unix".to_string()),
        directory: Some("/tmp".to_string()),
        osversion: Some("14.2".to_string()),
        release: Some("release-x".to_string()),
        version: Some("v1".to_string()),
        memavail: Some(1234.0),
        totmem: Some(5678.0),
        ..EngineInfo::default()
    });
    engine.set_sheet_origin("Sheet1", Some("$C$3")).unwrap();
    engine.set_sheet_origin("Sheet2", Some("$B$2")).unwrap();
    engine.recalculate_single_threaded();

    assert_eq!(
        engine.get_cell_value("Sheet1", "A1"),
        Value::Text("unix".to_string())
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "A2"),
        Value::Text("/tmp/".to_string())
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "A3"),
        Value::Text("14.2".to_string())
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "A4"),
        Value::Text("release-x".to_string())
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "A5"),
        Value::Text("v1".to_string())
    );
    assert_number(&engine.get_cell_value("Sheet1", "A6"), 1234.0);
    assert_number(&engine.get_cell_value("Sheet1", "A7"), 5678.0);
    assert_eq!(
        engine.get_cell_value("Sheet1", "A8"),
        Value::Text("$C$3".to_string())
    );
    assert_eq!(
        engine.get_cell_value("Sheet2", "A8"),
        Value::Text("$B$2".to_string())
    );
}

#[test]
fn cell_format_color_and_parentheses_reflect_number_format() {
    use formula_engine::Engine;
    use formula_model::Style;

    let mut engine = Engine::new();
    let style_id = engine.intern_style(Style {
        // Fractions are not part of the standard CELL("format") numeric families.
        number_format: Some("__builtin_numFmtId:12".to_string()),
        ..Style::default()
    });
    engine
        .set_cell_style_id("Sheet1", "A1", style_id)
        .expect("set style id");
    engine
        .set_cell_formula("Sheet1", "B1", "=CELL(\"format\",A1)")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Sheet1", "B1"),
        Value::Text("N".to_string())
    );

    // Color/parentheses flags are derived from the explicit negative section.
    let style_id = engine.intern_style(Style {
        number_format: Some("0;[Red](0)".to_string()),
        ..Style::default()
    });
    engine
        .set_cell_style_id("Sheet1", "A2", style_id)
        .expect("set style id");
    engine
        .set_cell_formula("Sheet1", "B2", "=CELL(\"color\",A2)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B3", "=CELL(\"parentheses\",A2)")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_number(&engine.get_cell_value("Sheet1", "B2"), 1.0);
    assert_number(&engine.get_cell_value("Sheet1", "B3"), 1.0);
}

#[test]
fn cell_format_color_and_parentheses_reflect_cell_number_format_override() {
    use formula_engine::Engine;

    let mut engine = Engine::new();
    engine
        .set_cell_number_format("Sheet1", "A1", Some("__builtin_numFmtId:12".to_string()))
        .expect("set number format");
    engine
        .set_cell_formula("Sheet1", "B1", "=CELL(\"format\",A1)")
        .unwrap();

    // Color/parentheses flags are derived from the explicit negative section.
    engine
        .set_cell_number_format("Sheet1", "A2", Some("0;[Red](0)".to_string()))
        .expect("set number format");
    engine
        .set_cell_formula("Sheet1", "B2", "=CELL(\"color\",A2)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B3", "=CELL(\"parentheses\",A2)")
        .unwrap();

    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Sheet1", "B1"),
        Value::Text("N".to_string())
    );
    assert_number(&engine.get_cell_value("Sheet1", "B2"), 1.0);
    assert_number(&engine.get_cell_value("Sheet1", "B3"), 1.0);
}

#[test]
fn cell_format_recalculates_when_cell_number_format_changes() {
    use formula_engine::Engine;

    let mut engine = Engine::new();
    engine
        .set_cell_number_format("Sheet1", "A1", Some("__builtin_numFmtId:12".to_string()))
        .expect("set number format");
    engine
        .set_cell_formula("Sheet1", "B1", "=CELL(\"format\",A1)")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Sheet1", "B1"),
        Value::Text("N".to_string())
    );

    // Change only the format metadata; dependent formulas should refresh on the next recalc.
    engine
        .set_cell_number_format("Sheet1", "A1", None)
        .expect("clear number format");
    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Sheet1", "B1"),
        Value::Text("G".to_string())
    );
}

#[test]
fn cell_format_color_and_parentheses_fallback_to_general_for_external_refs() {
    use formula_engine::Engine;

    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", r#"=CELL("format",[Book.xlsx]Sheet1!A1)"#)
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A2", r#"=CELL("color",[Book.xlsx]Sheet1!A1)"#)
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "A3",
            r#"=CELL("parentheses",[Book.xlsx]Sheet1!A1)"#,
        )
        .unwrap();
    engine.recalculate_single_threaded();

    // The engine does not track number formats for external workbooks, so these should fall back
    // to General semantics.
    assert_eq!(
        engine.get_cell_value("Sheet1", "A1"),
        Value::Text("G".to_string())
    );
    assert_number(&engine.get_cell_value("Sheet1", "A2"), 0.0);
    assert_number(&engine.get_cell_value("Sheet1", "A3"), 0.0);
}

#[test]
fn cell_errors_for_unknown_info_types() {
    let mut sheet = TestSheet::new();

    assert_eq!(
        sheet.eval("=CELL(\"no_such_info_type\",A1)"),
        Value::Error(ErrorKind::Value)
    );
}

#[test]
fn cell_filename_is_empty_for_unsaved_workbooks() {
    let mut sheet = TestSheet::new();

    // Excel returns "" until the workbook has been saved.
    assert_eq!(
        sheet.eval("=CELL(\"filename\")"),
        Value::Text(String::new())
    );

    // Excel returns #N/A until the workbook has been saved.
    assert_eq!(
        sheet.eval("=INFO(\"directory\")"),
        Value::Error(ErrorKind::NA)
    );
}

#[test]
fn cell_filename_is_empty_when_workbook_filename_is_empty_even_if_directory_is_set() {
    use formula_engine::Engine;

    // The engine models "unsaved workbook" as having an unknown filename. Directory metadata alone
    // should not be enough for `CELL("filename")` / `INFO("directory")`.
    let mut engine = Engine::new();
    engine.set_workbook_file_metadata(Some("/dir"), Some(""));
    engine
        .set_cell_formula("Sheet1", "A1", "=CELL(\"filename\")")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A2", "=INFO(\"directory\")")
        .unwrap();
    engine.recalculate_single_threaded();

    assert_eq!(
        engine.get_cell_value("Sheet1", "A1"),
        Value::Text(String::new())
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "A2"),
        Value::Error(ErrorKind::NA)
    );
}

#[test]
fn cell_filename_and_info_directory_use_workbook_file_metadata() {
    use formula_engine::Engine;

    // Windows-like path semantics: preserve host-supplied separators and ensure a trailing `\` is
    // present.
    let mut engine = Engine::new();
    engine.set_workbook_file_metadata(Some(r"C:\Dir"), Some("Book.xlsx"));
    engine
        .set_cell_formula("Sheet1", "A1", "=CELL(\"filename\")")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A2", "=INFO(\"directory\")")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Sheet1", "A1"),
        Value::Text(r"C:\Dir\[Book.xlsx]Sheet1".to_string())
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "A2"),
        Value::Text(r"C:\Dir\".to_string())
    );

    // POSIX-like path semantics: preserve host-supplied separators and ensure a trailing `/` is
    // present.
    let mut engine = Engine::new();
    engine.set_workbook_file_metadata(Some("/dir"), Some("Book.xlsx"));
    engine
        .set_cell_formula("Sheet1", "A1", "=CELL(\"filename\")")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A2", "=INFO(\"directory\")")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Sheet1", "A1"),
        Value::Text("/dir/[Book.xlsx]Sheet1".to_string())
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "A2"),
        Value::Text("/dir/".to_string())
    );
}

#[test]
fn workbook_file_metadata_updates_info_functions_in_automatic_mode() {
    use formula_engine::calc_settings::{CalcSettings, CalculationMode};
    use formula_engine::{Engine, EngineInfo};

    let mut engine = Engine::new();
    engine.set_calc_settings(CalcSettings {
        calculation_mode: CalculationMode::Automatic,
        ..CalcSettings::default()
    });

    engine
        .set_cell_formula("Sheet1", "A1", "=CELL(\"filename\")")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A2", "=INFO(\"directory\")")
        .unwrap();

    // Unsaved workbook defaults.
    assert_eq!(
        engine.get_cell_value("Sheet1", "A1"),
        Value::Text(String::new())
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "A2"),
        Value::Error(ErrorKind::NA)
    );

    // Workbook file metadata should drive both `CELL("filename")` and `INFO("directory")`.
    engine.set_workbook_file_metadata(Some("/dir"), Some("Book.xlsx"));
    assert_eq!(
        engine.get_cell_value("Sheet1", "A1"),
        Value::Text("/dir/[Book.xlsx]Sheet1".to_string())
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "A2"),
        Value::Text("/dir/".to_string())
    );

    // Host overrides should also propagate without requiring full-workbook dirtying.
    engine.set_engine_info(EngineInfo {
        directory: Some("/host".to_string()),
        ..EngineInfo::default()
    });
    assert_eq!(
        engine.get_cell_value("Sheet1", "A2"),
        Value::Text("/host/".to_string())
    );
}

#[test]
fn info_directory_prefers_host_override_over_workbook_file_metadata() {
    use formula_engine::{Engine, EngineInfo};

    let mut engine = Engine::new();
    engine.set_workbook_file_metadata(Some("/workbook"), Some("Book.xlsx"));
    engine.set_engine_info(EngineInfo {
        directory: Some("/host".to_string()),
        ..EngineInfo::default()
    });

    engine
        .set_cell_formula("Sheet1", "A1", "=INFO(\"directory\")")
        .unwrap();
    engine.recalculate_single_threaded();

    // `EngineInfo.directory` should override workbook file metadata when present.
    assert_eq!(
        engine.get_cell_value("Sheet1", "A1"),
        Value::Text("/host/".to_string())
    );
}

#[test]
fn cell_filename_supports_filename_only_metadata() {
    use formula_engine::Engine;

    // When only the filename is known (e.g. in a web environment), include the bracketed file name
    // but omit the directory prefix.
    let mut engine = Engine::new();
    engine.set_workbook_file_metadata(None, Some("Book.xlsx"));
    engine
        .set_cell_formula("Sheet1", "A1", "=CELL(\"filename\")")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A2", "=INFO(\"directory\")")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Sheet1", "A1"),
        Value::Text("[Book.xlsx]Sheet1".to_string())
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "A2"),
        Value::Error(ErrorKind::NA)
    );
}

#[test]
fn workbook_file_metadata_treats_empty_strings_as_unknown() {
    use formula_engine::Engine;

    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=CELL(\"filename\")")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A2", "=INFO(\"directory\")")
        .unwrap();

    engine.set_workbook_file_metadata(Some(""), Some(""));
    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Sheet1", "A1"),
        Value::Text(String::new())
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "A2"),
        Value::Error(ErrorKind::NA)
    );

    // Empty strings should behave like `None`: when only the filename is provided, omit the
    // directory prefix and continue to return #N/A for `INFO("directory")`.
    engine.set_workbook_file_metadata(Some(""), Some("Book.xlsx"));
    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Sheet1", "A1"),
        Value::Text("[Book.xlsx]Sheet1".to_string())
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "A2"),
        Value::Error(ErrorKind::NA)
    );
}

#[test]
fn cell_filename_uses_reference_sheet_name_without_quoting() {
    use formula_engine::Engine;

    let mut engine = Engine::new();
    engine.set_cell_value("Other Sheet", "A1", 1.0).unwrap();
    engine.set_workbook_file_metadata(Some("/dir/"), Some("Book.xlsx"));

    engine
        .set_cell_formula("Sheet1", "A1", "=CELL(\"filename\",'Other Sheet'!A1)")
        .unwrap();
    engine.recalculate_single_threaded();

    assert_eq!(
        engine.get_cell_value("Sheet1", "A1"),
        Value::Text("/dir/[Book.xlsx]Other Sheet".to_string())
    );
}

#[test]
fn cell_filename_uses_referenced_sheet_name_not_current_sheet_name() {
    use formula_engine::Engine;

    let mut engine = Engine::new();
    engine.set_cell_value("Sheet2", "A1", 1.0).unwrap();
    engine.set_workbook_file_metadata(None, Some("Book.xlsx"));

    engine
        .set_cell_formula("Sheet1", "A1", "=CELL(\"filename\",Sheet2!A1)")
        .unwrap();
    engine.recalculate_single_threaded();

    assert_eq!(
        engine.get_cell_value("Sheet1", "A1"),
        Value::Text("[Book.xlsx]Sheet2".to_string())
    );
}

#[test]
fn cell_filename_includes_filename_even_when_directory_is_unknown() {
    use formula_engine::Engine;

    let mut engine = Engine::new();
    engine.set_workbook_file_metadata(None, Some("Book1.xlsx"));

    engine
        .set_cell_formula("Sheet1", "A1", "=CELL(\"filename\")")
        .unwrap();
    engine.recalculate_single_threaded();

    assert_eq!(
        engine.get_cell_value("Sheet1", "A1"),
        Value::Text("[Book1.xlsx]Sheet1".to_string())
    );
}

#[test]
fn cell_filename_includes_workbook_path_when_metadata_is_set() {
    use formula_engine::Engine;

    let mut engine = Engine::new();
    engine.set_workbook_file_metadata(Some(r"C:\tmp\"), Some("Book1.xlsx"));

    engine
        .set_cell_formula("Sheet1", "A1", "=CELL(\"filename\")")
        .unwrap();
    engine.recalculate_single_threaded();

    assert_eq!(
        engine.get_cell_value("Sheet1", "A1"),
        Value::Text(r"C:\tmp\[Book1.xlsx]Sheet1".to_string())
    );
}

#[test]
fn cell_filename_uses_sheet_name_of_reference_argument() {
    use formula_engine::Engine;

    let mut engine = Engine::new();
    engine.set_workbook_file_metadata(Some(r"C:\tmp\"), Some("Book1.xlsx"));
    engine.ensure_sheet("Sheet2");

    engine
        .set_cell_formula("Sheet1", "A1", "=CELL(\"filename\",Sheet2!A1)")
        .unwrap();
    engine.recalculate_single_threaded();

    assert_eq!(
        engine.get_cell_value("Sheet1", "A1"),
        Value::Text(r"C:\tmp\[Book1.xlsx]Sheet2".to_string())
    );
}

#[test]
fn info_directory_returns_directory_when_metadata_is_set() {
    use formula_engine::Engine;

    let mut engine = Engine::new();
    engine.set_workbook_file_metadata(Some(r"C:\tmp\"), Some("Book1.xlsx"));

    engine
        .set_cell_formula("Sheet1", "A1", "=INFO(\"directory\")")
        .unwrap();
    engine.recalculate_single_threaded();

    assert_eq!(
        engine.get_cell_value("Sheet1", "A1"),
        Value::Text(r"C:\tmp\".to_string())
    );
}

#[test]
fn cell_protect_and_prefix_observe_row_col_style_layers() {
    use formula_engine::Engine;
    use formula_model::{Alignment, HorizontalAlignment, Protection, Style};

    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", "x").unwrap();

    let unlocked_style_id = engine.intern_style(Style {
        protection: Some(Protection {
            locked: false,
            hidden: false,
        }),
        ..Style::default()
    });
    let col_right_style_id = engine.intern_style(Style {
        alignment: Some(Alignment {
            horizontal: Some(HorizontalAlignment::Right),
            ..Alignment::default()
        }),
        ..Style::default()
    });
    let row_left_style_id = engine.intern_style(Style {
        alignment: Some(Alignment {
            horizontal: Some(HorizontalAlignment::Left),
            ..Alignment::default()
        }),
        ..Style::default()
    });

    // Row style layer affects CELL("protect").
    engine.set_row_style_id("Sheet1", 0, Some(unlocked_style_id));
    engine
        .set_cell_formula("Sheet1", "B1", "=CELL(\"protect\",A1)")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(0.0));

    // Clearing the row style should revert to Excel's default locked behavior.
    engine.set_row_style_id("Sheet1", 0, None);
    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(1.0));

    // Column style layer affects CELL("prefix") for label/text values.
    engine.set_col_style_id("Sheet1", 0, Some(col_right_style_id));
    engine
        .set_cell_formula("Sheet1", "B2", "=CELL(\"prefix\",A1)")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Sheet1", "B2"),
        Value::Text("\"".to_string())
    );

    // Row style overrides column style when both specify alignment (sheet < col < row < cell).
    engine.set_row_style_id("Sheet1", 0, Some(row_left_style_id));
    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Sheet1", "B2"),
        Value::Text("'".to_string())
    );
}

#[test]
fn info_and_cell_metadata_changes_refresh_after_recalc() {
    use formula_engine::Engine;

    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "B1", "=CELL(\"filename\")")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Sheet1", "B1"),
        Value::Text(String::new())
    );

    engine.set_workbook_file_metadata(Some("/tmp/"), Some("Book1.xlsx"));
    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Sheet1", "B1"),
        Value::Text("/tmp/[Book1.xlsx]Sheet1".to_string())
    );

    engine
        .set_cell_formula("Sheet1", "B2", "=CELL(\"width\")")
        .unwrap();
    engine.recalculate_single_threaded();
    let before = engine.get_cell_value("Sheet1", "B2");

    engine.set_col_width("Sheet1", 1, Some(42.0));
    engine.recalculate_single_threaded();
    let after = engine.get_cell_value("Sheet1", "B2");
    assert_number(&after, 42.1);
    assert_ne!(before, after);
}

#[test]
fn metadata_setters_trigger_auto_recalc_in_automatic_mode() {
    use formula_engine::calc_settings::{CalcSettings, CalculationMode};
    use formula_engine::Engine;

    let mut engine = Engine::new();
    engine.set_calc_settings(CalcSettings {
        calculation_mode: CalculationMode::Automatic,
        ..CalcSettings::default()
    });

    engine
        .set_cell_formula("Sheet1", "B1", "=CELL(\"filename\")")
        .unwrap();
    assert_eq!(
        engine.get_cell_value("Sheet1", "B1"),
        Value::Text(String::new())
    );

    engine
        .set_cell_formula("Sheet1", "B3", "=RECALC_TICK_TEST()")
        .unwrap();
    let tick_before = engine.get_cell_value("Sheet1", "B3");

    // Formatting-dependent CELL() keys should update in automatic mode without forcing unrelated
    // non-volatile formulas to recalculate.
    engine
        .set_cell_formula("Sheet1", "B4", "=CELL(\"format\",A1)")
        .unwrap();
    assert_eq!(
        engine.get_cell_value("Sheet1", "B4"),
        Value::Text("G".to_string())
    );

    // In automatic mode, metadata setters should eagerly recalculate.
    engine.set_workbook_file_metadata(Some("/tmp/"), Some("Book1.xlsx"));
    assert_eq!(
        engine.get_cell_value("Sheet1", "B1"),
        Value::Text("/tmp/[Book1.xlsx]Sheet1".to_string())
    );
    assert_eq!(engine.get_cell_value("Sheet1", "B3"), tick_before);

    engine
        .set_cell_formula("Sheet1", "B2", "=CELL(\"width\")")
        .unwrap();
    let before = engine.get_cell_value("Sheet1", "B2");

    engine.set_col_width("Sheet1", 1, Some(42.0));
    let after = engine.get_cell_value("Sheet1", "B2");
    assert_number(&after, 42.1);
    assert_ne!(before, after);

    // The column width edit should not dirty unrelated non-volatile formulas.
    assert_eq!(engine.get_cell_value("Sheet1", "B3"), tick_before);

    // Same semantics for hidden state.
    engine.set_col_hidden("Sheet1", 1, true);
    assert_number(&engine.get_cell_value("Sheet1", "B2"), 0.0);
    assert_eq!(engine.get_cell_value("Sheet1", "B3"), tick_before);

    // Pure style-table growth should not force recalculation.
    let fmt_style_id = engine.intern_style(Style {
        number_format: Some("0.00".to_string()),
        ..Style::default()
    });
    assert_eq!(engine.get_cell_value("Sheet1", "B3"), tick_before);
    assert_eq!(
        engine.get_cell_value("Sheet1", "B4"),
        Value::Text("G".to_string())
    );

    // Applying style metadata should update CELL() outputs but not dirty unrelated non-volatile
    // formulas.
    engine.set_col_style_id("Sheet1", 0, Some(fmt_style_id));
    assert_eq!(
        engine.get_cell_value("Sheet1", "B4"),
        Value::Text("F2".to_string())
    );
    assert_eq!(engine.get_cell_value("Sheet1", "B3"), tick_before);
}

#[test]
fn cell_implicit_reference_does_not_create_dynamic_dependency_cycles() {
    let mut sheet = TestSheet::new();

    // Including INDIRECT marks the formula as dynamic-deps even though the IF short-circuits
    // and the INDIRECT branch is never evaluated.
    //
    // CELL("contents") with no explicit reference should not record a self-reference as a
    // dynamic precedent; otherwise the engine's dynamic dependency update can introduce a
    // self-edge and force the cell into circular-reference handling.
    let formula = "=IF(FALSE,INDIRECT(\"A1\"),CELL(\"contents\"))";
    assert_eq!(sheet.eval(formula), Value::Text(formula.to_string()));
    assert_eq!(sheet.circular_reference_count(), 0);

    // Same idea, but for CELL("type") which also consults the referenced cell.
    assert_eq!(
        sheet.eval("=IF(FALSE,INDIRECT(\"A1\"),CELL(\"type\"))"),
        Value::Text("v".to_string())
    );
    assert_eq!(sheet.circular_reference_count(), 0);
}

#[test]
fn cell_implicit_reference_does_not_create_dynamic_dependency_cycles_for_metadata_keys() {
    let mut sheet = TestSheet::new();

    // Including INDIRECT marks the formula as dynamic-deps even though the IF short-circuits
    // and the INDIRECT branch is never evaluated.
    //
    // CELL metadata keys should not record an implicit self-reference when `reference` is omitted;
    // otherwise dynamic dependency updates can introduce a self-edge and force the cell into the
    // engine's circular-reference handling.
    match sheet.eval("=IF(FALSE,INDIRECT(\"A1\"),CELL(\"width\"))") {
        Value::Number(n) => assert!(n != 0.0, "expected non-zero width, got {n}"),
        other => panic!("expected number for CELL(\"width\"), got {other:?}"),
    }
    assert_eq!(sheet.circular_reference_count(), 0);

    match sheet.eval("=IF(FALSE,INDIRECT(\"A1\"),CELL(\"protect\"))") {
        Value::Number(n) => assert!(n != 0.0, "expected non-zero protect, got {n}"),
        other => panic!("expected number for CELL(\"protect\"), got {other:?}"),
    }
    assert_eq!(sheet.circular_reference_count(), 0);

    assert_eq!(
        sheet.eval("=IF(FALSE,INDIRECT(\"A1\"),CELL(\"prefix\"))"),
        Value::Text(String::new())
    );
    assert_eq!(sheet.circular_reference_count(), 0);

    // Format metadata keys consult number format style state, but should still avoid implicit
    // self-references for dynamic dependency tracing.
    assert_eq!(
        sheet.eval("=IF(FALSE,INDIRECT(\"A1\"),CELL(\"format\"))"),
        Value::Text("G".to_string())
    );
    assert_eq!(sheet.circular_reference_count(), 0);
    assert_number(
        &sheet.eval("=IF(FALSE,INDIRECT(\"A1\"),CELL(\"color\"))"),
        0.0,
    );
    assert_eq!(sheet.circular_reference_count(), 0);
    assert_number(
        &sheet.eval("=IF(FALSE,INDIRECT(\"A1\"),CELL(\"parentheses\"))"),
        0.0,
    );
    assert_eq!(sheet.circular_reference_count(), 0);
}

#[test]
fn cell_format_reports_excel_cell_format_codes_from_number_formats() {
    use formula_engine::Engine;
    use formula_format::builtin_format_code;
    use formula_model::Style;

    let mut engine = Engine::new();

    // Text format: literal code and builtin placeholder.
    let text_style = engine.intern_style(Style {
        number_format: Some("@".to_string()),
        ..Default::default()
    });
    engine
        .set_cell_style_id("Sheet1", "A1", text_style)
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B1", "=CELL(\"format\",A1)")
        .unwrap();

    let text_placeholder_style = engine.intern_style(Style {
        number_format: Some("__builtin_numFmtId:49".to_string()),
        ..Default::default()
    });
    engine
        .set_cell_style_id("Sheet1", "A2", text_placeholder_style)
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B2", "=CELL(\"format\",A2)")
        .unwrap();

    // Fractions: built-ins 12/13 are a non-standard numeric family and Excel reports `N`.
    let frac_style = engine.intern_style(Style {
        number_format: Some(builtin_format_code(12).unwrap().to_string()),
        ..Default::default()
    });
    engine
        .set_cell_style_id("Sheet1", "A3", frac_style)
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B3", "=CELL(\"format\",A3)")
        .unwrap();

    let frac_placeholder_style = engine.intern_style(Style {
        number_format: Some("__builtin_numFmtId:12".to_string()),
        ..Default::default()
    });
    engine
        .set_cell_style_id("Sheet1", "A4", frac_placeholder_style)
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B4", "=CELL(\"format\",A4)")
        .unwrap();

    // Accounting: 41/43 are accounting-without-currency (grouped number formats) => "N0"/"N2";
    // 42/44 include currency => "C0"/"C2".
    for (idx, (num_fmt_id, _expected)) in [(41u16, "N0"), (42, "C0"), (43, "N2"), (44, "C2")]
        .into_iter()
        .enumerate()
    {
        let row = 5 + idx as u32;
        let a = format!("A{row}");
        let b = format!("B{row}");
        let style = engine.intern_style(Style {
            number_format: Some(builtin_format_code(num_fmt_id).unwrap().to_string()),
            ..Default::default()
        });
        engine.set_cell_style_id("Sheet1", &a, style).unwrap();
        engine
            .set_cell_formula("Sheet1", &b, &format!("=CELL(\"format\",{a})"))
            .unwrap();
        let placeholder_row = row + 4;
        let a_ph = format!("A{placeholder_row}");
        let b_ph = format!("B{placeholder_row}");
        let style_ph = engine.intern_style(Style {
            number_format: Some(format!("__builtin_numFmtId:{num_fmt_id}")),
            ..Default::default()
        });
        engine.set_cell_style_id("Sheet1", &a_ph, style_ph).unwrap();
        engine
            .set_cell_formula("Sheet1", &b_ph, &format!("=CELL(\"format\",{a_ph})"))
            .unwrap();
    }

    engine.recalculate_single_threaded();

    assert_eq!(
        engine.get_cell_value("Sheet1", "B1"),
        Value::Text("@".to_string())
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "B2"),
        Value::Text("@".to_string())
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "B3"),
        Value::Text("N".to_string())
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "B4"),
        Value::Text("N".to_string())
    );

    for (idx, (_num_fmt_id, expected)) in [(41u16, "N0"), (42, "C0"), (43, "N2"), (44, "C2")]
        .into_iter()
        .enumerate()
    {
        let row = 5 + idx as u32;
        let b = format!("B{row}");
        assert_eq!(
            engine.get_cell_value("Sheet1", &b),
            Value::Text(expected.to_string())
        );
        let b_ph = format!("B{}", row + 4);
        assert_eq!(
            engine.get_cell_value("Sheet1", &b_ph),
            Value::Text(expected.to_string())
        );
    }
}

#[test]
fn cell_color_and_parentheses_reflect_number_format_sections() {
    use formula_engine::Engine;
    use formula_model::Style;

    let mut engine = Engine::new();

    // One-section formats do not set the CELL("color") / CELL("parentheses") flags.
    let one_section = engine.intern_style(Style {
        number_format: Some("[Red]0".to_string()),
        ..Default::default()
    });
    engine
        .set_cell_style_id("Sheet1", "A1", one_section)
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B1", "=CELL(\"color\",A1)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B2", "=CELL(\"parentheses\",A1)")
        .unwrap();

    // Two-section formats with explicit negative section.
    let two_section = engine.intern_style(Style {
        number_format: Some("0;[Red](0)".to_string()),
        ..Default::default()
    });
    engine
        .set_cell_style_id("Sheet1", "A2", two_section)
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B3", "=CELL(\"color\",A2)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B4", "=CELL(\"parentheses\",A2)")
        .unwrap();

    engine.recalculate_single_threaded();

    assert_number(&engine.get_cell_value("Sheet1", "B1"), 0.0);
    assert_number(&engine.get_cell_value("Sheet1", "B2"), 0.0);
    assert_number(&engine.get_cell_value("Sheet1", "B3"), 1.0);
    assert_number(&engine.get_cell_value("Sheet1", "B4"), 1.0);
}

#[test]
fn cell_protect_reflects_locked_style_and_precedence() {
    use formula_engine::Engine;
    use formula_model::{Protection, Style};

    let mut engine = Engine::new();

    let unlocked = engine.intern_style(Style {
        protection: Some(Protection {
            locked: false,
            hidden: false,
        }),
        ..Style::default()
    });
    let locked = engine.intern_style(Style {
        protection: Some(Protection {
            locked: true,
            hidden: false,
        }),
        ..Style::default()
    });

    // Row style should apply when cell has no explicit style id.
    engine.set_row_style_id("Sheet1", 0, Some(unlocked));
    // Column style is lower precedence than row style.
    engine.set_col_style_id("Sheet1", 0, Some(locked));

    engine
        .set_cell_formula("Sheet1", "B1", "=CELL(\"protect\",A1)")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_number(&engine.get_cell_value("Sheet1", "B1"), 0.0);

    // Cell style overrides row/column styles.
    engine.set_cell_style_id("Sheet1", "A1", locked).unwrap();
    engine.recalculate_single_threaded();
    assert_number(&engine.get_cell_value("Sheet1", "B1"), 1.0);
}

#[test]
fn cell_prefix_matches_horizontal_alignment() {
    use formula_engine::Engine;
    use formula_model::{Alignment, HorizontalAlignment, Style};

    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "B1", "=CELL(\"prefix\",A1)")
        .unwrap();

    let cases = [
        (HorizontalAlignment::Left, "'"),
        (HorizontalAlignment::Right, "\""),
        (HorizontalAlignment::Center, "^"),
        (HorizontalAlignment::Fill, "\\"),
        (HorizontalAlignment::General, ""),
        (HorizontalAlignment::Justify, ""),
    ];

    for (alignment, expected) in cases {
        let style_id = engine.intern_style(Style {
            alignment: Some(Alignment {
                horizontal: Some(alignment),
                ..Alignment::default()
            }),
            ..Style::default()
        });
        engine.set_cell_style_id("Sheet1", "A1", style_id).unwrap();
        engine.recalculate_single_threaded();
        assert_eq!(
            engine.get_cell_value("Sheet1", "B1"),
            Value::Text(expected.to_string())
        );
    }
}

#[test]
fn cell_width_reflects_column_width_and_hidden_flag() {
    use formula_engine::Engine;

    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "B1", "=CELL(\"width\",A1)")
        .unwrap();

    // Default Excel width is 8.43 characters, but CELL("width") returns the rounded integer part
    // plus an indicator for custom width.
    engine.recalculate_single_threaded();
    assert_number(&engine.get_cell_value("Sheet1", "B1"), 8.0);

    engine.set_col_width("Sheet1", 0, Some(10.0));
    engine.recalculate_single_threaded();
    assert_number(&engine.get_cell_value("Sheet1", "B1"), 10.1);

    engine.set_col_hidden("Sheet1", 0, true);
    engine.recalculate_single_threaded();
    assert_number(&engine.get_cell_value("Sheet1", "B1"), 0.0);
}

#[test]
fn cell_address_quotes_sheet_names_when_needed() {
    use formula_engine::Engine;

    let mut engine = Engine::new();
    engine.set_cell_value("My Sheet", "A1", 1.0).unwrap();
    engine.set_cell_value("A1", "A1", 1.0).unwrap();
    engine.set_cell_value("O'Brien", "A1", 1.0).unwrap();

    engine
        .set_cell_formula("Sheet1", "B1", "=CELL(\"address\",'My Sheet'!A1)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B2", "=CELL(\"address\",'A1'!A1)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B3", "=CELL(\"address\",'O''Brien'!A1)")
        .unwrap();

    engine.recalculate_single_threaded();

    assert_eq!(
        engine.get_cell_value("Sheet1", "B1"),
        Value::Text("'My Sheet'!$A$1".to_string())
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "B2"),
        Value::Text("'A1'!$A$1".to_string())
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "B3"),
        Value::Text("'O''Brien'!$A$1".to_string())
    );
}

#[test]
fn cell_format_classifies_currency_formats() {
    use formula_engine::Engine;

    let mut engine = Engine::new();
    let style_currency_bracket = engine.intern_style(Style {
        number_format: Some("[$€-407]#,##0.00".to_string()),
        ..Default::default()
    });
    let style_currency_plain = engine.intern_style(Style {
        number_format: Some("€#,##0.00".to_string()),
        ..Default::default()
    });
    let style_locale_only = engine.intern_style(Style {
        number_format: Some("[$-409]0.00".to_string()),
        ..Default::default()
    });

    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine
        .set_cell_style_id("Sheet1", "A1", style_currency_bracket)
        .unwrap();
    engine.set_cell_value("Sheet1", "A2", 1.0).unwrap();
    engine
        .set_cell_style_id("Sheet1", "A2", style_currency_plain)
        .unwrap();
    engine.set_cell_value("Sheet1", "A3", 1.0).unwrap();
    engine
        .set_cell_style_id("Sheet1", "A3", style_locale_only)
        .unwrap();

    engine
        .set_cell_formula("Sheet1", "B1", "=CELL(\"format\",A1)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B2", "=CELL(\"format\",A2)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B3", "=CELL(\"format\",A3)")
        .unwrap();
    engine.recalculate_single_threaded();

    assert_eq!(
        engine.get_cell_value("Sheet1", "B1"),
        Value::Text("C2".to_string())
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "B2"),
        Value::Text("C2".to_string())
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "B3"),
        Value::Text("F2".to_string())
    );
}

#[test]
fn cell_color_and_parentheses_match_excel_number_format_semantics() {
    use formula_engine::Engine;
    use formula_model::Style;

    let mut engine = Engine::new();

    let style_red_one_section = engine.intern_style(Style {
        number_format: Some("[Red]0".to_string()),
        ..Style::default()
    });
    let style_parens_one_section = engine.intern_style(Style {
        number_format: Some("(0)".to_string()),
        ..Style::default()
    });
    let style_parens_two_section = engine.intern_style(Style {
        number_format: Some("0;(0)".to_string()),
        ..Style::default()
    });
    let style_red_two_section = engine.intern_style(Style {
        number_format: Some("0;[Red]0".to_string()),
        ..Style::default()
    });
    let style_conditional_first_section = engine.intern_style(Style {
        number_format: Some("[<0][Red]0;0".to_string()),
        ..Style::default()
    });
    let style_conditional_second_section = engine.intern_style(Style {
        number_format: Some("[>=0]0;[Red]0".to_string()),
        ..Style::default()
    });
    let style_parentheses_quoted_literal = engine.intern_style(Style {
        number_format: Some(r#"0;"(neg)"0"#.to_string()),
        ..Style::default()
    });
    let style_parentheses_escaped = engine.intern_style(Style {
        number_format: Some(r#"0;\(0\)"#.to_string()),
        ..Style::default()
    });
    let style_bracket_token_parentheses = engine.intern_style(Style {
        number_format: Some(r#"0;[$(USD)-409]0"#.to_string()),
        ..Style::default()
    });

    // Attach the formats to a few cells (the values do not matter for these flags).
    for addr in ["A1", "A2", "A3", "A4", "A5", "A6", "A7", "A8", "A9"] {
        engine.set_cell_value("Sheet1", addr, 1.0).unwrap();
    }

    engine
        .set_cell_style_id("Sheet1", "A1", style_red_one_section)
        .unwrap();
    engine
        .set_cell_style_id("Sheet1", "A2", style_parens_one_section)
        .unwrap();
    engine
        .set_cell_style_id("Sheet1", "A3", style_parens_two_section)
        .unwrap();
    engine
        .set_cell_style_id("Sheet1", "A4", style_red_two_section)
        .unwrap();
    engine
        .set_cell_style_id("Sheet1", "A5", style_conditional_first_section)
        .unwrap();
    engine
        .set_cell_style_id("Sheet1", "A6", style_conditional_second_section)
        .unwrap();
    engine
        .set_cell_style_id("Sheet1", "A7", style_parentheses_quoted_literal)
        .unwrap();
    engine
        .set_cell_style_id("Sheet1", "A8", style_parentheses_escaped)
        .unwrap();
    engine
        .set_cell_style_id("Sheet1", "A9", style_bracket_token_parentheses)
        .unwrap();

    engine
        .set_cell_formula("Sheet1", "B1", "=CELL(\"color\",A1)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B2", "=CELL(\"parentheses\",A2)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B3", "=CELL(\"parentheses\",A3)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B4", "=CELL(\"color\",A4)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B5", "=CELL(\"color\",A5)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B6", "=CELL(\"color\",A6)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B7", "=CELL(\"parentheses\",A7)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B8", "=CELL(\"parentheses\",A8)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B9", "=CELL(\"parentheses\",A9)")
        .unwrap();

    engine.recalculate_single_threaded();

    // One-section formats: Excel reports 0/0 even if the only section contains a color token or
    // parentheses literals (negatives use the first section with an automatic '-' sign).
    assert_number(&engine.get_cell_value("Sheet1", "B1"), 0.0);
    assert_number(&engine.get_cell_value("Sheet1", "B2"), 0.0);

    // Two-section formats.
    assert_number(&engine.get_cell_value("Sheet1", "B3"), 1.0);
    assert_number(&engine.get_cell_value("Sheet1", "B4"), 1.0);

    // Conditional sections.
    assert_number(&engine.get_cell_value("Sheet1", "B5"), 1.0);
    assert_number(&engine.get_cell_value("Sheet1", "B6"), 1.0);

    // Parentheses inside quotes / escaped / inside bracket tokens should not count.
    assert_number(&engine.get_cell_value("Sheet1", "B7"), 0.0);
    assert_number(&engine.get_cell_value("Sheet1", "B8"), 0.0);
    assert_number(&engine.get_cell_value("Sheet1", "B9"), 0.0);
}

#[test]
fn cell_color_and_parentheses_resolve_row_and_column_styles() {
    use formula_engine::Engine;
    use formula_model::Style;

    let mut engine = Engine::new();

    // Row style makes negatives red; column style is default.
    let style_row = engine.intern_style(Style {
        number_format: Some("0;[Red]0".to_string()),
        ..Style::default()
    });
    let style_col = engine.intern_style(Style {
        number_format: Some("0".to_string()),
        ..Style::default()
    });
    let style_cell_override = engine.intern_style(Style {
        number_format: Some("0".to_string()),
        ..Style::default()
    });

    engine.set_row_style_id("Sheet1", 0, Some(style_row));
    engine.set_col_style_id("Sheet1", 0, Some(style_col));

    engine
        .set_cell_formula("Sheet1", "B1", "=CELL(\"color\",A1)")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_number(&engine.get_cell_value("Sheet1", "B1"), 1.0);

    // Explicit cell formatting should override row/column defaults.
    engine
        .set_cell_style_id("Sheet1", "A1", style_cell_override)
        .unwrap();
    engine.recalculate_single_threaded();
    assert_number(&engine.get_cell_value("Sheet1", "B1"), 0.0);

    // Column style applies when row style is not present.
    engine.set_row_style_id("Sheet1", 0, None);
    engine.set_cell_style_id("Sheet1", "A1", 0).unwrap();
    engine
        .set_cell_formula("Sheet1", "B2", "=CELL(\"color\",A1)")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_number(&engine.get_cell_value("Sheet1", "B2"), 0.0);

    // Now apply the red negative style to the column and ensure CELL picks it up.
    engine.set_col_style_id("Sheet1", 0, Some(style_row));
    engine.recalculate_single_threaded();
    assert_number(&engine.get_cell_value("Sheet1", "B2"), 1.0);
}

#[test]
fn nonvolatile_formulas_are_not_recalculated_when_nothing_is_dirty() {
    use formula_engine::Engine;

    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=RECALC_TICK_TEST()")
        .unwrap();

    engine.recalculate_single_threaded();
    let first = engine.get_cell_value("Sheet1", "A1");

    // With no dirty cells and no volatile inputs, the engine should short-circuit and keep the
    // previously computed value.
    engine.recalculate_single_threaded();
    let second = engine.get_cell_value("Sheet1", "A1");

    assert_eq!(first, second);
}

#[test]
fn cell_and_info_make_formulas_recalculate_each_tick() {
    use formula_engine::Engine;

    // CELL(...) should put the formula into the volatile closure, causing it to be evaluated on
    // each recalc tick even when nothing is dirty.
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine
        .set_cell_formula("Sheet1", "B1", "=RECALC_TICK_TEST()+0*CELL(\"row\",A1)")
        .unwrap();
    engine.recalculate_single_threaded();
    let first = engine.get_cell_value("Sheet1", "B1");
    engine.recalculate_single_threaded();
    let second = engine.get_cell_value("Sheet1", "B1");
    assert_ne!(first, second);

    // INFO(...) should also be treated as volatile for Excel compatibility.
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=RECALC_TICK_TEST()+0*INFO(\"numfile\")")
        .unwrap();
    engine.recalculate_single_threaded();
    let first = engine.get_cell_value("Sheet1", "A1");
    engine.recalculate_single_threaded();
    let second = engine.get_cell_value("Sheet1", "A1");
    assert_ne!(first, second);
}

#[test]
fn cell_format_color_and_parentheses_reflect_number_format_strings() {
    use formula_engine::Engine;
    use formula_model::Style;

    let mut engine = Engine::new();

    // Two-section format with explicit red parentheses for negatives.
    let style_id = engine.intern_style(Style {
        number_format: Some("0;[Red](0)".to_string()),
        ..Style::default()
    });
    engine.set_cell_style_id("Sheet1", "A1", style_id).unwrap();

    // One-section formats should report 0/0 even if the section contains red/parentheses literals.
    let one_section_red = engine.intern_style(Style {
        number_format: Some("[Red]0".to_string()),
        ..Style::default()
    });
    engine
        .set_cell_style_id("Sheet1", "A2", one_section_red)
        .unwrap();

    let one_section_paren = engine.intern_style(Style {
        number_format: Some("(0)".to_string()),
        ..Style::default()
    });
    engine
        .set_cell_style_id("Sheet1", "A3", one_section_paren)
        .unwrap();

    engine
        .set_cell_formula("Sheet1", "B1", "=CELL(\"color\",A1)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B2", "=CELL(\"parentheses\",A1)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B3", "=CELL(\"color\",A2)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B4", "=CELL(\"parentheses\",A3)")
        .unwrap();

    engine.recalculate_single_threaded();

    assert_number(&engine.get_cell_value("Sheet1", "B1"), 1.0);
    assert_number(&engine.get_cell_value("Sheet1", "B2"), 1.0);
    assert_number(&engine.get_cell_value("Sheet1", "B3"), 0.0);
    assert_number(&engine.get_cell_value("Sheet1", "B4"), 0.0);
}

#[test]
fn cell_format_classifies_locale_variant_datetime_formats() {
    use formula_engine::Engine;
    use formula_model::Style;

    let mut engine = Engine::new();

    let mdy = engine.intern_style(Style {
        number_format: Some("m/d/yyyy".to_string()),
        ..Style::default()
    });
    let dmy = engine.intern_style(Style {
        number_format: Some("dd/mm/yyyy".to_string()),
        ..Style::default()
    });
    let iso = engine.intern_style(Style {
        number_format: Some("yyyy-mm-dd".to_string()),
        ..Style::default()
    });
    let h = engine.intern_style(Style {
        number_format: Some("h:mm".to_string()),
        ..Style::default()
    });
    let hh = engine.intern_style(Style {
        number_format: Some("hh:mm".to_string()),
        ..Style::default()
    });
    let hh_ss = engine.intern_style(Style {
        number_format: Some("hh:mm:ss".to_string()),
        ..Style::default()
    });
    let system_long_date = engine.intern_style(Style {
        number_format: Some("[$-F800]dddd, mmmm dd, yyyy".to_string()),
        ..Style::default()
    });

    engine.set_cell_style_id("Sheet1", "A1", mdy).unwrap();
    engine.set_cell_style_id("Sheet1", "A2", dmy).unwrap();
    engine.set_cell_style_id("Sheet1", "A3", iso).unwrap();
    engine.set_cell_style_id("Sheet1", "A4", h).unwrap();
    engine.set_cell_style_id("Sheet1", "A5", hh).unwrap();
    engine.set_cell_style_id("Sheet1", "A6", hh_ss).unwrap();
    engine
        .set_cell_style_id("Sheet1", "A7", system_long_date)
        .unwrap();

    engine
        .set_cell_formula("Sheet1", "B1", "=CELL(\"format\",A1)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B2", "=CELL(\"format\",A2)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B3", "=CELL(\"format\",A3)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B4", "=CELL(\"format\",A4)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B5", "=CELL(\"format\",A5)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B6", "=CELL(\"format\",A6)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B7", "=CELL(\"format\",A7)")
        .unwrap();

    engine.recalculate_single_threaded();

    // Day-first dates + ISO-ish year-first dates should classify like Excel's canonical short
    // numeric date (`D4`, e.g. `m/d/yy`).
    let b1 = engine.get_cell_value("Sheet1", "B1");
    assert_eq!(b1, Value::Text("D4".to_string()));
    assert_eq!(engine.get_cell_value("Sheet1", "B2"), b1);
    assert_eq!(engine.get_cell_value("Sheet1", "B3"), b1);

    // hh:mm should classify like h:mm (`D9`).
    let b4 = engine.get_cell_value("Sheet1", "B4");
    assert_eq!(b4, Value::Text("D9".to_string()));
    assert_eq!(engine.get_cell_value("Sheet1", "B5"), b4);

    // hh:mm:ss should classify as a time-with-seconds (`D8`).
    assert_eq!(
        engine.get_cell_value("Sheet1", "B6"),
        Value::Text("D8".to_string())
    );

    // System long date tokens should classify as some date code (not currency).
    match engine.get_cell_value("Sheet1", "B7") {
        Value::Text(s) => assert!(
            s.starts_with('D'),
            "expected date classification for system long date, got {s:?}"
        ),
        other => panic!("expected text for CELL(\"format\"), got {other:?}"),
    }
}

#[test]
fn cell_format_uses_row_and_column_styles_when_cell_style_is_default() {
    use formula_engine::Engine;
    use formula_model::Style;

    let mut engine = Engine::new();

    // Row 1 default: short date (D4).
    let date_style = engine.intern_style(Style {
        number_format: Some("m/d/yyyy".to_string()),
        ..Style::default()
    });
    engine.set_row_style_id("Sheet1", 0, Some(date_style));

    // Column A default: time (D9).
    let time_style = engine.intern_style(Style {
        number_format: Some("h:mm".to_string()),
        ..Style::default()
    });
    engine.set_col_style_id("Sheet1", 0, Some(time_style));

    // A1 should inherit from the row style (row wins over column).
    engine
        .set_cell_formula("Sheet1", "C1", "=CELL(\"format\",A1)")
        .unwrap();
    // B1 should inherit from the row style.
    engine
        .set_cell_formula("Sheet1", "C2", "=CELL(\"format\",B1)")
        .unwrap();
    // A2 should inherit from the column style.
    engine
        .set_cell_formula("Sheet1", "C3", "=CELL(\"format\",A2)")
        .unwrap();
    // B2 has no style metadata and should default to General.
    engine
        .set_cell_formula("Sheet1", "C4", "=CELL(\"format\",B2)")
        .unwrap();

    engine.recalculate_single_threaded();

    assert_eq!(
        engine.get_cell_value("Sheet1", "C1"),
        Value::Text("D4".to_string())
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "C2"),
        Value::Text("D4".to_string())
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "C3"),
        Value::Text("D9".to_string())
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "C4"),
        Value::Text("G".to_string())
    );
}

#[test]
fn cell_width_matches_excel_encoding_for_default_custom_and_hidden_columns() {
    use formula_engine::Engine;

    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "B1", "=CELL(\"width\",A1)")
        .unwrap();
    engine.recalculate_single_threaded();

    // Default width (no explicit column metadata): Excel encodes whether the column uses an
    // explicit per-column width. The flag is `.0` for sheet-default, `.1` for explicit widths.
    assert_number(&engine.get_cell_value("Sheet1", "B1"), 8.0);

    // Explicit custom width override *even when equal to the standard width* sets the `.1` flag.
    engine.set_col_width("Sheet1", 0, Some(8.43_f32));
    engine.recalculate_single_threaded();
    assert_number(&engine.get_cell_value("Sheet1", "B1"), 8.1);

    // A wider custom width.
    engine.set_col_width("Sheet1", 0, Some(25.0_f32));
    engine.recalculate_single_threaded();
    assert_number(&engine.get_cell_value("Sheet1", "B1"), 25.1);

    // Hidden columns always report 0, regardless of the stored width.
    engine.set_col_hidden("Sheet1", 0, true);
    engine.recalculate_single_threaded();
    assert_number(&engine.get_cell_value("Sheet1", "B1"), 0.0);
}

#[test]
fn cell_width_recalculates_when_column_width_changes_in_auto_mode() {
    use formula_engine::calc_settings::{CalcSettings, CalculationMode};
    use formula_engine::Engine;

    let mut engine = Engine::new();
    engine.set_calc_settings(CalcSettings {
        calculation_mode: CalculationMode::Automatic,
        ..CalcSettings::default()
    });

    engine
        .set_cell_formula("Sheet1", "B1", "=CELL(\"width\",A1)")
        .unwrap();

    assert_number(&engine.get_cell_value("Sheet1", "B1"), 8.0);

    // Column width changes should trigger a recalculation in automatic mode.
    engine.set_col_width("Sheet1", 0, Some(25.0_f32));
    assert_number(&engine.get_cell_value("Sheet1", "B1"), 25.1);
}

#[test]
fn cell_width_recalculates_when_column_hidden_changes_in_auto_mode() {
    use formula_engine::calc_settings::{CalcSettings, CalculationMode};
    use formula_engine::Engine;

    let mut engine = Engine::new();
    engine.set_calc_settings(CalcSettings {
        calculation_mode: CalculationMode::Automatic,
        ..CalcSettings::default()
    });

    engine
        .set_cell_formula("Sheet1", "B1", "=CELL(\"width\",A1)")
        .unwrap();

    // Start with a custom width so we can verify that un-hiding restores the previous encoded
    // value.
    engine.set_col_width("Sheet1", 0, Some(25.0_f32));
    assert_number(&engine.get_cell_value("Sheet1", "B1"), 25.1);

    // Hiding the column should immediately recompute and return 0.
    engine.set_col_hidden("Sheet1", 0, true);
    assert_number(&engine.get_cell_value("Sheet1", "B1"), 0.0);

    // Unhiding should recompute again and restore the encoded width.
    engine.set_col_hidden("Sheet1", 0, false);
    assert_number(&engine.get_cell_value("Sheet1", "B1"), 25.1);
}
