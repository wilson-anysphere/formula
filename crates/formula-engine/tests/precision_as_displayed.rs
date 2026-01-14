use formula_engine::calc_settings::CalcSettings;
use formula_engine::metadata::FormatRun;
use formula_engine::{Engine, Value};
use formula_model::{CellRef, Font, Range, Style};

#[test]
fn precision_as_displayed_rounds_numeric_literals_fixed_decimals() {
    let mut engine = Engine::new();
    let mut settings: CalcSettings = engine.calc_settings().clone();
    settings.full_precision = false;
    engine.set_calc_settings(settings);

    engine
        .set_cell_number_format("Sheet1", "A1", Some("0.00".to_string()))
        .unwrap();
    engine.set_cell_value("Sheet1", "A1", 1.239).unwrap();

    // The stored value should be rounded to match the displayed precision.
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(1.24));

    // Downstream formulas should observe the rounded stored value.
    engine.set_cell_formula("Sheet1", "B1", "=A1").unwrap();
    engine.recalculate();
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(1.24));
}

#[test]
fn precision_as_displayed_rounds_numeric_literals_percent() {
    let mut engine = Engine::new();
    let mut settings: CalcSettings = engine.calc_settings().clone();
    settings.full_precision = false;
    engine.set_calc_settings(settings);

    engine
        .set_cell_number_format("Sheet1", "A1", Some("0%".to_string()))
        .unwrap();
    engine.set_cell_value("Sheet1", "A1", 0.1234).unwrap();

    // "0%" displays 12% for 0.1234, so the stored value should be 0.12.
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(0.12));

    engine.set_cell_formula("Sheet1", "B1", "=A1*100").unwrap();
    engine.recalculate();
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(12.0));
}

#[test]
fn full_precision_does_not_round_numeric_literals() {
    let mut engine = Engine::new();
    let mut settings: CalcSettings = engine.calc_settings().clone();
    settings.full_precision = true;
    engine.set_calc_settings(settings);

    engine
        .set_cell_number_format("Sheet1", "A1", Some("0.00".to_string()))
        .unwrap();
    engine.set_cell_value("Sheet1", "A1", 1.239).unwrap();

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(1.239));

    engine.set_cell_formula("Sheet1", "B1", "=A1").unwrap();
    engine.recalculate();
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(1.239));
}

#[test]
fn precision_as_displayed_rounds_numeric_literals_using_style_number_format() {
    let mut engine = Engine::new();
    let mut settings: CalcSettings = engine.calc_settings().clone();
    settings.full_precision = false;
    engine.set_calc_settings(settings);

    let style_id = engine.intern_style(Style {
        number_format: Some("0.00".to_string()),
        ..Style::default()
    });

    engine
        .set_cell_style_id("Sheet1", "A1", style_id)
        .expect("set style");
    engine.set_cell_value("Sheet1", "A1", 1.239).unwrap();

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(1.24));
}

#[test]
fn precision_as_displayed_rounds_numeric_literals_using_row_col_style_fallback() {
    let mut engine = Engine::new();
    let mut settings: CalcSettings = engine.calc_settings().clone();
    settings.full_precision = false;
    engine.set_calc_settings(settings);

    let row_style_id = engine.intern_style(Style {
        number_format: Some("0.00".to_string()),
        ..Style::default()
    });
    let col_style_id = engine.intern_style(Style {
        number_format: Some("0.0".to_string()),
        ..Style::default()
    });

    // Row styles take precedence over column styles when the cell has style_id 0.
    engine.set_row_style_id("Sheet1", 0, Some(row_style_id));
    engine.set_col_style_id("Sheet1", 0, Some(col_style_id));

    engine.set_cell_value("Sheet1", "A1", 1.239).unwrap();
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(1.24));

    // For other rows, fall back to column styles.
    engine.set_cell_value("Sheet1", "A2", 1.239).unwrap();
    assert_eq!(engine.get_cell_value("Sheet1", "A2"), Value::Number(1.2));
}

#[test]
fn precision_as_displayed_rounds_range_values_using_effective_style_format() {
    let mut engine = Engine::new();
    let mut settings: CalcSettings = engine.calc_settings().clone();
    settings.full_precision = false;
    engine.set_calc_settings(settings);

    // Apply a number format via a row style so cells without explicit style ids still pick it up.
    let row_style_id = engine.intern_style(Style {
        number_format: Some("0.00".to_string()),
        ..Style::default()
    });
    engine.set_row_style_id("Sheet1", 0, Some(row_style_id));

    let values = vec![vec![Value::Number(1.239), Value::Number(2.345)]];
    engine
        .set_range_values(
            "Sheet1",
            Range::new(CellRef::new(0, 0), CellRef::new(0, 1)),
            &values,
            false,
        )
        .unwrap();

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(1.24));
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(2.35));
}

#[test]
fn precision_as_displayed_rounds_numeric_literals_using_range_run_style_fallback() {
    let mut engine = Engine::new();
    let mut settings: CalcSettings = engine.calc_settings().clone();
    settings.full_precision = false;
    engine.set_calc_settings(settings);

    let run_style_id = engine.intern_style(Style {
        number_format: Some("0.00".to_string()),
        ..Style::default()
    });

    engine
        .set_format_runs_by_col(
            "Sheet1",
            0,
            vec![FormatRun {
                start_row: 0,
                end_row_exclusive: 1,
                style_id: run_style_id,
            }],
        )
        .unwrap();

    engine.set_cell_value("Sheet1", "A1", 1.239).unwrap();
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(1.24));
}

#[test]
fn precision_as_displayed_rounds_numeric_literals_using_sheet_default_style_fallback() {
    let mut engine = Engine::new();
    let mut settings: CalcSettings = engine.calc_settings().clone();
    settings.full_precision = false;
    engine.set_calc_settings(settings);

    let style_id = engine.intern_style(Style {
        number_format: Some("0.00".to_string()),
        ..Style::default()
    });

    engine.set_sheet_default_style_id("Sheet1", Some(style_id));
    engine.set_cell_value("Sheet1", "A1", 1.239).unwrap();

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(1.24));
}

#[test]
fn precision_as_displayed_rounds_using_range_run_number_format() {
    let mut engine = Engine::new();
    let mut settings: CalcSettings = engine.calc_settings().clone();
    settings.full_precision = false;
    engine.set_calc_settings(settings);

    let style_id = engine.intern_style(Style {
        number_format: Some("0.00".to_string()),
        ..Style::default()
    });

    // Apply the style via a range-run layer: rows 1-10 (inclusive) in column A.
    engine
        .set_format_runs_by_col(
            "Sheet1",
            0,
            vec![FormatRun {
                start_row: 0,
                end_row_exclusive: 10,
                style_id,
            }],
        )
        .unwrap();

    // In-run numeric literals should round.
    engine.set_cell_value("Sheet1", "A1", 1.239).unwrap();
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(1.24));

    // Outside the run, the default "General" format does not round 1.239.
    engine.set_cell_value("Sheet1", "A11", 1.239).unwrap();
    assert_eq!(engine.get_cell_value("Sheet1", "A11"), Value::Number(1.239));

    // Formula results should also be rounded using the run formatting.
    engine.set_cell_formula("Sheet1", "A2", "=1.239").unwrap();
    engine.recalculate();
    assert_eq!(engine.get_cell_value("Sheet1", "A2"), Value::Number(1.24));
}

#[test]
fn precision_as_displayed_number_format_inherits_through_cell_style_layer() {
    let mut engine = Engine::new();
    let mut settings: CalcSettings = engine.calc_settings().clone();
    settings.full_precision = false;
    engine.set_calc_settings(settings);

    // Simulate a cell having some explicit formatting (e.g. font) while the row provides the number
    // format. The number format should still apply for "precision as displayed" rounding.
    let cell_style_id = engine.intern_style(Style {
        font: Some(Font {
            bold: true,
            ..Default::default()
        }),
        ..Default::default()
    });
    let row_style_id = engine.intern_style(Style {
        number_format: Some("0.00".to_string()),
        ..Style::default()
    });

    engine
        .set_cell_style_id("Sheet1", "A1", cell_style_id)
        .expect("set style");
    engine.set_row_style_id("Sheet1", 0, Some(row_style_id));

    engine.set_cell_value("Sheet1", "A1", 1.239).unwrap();
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(1.24));
}

#[test]
fn precision_as_displayed_cell_number_format_override_wins_over_style() {
    let mut engine = Engine::new();
    let mut settings: CalcSettings = engine.calc_settings().clone();
    settings.full_precision = false;
    engine.set_calc_settings(settings);

    let style_id = engine.intern_style(Style {
        number_format: Some("0.00".to_string()),
        ..Style::default()
    });

    engine
        .set_cell_style_id("Sheet1", "A1", style_id)
        .expect("set style");
    engine
        .set_cell_number_format("Sheet1", "A1", Some("0.0".to_string()))
        .unwrap();
    engine.set_cell_value("Sheet1", "A1", 1.239).unwrap();

    // The explicit per-cell format override should be used for rounding.
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(1.2));
}

#[test]
fn precision_as_displayed_rounds_scalar_formula_results_using_effective_style_format() {
    let mut engine = Engine::new();
    let mut settings: CalcSettings = engine.calc_settings().clone();
    settings.full_precision = false;
    engine.set_calc_settings(settings);

    let style_id = engine.intern_style(Style {
        number_format: Some("0.00".to_string()),
        ..Style::default()
    });
    engine
        .set_cell_style_id("Sheet1", "A1", style_id)
        .expect("set style");
    engine.set_cell_formula("Sheet1", "A1", "=1.239").unwrap();
    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(1.24));
}

#[test]
fn precision_as_displayed_rounds_spilled_arrays_using_spill_origin_format() {
    let mut engine = Engine::new();
    let mut settings: CalcSettings = engine.calc_settings().clone();
    settings.full_precision = false;
    engine.set_calc_settings(settings);

    let origin_style_id = engine.intern_style(Style {
        number_format: Some("0.00".to_string()),
        ..Style::default()
    });
    let other_style_id = engine.intern_style(Style {
        number_format: Some("0.0".to_string()),
        ..Style::default()
    });

    engine
        .set_cell_style_id("Sheet1", "A1", origin_style_id)
        .expect("set style");

    // Set a conflicting style on column B to ensure rounding uses the spill origin's format.
    engine.set_col_style_id("Sheet1", 1, Some(other_style_id));

    engine
        .set_cell_formula("Sheet1", "A1", "=SEQUENCE(1,2,1.239,1.106)")
        .unwrap();
    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(1.24));
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(2.35));
}
