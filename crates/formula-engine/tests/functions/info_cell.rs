use formula_engine::{ErrorKind, Value};
use formula_model::Style;

use super::harness::{assert_number, TestSheet};

use formula_engine::eval::CompiledExpr;
use formula_engine::functions::{
    ArraySupport, FunctionContext, FunctionSpec, ThreadSafety, ValueType, Volatility,
};

fn recalc_tick_test(ctx: &dyn FunctionContext, _args: &[CompiledExpr]) -> Value {
    // Use only 53 bits so the f64 conversion is exact and comparisons remain deterministic.
    Value::Number((ctx.volatile_rand_u64() >> 11) as f64)
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
    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Sheet1", "A1"),
        Value::Text("pcdos".to_string())
    );
    assert_eq!(engine.get_cell_value("Sheet1", "A2"), Value::Error(ErrorKind::NA));
    assert_eq!(engine.get_cell_value("Sheet1", "A3"), Value::Error(ErrorKind::NA));
    assert_eq!(engine.get_cell_value("Sheet1", "A4"), Value::Error(ErrorKind::NA));
    assert_eq!(engine.get_cell_value("Sheet1", "A5"), Value::Error(ErrorKind::NA));
    assert_eq!(engine.get_cell_value("Sheet1", "A6"), Value::Error(ErrorKind::NA));
    assert_eq!(engine.get_cell_value("Sheet1", "A7"), Value::Error(ErrorKind::NA));
    assert_eq!(engine.get_cell_value("Sheet1", "A8"), Value::Error(ErrorKind::NA));
    assert_eq!(engine.get_cell_value("Sheet2", "A8"), Value::Error(ErrorKind::NA));

    engine.set_engine_info(EngineInfo {
        system: Some("unix".to_string()),
        directory: Some("/tmp".to_string()),
        osversion: Some("14.2".to_string()),
        release: Some("release-x".to_string()),
        version: Some("v1".to_string()),
        memavail: Some(1234.0),
        totmem: Some(5678.0),
        origin: Some("$A$1".to_string()),
        ..EngineInfo::default()
    });
    engine.set_info_origin_for_sheet("Sheet2", Some("$B$2"));
    engine.recalculate_single_threaded();

    assert_eq!(
        engine.get_cell_value("Sheet1", "A1"),
        Value::Text("unix".to_string())
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "A2"),
        Value::Text("/tmp".to_string())
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "A3"),
        Value::Text("14.2".to_string())
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "A4"),
        Value::Text("release-x".to_string())
    );
    assert_eq!(engine.get_cell_value("Sheet1", "A5"), Value::Text("v1".to_string()));
    assert_number(&engine.get_cell_value("Sheet1", "A6"), 1234.0);
    assert_number(&engine.get_cell_value("Sheet1", "A7"), 5678.0);
    assert_eq!(
        engine.get_cell_value("Sheet1", "A8"),
        Value::Text("$A$1".to_string())
    );
    assert_eq!(
        engine.get_cell_value("Sheet2", "A8"),
        Value::Text("$B$2".to_string())
    );
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

    // Same idea, but for CELL("type") which also consults the referenced cell.
    assert_eq!(
        sheet.eval("=IF(FALSE,INDIRECT(\"A1\"),CELL(\"type\"))"),
        Value::Text("v".to_string())
    );
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

    match sheet.eval("=IF(FALSE,INDIRECT(\"A1\"),CELL(\"protect\"))") {
        Value::Number(n) => assert!(n != 0.0, "expected non-zero protect, got {n}"),
        other => panic!("expected number for CELL(\"protect\"), got {other:?}"),
    }

    assert_eq!(
        sheet.eval("=IF(FALSE,INDIRECT(\"A1\"),CELL(\"prefix\"))"),
        Value::Text(String::new())
    );
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
    engine.set_cell_style_id("Sheet1", "A1", style_currency_bracket).unwrap();
    engine.set_cell_value("Sheet1", "A2", 1.0).unwrap();
    engine.set_cell_style_id("Sheet1", "A2", style_currency_plain).unwrap();
    engine.set_cell_value("Sheet1", "A3", 1.0).unwrap();
    engine.set_cell_style_id("Sheet1", "A3", style_locale_only).unwrap();

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
    engine.set_cell_style_id("Sheet1", "A2", one_section_red).unwrap();

    let one_section_paren = engine.intern_style(Style {
        number_format: Some("(0)".to_string()),
        ..Style::default()
    });
    engine.set_cell_style_id("Sheet1", "A3", one_section_paren).unwrap();

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

    // Day-first dates + ISO-ish year-first dates should classify like the canonical short date.
    let b1 = engine.get_cell_value("Sheet1", "B1");
    // Excel uses `D4` for numeric short dates like `m/d/yy` (ordering is locale-dependent).
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
