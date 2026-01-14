use formula_engine::date::{ymd_to_serial, ExcelDate, ExcelDateSystem};
use formula_engine::locale::ValueLocaleConfig;
use formula_engine::value::{EntityValue, RecordValue};
use formula_engine::{Engine, ErrorKind, Value};

fn eval(engine: &mut Engine, formula: &str) -> Value {
    engine
        .set_cell_formula("Sheet1", "Z1", formula)
        .expect("set formula");
    engine.recalculate_single_threaded();
    engine.get_cell_value("Sheet1", "Z1")
}

#[test]
fn countif_numeric_operator_criteria() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 6.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", 10.0).unwrap();

    assert_eq!(
        eval(&mut engine, r#"=COUNTIF(A1:A3, ">5")"#),
        Value::Number(2.0)
    );
    assert_eq!(
        eval(&mut engine, r#"=COUNTIF(A1:A3, "<=3")"#),
        Value::Number(1.0)
    );
}

#[test]
fn countif_text_wildcards() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", "apple").unwrap();
    engine.set_cell_value("Sheet1", "A2", "apricot").unwrap();
    engine.set_cell_value("Sheet1", "A3", "banana").unwrap();
    engine.set_cell_value("Sheet1", "A4", "*").unwrap();
    engine.set_cell_value("Sheet1", "A5", "ab").unwrap();
    engine.set_cell_value("Sheet1", "A6", "a").unwrap();

    assert_eq!(
        eval(&mut engine, r#"=COUNTIF(A1:A6, "ap*")"#),
        Value::Number(2.0)
    );
    assert_eq!(
        eval(&mut engine, r#"=COUNTIF(A1:A6, "~*")"#),
        Value::Number(1.0)
    );
    assert_eq!(
        eval(&mut engine, r#"=COUNTIF(A1:A6, "??")"#),
        Value::Number(1.0)
    );
}

#[test]
fn countif_blank_criteria() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", "").unwrap();
    // A3 left unset (blank).

    assert_eq!(
        eval(&mut engine, r#"=COUNTIF(A1:A3, "")"#),
        Value::Number(2.0)
    );
    assert_eq!(
        eval(&mut engine, r#"=COUNTIF(A1:A3, "=")"#),
        Value::Number(2.0)
    );
    assert_eq!(
        eval(&mut engine, r#"=COUNTIF(A1:A3, "<>")"#),
        Value::Number(1.0)
    );
}

#[test]
fn countif_blank_criteria_equal_empty_string_literal_counts_blanks() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", "").unwrap();
    // A3 left unset (blank).

    // A criteria string of `=""` should match blank cells (same as `""` / `"="`).
    // Build it via concatenation to avoid hard-to-read quoting.
    assert_eq!(
        eval(&mut engine, r#"=COUNTIF(A1:A3, "="&""""&"""")"#),
        Value::Number(2.0)
    );
}

#[test]
fn countif_blank_criteria_equal_empty_string_literal_compiles_to_bytecode() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", "").unwrap();
    // A3 left unset (blank).

    engine
        .set_cell_formula("Sheet1", "Z1", r#"=COUNTIF(A1:A3, "=""""")"#)
        .unwrap();

    assert!(
        engine.bytecode_program_count() > 0,
        "expected COUNTIF formula to compile to bytecode for this test"
    );

    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "Z1"), Value::Number(2.0));
}

#[test]
fn countif_error_criteria_counts_errors_and_criteria_errors_propagate() {
    let mut engine = Engine::new();
    engine
        .set_cell_value("Sheet1", "A1", Value::Error(ErrorKind::Div0))
        .unwrap();
    engine
        .set_cell_value("Sheet1", "A2", Value::Error(ErrorKind::Div0))
        .unwrap();
    engine.set_cell_value("Sheet1", "A3", 1.0).unwrap();

    assert_eq!(
        eval(&mut engine, r##"=COUNTIF(A1:A3, "#DIV/0!")"##),
        Value::Number(2.0)
    );

    // Range errors do not propagate, but an error *criteria argument* does.
    assert_eq!(
        eval(&mut engine, "=COUNTIF(A1:A3, 1/0)"),
        Value::Error(ErrorKind::Div0)
    );

    // Candidate cell errors must not propagate.
    assert_eq!(
        eval(&mut engine, r#"=COUNTIF(A1:A3, ">0")"#),
        Value::Number(1.0)
    );
}

#[test]
fn countif_error_literal_criteria_compiles_to_bytecode_and_propagates() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", 3.0).unwrap();

    engine
        .set_cell_formula("Sheet1", "Z1", "=COUNTIF(A1:A3, #DIV/0!)")
        .unwrap();

    assert!(
        engine.bytecode_program_count() > 0,
        "expected COUNTIF formula with error literal criteria to compile to bytecode"
    );

    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Sheet1", "Z1"),
        Value::Error(ErrorKind::Div0)
    );
}

#[test]
fn countif_field_error_criteria_counts_field_errors() {
    let mut engine = Engine::new();
    engine
        .set_cell_value("Sheet1", "A1", Value::Error(ErrorKind::Field))
        .unwrap();
    engine
        .set_cell_value("Sheet1", "A2", Value::Error(ErrorKind::Field))
        .unwrap();
    engine
        .set_cell_value("Sheet1", "A3", Value::Error(ErrorKind::Div0))
        .unwrap();

    assert_eq!(
        eval(&mut engine, r##"=COUNTIF(A1:A3, "#FIELD!")"##),
        Value::Number(2.0)
    );
}

#[test]
fn countif_text_wildcards_match_entity_and_record_display_strings() {
    let mut engine = Engine::new();
    // Force the AST evaluator so the criteria matcher sees Entity/Record values directly.
    engine.set_bytecode_enabled(false);

    engine
        .set_cell_value("Sheet1", "A1", Value::Entity(EntityValue::new("Apple")))
        .unwrap();
    engine
        .set_cell_value("Sheet1", "A2", Value::Record(RecordValue::new("Apple")))
        .unwrap();
    engine.set_cell_value("Sheet1", "A3", "Banana").unwrap();

    assert_eq!(
        eval(&mut engine, r#"=COUNTIF(A1:A3, "*pp*")"#),
        Value::Number(2.0)
    );
}

#[test]
fn countif_accepts_entity_record_criteria_argument_as_text() {
    let mut engine = Engine::new();
    // Force the AST evaluator so the criteria parser sees Entity/Record criteria inputs directly.
    engine.set_bytecode_enabled(false);

    engine.set_cell_value("Sheet1", "A1", "Apple").unwrap();
    engine.set_cell_value("Sheet1", "A2", "Apple").unwrap();
    engine.set_cell_value("Sheet1", "A3", "Banana").unwrap();

    engine
        .set_cell_value("Sheet1", "B1", Value::Entity(EntityValue::new("Apple")))
        .unwrap();
    engine
        .set_cell_value("Sheet1", "C1", Value::Record(RecordValue::new("Apple")))
        .unwrap();

    assert_eq!(eval(&mut engine, "=COUNTIF(A1:A3, B1)"), Value::Number(2.0));
    assert_eq!(eval(&mut engine, "=COUNTIF(A1:A3, C1)"), Value::Number(2.0));
}

#[test]
fn countif_boolean_criteria() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", true).unwrap();
    engine.set_cell_value("Sheet1", "A2", false).unwrap();
    engine.set_cell_value("Sheet1", "A3", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A4", 0.0).unwrap();
    engine.set_cell_value("Sheet1", "A5", "TRUE").unwrap();
    // A6 left unset (blank).

    assert_eq!(
        eval(&mut engine, "=COUNTIF(A1:A6, TRUE)"),
        Value::Number(2.0)
    );
    assert_eq!(
        eval(&mut engine, "=COUNTIF(A1:A6, FALSE)"),
        Value::Number(3.0)
    );
}

#[test]
fn countif_numeric_criteria_does_not_treat_text_as_zero() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", "x").unwrap();
    engine.set_cell_value("Sheet1", "A2", 0.0).unwrap();
    // A3 left unset (blank) -> treated as 0 for numeric COUNTIF criteria.

    engine
        .set_cell_formula("Sheet1", "Z1", "=COUNTIF(A1:A3, 0)")
        .unwrap();
    assert!(
        engine.bytecode_program_count() > 0,
        "expected COUNTIF formula to compile to bytecode for this test"
    );
    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "Z1"), Value::Number(2.0));
}

#[test]
fn countifs_numeric_criteria_does_not_treat_text_as_zero() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", "x").unwrap();
    engine.set_cell_value("Sheet1", "A2", 0.0).unwrap();
    // A3 left unset (blank) -> treated as 0 for numeric COUNTIFS criteria.

    // Ensure the second criteria always matches so the result depends only on the numeric criteria.
    engine.set_cell_value("Sheet1", "B1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "B2", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "B3", 1.0).unwrap();

    engine
        .set_cell_formula("Sheet1", "Z1", r#"=COUNTIFS(A1:A3, 0, B1:B3, ">0")"#)
        .unwrap();
    assert!(
        engine.bytecode_program_count() > 0,
        "expected COUNTIFS formula to compile to bytecode for this test"
    );
    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "Z1"), Value::Number(2.0));
}

#[test]
fn countif_date_criteria_parses_date_strings() {
    let mut engine = Engine::new();
    engine.set_date_system(ExcelDateSystem::EXCEL_1900);

    let system = ExcelDateSystem::EXCEL_1900;
    let d2019 = ymd_to_serial(ExcelDate::new(2019, 12, 31), system).unwrap();
    let d2020 = ymd_to_serial(ExcelDate::new(2020, 1, 1), system).unwrap();
    let d2020_next = ymd_to_serial(ExcelDate::new(2020, 1, 2), system).unwrap();

    engine.set_cell_value("Sheet1", "A1", d2019 as f64).unwrap();
    engine.set_cell_value("Sheet1", "A2", d2020 as f64).unwrap();
    engine
        .set_cell_value("Sheet1", "A3", d2020_next as f64)
        .unwrap();

    assert_eq!(
        eval(&mut engine, r#"=COUNTIF(A1:A3, ">1/1/2020")"#),
        Value::Number(1.0)
    );
    assert_eq!(
        eval(&mut engine, r#"=COUNTIF(A1:A3, "=1/1/2020")"#),
        Value::Number(1.0)
    );
}

#[test]
fn countif_criteria_parses_numbers_using_value_locale() {
    let mut engine = Engine::new();
    engine.set_value_locale(ValueLocaleConfig::de_de());

    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", 3.0).unwrap();

    engine
        .set_cell_formula("Sheet1", "Z1", r#"=COUNTIF(A1:A3, ">1,5")"#)
        .unwrap();
    assert!(
        engine.bytecode_program_count() > 0,
        "expected COUNTIF formula to compile to bytecode for this test"
    );

    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "Z1"), Value::Number(2.0));
}

#[test]
fn countif_criteria_parses_numbers_using_value_locale_equality() {
    let mut engine = Engine::new();
    engine.set_value_locale(ValueLocaleConfig::de_de());
    engine.set_cell_value("Sheet1", "A1", 1.5).unwrap();
    engine.set_cell_value("Sheet1", "A2", "1,5").unwrap();
    // A3 left unset (blank).

    engine
        .set_cell_formula("Sheet1", "Z1", r#"=COUNTIF(A1:A3, "1,5")"#)
        .unwrap();
    assert!(
        engine.bytecode_program_count() > 0,
        "expected COUNTIF formula to compile to bytecode for this test"
    );

    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "Z1"), Value::Number(2.0));
}

#[test]
fn countif_text_wildcards_coerce_numbers_using_value_locale() {
    let mut engine = Engine::new();
    engine.set_value_locale(ValueLocaleConfig::de_de());
    engine.set_cell_value("Sheet1", "A1", 1.5).unwrap();
    engine.set_cell_value("Sheet1", "A2", 1.6).unwrap();
    engine
        .set_cell_formula("Sheet1", "Z1", r#"=COUNTIF(A1:A2, "*,5")"#)
        .unwrap();
    assert!(
        engine.bytecode_program_count() > 0,
        "expected COUNTIF formula to compile to bytecode for this test"
    );
    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "Z1"), Value::Number(1.0));
}

#[test]
fn countif_text_criteria_coerces_record_display_field_numbers_using_value_locale() {
    let mut engine = Engine::new();
    engine.set_value_locale(ValueLocaleConfig::de_de());

    let mut record = RecordValue::new("Fallback").field("V", 1.5);
    record.display_field = Some("V".to_string());
    engine
        .set_cell_value("Sheet1", "A1", Value::Record(record))
        .unwrap();

    // Build a criteria string of `="1,5"` so it is parsed as a *text* criteria (quoted RHS),
    // rather than a locale-aware numeric criteria.
    engine
        .set_cell_formula("Sheet1", "Z1", r#"=COUNTIF(A1:A1, "=""1,5""")"#)
        .unwrap();
    assert!(
        engine.bytecode_program_count() > 0,
        "expected COUNTIF formula to compile to bytecode for this test"
    );

    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "Z1"), Value::Number(1.0));
}

#[test]
fn countif_criteria_parses_dates_using_value_locale_date_order() {
    let mut engine = Engine::new();
    engine.set_date_system(ExcelDateSystem::EXCEL_1900);
    engine.set_value_locale(ValueLocaleConfig::de_de());

    let system = ExcelDateSystem::EXCEL_1900;
    let feb_1_2020 = ymd_to_serial(ExcelDate::new(2020, 2, 1), system).unwrap();
    engine
        .set_cell_value("Sheet1", "A1", feb_1_2020 as f64)
        .unwrap();

    // Under DMY locales like de-DE, `1/2/2020` is interpreted as 1-Feb-2020.
    assert_eq!(
        eval(&mut engine, r#"=COUNTIF(A1:A1, "1/2/2020")"#),
        Value::Number(1.0)
    );
}

#[test]
fn countif_sparse_blank_counting_counts_missing_cells() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A1048576", 2.0).unwrap();

    assert_eq!(
        eval(&mut engine, r#"=COUNTIF(A1:A1048576, "")"#),
        Value::Number(1_048_574.0)
    );
}

#[test]
fn countifs_sparse_blank_counting_counts_missing_cells() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A1048576", 2.0).unwrap();

    assert_eq!(
        eval(&mut engine, r#"=COUNTIFS(A1:A1048576, "")"#),
        Value::Number(1_048_574.0)
    );
}

#[test]
fn countifs_sparse_driver_iteration_skips_implicit_blanks() {
    let mut engine = Engine::new();
    // A1 is non-blank, A2 is implicit blank.
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();

    // B>0 matches rows 1 and 2, but only row 2 should be counted because A2 is blank.
    engine.set_cell_value("Sheet1", "B1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "B2", 1.0).unwrap();

    assert_eq!(
        eval(
            &mut engine,
            r#"=COUNTIFS(A1:A1048576, "", B1:B1048576, ">0")"#
        ),
        Value::Number(1.0)
    );
}

#[test]
fn countif_reference_union_dedupes_overlaps() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", 1.0).unwrap();

    assert_eq!(
        eval(&mut engine, "=COUNTIF((A1:A2,A2:A3), 1)"),
        Value::Number(3.0)
    );
}

#[test]
fn countif_reference_union_blank_criteria_equal_empty_string_counts_missing_cells() {
    let mut engine = Engine::new();
    // A1/A3 left unset (blank).
    engine.set_cell_value("Sheet1", "A2", "").unwrap();

    // Criteria is the string `=""`, built via concatenation.
    assert_eq!(
        eval(&mut engine, r#"=COUNTIF((A1:A2,A2:A3), "="&""""&"""")"#),
        Value::Number(3.0)
    );
}

#[test]
fn countif_reference_union_blank_criteria_equal_empty_string_literal_counts_missing_cells() {
    let mut engine = Engine::new();
    // A1/A3 left unset (blank).
    engine.set_cell_value("Sheet1", "A2", "").unwrap();

    // Same as above, but expressed as a single criteria string literal `=""` using Excel string
    // escaping (`"=""""` inside the formula source).
    assert_eq!(
        eval(&mut engine, r#"=COUNTIF((A1:A2,A2:A3), "=""""")"#),
        Value::Number(3.0)
    );
}

#[test]
fn countifs_multiple_criteria_pairs() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", 3.0).unwrap();
    engine.set_cell_value("Sheet1", "A4", -1.0).unwrap();
    engine.set_cell_value("Sheet1", "A5", 0.0).unwrap();

    engine.set_cell_value("Sheet1", "B1", "x").unwrap();
    engine.set_cell_value("Sheet1", "B2", "y").unwrap();
    engine.set_cell_value("Sheet1", "B3", "x").unwrap();
    engine.set_cell_value("Sheet1", "B4", "x").unwrap();
    engine.set_cell_value("Sheet1", "B5", "x").unwrap();

    assert_eq!(
        eval(&mut engine, r#"=COUNTIFS(A1:A5, ">0", B1:B5, "x")"#),
        Value::Number(2.0)
    );
}

#[test]
fn countifs_shape_mismatch_returns_value() {
    let mut engine = Engine::new();
    // Same number of cells, different shapes (2x2 vs 4x1).
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "B1", 3.0).unwrap();
    engine.set_cell_value("Sheet1", "B2", 4.0).unwrap();

    engine.set_cell_value("Sheet1", "C1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "C2", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "C3", 3.0).unwrap();
    engine.set_cell_value("Sheet1", "C4", 4.0).unwrap();

    assert_eq!(
        eval(&mut engine, r#"=COUNTIFS(A1:B2, ">0", C1:C4, ">0")"#),
        Value::Error(ErrorKind::Value)
    );
}

#[test]
fn countifs_odd_arg_count_returns_value() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", 3.0).unwrap();

    engine.set_cell_value("Sheet1", "B1", "x").unwrap();
    engine.set_cell_value("Sheet1", "B2", "x").unwrap();
    engine.set_cell_value("Sheet1", "B3", "x").unwrap();

    assert_eq!(
        eval(&mut engine, r#"=COUNTIFS(A1:A3, ">0", B1:B3)"#),
        Value::Error(ErrorKind::Value)
    );
}
