use formula_engine::coercion::ValueLocaleConfig;
use formula_engine::date::{ymd_to_serial, ExcelDate, ExcelDateSystem};
use formula_engine::{Engine, Value};

fn eval_single_cell_with_date_system(
    formula: &str,
    bytecode_enabled: bool,
    locale: ValueLocaleConfig,
    date_system: ExcelDateSystem,
) -> (Value, usize) {
    let mut engine = Engine::new();
    engine.set_value_locale(locale);
    engine.set_date_system(date_system);
    engine.set_bytecode_enabled(bytecode_enabled);
    engine.set_cell_formula("Sheet1", "A1", formula).unwrap();
    engine.recalculate_single_threaded();
    (
        engine.get_cell_value("Sheet1", "A1"),
        engine.bytecode_program_count(),
    )
}

fn eval_single_cell(
    formula: &str,
    bytecode_enabled: bool,
    locale: ValueLocaleConfig,
) -> (Value, usize) {
    eval_single_cell_with_date_system(
        formula,
        bytecode_enabled,
        locale,
        ExcelDateSystem::EXCEL_1900,
    )
}

fn assert_number_close(value: &Value, expected: f64) {
    let Value::Number(n) = value else {
        panic!("expected Value::Number({expected}), got {value:?}");
    };
    assert!(
        (*n - expected).abs() < 1e-9,
        "expected number {expected}, got {n}"
    );
}

#[test]
fn bytecode_coercion_empty_string_to_number_matches_ast() {
    let formula = "=\"\"+1";

    let (ast_val, ast_programs) = eval_single_cell(formula, false, ValueLocaleConfig::en_us());
    assert_eq!(ast_programs, 0);
    assert_eq!(ast_val, Value::Number(1.0));

    let (bc_val, bc_programs) = eval_single_cell(formula, true, ValueLocaleConfig::en_us());
    assert!(bc_programs > 0, "expected formula to compile to bytecode");
    assert_eq!(bc_val, ast_val);
}

#[test]
fn bytecode_coercion_not_empty_string_matches_ast() {
    let formula = "=NOT(\"\")";

    let (ast_val, _) = eval_single_cell(formula, false, ValueLocaleConfig::en_us());
    assert_eq!(ast_val, Value::Bool(true));

    let (bc_val, bc_programs) = eval_single_cell(formula, true, ValueLocaleConfig::en_us());
    assert!(bc_programs > 0, "expected formula to compile to bytecode");
    assert_eq!(bc_val, ast_val);
}

#[test]
fn bytecode_coercion_thousands_separator_matches_ast() {
    let formula = "=\"1,234\"+1";

    let (ast_val, _) = eval_single_cell(formula, false, ValueLocaleConfig::en_us());
    assert_eq!(ast_val, Value::Number(1235.0));

    let (bc_val, bc_programs) = eval_single_cell(formula, true, ValueLocaleConfig::en_us());
    assert!(bc_programs > 0, "expected formula to compile to bytecode");
    assert_eq!(bc_val, ast_val);
}

#[test]
fn bytecode_coercion_locale_aware_decimal_and_group_separators_match_ast() {
    // de-DE uses '.' as group separator and ',' as decimal separator.
    let formula = "=\"1.234,56\"+1";

    let (ast_val, _) = eval_single_cell(formula, false, ValueLocaleConfig::de_de());
    assert_number_close(&ast_val, 1235.56);

    let (bc_val, bc_programs) = eval_single_cell(formula, true, ValueLocaleConfig::de_de());
    assert!(bc_programs > 0, "expected formula to compile to bytecode");
    assert_number_close(&bc_val, 1235.56);

    let ast_n = match ast_val {
        Value::Number(n) => n,
        other => panic!("expected Value::Number from AST, got {other:?}"),
    };
    let bc_n = match bc_val {
        Value::Number(n) => n,
        other => panic!("expected Value::Number from bytecode, got {other:?}"),
    };
    assert!(
        (bc_n - ast_n).abs() < 1e-9,
        "expected AST and bytecode values to match; ast={ast_n}, bytecode={bc_n}"
    );
}

#[test]
fn bytecode_coercion_locale_aware_decimal_separator_matches_ast() {
    // de-DE uses ',' as decimal separator.
    let formula = "=\"1,5\"+1";
    let expected = 2.5;

    let (ast_val, _) = eval_single_cell(formula, false, ValueLocaleConfig::de_de());
    assert_number_close(&ast_val, expected);

    let (bc_val, bc_programs) = eval_single_cell(formula, true, ValueLocaleConfig::de_de());
    assert!(bc_programs > 0, "expected formula to compile to bytecode");
    assert_number_close(&bc_val, expected);

    let ast_n = match ast_val {
        Value::Number(n) => n,
        other => panic!("expected Value::Number from AST, got {other:?}"),
    };
    let bc_n = match bc_val {
        Value::Number(n) => n,
        other => panic!("expected Value::Number from bytecode, got {other:?}"),
    };
    assert!(
        (bc_n - ast_n).abs() < 1e-9,
        "expected AST and bytecode values to match; ast={ast_n}, bytecode={bc_n}"
    );
}

#[test]
fn bytecode_coercion_function_args_use_value_locale_for_text_to_number() {
    let formula = "=AVERAGE(\"1,5\",1)";
    let expected = 1.25;

    let (ast_val, _) = eval_single_cell(formula, false, ValueLocaleConfig::de_de());
    assert_number_close(&ast_val, expected);

    let (bc_val, bc_programs) = eval_single_cell(formula, true, ValueLocaleConfig::de_de());
    assert!(bc_programs > 0, "expected formula to compile to bytecode");
    assert_number_close(&bc_val, expected);

    let ast_n = match ast_val {
        Value::Number(n) => n,
        other => panic!("expected Value::Number from AST, got {other:?}"),
    };
    let bc_n = match bc_val {
        Value::Number(n) => n,
        other => panic!("expected Value::Number from bytecode, got {other:?}"),
    };
    assert!(
        (bc_n - ast_n).abs() < 1e-9,
        "expected AST and bytecode values to match; ast={ast_n}, bytecode={bc_n}"
    );
}

#[test]
fn bytecode_coercion_function_args_use_value_locale_for_text_to_bool() {
    let formula = "=IFS(\"1,5\",10,TRUE,20)";

    let (ast_val, _) = eval_single_cell(formula, false, ValueLocaleConfig::de_de());
    assert_eq!(ast_val, Value::Number(10.0));

    let (bc_val, bc_programs) = eval_single_cell(formula, true, ValueLocaleConfig::de_de());
    assert!(bc_programs > 0, "expected formula to compile to bytecode");
    assert_eq!(bc_val, ast_val);
}

#[test]
fn bytecode_coercion_xor_uses_value_locale_for_text_to_bool() {
    let formula = "=XOR(\"1,5\")";

    let (ast_val, _) = eval_single_cell(formula, false, ValueLocaleConfig::de_de());
    assert_eq!(ast_val, Value::Bool(true));

    let (bc_val, bc_programs) = eval_single_cell(formula, true, ValueLocaleConfig::de_de());
    assert!(bc_programs > 0, "expected formula to compile to bytecode");
    assert_eq!(bc_val, ast_val);
}

#[test]
fn bytecode_coercion_address_args_use_value_locale_for_text_to_number() {
    let formula = "=ADDRESS(\"1,0\",\"1,0\")";

    let (ast_val, _) = eval_single_cell(formula, false, ValueLocaleConfig::de_de());
    assert_eq!(ast_val, Value::Text("$A$1".to_string()));

    let (bc_val, bc_programs) = eval_single_cell(formula, true, ValueLocaleConfig::de_de());
    assert!(bc_programs > 0, "expected formula to compile to bytecode");
    assert_eq!(bc_val, ast_val);
}

#[test]
fn bytecode_coercion_date_string_to_number_matches_ast() {
    // Text dates should coerce via DATEVALUE/TIMEVALUE-like rules in both AST and bytecode paths.
    let formula = "=\"2020-01-01\"+0";

    let (ast_val, _) = eval_single_cell(formula, false, ValueLocaleConfig::en_us());
    assert_eq!(ast_val, Value::Number(43831.0));

    let (bc_val, bc_programs) = eval_single_cell(formula, true, ValueLocaleConfig::en_us());
    assert!(bc_programs > 0, "expected formula to compile to bytecode");
    assert_eq!(bc_val, ast_val);
}

#[test]
fn bytecode_coercion_not_date_string_matches_ast() {
    let formula = "=NOT(\"2020-01-01\")";

    let (ast_val, _) = eval_single_cell(formula, false, ValueLocaleConfig::en_us());
    assert_eq!(ast_val, Value::Bool(false));

    let (bc_val, bc_programs) = eval_single_cell(formula, true, ValueLocaleConfig::en_us());
    assert!(bc_programs > 0, "expected formula to compile to bytecode");
    assert_eq!(bc_val, ast_val);
}

#[test]
fn bytecode_coercion_date_order_matches_ast() {
    let formula = "=\"1/2/2020\"+0";

    let expected_mdy =
        ymd_to_serial(ExcelDate::new(2020, 1, 2), ExcelDateSystem::EXCEL_1900).unwrap() as f64;
    let expected_dmy =
        ymd_to_serial(ExcelDate::new(2020, 2, 1), ExcelDateSystem::EXCEL_1900).unwrap() as f64;
    assert_ne!(expected_mdy, expected_dmy);

    // en-US: MDY => Jan 2, 2020
    let (ast_mdy, _) = eval_single_cell(formula, false, ValueLocaleConfig::en_us());
    assert_number_close(&ast_mdy, expected_mdy);
    let (bc_mdy, bc_programs) = eval_single_cell(formula, true, ValueLocaleConfig::en_us());
    assert!(bc_programs > 0, "expected formula to compile to bytecode");
    assert_number_close(&bc_mdy, expected_mdy);
    assert_eq!(bc_mdy, ast_mdy);

    // de-DE: DMY => Feb 1, 2020
    let (ast_dmy, _) = eval_single_cell(formula, false, ValueLocaleConfig::de_de());
    assert_number_close(&ast_dmy, expected_dmy);
    let (bc_dmy, bc_programs) = eval_single_cell(formula, true, ValueLocaleConfig::de_de());
    assert!(bc_programs > 0, "expected formula to compile to bytecode");
    assert_number_close(&bc_dmy, expected_dmy);
    assert_eq!(bc_dmy, ast_dmy);
}

#[test]
fn bytecode_coercion_respects_excel_1904_date_system() {
    let formula = "=\"2020-01-01\"+0";

    let expected =
        ymd_to_serial(ExcelDate::new(2020, 1, 1), ExcelDateSystem::Excel1904).unwrap() as f64;

    let (ast_val, _) = eval_single_cell_with_date_system(
        formula,
        false,
        ValueLocaleConfig::en_us(),
        ExcelDateSystem::Excel1904,
    );
    assert_number_close(&ast_val, expected);

    let (bc_val, bc_programs) = eval_single_cell_with_date_system(
        formula,
        true,
        ValueLocaleConfig::en_us(),
        ExcelDateSystem::Excel1904,
    );
    assert!(bc_programs > 0, "expected formula to compile to bytecode");
    assert_number_close(&bc_val, expected);
    assert_eq!(bc_val, ast_val);
}

#[test]
fn bytecode_coercion_number_to_text_matches_ast() {
    // Excel's "General" formatting switches to scientific notation for large magnitudes.
    let formula = "=CONCAT(100000000000)";

    let (ast_val, _) = eval_single_cell(formula, false, ValueLocaleConfig::en_us());
    assert_eq!(ast_val, Value::Text("1E+11".to_string()));

    let (bc_val, bc_programs) = eval_single_cell(formula, true, ValueLocaleConfig::en_us());
    assert!(bc_programs > 0, "expected formula to compile to bytecode");
    assert_eq!(bc_val, ast_val);
}

#[test]
fn bytecode_coercion_number_to_text_respects_value_locale() {
    // Locale-specific number->text coercion should match between AST and bytecode backends.
    // de-DE uses ',' for decimals.
    let formula = "=CONCAT(1.5)";

    let (ast_val, _) = eval_single_cell(formula, false, ValueLocaleConfig::de_de());
    assert_eq!(ast_val, Value::Text("1,5".to_string()));

    let (bc_val, bc_programs) = eval_single_cell(formula, true, ValueLocaleConfig::de_de());
    assert!(bc_programs > 0, "expected formula to compile to bytecode");
    assert_eq!(bc_val, ast_val);
}
