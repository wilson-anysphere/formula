use formula_engine::coercion::ValueLocaleConfig;
use formula_engine::{Engine, Value};

fn eval_single_cell(formula: &str, bytecode_enabled: bool, locale: ValueLocaleConfig) -> (Value, usize) {
    let mut engine = Engine::new();
    engine.set_value_locale(locale);
    engine.set_bytecode_enabled(bytecode_enabled);
    engine.set_cell_formula("Sheet1", "A1", formula).unwrap();
    engine.recalculate_single_threaded();
    (engine.get_cell_value("Sheet1", "A1"), engine.bytecode_program_count())
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
