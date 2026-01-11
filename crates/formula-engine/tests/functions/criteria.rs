use formula_engine::date::ExcelDateSystem;
use formula_engine::functions::math::criteria::Criteria;
use formula_engine::simd::CmpOp;
use formula_engine::{ErrorKind, Value};

#[test]
fn criteria_blank_semantics() {
    let c = Criteria::parse(&Value::from("")).unwrap();
    assert!(c.matches(&Value::Blank));
    assert!(c.matches(&Value::from("")));
    assert!(!c.matches(&Value::from("x")));
    assert!(!c.matches(&Value::Number(0.0)));

    let c = Criteria::parse(&Value::from("=")).unwrap();
    assert!(c.matches(&Value::Blank));
    assert!(c.matches(&Value::from("")));
    assert!(!c.matches(&Value::from("x")));

    let c = Criteria::parse(&Value::from("<>")).unwrap();
    assert!(!c.matches(&Value::Blank));
    assert!(!c.matches(&Value::from("")));
    assert!(c.matches(&Value::from("x")));
    assert!(c.matches(&Value::Number(0.0)));
    // Errors never match non-error criteria.
    assert!(!c.matches(&Value::Error(ErrorKind::Div0)));
}

#[test]
fn criteria_wildcard_escapes() {
    // "~a" is a literal "~a" (tilde only escapes *, ?, and ~).
    let c = Criteria::parse(&Value::from("~a")).unwrap();
    assert!(c.matches(&Value::from("~A")));
    assert!(!c.matches(&Value::from("a")));

    // "~~" is a literal "~".
    let c = Criteria::parse(&Value::from("~~")).unwrap();
    assert!(c.matches(&Value::from("~")));
    assert!(!c.matches(&Value::from("~~")));

    // "~*" is a literal "*".
    let c = Criteria::parse(&Value::from("~*")).unwrap();
    assert!(c.matches(&Value::from("*")));
    assert!(!c.matches(&Value::from("anything")));

    // "~?" is a literal "?".
    let c = Criteria::parse(&Value::from("~?")).unwrap();
    assert!(c.matches(&Value::from("?")));
    assert!(!c.matches(&Value::from("a")));
}

#[test]
fn criteria_wildcard_matching() {
    let c = Criteria::parse(&Value::from("a?c")).unwrap();
    assert!(c.matches(&Value::from("abc")));
    assert!(!c.matches(&Value::from("ac")));
    assert!(!c.matches(&Value::from("abbc")));

    let c = Criteria::parse(&Value::from("a*c")).unwrap();
    assert!(c.matches(&Value::from("ac")));
    assert!(c.matches(&Value::from("abbbbbc")));
    assert!(!c.matches(&Value::from("abbbbbd")));
}

#[test]
fn criteria_is_case_insensitive_unicode() {
    // Uses Unicode-aware uppercasing: ß -> SS.
    let c = Criteria::parse(&Value::from("straße")).unwrap();
    assert!(c.matches(&Value::from("STRASSE")));
    assert!(c.matches(&Value::from("Straße")));
}

#[test]
fn criteria_parses_bool_keywords() {
    let c = Criteria::parse(&Value::from("TRUE")).unwrap();
    assert!(c.matches(&Value::Bool(true)));
    assert!(c.matches(&Value::Number(1.0)));
    assert!(!c.matches(&Value::Number(2.0)));
    assert!(!c.matches(&Value::Blank));
    // Literal text "TRUE" does not match the boolean criteria.
    assert!(!c.matches(&Value::from("TRUE")));
}

#[test]
fn criteria_parses_error_tokens() {
    let c = Criteria::parse(&Value::from("#DIV/0!")).unwrap();
    assert!(c.matches(&Value::Error(ErrorKind::Div0)));
    assert!(!c.matches(&Value::Error(ErrorKind::Value)));
    assert!(!c.matches(&Value::Number(0.0)));

    let c = Criteria::parse(&Value::from("<>#DIV/0!")).unwrap();
    assert!(!c.matches(&Value::Error(ErrorKind::Div0)));
    assert!(c.matches(&Value::Error(ErrorKind::Value)));
    assert!(c.matches(&Value::Number(123.0)));
}

#[test]
fn criteria_parses_scientific_numbers() {
    let c = Criteria::parse(&Value::from("1e2")).unwrap();
    let numeric = c.as_numeric_criteria().expect("should be numeric criteria");
    assert_eq!(numeric.op, CmpOp::Eq);
    assert_eq!(numeric.rhs, 100.0);
}

#[test]
fn criteria_numeric_coercion_matches_text_and_blanks() {
    let c = Criteria::parse(&Value::from("=1")).unwrap();
    assert!(c.matches(&Value::Number(1.0)));
    assert!(c.matches(&Value::from("1")));
    assert!(c.matches(&Value::from("  1  ")));

    // Criteria functions treat blank / empty string as zero for numeric comparisons.
    let c = Criteria::parse(&Value::from("0")).unwrap();
    assert!(c.matches(&Value::Blank));
    assert!(c.matches(&Value::from("")));
    assert!(c.matches(&Value::Number(0.0)));

    let c = Criteria::parse(&Value::from(">0")).unwrap();
    assert!(!c.matches(&Value::Blank));
    assert!(!c.matches(&Value::from("")));
    assert!(c.matches(&Value::Number(0.1)));
    assert!(c.matches(&Value::from("2")));
}

#[test]
fn criteria_parses_dates_and_times_with_workbook_date_system() {
    let c =
        Criteria::parse_with_date_system(&Value::from("1900-01-01"), ExcelDateSystem::EXCEL_1900)
            .unwrap();
    let numeric = c.as_numeric_criteria().unwrap();
    assert_eq!(numeric.rhs, 1.0);

    let c = Criteria::parse_with_date_system(&Value::from("12:00"), ExcelDateSystem::EXCEL_1900)
        .unwrap();
    let numeric = c.as_numeric_criteria().unwrap();
    assert_eq!(numeric.rhs, 0.5);

    let c = Criteria::parse_with_date_system(
        &Value::from("1900-01-01 12:00"),
        ExcelDateSystem::EXCEL_1900,
    )
    .unwrap();
    let numeric = c.as_numeric_criteria().unwrap();
    assert_eq!(numeric.rhs, 1.5);

    let c =
        Criteria::parse_with_date_system(&Value::from("1904-01-01"), ExcelDateSystem::Excel1904)
            .unwrap();
    let numeric = c.as_numeric_criteria().unwrap();
    assert_eq!(numeric.rhs, 0.0);
}
