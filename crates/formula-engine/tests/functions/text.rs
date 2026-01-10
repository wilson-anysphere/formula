use formula_engine::functions::text;
use formula_engine::{ErrorKind, ExcelError, Value};

#[test]
fn exact_is_case_sensitive() {
    assert!(text::exact("Hello", "Hello"));
    assert!(!text::exact("Hello", "hello"));
}

#[test]
fn clean_strips_control_codes() {
    let input = "a\u{0000}\u{0009}b\u{001F}c\u{007F}d";
    assert_eq!(text::clean(input), "abcd");
}

#[test]
fn proper_capitalizes_words() {
    assert_eq!(text::proper("hello world"), "Hello World");
    assert_eq!(text::proper("hELLO wORLD"), "Hello World");
    assert_eq!(text::proper("123abc"), "123Abc");
    assert_eq!(text::proper("O'CONNOR"), "O'Connor");
}

#[test]
fn substitute_replaces_all_or_nth_instance() {
    assert_eq!(text::substitute("abab", "ab", "X", None).unwrap(), "XX");
    assert_eq!(
        text::substitute("abab", "ab", "X", Some(2)).unwrap(),
        "abX"
    );
    assert_eq!(text::substitute("abab", "ab", "X", Some(0)).unwrap_err(), ExcelError::Value);
}

#[test]
fn replace_replaces_by_character_positions() {
    assert_eq!(text::replace("abcdef", 2, 3, "X").unwrap(), "aXef");
    assert_eq!(text::replace("abc", 5, 1, "X").unwrap(), "abcX");
    assert_eq!(text::replace("abc", 0, 1, "X").unwrap_err(), ExcelError::Value);
}

#[test]
fn textjoin_concatenates_and_can_ignore_empty() {
    let values = vec![Value::from("a"), Value::Blank, Value::from(""), Value::Number(1.0)];
    assert_eq!(text::textjoin(",", true, &values).unwrap(), "a,1");
    assert_eq!(text::textjoin(",", false, &values).unwrap(), "a,,,1");
}

#[test]
fn value_and_numbervalue_parse_common_inputs() {
    assert_eq!(text::value("1,234.5").unwrap(), 1234.5);
    assert_eq!(text::value("(1,000)").unwrap(), -1000.0);
    assert_eq!(text::value("10%").unwrap(), 0.1);

    assert_eq!(
        text::numbervalue("1.234,5", Some(','), Some('.')).unwrap(),
        1234.5
    );
    assert_eq!(
        text::numbervalue("1,23", Some(','), Some(',')).unwrap_err(),
        ExcelError::Value
    );
}

#[test]
fn dollar_formats_currency() {
    assert_eq!(text::dollar(1234.567, Some(2)).unwrap(), "$1,234.57");
    assert_eq!(text::dollar(-1234.567, Some(2)).unwrap(), "($1,234.57)");
    assert_eq!(text::dollar(1234.0, Some(-1)).unwrap(), "$1,230");
}

#[test]
fn text_formats_numbers_with_simple_patterns() {
    assert_eq!(
        text::text(&Value::Number(1234.567), "#,##0.00").unwrap(),
        "1,234.57"
    );
    assert_eq!(
        text::text(&Value::Number(1.23), "0%").unwrap(),
        "123%"
    );
    assert_eq!(
        text::text(&Value::Number(-1.0), "$0.00").unwrap(),
        "-$1.00"
    );
    assert_eq!(text::text(&Value::from("x"), "0.00").unwrap(), "x");
}

#[test]
fn textjoin_propagates_errors() {
    let values = vec![Value::from("a"), Value::Error(ErrorKind::Div0), Value::from("b")];
    assert_eq!(text::textjoin(",", true, &values).unwrap_err(), ErrorKind::Div0);
}
