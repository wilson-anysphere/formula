use formula_engine::date::ExcelDateSystem;
use formula_engine::eval::parse_a1;
use formula_engine::functions::text;
use formula_engine::locale::ValueLocaleConfig;
use formula_engine::value::{EntityValue, RecordValue};
use formula_engine::{Engine, ErrorKind, ExcelError, Value};

use super::harness::TestSheet;

#[test]
fn exact_is_case_sensitive() {
    assert!(text::exact("Hello", "Hello"));
    assert!(!text::exact("Hello", "hello"));
}

#[test]
fn clean_strips_control_codes() {
    let input = "a\u{0000}\u{0009}b\u{001F}c\u{007F}d";
    assert_eq!(text::clean(input).unwrap(), "abcd");
}

#[test]
fn proper_capitalizes_words() {
    assert_eq!(text::proper("hello world").unwrap(), "Hello World");
    assert_eq!(text::proper("hELLO wORLD").unwrap(), "Hello World");
    assert_eq!(text::proper("123abc").unwrap(), "123Abc");
    assert_eq!(text::proper("O'CONNOR").unwrap(), "O'Connor");
}

#[test]
fn substitute_replaces_all_or_nth_instance() {
    assert_eq!(text::substitute("abab", "ab", "X", None).unwrap(), "XX");
    assert_eq!(text::substitute("abab", "ab", "X", Some(2)).unwrap(), "abX");
    assert_eq!(
        text::substitute("abab", "ab", "X", Some(0)).unwrap_err(),
        ExcelError::Value
    );
}

#[test]
fn replace_replaces_by_character_positions() {
    assert_eq!(text::replace("abcdef", 2, 3, "X").unwrap(), "aXef");
    assert_eq!(text::replace("abc", 5, 1, "X").unwrap(), "abcX");
    assert_eq!(
        text::replace("abc", 0, 1, "X").unwrap_err(),
        ExcelError::Value
    );
}

#[test]
fn textjoin_concatenates_and_can_ignore_empty() {
    let values = vec![
        Value::from("a"),
        Value::Blank,
        Value::from(""),
        Value::Number(1.0),
    ];
    assert_eq!(
        text::textjoin(
            ",",
            true,
            &values,
            ExcelDateSystem::EXCEL_1900,
            ValueLocaleConfig::en_us()
        )
        .unwrap(),
        "a,1"
    );
    assert_eq!(
        text::textjoin(
            ",",
            false,
            &values,
            ExcelDateSystem::EXCEL_1900,
            ValueLocaleConfig::en_us()
        )
        .unwrap(),
        "a,,,1"
    );
}

#[test]
fn value_and_numbervalue_parse_common_inputs() {
    assert_eq!(text::value("1,234.5").unwrap(), 1234.5);
    assert_eq!(text::value("(1,000)").unwrap(), -1000.0);
    assert_eq!(text::value("10%").unwrap(), 0.1);
    assert_eq!(text::value("10%%").unwrap(), 0.001);
    assert_eq!(text::value("$1,234.50").unwrap(), 1234.5);
    assert_eq!(text::value("$ 1,234.50").unwrap(), 1234.5);
    assert_eq!(text::value("1\u{00A0}234.5").unwrap(), 1234.5);
    assert_eq!(text::value("").unwrap_err(), ExcelError::Value);
    assert_eq!(text::value("1e9999").unwrap_err(), ExcelError::Num);

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
    assert_eq!(
        text::dollar(1234.567, Some(2), ValueLocaleConfig::en_us()).unwrap(),
        "$1,234.57"
    );
    assert_eq!(
        text::dollar(-1234.567, Some(2), ValueLocaleConfig::en_us()).unwrap(),
        "($1,234.57)"
    );
    assert_eq!(
        text::dollar(1234.0, Some(-1), ValueLocaleConfig::en_us()).unwrap(),
        "$1,230"
    );
}

#[test]
fn text_formats_numbers_with_simple_patterns() {
    assert_eq!(
        text::text(
            &Value::Number(1234.567),
            "#,##0.00",
            ExcelDateSystem::EXCEL_1900,
            ValueLocaleConfig::en_us()
        )
        .unwrap(),
        "1,234.57"
    );
    assert_eq!(
        text::text(
            &Value::Number(1.23),
            "0%",
            ExcelDateSystem::EXCEL_1900,
            ValueLocaleConfig::en_us()
        )
        .unwrap(),
        "123%"
    );
    assert_eq!(
        text::text(
            &Value::Number(-1.0),
            "$0.00",
            ExcelDateSystem::EXCEL_1900,
            ValueLocaleConfig::en_us()
        )
        .unwrap(),
        "-$1.00"
    );
    assert_eq!(
        text::text(
            &Value::from("x"),
            "0.00",
            ExcelDateSystem::EXCEL_1900,
            ValueLocaleConfig::en_us()
        )
        .unwrap(),
        "x"
    );
}

#[test]
fn text_formats_entity_display_string() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", Value::Entity(EntityValue::new("hello")));
    assert_eq!(
        sheet.eval(r#"=TEXT(A1,"@")"#),
        Value::Text("hello".to_string())
    );
}

#[test]
fn text_formats_record_display_string() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", Value::Record(RecordValue::new("rec")));
    assert_eq!(
        sheet.eval(r#"=TEXT(A1,"@")"#),
        Value::Text("rec".to_string())
    );
}

#[test]
fn text_formats_entity_display_string_with_numeric_format() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", Value::Entity(EntityValue::new("hello")));
    assert_eq!(
        sheet.eval(r#"=TEXT(A1,"0.00")"#),
        Value::Text("hello".to_string())
    );
}

#[test]
fn text_formats_record_display_string_with_numeric_format() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", Value::Record(RecordValue::new("rec")));
    assert_eq!(
        sheet.eval(r#"=TEXT(A1,"0.00")"#),
        Value::Text("rec".to_string())
    );
}

#[test]
fn textjoin_includes_entity_display_string() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", Value::Entity(EntityValue::new("hello")));
    assert_eq!(
        sheet.eval(r#"=TEXTJOIN(",",TRUE,A1,"x")"#),
        Value::Text("hello,x".to_string())
    );
}

#[test]
fn textjoin_includes_record_display_string() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", Value::Record(RecordValue::new("rec")));
    assert_eq!(
        sheet.eval(r#"=TEXTJOIN(",",TRUE,A1,"x")"#),
        Value::Text("rec,x".to_string())
    );
}

#[test]
fn textjoin_ignores_empty_entity_display_string_when_requested() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", Value::Entity(EntityValue::new("")));
    assert_eq!(
        sheet.eval(r#"=TEXTJOIN(",",TRUE,A1,"x")"#),
        Value::Text("x".to_string())
    );
}

#[test]
fn textjoin_ignores_empty_record_display_string_when_requested() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", Value::Record(RecordValue::new("")));
    assert_eq!(
        sheet.eval(r#"=TEXTJOIN(",",TRUE,A1,"x")"#),
        Value::Text("x".to_string())
    );
}

#[test]
fn concat_operator_concatenates_entity_display_string() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", Value::Entity(EntityValue::new("hello")));
    assert_eq!(sheet.eval(r#"=A1&"x""#), Value::Text("hellox".to_string()));
}

#[test]
fn concat_function_concatenates_record_display_string() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", Value::Record(RecordValue::new("rec")));
    assert_eq!(
        sheet.eval(r#"=CONCAT(A1,"x")"#),
        Value::Text("recx".to_string())
    );
}

#[test]
fn textjoin_propagates_errors() {
    let values = vec![
        Value::from("a"),
        Value::Error(ErrorKind::Div0),
        Value::from("b"),
    ];
    assert_eq!(
        text::textjoin(
            ",",
            true,
            &values,
            ExcelDateSystem::EXCEL_1900,
            ValueLocaleConfig::en_us()
        )
        .unwrap_err(),
        ErrorKind::Div0
    );
}

#[test]
fn concat_and_concatenate_ranges() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", Value::Text("a".to_string()));
    sheet.set("A2", Value::Text("b".to_string()));

    assert_eq!(
        sheet.eval("=CONCAT(A1:A2, \"c\")"),
        Value::Text("abc".to_string())
    );
    assert_eq!(
        sheet.eval("=CONCATENATE(A1:A2, \"c\")"),
        Value::Text("ac".to_string())
    );
}

#[test]
fn hyperlink_returns_friendly_name_or_link_location() {
    let mut sheet = TestSheet::new();

    assert_eq!(
        sheet.eval("=HYPERLINK(\"https://example.com\")"),
        Value::Text("https://example.com".to_string())
    );
    assert_eq!(
        sheet.eval("=HYPERLINK(\"https://example.com\",\"Example\")"),
        Value::Text("Example".to_string())
    );
    assert_eq!(
        sheet.eval("=HYPERLINK(123)"),
        Value::Text("123".to_string())
    );
    // Numeric text coercion should respect the workbook value locale, like other text-producing
    // functions/operators.
    sheet.set_value_locale(ValueLocaleConfig::de_de());
    assert_eq!(
        sheet.eval("=HYPERLINK(1.5)"),
        Value::Text("1,5".to_string())
    );

    // Errors propagate from either argument.
    assert_eq!(
        sheet.eval("=HYPERLINK(1/0,\"x\")"),
        Value::Error(ErrorKind::Div0)
    );
    assert_eq!(
        sheet.eval("=HYPERLINK(\"x\",1/0)"),
        Value::Error(ErrorKind::Div0)
    );
}

#[test]
fn hyperlink_respects_value_locale_for_numeric_text_coercion() {
    let mut sheet = TestSheet::new();
    sheet.set_value_locale(ValueLocaleConfig::de_de());

    assert_eq!(
        sheet.eval("=HYPERLINK(1.5)"),
        Value::Text("1,5".to_string())
    );
    assert_eq!(
        sheet.eval("=HYPERLINK(\"https://example.com\",1.5)"),
        Value::Text("1,5".to_string())
    );
}

#[test]
fn left_right_mid_len() {
    let mut sheet = TestSheet::new();
    assert_eq!(
        sheet.eval("=LEFT(\"hello\",2)"),
        Value::Text("he".to_string())
    );
    assert_eq!(sheet.eval("=LEFT(\"hello\")"), Value::Text("h".to_string()));
    assert_eq!(
        sheet.eval("=LEFT(\"hello\",-1)"),
        Value::Error(ErrorKind::Value)
    );

    assert_eq!(
        sheet.eval("=RIGHT(\"hello\",3)"),
        Value::Text("llo".to_string())
    );

    assert_eq!(
        sheet.eval("=MID(\"hello\",2,3)"),
        Value::Text("ell".to_string())
    );
    assert_eq!(
        sheet.eval("=MID(\"hello\",6,3)"),
        Value::Text(String::new())
    );
    assert_eq!(
        sheet.eval("=MID(\"hello\",0,1)"),
        Value::Error(ErrorKind::Value)
    );

    assert_eq!(sheet.eval("=LEN(\"hello\")"), Value::Number(5.0));
}

#[test]
fn trim_upper_lower() {
    let mut sheet = TestSheet::new();
    assert_eq!(
        sheet.eval("=TRIM(\"  a   b  \")"),
        Value::Text("a b".to_string())
    );
    assert_eq!(
        sheet.eval("=TRIM(\"\ta  b\")"),
        Value::Text("\ta b".to_string())
    );
    assert_eq!(
        sheet.eval("=UPPER(\"Abc\")"),
        Value::Text("ABC".to_string())
    );
    assert_eq!(
        sheet.eval("=LOWER(\"AbC\")"),
        Value::Text("abc".to_string())
    );
}

#[test]
fn find_and_search() {
    let mut sheet = TestSheet::new();
    assert_eq!(sheet.eval("=FIND(\"b\",\"abc\")"), Value::Number(2.0));
    assert_eq!(
        sheet.eval("=FIND(\"B\",\"abc\")"),
        Value::Error(ErrorKind::Value)
    );
    assert_eq!(
        sheet.eval("=FIND(\"Ö\",\"ö\")"),
        Value::Error(ErrorKind::Value)
    );
    assert_eq!(sheet.eval("=SEARCH(\"B\",\"abc\")"), Value::Number(2.0));
    assert_eq!(sheet.eval("=SEARCH(\"Ö\",\"ö\")"), Value::Number(1.0));
    assert_eq!(sheet.eval("=SEARCH(\"A\",\"aö\")"), Value::Number(1.0));

    assert_eq!(sheet.eval("=SEARCH(\"a?c\",\"abc\")"), Value::Number(1.0));
    assert_eq!(
        sheet.eval("=SEARCH(\"a*c\",\"abbbbbc\")"),
        Value::Number(1.0)
    );
    assert_eq!(
        sheet.eval("=SEARCH(\"a**c\",\"abbbbbc\")"),
        Value::Number(1.0)
    );
    assert_eq!(sheet.eval("=SEARCH(\"~*\",\"a*b\")"), Value::Number(2.0));
    assert_eq!(sheet.eval("=SEARCH(\"~?\",\"a?b\")"), Value::Number(2.0));
    assert_eq!(sheet.eval("=SEARCH(\"~\",\"a~b\")"), Value::Number(2.0));
    assert_eq!(sheet.eval("=SEARCH(\"SS\",\"ß\")"), Value::Number(1.0));
    assert_eq!(
        sheet.eval("=SEARCH(\"b\",\"abc\",3)"),
        Value::Error(ErrorKind::Value)
    );
}

#[test]
fn substitute_worksheet_function_replaces_all_or_nth_instance() {
    let mut sheet = TestSheet::new();
    assert_eq!(
        sheet.eval("=SUBSTITUTE(\"foo bar foo\",\"foo\",\"x\")"),
        Value::Text("x bar x".to_string())
    );
    assert_eq!(
        sheet.eval("=SUBSTITUTE(\"abab\",\"ab\",\"X\",2)"),
        Value::Text("abX".to_string())
    );
    assert_eq!(
        sheet.eval("=SUBSTITUTE(\"abab\",\"ab\",\"X\",0)"),
        Value::Error(ErrorKind::Value)
    );
}

#[test]
fn substitute_accepts_xlfn_prefix() {
    let mut sheet = TestSheet::new();
    assert_eq!(
        sheet.eval("=_xlfn.SUBSTITUTE(\"foo bar foo\",\"foo\",\"x\")"),
        Value::Text("x bar x".to_string())
    );
}

#[test]
fn value_and_numbervalue_worksheet_functions_parse_common_inputs() {
    let mut sheet = TestSheet::new();
    assert_eq!(sheet.eval(r#"=VALUE("1,234.5")"#), Value::Number(1234.5));
    assert_eq!(
        sheet.eval(r#"=NUMBERVALUE("1.234,5", ",", ".")"#),
        Value::Number(1234.5)
    );
}

#[test]
fn numbervalue_worksheet_function_separator_semantics() {
    let mut sheet = TestSheet::new();

    assert_eq!(
        sheet.eval(r#"=NUMBERVALUE("1 234,5", ",", " ")"#),
        Value::Number(1234.5)
    );
    assert_eq!(
        sheet.eval(r#"=NUMBERVALUE("1234,5", ",", "")"#),
        Value::Number(1234.5)
    );

    assert_eq!(
        sheet.eval(r#"=NUMBERVALUE("1,23", ",", ",")"#),
        Value::Error(ErrorKind::Value)
    );
    assert_eq!(
        sheet.eval(r#"=NUMBERVALUE("1", "", ",")"#),
        Value::Error(ErrorKind::Value)
    );
    assert_eq!(
        sheet.eval(r#"=NUMBERVALUE("1", "..", ",")"#),
        Value::Error(ErrorKind::Value)
    );
}

#[test]
fn numbervalue_defaults_to_engine_value_locale_separators() {
    let mut engine = Engine::new();
    engine.set_value_locale(ValueLocaleConfig::de_de());
    engine
        .set_cell_formula("Sheet1", "A1", r#"=NUMBERVALUE("1.234,5")"#)
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(1234.5));
}

#[test]
fn value_locale_affects_concat_text_coercion() {
    let mut sheet = TestSheet::new();
    sheet.set_value_locale(ValueLocaleConfig::de_de());
    assert_eq!(sheet.eval("=1.5&\"\""), Value::Text("1,5".to_string()));
}

#[test]
fn value_locale_affects_text_formatting() {
    let mut sheet = TestSheet::new();
    sheet.set_value_locale(ValueLocaleConfig::de_de());
    assert_eq!(
        sheet.eval(r##"=TEXT(1234.567,"#,##0.00")"##),
        Value::Text("1.234,57".to_string())
    );
}

#[test]
fn value_locale_affects_dollar_formatting() {
    let mut sheet = TestSheet::new();
    sheet.set_value_locale(ValueLocaleConfig::de_de());
    assert_eq!(
        sheet.eval("=DOLLAR(1234.567,2)"),
        Value::Text("$1.234,57".to_string())
    );
}

#[test]
fn text_and_dollar_worksheet_functions_format_values() {
    let mut sheet = TestSheet::new();
    assert_eq!(
        sheet.eval(r##"=TEXT(1234.567,"#,##0.00")"##),
        Value::Text("1,234.57".to_string())
    );
    assert_eq!(
        sheet.eval(r#"=DOLLAR(-1234.567,2)"#),
        Value::Text("($1,234.57)".to_string())
    );
}

#[test]
fn text_supports_sections_conditions_and_text_placeholders() {
    let mut sheet = TestSheet::new();

    // pos;neg;zero;text
    let fmt = r#""0.00;(0.00);""zero"";""text:""@""#;
    assert_eq!(
        sheet.eval(&format!("=TEXT(1.2,{fmt})")),
        Value::Text("1.20".to_string())
    );
    assert_eq!(
        sheet.eval(&format!("=TEXT(-1.2,{fmt})")),
        Value::Text("(1.20)".to_string())
    );
    assert_eq!(
        sheet.eval(&format!("=TEXT(0,{fmt})")),
        Value::Text("zero".to_string())
    );
    assert_eq!(
        sheet.eval(&format!(r#"=TEXT("hi",{fmt})"#)),
        Value::Text("text:hi".to_string())
    );

    // Conditions: first matching conditional section, else first unconditional.
    assert_eq!(
        sheet.eval(r#"=TEXT(-1,"[<0]""neg"";""pos""")"#),
        Value::Text("neg".to_string())
    );
    assert_eq!(
        sheet.eval(r#"=TEXT(1,"[<0]""neg"";""pos""")"#),
        Value::Text("pos".to_string())
    );

    // `@` placeholder can appear in a non-4th section and still apply to text.
    assert_eq!(
        sheet.eval(r#"=TEXT("x","""pre-""@")"#),
        Value::Text("pre-x".to_string())
    );
}

#[test]
fn text_formats_dates_using_workbook_date_system() {
    let mut sheet = TestSheet::new();

    sheet.set_date_system(ExcelDateSystem::EXCEL_1900);
    assert_eq!(
        sheet.eval(r#"=TEXT(1,"m/d/yyyy")"#),
        Value::Text("1/1/1900".to_string())
    );

    sheet.set_date_system(ExcelDateSystem::Excel1904);
    assert_eq!(
        sheet.eval(r#"=TEXT(1,"m/d/yyyy")"#),
        Value::Text("1/2/1904".to_string())
    );
}

#[test]
fn text_empty_format_code_falls_back_to_general() {
    let mut sheet = TestSheet::new();
    assert_eq!(
        sheet.eval(r#"=TEXT(1234.5,"")"#),
        Value::Text("1234.5".to_string())
    );
}

#[test]
fn text_spills_arrays_elementwise() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", r#"=TEXT({1;2},"0")"#)
        .unwrap();
    engine.recalculate_single_threaded();

    assert_eq!(
        engine.get_cell_value("Sheet1", "A1"),
        Value::Text("1".to_string())
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "A2"),
        Value::Text("2".to_string())
    );
}

#[test]
fn textjoin_flattens_ranges_and_array_literals() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", Value::Text("a".to_string()));
    sheet.set("A2", Value::Blank);
    sheet.set("A3", Value::Text("b".to_string()));

    assert_eq!(
        sheet.eval(r#"=TEXTJOIN(",", TRUE, A1:A3, {"x",""})"#),
        Value::Text("a,b,x".to_string())
    );
    assert_eq!(
        sheet.eval(r#"=TEXTJOIN(",", FALSE, A1:A3, {"x",""})"#),
        Value::Text("a,,b,x,".to_string())
    );
}

#[test]
fn textsplit_splits_into_columns() {
    let mut sheet = TestSheet::new();
    sheet.set_formula("Z1", r#"=TEXTSPLIT("a,b,c", ",")"#);
    sheet.recalc();

    assert_eq!(sheet.get("Z1"), Value::Text("a".to_string()));
    assert_eq!(sheet.get("AA1"), Value::Text("b".to_string()));
    assert_eq!(sheet.get("AB1"), Value::Text("c".to_string()));
}

#[test]
fn textsplit_splits_rows_and_columns_and_pads_with_na() {
    let mut sheet = TestSheet::new();
    sheet.set_formula("Z1", r#"=TEXTSPLIT("a,b;c", ",", ";")"#);
    sheet.recalc();

    assert_eq!(sheet.get("Z1"), Value::Text("a".to_string()));
    assert_eq!(sheet.get("AA1"), Value::Text("b".to_string()));
    assert_eq!(sheet.get("Z2"), Value::Text("c".to_string()));
    assert_eq!(sheet.get("AA2"), Value::Error(ErrorKind::NA));
}

#[test]
fn textsplit_respects_ignore_empty() {
    let mut sheet = TestSheet::new();
    sheet.set_formula("Z1", r#"=TEXTSPLIT("a,,b", ",", , TRUE)"#);
    sheet.recalc();

    assert_eq!(sheet.get("Z1"), Value::Text("a".to_string()));
    assert_eq!(sheet.get("AA1"), Value::Text("b".to_string()));
}

#[test]
fn textsplit_respects_value_locale_when_coercing_numeric_text() {
    let mut sheet = TestSheet::new();
    sheet.set_value_locale(ValueLocaleConfig::de_de());
    sheet.set_formula("Z1", r#"=TEXTSPLIT(1.5, ",")"#);
    sheet.recalc();

    // de-DE uses ',' as decimal separator, so TEXTSPLIT sees the coerced text "1,5".
    assert_eq!(sheet.get("Z1"), Value::Text("1".to_string()));
    assert_eq!(sheet.get("AA1"), Value::Text("5".to_string()));
}

#[test]
fn textsplit_accepts_xlfn_prefix() {
    let mut sheet = TestSheet::new();
    sheet.set_formula("Z1", r#"=_xlfn.TEXTSPLIT("a,b", ",")"#);
    sheet.recalc();

    assert_eq!(sheet.get("Z1"), Value::Text("a".to_string()));
    assert_eq!(sheet.get("AA1"), Value::Text("b".to_string()));
}

#[test]
fn clean_exact_proper_replace_worksheet_functions() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", Value::Text(format!("a\u{0000}b")));

    assert_eq!(sheet.eval("=CLEAN(A1)"), Value::Text("ab".to_string()));
    assert_eq!(sheet.eval(r#"=EXACT("Hello","hello")"#), Value::Bool(false));
    assert_eq!(
        sheet.eval(r#"=PROPER("hELLO wORLD")"#),
        Value::Text("Hello World".to_string())
    );
    assert_eq!(
        sheet.eval(r#"=REPLACE("abcdef",2,3,"X")"#),
        Value::Text("aXef".to_string())
    );
}

#[test]
fn value_spills_arrays_elementwise() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", r#"=VALUE({"1";"2"})"#)
        .unwrap();
    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "A2"), Value::Number(2.0));
}

#[test]
fn value_on_scalar_reference_returns_scalar() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", "1").unwrap();
    engine
        .set_cell_formula("Sheet1", "B1", "=VALUE(A1)")
        .unwrap();
    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(1.0));
    assert!(engine.spill_range("Sheet1", "B1").is_none());
}

#[test]
fn value_spills_singleton_array_literal() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", r#"=VALUE({"1"})"#)
        .unwrap();
    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(1.0));
    assert_eq!(
        engine.spill_range("Sheet1", "A1"),
        Some((parse_a1("A1").unwrap(), parse_a1("A1").unwrap()))
    );
}
