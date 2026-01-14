use formula_engine::locale::ValueLocaleConfig;
use formula_engine::Value;

use super::harness::{assert_number, TestSheet};

#[test]
fn sumif_parses_locale_numeric_criteria_strings() {
    let mut sheet = TestSheet::new();
    sheet.set_value_locale(ValueLocaleConfig::de_de());

    sheet.set("A1", 1.4);
    sheet.set("A2", 1.6);
    sheet.set("A3", 2.0);

    assert_number(&sheet.eval(r#"=SUMIF(A1:A3,">1,5",A1:A3)"#), 3.6);
}

#[test]
fn sumif_parses_nbsp_thousands_separator_in_fr_fr_criteria_strings() {
    let mut sheet = TestSheet::new();
    sheet.set_value_locale(ValueLocaleConfig::fr_fr());

    sheet.set("A1", 1234.4);
    sheet.set("A2", 1234.6);
    sheet.set("A3", 2000.0);

    // Grouping separator: U+00A0 NO-BREAK SPACE.
    assert_number(
        &sheet.eval("=SUMIF(A1:A3,\">1\u{00A0}234,5\",A1:A3)"),
        3234.6,
    );
}

#[test]
fn sumif_parses_narrow_nbsp_thousands_separator_in_fr_fr_criteria_strings() {
    let mut sheet = TestSheet::new();
    sheet.set_value_locale(ValueLocaleConfig::fr_fr());

    sheet.set("A1", 1234.4);
    sheet.set("A2", 1234.6);
    sheet.set("A3", 2000.0);

    // Some French locales/spreadsheets use U+202F NARROW NO-BREAK SPACE for grouping.
    assert_number(
        &sheet.eval("=SUMIF(A1:A3,\">1\u{202F}234,5\",A1:A3)"),
        3234.6,
    );
}

#[test]
fn countif_parses_iso_dates_in_criteria_strings() {
    let mut sheet = TestSheet::new();

    sheet.set_formula("A1", "=DATE(2020,1,1)");
    sheet.set_formula("A2", "=DATE(2020,1,2)");
    sheet.set_formula("A3", "=DATE(2020,1,3)");

    assert_number(&sheet.eval(r#"=COUNTIF(A1:A3,">2020-01-01")"#), 2.0);
}

#[test]
fn countif_treats_invalid_numeric_or_date_rhs_as_text() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", 1.0);
    sheet.set("A2", 2.0);
    sheet.set("A3", 3.0);

    assert_eq!(
        sheet.eval(r#"=COUNTIF(A1:A3,">not-a-date")"#),
        Value::Number(0.0)
    );
}
