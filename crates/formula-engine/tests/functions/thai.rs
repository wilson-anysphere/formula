use formula_engine::locale::ValueLocaleConfig;
use formula_engine::Value;

use super::harness::{assert_number, TestSheet};

#[test]
fn bahttext_examples() {
    let mut sheet = TestSheet::new();
    assert_eq!(
        sheet.eval("=BAHTTEXT(0)"),
        Value::Text("ศูนย์บาทถ้วน".to_string())
    );
    assert_eq!(
        sheet.eval("=BAHTTEXT(1)"),
        Value::Text("หนึ่งบาทถ้วน".to_string())
    );
    assert_eq!(
        sheet.eval("=BAHTTEXT(21)"),
        Value::Text("ยี่สิบเอ็ดบาทถ้วน".to_string())
    );
    assert_eq!(
        sheet.eval("=BAHTTEXT(11.25)"),
        Value::Text("สิบเอ็ดบาทยี่สิบห้าสตางค์".to_string())
    );
    assert_eq!(
        sheet.eval("=BAHTTEXT(-11.25)"),
        Value::Text("ลบสิบเอ็ดบาทยี่สิบห้าสตางค์".to_string())
    );
}

#[test]
fn bahttext_rounds_satang_and_supports_million_groups() {
    let mut sheet = TestSheet::new();
    assert_eq!(
        sheet.eval("=BAHTTEXT(1.999)"),
        Value::Text("สองบาทถ้วน".to_string())
    );
    assert_eq!(
        sheet.eval("=BAHTTEXT(0.01)"),
        Value::Text("ศูนย์บาทหนึ่งสตางค์".to_string())
    );
    assert_eq!(
        sheet.eval("=BAHTTEXT(1000000)"),
        Value::Text("หนึ่งล้านบาทถ้วน".to_string())
    );
    assert_eq!(
        sheet.eval("=BAHTTEXT(1000001)"),
        Value::Text("หนึ่งล้านหนึ่งบาทถ้วน".to_string())
    );
}

#[test]
fn thainumstring_and_thainumsound_examples() {
    let mut sheet = TestSheet::new();
    assert_eq!(
        sheet.eval("=THAINUMSTRING(1234.5)"),
        Value::Text("๑๒๓๔.๕".to_string())
    );
    assert_eq!(
        sheet.eval("=THAINUMSTRING(-1234.5)"),
        Value::Text("-๑๒๓๔.๕".to_string())
    );
    assert_eq!(
        sheet.eval("=THAINUMSOUND(1234.5)"),
        Value::Text("หนึ่งพันสองร้อยสามสิบสี่จุดห้า".to_string())
    );
    assert_eq!(
        sheet.eval("=THAINUMSOUND(-1234.5)"),
        Value::Text("ลบหนึ่งพันสองร้อยสามสิบสี่จุดห้า".to_string())
    );
}

#[test]
fn thai_date_functions() {
    let mut sheet = TestSheet::new();
    assert_eq!(
        sheet.eval("=THAIYEAR(DATE(2020,1,1))"),
        Value::Number(2563.0)
    );
    assert_eq!(
        sheet.eval("=THAIMONTHOFYEAR(DATE(2020,1,1))"),
        Value::Text("มกราคม".to_string())
    );
    assert_eq!(
        sheet.eval("=THAIDAYOFWEEK(DATE(2020,1,1))"),
        Value::Text("วันพุธ".to_string())
    );
    assert_eq!(
        sheet.eval("=THAIDAYOFWEEK(DATE(2020,1,5))"),
        Value::Text("วันอาทิตย์".to_string())
    );
    assert_eq!(
        sheet.eval("=THAIMONTHOFYEAR(DATE(2020,12,31))"),
        Value::Text("ธันวาคม".to_string())
    );
    assert_eq!(
        sheet.eval("=THAIYEAR(DATE(1900,1,1))"),
        Value::Number(2443.0)
    );
}

#[test]
fn roundbaht_examples() {
    let mut sheet = TestSheet::new();
    assert_number(&sheet.eval("=ROUNDBAHTDOWN(1.26)"), 1.25);
    assert_number(&sheet.eval("=ROUNDBAHTUP(1.26)"), 1.5);
    assert_number(&sheet.eval("=ROUNDBAHTDOWN(-1.26)"), -1.25);
    assert_number(&sheet.eval("=ROUNDBAHTUP(-1.26)"), -1.5);
}

#[test]
fn isthaidigit_and_thaidigit_roundtrip() {
    let mut sheet = TestSheet::new();
    assert_eq!(
        sheet.eval("=THAIDIGIT(\"123\")"),
        Value::Text("๑๒๓".to_string())
    );
    assert_eq!(
        sheet.eval("=THAIDIGIT(1234)"),
        Value::Text("๑๒๓๔".to_string())
    );
    assert_eq!(
        sheet.eval("=THAIDIGIT(\"A1B2\")"),
        Value::Text("A๑B๒".to_string())
    );
    assert_eq!(
        sheet.eval("=ISTHAIDIGIT(THAIDIGIT(\"123\"))"),
        Value::Bool(true)
    );
    assert_eq!(sheet.eval("=ISTHAIDIGIT(\"๑๒๓\")"), Value::Bool(true));
    assert_eq!(sheet.eval("=ISTHAIDIGIT(\"123\")"), Value::Bool(false));
    assert_eq!(sheet.eval("=ISTHAIDIGIT(\"\")"), Value::Bool(false));
}

#[test]
fn thaidigit_coerces_numbers_using_value_locale() {
    let mut sheet = TestSheet::new();
    sheet.set_value_locale(ValueLocaleConfig::de_de());
    // de-DE numeric -> text coercion uses ',' as decimal separator.
    assert_eq!(
        sheet.eval("=THAIDIGIT(1.5)"),
        Value::Text("๑,๕".to_string())
    );
}

#[test]
fn thaistringlength_counts_graphemes() {
    let mut sheet = TestSheet::new();
    assert_eq!(sheet.eval("=THAISTRINGLENGTH(\"เก้า\")"), Value::Number(3.0));
    // Thai combining marks should not increase the grapheme cluster count.
    assert_eq!(sheet.eval("=THAISTRINGLENGTH(\"ก้\")"), Value::Number(1.0));
}
