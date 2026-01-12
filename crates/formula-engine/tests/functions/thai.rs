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
}

#[test]
fn thainumstring_and_thainumsound_examples() {
    let mut sheet = TestSheet::new();
    assert_eq!(
        sheet.eval("=THAINUMSTRING(1234.5)"),
        Value::Text("๑๒๓๔.๕".to_string())
    );
    assert_eq!(
        sheet.eval("=THAINUMSOUND(1234.5)"),
        Value::Text("หนึ่งพันสองร้อยสามสิบสี่จุดห้า".to_string())
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
    assert_eq!(sheet.eval("=THAIDIGIT(\"123\")"), Value::Text("๑๒๓".to_string()));
    assert_eq!(sheet.eval("=ISTHAIDIGIT(THAIDIGIT(\"123\"))"), Value::Bool(true));
    assert_eq!(sheet.eval("=ISTHAIDIGIT(\"123\")"), Value::Bool(false));
}

#[test]
fn thaistringlength_counts_graphemes() {
    let mut sheet = TestSheet::new();
    assert_eq!(sheet.eval("=THAISTRINGLENGTH(\"เก้า\")"), Value::Number(3.0));
}

