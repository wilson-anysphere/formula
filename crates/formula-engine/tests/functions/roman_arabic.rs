use super::harness::{assert_number, TestSheet};
use formula_engine::value::{ErrorKind, Value};

#[test]
fn roman_known_conversions() {
    let mut sheet = TestSheet::new();
    assert_eq!(
        sheet.eval("=ROMAN(1999)"),
        Value::Text("MCMXCIX".to_string())
    );
    assert_eq!(sheet.eval("=ROMAN(0)"), Value::Text("".to_string()));
}

#[test]
fn roman_form_variants_match_excel_docs() {
    let mut sheet = TestSheet::new();
    assert_eq!(
        sheet.eval("=ROMAN(499,0)"),
        Value::Text("CDXCIX".to_string())
    );
    assert_eq!(
        sheet.eval("=ROMAN(499,1)"),
        Value::Text("LDVLIV".to_string())
    );
    assert_eq!(sheet.eval("=ROMAN(499,2)"), Value::Text("XDIX".to_string()));
    assert_eq!(sheet.eval("=ROMAN(499,3)"), Value::Text("VDIV".to_string()));
    assert_eq!(sheet.eval("=ROMAN(499,4)"), Value::Text("ID".to_string()));
}

#[test]
fn roman_truncates_number_and_form() {
    let mut sheet = TestSheet::new();
    // Excel-style coercion: truncate both arguments.
    assert_eq!(
        sheet.eval("=ROMAN(499.9,1.9)"),
        Value::Text("LDVLIV".to_string())
    );
}

#[test]
fn roman_invalid_inputs() {
    let mut sheet = TestSheet::new();
    assert_eq!(sheet.eval("=ROMAN(-1)"), Value::Error(ErrorKind::Value));
    assert_eq!(sheet.eval("=ROMAN(4000)"), Value::Error(ErrorKind::Value));
    assert_eq!(sheet.eval("=ROMAN(1,5)"), Value::Error(ErrorKind::Value));
}

#[test]
fn arabic_parses_roman_strings() {
    let mut sheet = TestSheet::new();
    assert_number(&sheet.eval("=ARABIC(\"MCMXCIX\")"), 1999.0);
    assert_number(&sheet.eval("=ARABIC(\"id\")"), 499.0);
    assert_number(&sheet.eval("=ARABIC(\"  MCMXCIX  \")"), 1999.0);
    assert_number(&sheet.eval("=ARABIC(ROMAN(499,4))"), 499.0);
    assert_number(&sheet.eval("=ARABIC(\"\")"), 0.0);
}

#[test]
fn arabic_rejects_invalid_numerals() {
    let mut sheet = TestSheet::new();
    assert_eq!(
        sheet.eval("=ARABIC(\"VV\")"),
        Value::Error(ErrorKind::Value)
    );
    assert_eq!(
        sheet.eval("=ARABIC(\"IIV\")"),
        Value::Error(ErrorKind::Value)
    );
}
