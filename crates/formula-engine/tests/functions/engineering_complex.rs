use formula_engine::locale::ValueLocaleConfig;
use formula_engine::{ErrorKind, Value};

use super::harness::{assert_number, TestSheet};

#[test]
fn complex_round_trip_real_imaginary() {
    let mut sheet = TestSheet::new();
    assert_eq!(sheet.eval("=COMPLEX(3,4)"), Value::Text("3+4i".to_string()));
    assert_number(&sheet.eval("=IMREAL(COMPLEX(3,4))"), 3.0);
    assert_number(&sheet.eval("=IMAGINARY(COMPLEX(3,4))"), 4.0);
}

#[test]
fn complex_parses_shorthand_and_whitespace() {
    let mut sheet = TestSheet::new();
    assert_number(&sheet.eval(r#"=IMAGINARY("i")"#), 1.0);
    assert_number(&sheet.eval(r#"=IMAGINARY("-i")"#), -1.0);
    assert_number(&sheet.eval(r#"=IMREAL("4i")"#), 0.0);
    assert_number(&sheet.eval(r#"=IMREAL(" 3 + 4i ")"#), 3.0);
    assert_number(&sheet.eval(r#"=IMAGINARY(" 3 + 4i ")"#), 4.0);
}

#[test]
fn complex_suffix_preserved_for_operations() {
    let mut sheet = TestSheet::new();
    assert_eq!(
        sheet.eval(r#"=IMSUM("1+2j","3+4i")"#),
        Value::Text("4+6j".to_string())
    );
    assert_eq!(
        sheet.eval(r#"=IMPRODUCT("1+j","1-j")"#),
        Value::Text("2".to_string())
    );
}

#[test]
fn complex_error_mapping() {
    let mut sheet = TestSheet::new();
    assert_eq!(
        sheet.eval(r#"=IMREAL("nope")"#),
        Value::Error(ErrorKind::Num)
    );
    assert_eq!(
        sheet.eval(r#"=IMREAL("3+4")"#),
        Value::Error(ErrorKind::Num)
    );
    assert_eq!(
        sheet.eval(r#"=IMREAL("1e9999+0i")"#),
        Value::Error(ErrorKind::Num)
    );
    assert_eq!(
        sheet.eval(r#"=COMPLEX(1,2,"k")"#),
        Value::Error(ErrorKind::Value)
    );
    assert_eq!(
        sheet.eval(r#"=IMDIV("1+i","0")"#),
        Value::Error(ErrorKind::Div0)
    );
    assert_eq!(
        sheet.eval(r#"=IMARGUMENT("0")"#),
        Value::Error(ErrorKind::Div0)
    );
}

#[test]
fn complex_power_and_sqrt() {
    let mut sheet = TestSheet::new();
    assert_eq!(
        sheet.eval(r#"=IMPOWER("i",2)"#),
        Value::Text("-1".to_string())
    );
    assert_eq!(sheet.eval(r#"=IMSQRT("-1")"#), Value::Text("i".to_string()));
}

#[test]
fn complex_respects_value_locale_for_parsing_and_formatting() {
    let mut sheet = TestSheet::new();
    sheet.set_value_locale(ValueLocaleConfig::de_de());

    assert_eq!(
        sheet.eval("=COMPLEX(1.5,0)"),
        Value::Text("1,5".to_string())
    );
    assert_eq!(
        sheet.eval("=COMPLEX(0,1.5)"),
        Value::Text("1,5i".to_string())
    );

    assert_number(&sheet.eval(r#"=IMREAL("1,5+0i")"#), 1.5);
    assert_number(&sheet.eval(r#"=IMREAL("1.5+0i")"#), 1.5);
}
