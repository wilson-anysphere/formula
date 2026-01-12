use formula_engine::value::RecordValue;
use formula_engine::{ErrorKind, Value};

use super::harness::TestSheet;

#[test]
fn image_returns_placeholder_record() {
    let mut sheet = TestSheet::new();
    let v = sheet.eval(r#"=IMAGE("https://example.com/cat.png")"#);
    assert_eq!(
        v,
        Value::Record(RecordValue::new("https://example.com/cat.png"))
    );
}

#[test]
fn image_uses_alt_text_for_display() {
    let mut sheet = TestSheet::new();
    let v = sheet.eval(r#"=IMAGE("https://example.com/cat.png","cat")"#);
    assert_eq!(v, Value::Record(RecordValue::new("cat")));
}

#[test]
fn image_accepts_xlfn_prefix() {
    let mut sheet = TestSheet::new();
    assert_eq!(
        sheet.eval(r#"=_xlfn.IMAGE("x")"#),
        sheet.eval(r#"=IMAGE("x")"#),
    );
}

#[test]
fn image_rejects_invalid_argument_counts() {
    let mut sheet = TestSheet::new();
    assert_eq!(sheet.eval("=IMAGE()"), Value::Error(ErrorKind::Value));
    assert_eq!(
        sheet.eval(r#"=IMAGE("x",1,2,3,4,5)"#),
        Value::Error(ErrorKind::Value)
    );
}

#[test]
fn image_validates_sizing_mode() {
    let mut sheet = TestSheet::new();
    assert_eq!(
        sheet.eval(r#"=IMAGE("x","",4)"#),
        Value::Error(ErrorKind::Value)
    );
    assert_eq!(
        sheet.eval(r#"=IMAGE("x","",-1)"#),
        Value::Error(ErrorKind::Value)
    );
}

#[test]
fn image_custom_sizing_requires_and_validates_dimensions() {
    let mut sheet = TestSheet::new();
    assert_eq!(
        sheet.eval(r#"=IMAGE("x","",3)"#),
        Value::Error(ErrorKind::Value)
    );
    assert_eq!(
        sheet.eval(r#"=IMAGE("x","",3,10)"#),
        Value::Error(ErrorKind::Value)
    );

    assert_eq!(
        sheet.eval(r#"=IMAGE("x","",3,0,10)"#),
        Value::Error(ErrorKind::Num)
    );
    assert_eq!(
        sheet.eval(r#"=IMAGE("x","",3,10,0)"#),
        Value::Error(ErrorKind::Num)
    );
}

#[test]
fn image_dimensions_are_ignored_when_not_custom_sizing() {
    let mut sheet = TestSheet::new();
    let v = sheet.eval(r#"=IMAGE("x","",0,-1,-1)"#);
    assert_eq!(v, Value::Record(RecordValue::new("")));
}

#[test]
fn image_placeholder_does_not_coerce_to_number() {
    let mut sheet = TestSheet::new();
    assert_eq!(sheet.eval(r#"=1+IMAGE("x")"#), Value::Error(ErrorKind::Value));
}

