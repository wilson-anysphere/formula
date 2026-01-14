use formula_engine::value::RecordValue;
use formula_engine::{ErrorKind, Value};

use super::harness::TestSheet;

#[test]
fn image_returns_record_with_standard_fields() {
    let mut sheet = TestSheet::new();
    let v = sheet.eval(r#"=IMAGE("https://example.com/cat.png")"#);
    let mut expected = RecordValue::with_fields_iter(
        "https://example.com/cat.png",
        [
            (
                "source",
                Value::Text("https://example.com/cat.png".to_string()),
            ),
            ("alt_text", Value::Blank),
            ("sizing", Value::Number(0.0)),
            ("height", Value::Blank),
            ("width", Value::Blank),
        ],
    );
    expected.display_field = Some("source".to_string());
    assert_eq!(v, Value::Record(expected));
}

#[test]
fn image_uses_alt_text_for_display() {
    let mut sheet = TestSheet::new();
    let v = sheet.eval(r#"=IMAGE("https://example.com/cat.png","cat")"#);
    let mut expected = RecordValue::with_fields_iter(
        "cat",
        [
            (
                "source",
                Value::Text("https://example.com/cat.png".to_string()),
            ),
            ("alt_text", Value::Text("cat".to_string())),
            ("sizing", Value::Number(0.0)),
            ("height", Value::Blank),
            ("width", Value::Blank),
        ],
    );
    expected.display_field = Some("alt_text".to_string());
    assert_eq!(v, Value::Record(expected));
}

#[test]
fn image_record_supports_field_access() {
    let mut sheet = TestSheet::new();
    assert_eq!(
        sheet.eval(r#"=IMAGE("url","alt").source"#),
        Value::Text("url".to_string())
    );
    assert_eq!(
        sheet.eval(r#"=IMAGE("url","alt").alt_text"#),
        Value::Text("alt".to_string())
    );
    assert_eq!(
        sheet.eval(r#"=IMAGE("url","alt",2).sizing"#),
        Value::Number(2.0)
    );
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
    let mut expected = RecordValue::with_fields_iter(
        "",
        [
            ("source", Value::Text("x".to_string())),
            ("alt_text", Value::Text(String::new())),
            ("sizing", Value::Number(0.0)),
            ("height", Value::Number(-1.0)),
            ("width", Value::Number(-1.0)),
        ],
    );
    expected.display_field = Some("alt_text".to_string());
    assert_eq!(v, Value::Record(expected));
}

#[test]
fn image_record_does_not_coerce_to_number() {
    let mut sheet = TestSheet::new();
    assert_eq!(
        sheet.eval(r#"=1+IMAGE("x")"#),
        Value::Error(ErrorKind::Value)
    );
}
