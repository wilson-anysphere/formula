use formula_engine::functions::information;
use formula_engine::{Entity, ErrorKind, Record, Value};

#[test]
fn isblank_only_matches_empty_cells() {
    assert!(information::isblank(&Value::Blank));
    assert!(!information::isblank(&Value::from("")));
}

#[test]
fn iserror_isnumber_istext_behave_like_excel() {
    assert!(information::iserror(&Value::Error(ErrorKind::Div0)));
    assert!(!information::iserror(&Value::Number(1.0)));

    assert!(information::isnumber(&Value::Number(1.0)));
    assert!(!information::isnumber(&Value::from("1")));

    assert!(information::istext(&Value::from("hello")));
    assert!(!information::istext(&Value::Bool(true)));

    assert!(information::istext(&Value::Entity(Entity::new("Hello"))));
    assert!(information::istext(&Value::Record(Record::new("Hello"))));
}

#[test]
fn type_returns_excel_type_codes() {
    assert_eq!(information::r#type(&Value::Blank), 1);
    assert_eq!(information::r#type(&Value::Number(1.0)), 1);
    assert_eq!(information::r#type(&Value::from("x")), 2);
    assert_eq!(information::r#type(&Value::Entity(Entity::new("x"))), 2);
    assert_eq!(information::r#type(&Value::Record(Record::new("x"))), 2);
    assert_eq!(information::r#type(&Value::Bool(false)), 4);
    assert_eq!(information::r#type(&Value::Error(ErrorKind::NA)), 16);
}
