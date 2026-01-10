use formula_engine::functions::lookup;
use formula_engine::{ErrorKind, Value};

#[test]
fn xmatch_finds_case_insensitive_text() {
    let array = vec![Value::from("A"), Value::from("b"), Value::Number(1.0)];
    assert_eq!(lookup::xmatch(&Value::from("B"), &array).unwrap(), 2);
    assert_eq!(lookup::xmatch(&Value::Number(1.0), &array).unwrap(), 3);
    assert_eq!(lookup::xmatch(&Value::from("missing"), &array).unwrap_err(), ErrorKind::NA);
}

#[test]
fn xlookup_returns_if_not_found_when_provided() {
    let lookup_array = vec![Value::from("A"), Value::from("B")];
    let return_array = vec![Value::Number(10.0), Value::Number(20.0)];

    assert_eq!(
        lookup::xlookup(&Value::from("B"), &lookup_array, &return_array, None).unwrap(),
        Value::Number(20.0)
    );

    assert_eq!(
        lookup::xlookup(
            &Value::from("C"),
            &lookup_array,
            &return_array,
            Some(Value::from("not found"))
        )
        .unwrap(),
        Value::from("not found")
    );
}
