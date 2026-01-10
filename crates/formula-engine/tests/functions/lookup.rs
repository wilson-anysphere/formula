use formula_engine::functions::lookup;
use formula_engine::{ErrorKind, Value};

use super::harness::TestSheet;

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

#[test]
fn xmatch_and_xlookup_work_in_formulas_and_accept_xlfn_prefix() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", Value::Text("A".to_string()));
    sheet.set("A2", Value::Text("b".to_string()));
    sheet.set("A3", Value::Text("C".to_string()));
    sheet.set("B1", 10.0);
    sheet.set("B2", 20.0);
    sheet.set("B3", 30.0);

    assert_eq!(sheet.eval("=XMATCH(\"B\", A1:A3)"), Value::Number(2.0));
    assert_eq!(sheet.eval("=_xlfn.XMATCH(\"B\", A1:A3)"), Value::Number(2.0));

    assert_eq!(
        sheet.eval("=XLOOKUP(\"B\", A1:A3, B1:B3)"),
        Value::Number(20.0)
    );
    assert_eq!(
        sheet.eval("=_xlfn.XLOOKUP(\"B\", A1:A3, B1:B3)"),
        Value::Number(20.0)
    );

    assert_eq!(
        sheet.eval("=XLOOKUP(\"missing\", A1:A3, B1:B3, \"no\")"),
        Value::Text("no".to_string())
    );
    assert_eq!(
        sheet.eval("=XLOOKUP(\"missing\", A1:A3, B1:B3)"),
        Value::Error(ErrorKind::NA)
    );
}

#[test]
fn vlookup_exact_match_and_errors() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", 1.0);
    sheet.set("B1", Value::Text("a".to_string()));
    sheet.set("A2", 2.0);
    sheet.set("B2", Value::Text("b".to_string()));
    sheet.set("A3", 3.0);
    sheet.set("B3", Value::Text("c".to_string()));

    assert_eq!(
        sheet.eval("=VLOOKUP(2, A1:B3, 2, FALSE)"),
        Value::Text("b".to_string())
    );
    assert_eq!(sheet.eval("=VLOOKUP(4, A1:B3, 2, FALSE)"), Value::Error(ErrorKind::NA));
    assert_eq!(sheet.eval("=VLOOKUP(2, A1:B3, 3, FALSE)"), Value::Error(ErrorKind::Ref));
}

#[test]
fn vlookup_approximate_match() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", 1.0);
    sheet.set("B1", Value::Text("a".to_string()));
    sheet.set("A2", 3.0);
    sheet.set("B2", Value::Text("b".to_string()));
    sheet.set("A3", 5.0);
    sheet.set("B3", Value::Text("c".to_string()));

    assert_eq!(sheet.eval("=VLOOKUP(4, A1:B3, 2)"), Value::Text("b".to_string()));
    assert_eq!(sheet.eval("=VLOOKUP(0, A1:B3, 2)"), Value::Error(ErrorKind::NA));
}

#[test]
fn hlookup_exact_match() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", 1.0);
    sheet.set("B1", 2.0);
    sheet.set("C1", 3.0);
    sheet.set("A2", Value::Text("a".to_string()));
    sheet.set("B2", Value::Text("b".to_string()));
    sheet.set("C2", Value::Text("c".to_string()));

    assert_eq!(
        sheet.eval("=HLOOKUP(2, A1:C2, 2, FALSE)"),
        Value::Text("b".to_string())
    );
}

#[test]
fn index_and_match() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", Value::Text("A".to_string()));
    sheet.set("B1", Value::Text("b".to_string()));
    sheet.set("C1", Value::Text("C".to_string()));

    assert_eq!(sheet.eval("=INDEX(A1:C1,1,2)"), Value::Text("b".to_string()));
    assert_eq!(sheet.eval("=MATCH(\"B\", A1:C1, 0)"), Value::Number(2.0));

    sheet.set("A2", 1.0);
    sheet.set("A3", 3.0);
    sheet.set("A4", 5.0);
    sheet.set("A5", 7.0);
    assert_eq!(sheet.eval("=MATCH(4, A2:A5, 1)"), Value::Number(2.0));

    sheet.set("B2", 7.0);
    sheet.set("B3", 5.0);
    sheet.set("B4", 3.0);
    sheet.set("B5", 1.0);
    assert_eq!(sheet.eval("=MATCH(4, B2:B5, -1)"), Value::Number(2.0));
    assert_eq!(sheet.eval("=MATCH(11, B2:B5, -1)"), Value::Error(ErrorKind::NA));
}
