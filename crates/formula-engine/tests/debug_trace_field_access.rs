use formula_engine::debug::{Span, TraceKind};
use formula_engine::value::RecordValue;
use formula_engine::{Engine, ErrorKind, Value};
use std::collections::HashMap;

fn slice(formula: &str, span: Span) -> &str {
    &formula[span.start..span.end]
}

#[test]
fn debug_trace_supports_field_access() {
    let mut engine = Engine::new();

    let mut fields = HashMap::new();
    fields.insert("Price".to_string(), Value::Number(19.99));
    engine
        .set_cell_value(
            "Sheet1",
            "A1",
            Value::Record(RecordValue::with_fields("Record", fields)),
        )
        .unwrap();

    engine
        .set_cell_formula("Sheet1", "B1", "=A1.Price")
        .unwrap();
    engine.recalculate();

    let dbg = engine.debug_evaluate("Sheet1", "B1").unwrap();
    assert_eq!(dbg.value, Value::Number(19.99));

    assert_eq!(slice(&dbg.formula, dbg.trace.span), "A1.Price");
    assert!(matches!(
        dbg.trace.kind,
        TraceKind::FieldAccess { ref field } if field == "Price"
    ));
    assert_eq!(dbg.trace.children.len(), 1);
    assert_eq!(slice(&dbg.formula, dbg.trace.children[0].span), "A1");
}

#[test]
fn debug_trace_supports_bracket_field_access() {
    let mut engine = Engine::new();

    let mut fields = HashMap::new();
    fields.insert("List Price".to_string(), Value::Number(42.0));
    engine
        .set_cell_value(
            "Sheet1",
            "A1",
            Value::Record(RecordValue::with_fields("Record", fields)),
        )
        .unwrap();

    engine
        .set_cell_formula("Sheet1", "B1", "=A1.[\"List Price\"]")
        .unwrap();
    engine.recalculate();

    let computed = engine.get_cell_value("Sheet1", "B1");
    assert_eq!(computed, Value::Number(42.0));

    let dbg = engine.debug_evaluate("Sheet1", "B1").unwrap();
    assert_eq!(dbg.value, computed);
    assert_eq!(slice(&dbg.formula, dbg.trace.span), "A1.[\"List Price\"]");
    assert!(matches!(
        dbg.trace.kind,
        TraceKind::FieldAccess { ref field } if field == "List Price"
    ));
}

#[test]
fn debug_trace_field_access_on_non_record_returns_value_error() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine
        .set_cell_formula("Sheet1", "B1", "=A1.Price")
        .unwrap();
    engine.recalculate();

    let computed = engine.get_cell_value("Sheet1", "B1");
    assert_eq!(computed, Value::Error(ErrorKind::Value));

    let dbg = engine.debug_evaluate("Sheet1", "B1").unwrap();
    assert_eq!(dbg.value, computed);
    assert_eq!(slice(&dbg.formula, dbg.trace.span), "A1.Price");
    assert!(matches!(dbg.trace.kind, TraceKind::FieldAccess { .. }));
}
