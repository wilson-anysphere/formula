use formula_engine::debug::{Span, TraceKind};
use formula_engine::value::RecordValue;
use formula_engine::{Engine, Value};
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
        .set_cell_value("Sheet1", "A1", Value::Record(RecordValue::with_fields("Record", fields)))
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
