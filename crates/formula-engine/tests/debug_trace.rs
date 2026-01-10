use formula_engine::debug::{Span, TraceKind, TraceRef};
use formula_engine::{Engine, Value};

fn slice(formula: &str, span: Span) -> &str {
    &formula[span.start..span.end]
}

#[test]
fn trace_spans_map_to_formula_and_values_match() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=1+2*3")
        .unwrap();
    engine.recalculate();

    let computed = engine.get_cell_value("Sheet1", "A1");
    assert_eq!(computed, Value::Number(7.0));

    let dbg = engine.debug_evaluate("Sheet1", "A1").unwrap();
    assert_eq!(dbg.value, computed);

    assert_eq!(slice(&dbg.formula, dbg.trace.span), "1+2*3");
    assert_eq!(
        dbg.trace.kind,
        TraceKind::Binary {
            op: formula_engine::eval::BinaryOp::Add
        }
    );

    // Left is `1`.
    assert_eq!(slice(&dbg.formula, dbg.trace.children[0].span), "1");
    assert_eq!(dbg.trace.children[0].value, Value::Number(1.0));

    // Right is `2*3`.
    assert_eq!(slice(&dbg.formula, dbg.trace.children[1].span), "2*3");
    assert_eq!(dbg.trace.children[1].value, Value::Number(6.0));
}

#[test]
fn trace_preserves_reference_context_for_sum() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 2.0).unwrap();
    engine
        .set_cell_formula("Sheet1", "B1", "=SUM(A1:A2)")
        .unwrap();
    engine.recalculate();

    let dbg = engine.debug_evaluate("Sheet1", "B1").unwrap();
    assert_eq!(dbg.value, Value::Number(3.0));
    assert_eq!(slice(&dbg.formula, dbg.trace.span), "SUM(A1:A2)");
    assert!(matches!(dbg.trace.kind, TraceKind::FunctionCall { .. }));

    // The range is evaluated as a reference inside SUM, so the trace keeps the range metadata
    // without forcing scalar dereference (which would yield #SPILL!).
    let range_node = &dbg.trace.children[0];
    assert_eq!(slice(&dbg.formula, range_node.span), "A1:A2");
    assert!(matches!(range_node.kind, TraceKind::RangeRef));
    assert_eq!(range_node.value, Value::Blank);
    assert_eq!(
        range_node.reference,
        Some(TraceRef::Range {
            sheet: 0,
            start: formula_engine::eval::CellAddr { row: 0, col: 0 },
            end: formula_engine::eval::CellAddr { row: 1, col: 0 }
        })
    );
}

