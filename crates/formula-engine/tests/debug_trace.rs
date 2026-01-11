use formula_engine::debug::{Span, TraceKind, TraceRef};
use formula_engine::eval::CellAddr;
use formula_engine::{Engine, ExternalValueProvider, NameDefinition, NameScope, Value};
use std::collections::HashMap;
use std::sync::{Arc, Mutex};

#[derive(Default)]
struct TestExternalProvider {
    values: Mutex<HashMap<(String, CellAddr), Value>>,
}

impl TestExternalProvider {
    fn set(&self, sheet: &str, addr: CellAddr, value: impl Into<Value>) {
        self.values
            .lock()
            .expect("lock poisoned")
            .insert((sheet.to_string(), addr), value.into());
    }
}

impl ExternalValueProvider for TestExternalProvider {
    fn get(&self, sheet: &str, addr: CellAddr) -> Option<Value> {
        self.values
            .lock()
            .expect("lock poisoned")
            .get(&(sheet.to_string(), addr))
            .cloned()
    }
}

fn slice(formula: &str, span: Span) -> &str {
    &formula[span.start..span.end]
}

#[test]
fn trace_spans_map_to_formula_and_values_match() {
    let mut engine = Engine::new();
    engine.set_cell_formula("Sheet1", "A1", "=1+2*3").unwrap();
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
            sheet: formula_engine::functions::SheetId::Local(0),
            start: formula_engine::eval::CellAddr { row: 0, col: 0 },
            end: formula_engine::eval::CellAddr { row: 1, col: 0 }
        })
    );
}

#[test]
fn debug_trace_for_vlookup_includes_reference_arg_and_matches_result() {
    let mut engine = Engine::new();
    // Lookup key.
    engine.set_cell_value("Sheet1", "A1", "Key-123").unwrap();

    // Table: B1:C2
    engine.set_cell_value("Sheet1", "B1", "Key-123").unwrap();
    engine.set_cell_value("Sheet1", "C1", 19.99).unwrap();
    engine.set_cell_value("Sheet1", "B2", "Key-456").unwrap();
    engine.set_cell_value("Sheet1", "C2", 29.99).unwrap();

    engine
        .set_cell_formula("Sheet1", "D1", "=VLOOKUP(A1,B1:C2,2,FALSE)")
        .unwrap();
    engine.recalculate();

    let dbg = engine.debug_evaluate("Sheet1", "D1").unwrap();
    assert_eq!(dbg.value, Value::Number(19.99));

    assert_eq!(
        slice(&dbg.formula, dbg.trace.span),
        "VLOOKUP(A1,B1:C2,2,FALSE)"
    );
    assert!(matches!(
        dbg.trace.kind,
        TraceKind::FunctionCall { ref name } if name == "VLOOKUP"
    ));
    assert_eq!(dbg.trace.children.len(), 4);

    // The lookup value dereferences A1 and keeps the cell reference metadata.
    let lookup_node = &dbg.trace.children[0];
    assert_eq!(slice(&dbg.formula, lookup_node.span), "A1");
    assert_eq!(lookup_node.value, Value::Text("Key-123".to_string()));
    assert!(matches!(lookup_node.reference, Some(TraceRef::Cell { .. })));

    // The table array is evaluated as a reference (not spilled/dereferenced).
    let table_node = &dbg.trace.children[1];
    assert_eq!(slice(&dbg.formula, table_node.span), "B1:C2");
    assert_eq!(table_node.value, Value::Blank);
    assert!(matches!(table_node.reference, Some(TraceRef::Range { .. })));
}

#[test]
fn trace_respects_if_short_circuiting() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=IF(TRUE,1,1/0)")
        .unwrap();
    engine.recalculate();

    let dbg = engine.debug_evaluate("Sheet1", "A1").unwrap();
    assert_eq!(dbg.value, Value::Number(1.0));
    assert_eq!(slice(&dbg.formula, dbg.trace.span), "IF(TRUE,1,1/0)");

    // The trace should include only the condition and the chosen branch.
    assert!(matches!(dbg.trace.kind, TraceKind::FunctionCall { ref name } if name == "IF"));
    assert_eq!(dbg.trace.children.len(), 2);
    assert_eq!(slice(&dbg.formula, dbg.trace.children[0].span), "TRUE");
    assert_eq!(dbg.trace.children[0].value, Value::Bool(true));
    assert_eq!(slice(&dbg.formula, dbg.trace.children[1].span), "1");
    assert_eq!(dbg.trace.children[1].value, Value::Number(1.0));
}

#[test]
fn trace_preserves_reference_context_for_named_ranges() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 2.0).unwrap();
    engine
        .define_name(
            "MyRange",
            NameScope::Workbook,
            NameDefinition::Reference("Sheet1!A1:A2".to_string()),
        )
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B1", "=SUM(MyRange)")
        .unwrap();
    engine.recalculate();

    let dbg = engine.debug_evaluate("Sheet1", "B1").unwrap();
    assert_eq!(dbg.value, Value::Number(3.0));
    assert_eq!(slice(&dbg.formula, dbg.trace.span), "SUM(MyRange)");
    assert!(matches!(dbg.trace.kind, TraceKind::FunctionCall { .. }));

    let arg_node = &dbg.trace.children[0];
    assert_eq!(slice(&dbg.formula, arg_node.span), "MyRange");
    assert!(matches!(arg_node.kind, TraceKind::NameRef { .. }));
    assert_eq!(arg_node.value, Value::Blank);
    assert_eq!(
        arg_node.reference,
        Some(TraceRef::Range {
            sheet: formula_engine::functions::SheetId::Local(0),
            start: formula_engine::eval::CellAddr { row: 0, col: 0 },
            end: formula_engine::eval::CellAddr { row: 1, col: 0 }
        })
    );
}

#[test]
fn debug_trace_supports_array_literals() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "={1,2;3,4}")
        .unwrap();

    let dbg = engine.debug_evaluate("Sheet1", "A1").unwrap();
    assert_eq!(slice(&dbg.formula, dbg.trace.span), "{1,2;3,4}");
    assert!(matches!(
        dbg.trace.kind,
        TraceKind::ArrayLiteral { rows: 2, cols: 2 }
    ));

    let Value::Array(arr) = dbg.value else {
        panic!(
            "expected Value::Array from debug evaluation, got {:?}",
            dbg.value
        );
    };
    assert_eq!(arr.rows, 2);
    assert_eq!(arr.cols, 2);
    assert_eq!(arr.get(0, 0), Some(&Value::Number(1.0)));
    assert_eq!(arr.get(0, 1), Some(&Value::Number(2.0)));
    assert_eq!(arr.get(1, 0), Some(&Value::Number(3.0)));
    assert_eq!(arr.get(1, 1), Some(&Value::Number(4.0)));
}

#[test]
fn debug_trace_supports_3d_sheet_range_refs() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet2", "A1", 2.0).unwrap();
    engine.set_cell_value("Sheet3", "A1", 3.0).unwrap();
    engine
        .set_cell_formula("Summary", "A1", "=SUM(Sheet1:Sheet3!A1)")
        .unwrap();
    engine.recalculate();

    let dbg = engine.debug_evaluate("Summary", "A1").unwrap();
    assert_eq!(dbg.value, Value::Number(6.0));
    assert_eq!(slice(&dbg.formula, dbg.trace.span), "SUM(Sheet1:Sheet3!A1)");
    assert!(matches!(
        dbg.trace.kind,
        TraceKind::FunctionCall { ref name } if name == "SUM"
    ));
    assert_eq!(dbg.trace.children.len(), 1);

    let arg = &dbg.trace.children[0];
    assert_eq!(slice(&dbg.formula, arg.span), "Sheet1:Sheet3!A1");
    assert!(matches!(arg.kind, TraceKind::CellRef));
    assert_eq!(arg.value, Value::Blank);
    assert!(
        arg.reference.is_none(),
        "3D sheet-range refs resolve to multiple sheets"
    );
}

#[test]
fn debug_trace_supports_single_quoted_3d_sheet_range_refs() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet 1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet 2", "A1", 2.0).unwrap();
    engine.set_cell_value("Sheet 3", "A1", 3.0).unwrap();
    engine
        .set_cell_formula("Summary", "A1", "=SUM('Sheet 1:Sheet 3'!A1)")
        .unwrap();
    engine.recalculate();

    let dbg = engine.debug_evaluate("Summary", "A1").unwrap();
    assert_eq!(dbg.value, Value::Number(6.0));
    assert_eq!(
        slice(&dbg.formula, dbg.trace.span),
        "SUM('Sheet 1:Sheet 3'!A1)"
    );

    let arg = &dbg.trace.children[0];
    assert_eq!(slice(&dbg.formula, arg.span), "'Sheet 1:Sheet 3'!A1");
    assert!(matches!(arg.kind, TraceKind::CellRef));
    assert_eq!(arg.value, Value::Blank);
    assert!(
        arg.reference.is_none(),
        "3D sheet-range refs resolve to multiple sheets"
    );
}

#[test]
fn debug_trace_supports_reversed_3d_sheet_range_refs() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet2", "A1", 2.0).unwrap();
    engine.set_cell_value("Sheet3", "A1", 3.0).unwrap();
    engine
        .set_cell_formula("Summary", "A1", "=SUM(Sheet3:Sheet1!A1)")
        .unwrap();
    engine.recalculate();

    let dbg = engine.debug_evaluate("Summary", "A1").unwrap();
    assert_eq!(dbg.value, Value::Number(6.0));
    assert_eq!(slice(&dbg.formula, dbg.trace.span), "SUM(Sheet3:Sheet1!A1)");

    let arg = &dbg.trace.children[0];
    assert_eq!(slice(&dbg.formula, arg.span), "Sheet3:Sheet1!A1");
    assert!(matches!(arg.kind, TraceKind::CellRef));
    assert_eq!(arg.value, Value::Blank);
    assert!(
        arg.reference.is_none(),
        "3D sheet-range refs resolve to multiple sheets"
    );
}

#[test]
fn debug_trace_supports_external_workbook_cell_refs() {
    let provider = Arc::new(TestExternalProvider::default());
    provider.set(
        "[Book.xlsx]Sheet1",
        CellAddr { row: 0, col: 0 },
        41.0,
    );

    let mut engine = Engine::new();
    engine.set_external_value_provider(Some(provider));
    engine
        .set_cell_formula("Sheet1", "A1", "=[Book.xlsx]Sheet1!A1")
        .unwrap();
    engine.recalculate();

    let computed = engine.get_cell_value("Sheet1", "A1");
    assert_eq!(computed, Value::Number(41.0));

    let dbg = engine.debug_evaluate("Sheet1", "A1").unwrap();
    assert_eq!(dbg.value, computed);
    assert_eq!(slice(&dbg.formula, dbg.trace.span), "[Book.xlsx]Sheet1!A1");
    assert!(matches!(dbg.trace.kind, TraceKind::CellRef));
    assert_eq!(
        dbg.trace.reference,
        Some(TraceRef::Cell {
            sheet: formula_engine::functions::SheetId::External("[Book.xlsx]Sheet1".to_string()),
            addr: CellAddr { row: 0, col: 0 }
        })
    );
}

#[test]
fn debug_trace_supports_unquoted_external_refs_with_non_ident_workbook_names() {
    let provider = Arc::new(TestExternalProvider::default());
    provider.set(
        "[Work Book-1.xlsx]Sheet1",
        CellAddr { row: 0, col: 0 },
        9.0,
    );

    let mut engine = Engine::new();
    engine.set_external_value_provider(Some(provider));
    engine
        .set_cell_formula("Sheet1", "A1", "=[Work Book-1.xlsx]Sheet1!A1")
        .unwrap();
    engine.recalculate();

    let dbg = engine.debug_evaluate("Sheet1", "A1").unwrap();
    assert_eq!(dbg.value, Value::Number(9.0));
    assert_eq!(
        dbg.trace.reference,
        Some(TraceRef::Cell {
            sheet: formula_engine::functions::SheetId::External(
                "[Work Book-1.xlsx]Sheet1".to_string()
            ),
            addr: CellAddr { row: 0, col: 0 }
        })
    );
}

#[test]
fn trace_preserves_reference_context_for_sum_over_external_ranges() {
    let provider = Arc::new(TestExternalProvider::default());
    provider.set(
        "[Book.xlsx]Sheet1",
        CellAddr { row: 0, col: 0 },
        1.0,
    );
    provider.set(
        "[Book.xlsx]Sheet1",
        CellAddr { row: 1, col: 0 },
        2.0,
    );

    let mut engine = Engine::new();
    engine.set_external_value_provider(Some(provider));
    engine
        .set_cell_formula("Sheet1", "B1", "=SUM([Book.xlsx]Sheet1!A1:A2)")
        .unwrap();
    engine.recalculate();

    let dbg = engine.debug_evaluate("Sheet1", "B1").unwrap();
    assert_eq!(dbg.value, Value::Number(3.0));

    let range_node = &dbg.trace.children[0];
    assert_eq!(slice(&dbg.formula, range_node.span), "[Book.xlsx]Sheet1!A1:A2");
    assert!(matches!(range_node.kind, TraceKind::RangeRef));
    assert_eq!(range_node.value, Value::Blank);
    assert_eq!(
        range_node.reference,
        Some(TraceRef::Range {
            sheet: formula_engine::functions::SheetId::External("[Book.xlsx]Sheet1".to_string()),
            start: CellAddr { row: 0, col: 0 },
            end: CellAddr { row: 1, col: 0 }
        })
    );
}

#[test]
fn debug_trace_collapses_degenerate_external_3d_sheet_spans() {
    let provider = Arc::new(TestExternalProvider::default());
    provider.set(
        "[Book.xlsx]Sheet1",
        CellAddr { row: 0, col: 0 },
        7.0,
    );

    let mut engine = Engine::new();
    engine.set_external_value_provider(Some(provider));
    engine
        .set_cell_formula("Sheet1", "A1", "=[Book.xlsx]Sheet1:Sheet1!A1")
        .unwrap();
    engine.recalculate();

    let computed = engine.get_cell_value("Sheet1", "A1");
    assert_eq!(computed, Value::Number(7.0));

    let dbg = engine.debug_evaluate("Sheet1", "A1").unwrap();
    assert_eq!(dbg.value, computed);
    assert_eq!(
        slice(&dbg.formula, dbg.trace.span),
        "[Book.xlsx]Sheet1:Sheet1!A1"
    );
    assert_eq!(
        dbg.trace.reference,
        Some(TraceRef::Cell {
            sheet: formula_engine::functions::SheetId::External("[Book.xlsx]Sheet1".to_string()),
            addr: CellAddr { row: 0, col: 0 }
        })
    );
}
