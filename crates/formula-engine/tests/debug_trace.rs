use formula_engine::date::ExcelDateSystem;
use formula_engine::debug::{Span, TraceKind, TraceRef};
use formula_engine::eval::CellAddr;
use formula_engine::{
    Engine, ExternalDataProvider, ExternalValueProvider, NameDefinition, NameScope, Value,
};
use formula_model::table::TableColumn;
use formula_model::{Range, Table};
use std::collections::HashMap;
use std::sync::{Arc, Mutex};

#[derive(Default)]
struct TestExternalProvider {
    values: Mutex<HashMap<(String, CellAddr), Value>>,
    tables: Mutex<HashMap<(String, String), (String, Table)>>,
    sheet_order: Mutex<HashMap<String, Vec<String>>>,
}

impl TestExternalProvider {
    fn set(&self, sheet: &str, addr: CellAddr, value: impl Into<Value>) {
        self.values
            .lock()
            .expect("lock poisoned")
            .insert((sheet.to_string(), addr), value.into());
    }

    fn set_table(&self, workbook: &str, sheet: &str, table: Table) {
        self.tables.lock().expect("lock poisoned").insert(
            (workbook.to_string(), table.name.clone()),
            (sheet.to_string(), table),
        );
    }

    fn set_sheet_order(&self, workbook: &str, order: &[&str]) {
        self.sheet_order.lock().expect("lock poisoned").insert(
            workbook.to_string(),
            order.iter().map(|s| s.to_string()).collect(),
        );
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

    fn workbook_table(&self, workbook: &str, table_name: &str) -> Option<(String, Table)> {
        self.tables
            .lock()
            .expect("lock poisoned")
            .get(&(workbook.to_string(), table_name.to_string()))
            .cloned()
    }

    fn sheet_order(&self, workbook: &str) -> Option<Vec<String>> {
        self.sheet_order
            .lock()
            .expect("lock poisoned")
            .get(workbook)
            .cloned()
    }
}

#[derive(Default)]
struct TestExternalDataProvider {
    calls: Mutex<Vec<(String, Vec<Value>)>>,
}

impl TestExternalDataProvider {
    fn record_call(&self, function: &str, args: Vec<Value>) {
        self.calls
            .lock()
            .expect("lock poisoned")
            .push((function.to_string(), args));
    }

    fn calls(&self) -> Vec<(String, Vec<Value>)> {
        self.calls.lock().expect("lock poisoned").clone()
    }
}

impl ExternalDataProvider for TestExternalDataProvider {
    fn rtd(&self, prog_id: &str, server: &str, topics: &[String]) -> Value {
        let mut args = vec![
            Value::Text(prog_id.to_string()),
            Value::Text(server.to_string()),
        ];
        args.extend(topics.iter().cloned().map(Value::Text));
        self.record_call("RTD", args);
        Value::Number(123.0)
    }

    fn cube_value(&self, connection: &str, tuples: &[String]) -> Value {
        let mut args = vec![Value::Text(connection.to_string())];
        args.extend(tuples.iter().cloned().map(Value::Text));
        self.record_call("CUBEVALUE", args);
        Value::Number(456.0)
    }

    fn cube_member(
        &self,
        connection: &str,
        member_expression: &str,
        caption: Option<&str>,
    ) -> Value {
        let mut args = vec![
            Value::Text(connection.to_string()),
            Value::Text(member_expression.to_string()),
        ];
        if let Some(caption) = caption {
            args.push(Value::Text(caption.to_string()));
        }
        self.record_call("CUBEMEMBER", args);
        Value::Error(formula_engine::ErrorKind::NA)
    }

    fn cube_member_property(
        &self,
        connection: &str,
        member_expression_or_handle: &str,
        property: &str,
    ) -> Value {
        self.record_call(
            "CUBEMEMBERPROPERTY",
            vec![
                Value::Text(connection.to_string()),
                Value::Text(member_expression_or_handle.to_string()),
                Value::Text(property.to_string()),
            ],
        );
        Value::Error(formula_engine::ErrorKind::NA)
    }

    fn cube_ranked_member(
        &self,
        connection: &str,
        set_expression_or_handle: &str,
        rank: i64,
        caption: Option<&str>,
    ) -> Value {
        let mut args = vec![
            Value::Text(connection.to_string()),
            Value::Text(set_expression_or_handle.to_string()),
            Value::Number(rank as f64),
        ];
        if let Some(caption) = caption {
            args.push(Value::Text(caption.to_string()));
        }
        self.record_call("CUBERANKEDMEMBER", args);
        Value::Number(rank as f64)
    }

    fn cube_set(
        &self,
        connection: &str,
        set_expression: &str,
        caption: Option<&str>,
        sort_order: Option<i64>,
        sort_by: Option<&str>,
    ) -> Value {
        let mut args = vec![
            Value::Text(connection.to_string()),
            Value::Text(set_expression.to_string()),
        ];
        if let Some(caption) = caption {
            args.push(Value::Text(caption.to_string()));
        }
        if let Some(sort_order) = sort_order {
            args.push(Value::Number(sort_order as f64));
        }
        if let Some(sort_by) = sort_by {
            args.push(Value::Text(sort_by.to_string()));
        }
        self.record_call("CUBESET", args);
        Value::Error(formula_engine::ErrorKind::NA)
    }

    fn cube_set_count(&self, set_expression_or_handle: &str) -> Value {
        self.record_call(
            "CUBESETCOUNT",
            vec![Value::Text(set_expression_or_handle.to_string())],
        );
        Value::Error(formula_engine::ErrorKind::NA)
    }

    fn cube_kpi_member(
        &self,
        connection: &str,
        kpi_name: &str,
        kpi_property: &str,
        caption: Option<&str>,
    ) -> Value {
        let mut args = vec![
            Value::Text(connection.to_string()),
            Value::Text(kpi_name.to_string()),
            Value::Text(kpi_property.to_string()),
        ];
        if let Some(caption) = caption {
            args.push(Value::Text(caption.to_string()));
        }
        self.record_call("CUBEKPIMEMBER", args);
        Value::Error(formula_engine::ErrorKind::NA)
    }
}

fn table_fixture_multi_col() -> Table {
    Table {
        id: 1,
        name: "Table1".into(),
        display_name: "Table1".into(),
        range: Range::from_a1("A1:D4").unwrap(),
        header_row_count: 1,
        totals_row_count: 0,
        columns: vec![
            TableColumn {
                id: 1,
                name: "Col1".into(),
                formula: None,
                totals_formula: None,
            },
            TableColumn {
                id: 2,
                name: "Col2".into(),
                formula: None,
                totals_formula: None,
            },
            TableColumn {
                id: 3,
                name: "Col3".into(),
                formula: None,
                totals_formula: None,
            },
            TableColumn {
                id: 4,
                name: "Col4".into(),
                formula: None,
                totals_formula: None,
            },
        ],
        style: None,
        auto_filter: None,
        relationship_id: None,
        part_path: None,
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
fn debug_trace_supports_getting_data_error_literal() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=#GETTING_DATA")
        .unwrap();
    engine.recalculate();

    let computed = engine.get_cell_value("Sheet1", "A1");
    assert_eq!(
        computed,
        Value::Error(formula_engine::ErrorKind::GettingData)
    );

    let dbg = engine.debug_evaluate("Sheet1", "A1").unwrap();
    assert_eq!(dbg.value, computed);
}

#[test]
fn debug_trace_supports_rtd_calls() {
    let provider = Arc::new(TestExternalDataProvider::default());

    let mut engine = Engine::new();
    engine.set_external_data_provider(Some(provider.clone()));
    engine
        .set_cell_formula("Sheet1", "A1", "=RTD(\"my.prog\",\"\",42)")
        .unwrap();
    engine.recalculate();

    let computed = engine.get_cell_value("Sheet1", "A1");
    assert_eq!(computed, Value::Number(123.0));

    let dbg = engine.debug_evaluate("Sheet1", "A1").unwrap();
    assert_eq!(dbg.value, computed);
    assert_eq!(dbg.trace.children.len(), 3);

    assert!(
        provider.calls().iter().any(|(name, _args)| name == "RTD"),
        "expected provider to record an RTD call"
    );
}

#[test]
fn debug_trace_supports_cube_calls() {
    let provider = Arc::new(TestExternalDataProvider::default());

    let mut engine = Engine::new();
    engine.set_external_data_provider(Some(provider.clone()));
    engine
        .set_cell_formula("Sheet1", "A1", "=CUBEVALUE(\"conn\",\"[Measures].[X]\")")
        .unwrap();
    engine.recalculate();

    let computed = engine.get_cell_value("Sheet1", "A1");
    assert_eq!(computed, Value::Number(456.0));

    let dbg = engine.debug_evaluate("Sheet1", "A1").unwrap();
    assert_eq!(dbg.value, computed);
    assert_eq!(dbg.trace.children.len(), 2);

    assert!(
        provider
            .calls()
            .iter()
            .any(|(name, _args)| name == "CUBEVALUE"),
        "expected provider to record a CUBEVALUE call"
    );
}

#[test]
fn debug_trace_respects_value_locale_for_cube_numeric_args() {
    let provider = Arc::new(TestExternalDataProvider::default());

    let mut engine = Engine::new();
    assert!(engine.set_value_locale_id("de-DE"));
    engine.set_external_data_provider(Some(provider.clone()));

    // Rank argument is text that depends on value-locale numeric parsing ("," decimal separator
    // for de-DE). The provider returns the parsed rank so we can assert debug tracing uses the
    // same coercion semantics as normal evaluation.
    engine
        .set_cell_formula(
            "Sheet1",
            "A1",
            "=CUBERANKEDMEMBER(\"conn\",\"set\",\"1,5\")",
        )
        .unwrap();
    engine.recalculate();

    let computed = engine.get_cell_value("Sheet1", "A1");
    assert_eq!(computed, Value::Number(1.0));

    let dbg = engine.debug_evaluate("Sheet1", "A1").unwrap();
    assert_eq!(dbg.value, computed);
    assert_eq!(dbg.trace.children.len(), 3);

    assert!(
        provider
            .calls()
            .iter()
            .any(|(name, args)| name == "CUBERANKEDMEMBER"
                && args.get(2) == Some(&Value::Number(1.0))),
        "expected provider to record a CUBERANKEDMEMBER call with rank=1"
    );
}

#[test]
fn debug_trace_respects_value_locale_for_concat_number_formatting() {
    let mut engine = Engine::new();
    assert!(engine.set_value_locale_id("de-DE"));

    // Formula is stored canonically with '.', but de-DE value locale should render ',' when
    // coercing numbers to text via concatenation.
    engine
        .set_cell_formula("Sheet1", "A1", "=1.5&\"\"")
        .unwrap();
    engine.recalculate();

    let computed = engine.get_cell_value("Sheet1", "A1");
    assert_eq!(computed, Value::Text("1,5".to_string()));

    let dbg = engine.debug_evaluate("Sheet1", "A1").unwrap();
    assert_eq!(dbg.value, computed);
}

#[test]
fn debug_trace_respects_value_locale_for_numeric_coercion_in_operators() {
    let mut engine = Engine::new();
    assert!(engine.set_value_locale_id("de-DE"));

    engine
        .set_cell_formula("Sheet1", "A1", "=\"1,5\"+0")
        .unwrap();
    engine.recalculate();

    let computed = engine.get_cell_value("Sheet1", "A1");
    assert_eq!(computed, Value::Number(1.5));

    let dbg = engine.debug_evaluate("Sheet1", "A1").unwrap();
    assert_eq!(dbg.value, computed);
}

#[test]
fn debug_trace_respects_date_system_for_text_date_coercion() {
    let mut engine = Engine::new();
    engine.set_date_system(ExcelDateSystem::Excel1904);

    engine
        .set_cell_formula("Sheet1", "A1", "=\"1/1/1904\"+0")
        .unwrap();
    engine.recalculate();

    let computed = engine.get_cell_value("Sheet1", "A1");
    assert_eq!(computed, Value::Number(0.0));

    let dbg = engine.debug_evaluate("Sheet1", "A1").unwrap();
    assert_eq!(dbg.value, computed);
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
fn debug_trace_supports_row_and_column_ranges_with_sheet_dimensions() {
    let mut engine = Engine::new();
    engine.set_sheet_dimensions("Sheet1", 3, 5).unwrap(); // 3 rows, 5 cols (A..E)
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "B1", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "C1", 3.0).unwrap();
    engine.set_cell_value("Sheet1", "D1", 4.0).unwrap();
    engine.set_cell_value("Sheet1", "E1", 5.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 10.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", 100.0).unwrap();

    engine
        .set_cell_formula("Sheet1", "B2", "=SUM(A:A)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B3", "=SUM(1:1)")
        .unwrap();
    engine.recalculate();

    let dbg_col = engine.debug_evaluate("Sheet1", "B2").unwrap();
    assert_eq!(dbg_col.value, Value::Number(111.0));
    assert_eq!(slice(&dbg_col.formula, dbg_col.trace.span), "SUM(A:A)");
    assert!(matches!(dbg_col.trace.kind, TraceKind::FunctionCall { .. }));
    let range_node = &dbg_col.trace.children[0];
    assert_eq!(slice(&dbg_col.formula, range_node.span), "A:A");
    assert_eq!(
        range_node.reference,
        Some(TraceRef::Range {
            sheet: formula_engine::functions::SheetId::Local(0),
            start: CellAddr { row: 0, col: 0 },
            end: CellAddr { row: 2, col: 0 },
        })
    );

    let dbg_row = engine.debug_evaluate("Sheet1", "B3").unwrap();
    assert_eq!(dbg_row.value, Value::Number(15.0));
    assert_eq!(slice(&dbg_row.formula, dbg_row.trace.span), "SUM(1:1)");
    assert!(matches!(dbg_row.trace.kind, TraceKind::FunctionCall { .. }));
    let range_node = &dbg_row.trace.children[0];
    assert_eq!(slice(&dbg_row.formula, range_node.span), "1:1");
    assert_eq!(
        range_node.reference,
        Some(TraceRef::Range {
            sheet: formula_engine::functions::SheetId::Local(0),
            start: CellAddr { row: 0, col: 0 },
            end: CellAddr { row: 0, col: 4 },
        })
    );
}

#[test]
fn debug_trace_does_not_materialize_huge_references() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=A1:XFD1048576")
        .unwrap();

    let dbg = engine.debug_evaluate("Sheet1", "A1").unwrap();
    assert_eq!(dbg.value, Value::Error(formula_engine::ErrorKind::Spill));
    assert!(matches!(dbg.trace.kind, TraceKind::RangeRef));
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
fn debug_trace_supports_array_arithmetic_broadcasting() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "={1;2}+{10,20,30}")
        .unwrap();

    let dbg = engine.debug_evaluate("Sheet1", "A1").unwrap();
    assert_eq!(slice(&dbg.formula, dbg.trace.span), "{1;2}+{10,20,30}");
    assert_eq!(
        dbg.trace.kind,
        TraceKind::Binary {
            op: formula_engine::eval::BinaryOp::Add
        }
    );

    let Value::Array(arr) = dbg.value else {
        panic!(
            "expected Value::Array from debug evaluation, got {:?}",
            dbg.value
        );
    };
    assert_eq!(arr.rows, 2);
    assert_eq!(arr.cols, 3);
    assert_eq!(arr.get(0, 0), Some(&Value::Number(11.0)));
    assert_eq!(arr.get(0, 1), Some(&Value::Number(21.0)));
    assert_eq!(arr.get(0, 2), Some(&Value::Number(31.0)));
    assert_eq!(arr.get(1, 0), Some(&Value::Number(12.0)));
    assert_eq!(arr.get(1, 1), Some(&Value::Number(22.0)));
    assert_eq!(arr.get(1, 2), Some(&Value::Number(32.0)));
}

#[test]
fn debug_trace_supports_spill_range_arithmetic_broadcasting() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=SEQUENCE(3,1)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "C1", "=SEQUENCE(1,4)")
        .unwrap();
    engine.set_cell_formula("Sheet1", "A5", "=A1#+C1#").unwrap();
    engine.recalculate();

    let dbg = engine.debug_evaluate("Sheet1", "A5").unwrap();
    assert_eq!(slice(&dbg.formula, dbg.trace.span), "A1#+C1#");
    assert_eq!(
        dbg.trace.kind,
        TraceKind::Binary {
            op: formula_engine::eval::BinaryOp::Add
        }
    );

    assert_eq!(dbg.trace.children.len(), 2);
    assert_eq!(dbg.trace.children[0].kind, TraceKind::SpillRange);
    assert_eq!(slice(&dbg.formula, dbg.trace.children[0].span), "A1#");
    assert_eq!(
        dbg.trace.children[0].reference,
        Some(TraceRef::Range {
            sheet: formula_engine::functions::SheetId::Local(0),
            start: CellAddr { row: 0, col: 0 },
            end: CellAddr { row: 2, col: 0 },
        })
    );

    assert_eq!(dbg.trace.children[1].kind, TraceKind::SpillRange);
    assert_eq!(slice(&dbg.formula, dbg.trace.children[1].span), "C1#");
    assert_eq!(
        dbg.trace.children[1].reference,
        Some(TraceRef::Range {
            sheet: formula_engine::functions::SheetId::Local(0),
            start: CellAddr { row: 0, col: 2 },
            end: CellAddr { row: 0, col: 5 },
        })
    );

    let Value::Array(arr) = dbg.value else {
        panic!(
            "expected Value::Array from debug evaluation, got {:?}",
            dbg.value
        );
    };
    assert_eq!(arr.rows, 3);
    assert_eq!(arr.cols, 4);
    assert_eq!(arr.get(0, 0), Some(&Value::Number(2.0)));
    assert_eq!(arr.get(0, 3), Some(&Value::Number(5.0)));
    assert_eq!(arr.get(2, 0), Some(&Value::Number(4.0)));
    assert_eq!(arr.get(2, 3), Some(&Value::Number(7.0)));
}

#[test]
fn debug_trace_supports_concat_operator_and_precedence() {
    let mut engine = Engine::new();
    engine.set_cell_formula("Sheet1", "A1", "=1+2&3").unwrap();
    engine.recalculate();

    let dbg = engine.debug_evaluate("Sheet1", "A1").unwrap();
    assert_eq!(dbg.value, Value::Text("33".to_string()));
    assert_eq!(slice(&dbg.formula, dbg.trace.span), "1+2&3");
    assert_eq!(
        dbg.trace.kind,
        TraceKind::Binary {
            op: formula_engine::eval::BinaryOp::Concat
        }
    );

    // Left side should be `1+2` (add has higher precedence than `&`).
    assert_eq!(slice(&dbg.formula, dbg.trace.children[0].span), "1+2");
    assert_eq!(
        dbg.trace.children[0].kind,
        TraceKind::Binary {
            op: formula_engine::eval::BinaryOp::Add
        }
    );
    assert_eq!(dbg.trace.children[0].value, Value::Number(3.0));

    // Right side is `3`.
    assert_eq!(slice(&dbg.formula, dbg.trace.children[1].span), "3");
    assert_eq!(dbg.trace.children[1].value, Value::Number(3.0));
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
fn debug_trace_supports_double_quoted_3d_sheet_range_refs() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet 1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet 2", "A1", 2.0).unwrap();
    engine.set_cell_value("Sheet 3", "A1", 3.0).unwrap();
    engine
        .set_cell_formula("Summary", "A1", "=SUM('Sheet 1':'Sheet 3'!A1)")
        .unwrap();
    engine.recalculate();

    let dbg = engine.debug_evaluate("Summary", "A1").unwrap();
    assert_eq!(dbg.value, Value::Number(6.0));
    assert_eq!(
        slice(&dbg.formula, dbg.trace.span),
        "SUM('Sheet 1':'Sheet 3'!A1)"
    );

    let arg = &dbg.trace.children[0];
    assert_eq!(slice(&dbg.formula, arg.span), "'Sheet 1':'Sheet 3'!A1");
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
    provider.set("[Book.xlsx]Sheet1", CellAddr { row: 0, col: 0 }, 41.0);

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
    provider.set("[Work Book-1.xlsx]Sheet1", CellAddr { row: 0, col: 0 }, 9.0);

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
fn debug_trace_supports_external_refs_with_quoted_sheet_name_after_workbook_prefix() {
    let provider = Arc::new(TestExternalProvider::default());
    provider.set("[Book.xlsx]My Sheet", CellAddr { row: 0, col: 0 }, 5.0);

    let mut engine = Engine::new();
    engine.set_external_value_provider(Some(provider));
    engine
        .set_cell_formula("Sheet1", "A1", "=[Book.xlsx]'My Sheet'!A1")
        .unwrap();
    engine.recalculate();

    let computed = engine.get_cell_value("Sheet1", "A1");
    assert_eq!(computed, Value::Number(5.0));

    let dbg = engine.debug_evaluate("Sheet1", "A1").unwrap();
    assert_eq!(dbg.value, computed);
    assert_eq!(
        slice(&dbg.formula, dbg.trace.span),
        "[Book.xlsx]'My Sheet'!A1"
    );
    assert_eq!(
        dbg.trace.reference,
        Some(TraceRef::Cell {
            sheet: formula_engine::functions::SheetId::External("[Book.xlsx]My Sheet".to_string()),
            addr: CellAddr { row: 0, col: 0 }
        })
    );
}

#[test]
fn debug_trace_supports_path_qualified_external_workbook_cell_refs() {
    let provider = Arc::new(TestExternalProvider::default());
    provider.set(
        r"[C:\path\Book.xlsx]Sheet1",
        CellAddr { row: 0, col: 0 },
        11.0,
    );

    let mut engine = Engine::new();
    engine.set_external_value_provider(Some(provider));
    engine
        .set_cell_formula("Sheet1", "A1", r#"='C:\path\[Book.xlsx]Sheet1'!A1"#)
        .unwrap();
    engine.recalculate();

    let computed = engine.get_cell_value("Sheet1", "A1");
    assert_eq!(computed, Value::Number(11.0));

    let dbg = engine.debug_evaluate("Sheet1", "A1").unwrap();
    assert_eq!(dbg.value, computed);
    assert_eq!(
        slice(&dbg.formula, dbg.trace.span),
        r#"'C:\path\[Book.xlsx]Sheet1'!A1"#
    );
    assert!(matches!(dbg.trace.kind, TraceKind::CellRef));
    assert_eq!(
        dbg.trace.reference,
        Some(TraceRef::Cell {
            sheet: formula_engine::functions::SheetId::External(
                r"[C:\path\Book.xlsx]Sheet1".to_string()
            ),
            addr: CellAddr { row: 0, col: 0 }
        })
    );
}

#[test]
fn debug_trace_supports_path_qualified_external_workbook_range_refs() {
    let provider = Arc::new(TestExternalProvider::default());
    provider.set(
        r"[C:\path\Book.xlsx]Sheet1",
        CellAddr { row: 0, col: 0 },
        1.0,
    );
    provider.set(
        r"[C:\path\Book.xlsx]Sheet1",
        CellAddr { row: 1, col: 0 },
        2.0,
    );

    let mut engine = Engine::new();
    engine.set_external_value_provider(Some(provider));
    engine
        .set_cell_formula("Sheet1", "A1", r#"='C:\path\[Book.xlsx]Sheet1'!A1:A2"#)
        .unwrap();

    let dbg = engine.debug_evaluate("Sheet1", "A1").unwrap();
    assert_eq!(
        slice(&dbg.formula, dbg.trace.span),
        r#"'C:\path\[Book.xlsx]Sheet1'!A1:A2"#
    );
    assert!(matches!(dbg.trace.kind, TraceKind::RangeRef));
    assert_eq!(
        dbg.trace.reference,
        Some(TraceRef::Range {
            sheet: formula_engine::functions::SheetId::External(
                r"[C:\path\Book.xlsx]Sheet1".into()
            ),
            start: CellAddr { row: 0, col: 0 },
            end: CellAddr { row: 1, col: 0 },
        })
    );

    let Value::Array(arr) = dbg.value else {
        panic!(
            "expected Value::Array from debug evaluation, got {:?}",
            dbg.value
        );
    };
    assert_eq!(arr.rows, 2);
    assert_eq!(arr.cols, 1);
    assert_eq!(arr.get(0, 0), Some(&Value::Number(1.0)));
    assert_eq!(arr.get(1, 0), Some(&Value::Number(2.0)));
}

#[test]
fn debug_trace_supports_structured_references() {
    let mut engine = Engine::new();
    engine.set_sheet_tables("Sheet1", vec![table_fixture_multi_col()]);
    engine.set_cell_value("Sheet1", "A1", "Col1").unwrap();
    engine.set_cell_value("Sheet1", "B1", "Col2").unwrap();
    engine.set_cell_value("Sheet1", "C1", "Col3").unwrap();
    engine.set_cell_value("Sheet1", "D1", "Col4").unwrap();
    engine.set_cell_value("Sheet1", "B2", 10.0).unwrap();
    engine.set_cell_formula("Sheet1", "D2", "=[@Col2]").unwrap();
    engine.recalculate_single_threaded();

    let computed = engine.get_cell_value("Sheet1", "D2");
    assert_eq!(computed, Value::Number(10.0));

    let dbg = engine.debug_evaluate("Sheet1", "D2").unwrap();
    assert_eq!(dbg.value, computed);
    assert_eq!(slice(&dbg.formula, dbg.trace.span), "[@Col2]");
    assert!(matches!(dbg.trace.kind, TraceKind::StructuredRef));
    assert_eq!(
        dbg.trace.reference,
        Some(TraceRef::Cell {
            sheet: formula_engine::functions::SheetId::Local(0),
            addr: CellAddr { row: 1, col: 1 }
        })
    );
}

#[test]
fn debug_trace_supports_sheet_prefixed_structured_references() {
    let mut engine = Engine::new();
    engine.set_sheet_tables("Sheet1", vec![table_fixture_multi_col()]);
    engine.set_cell_value("Sheet1", "B2", 10.0).unwrap();
    engine.set_cell_value("Sheet1", "B3", 20.0).unwrap();
    engine.set_cell_value("Sheet1", "B4", 30.0).unwrap();

    engine
        .set_cell_formula("Summary", "A1", "=SUM(Sheet1!Table1[Col2])")
        .unwrap();
    engine.recalculate_single_threaded();

    let computed = engine.get_cell_value("Summary", "A1");
    assert_eq!(computed, Value::Number(60.0));

    let dbg = engine.debug_evaluate("Summary", "A1").unwrap();
    assert_eq!(dbg.value, computed);
    assert_eq!(
        slice(&dbg.formula, dbg.trace.span),
        "SUM(Sheet1!Table1[Col2])"
    );

    let arg = &dbg.trace.children[0];
    assert_eq!(slice(&dbg.formula, arg.span), "Sheet1!Table1[Col2]");
    assert!(matches!(arg.kind, TraceKind::StructuredRef));
    assert_eq!(
        arg.reference,
        Some(TraceRef::Range {
            sheet: formula_engine::functions::SheetId::Local(0),
            start: CellAddr { row: 1, col: 1 },
            end: CellAddr { row: 3, col: 1 }
        })
    );
}

#[test]
fn debug_trace_supports_external_workbook_structured_references() {
    let provider = Arc::new(TestExternalProvider::default());
    provider.set_table("Book.xlsx", "Sheet1", table_fixture_multi_col());
    provider.set("[Book.xlsx]Sheet1", CellAddr { row: 1, col: 1 }, 10.0);
    provider.set("[Book.xlsx]Sheet1", CellAddr { row: 2, col: 1 }, 20.0);
    provider.set("[Book.xlsx]Sheet1", CellAddr { row: 3, col: 1 }, 30.0);

    let mut engine = Engine::new();
    engine.set_external_value_provider(Some(provider));
    engine
        .set_cell_formula("Summary", "A1", "=SUM([Book.xlsx]Sheet1!Table1[Col2])")
        .unwrap();
    engine.recalculate_single_threaded();

    let computed = engine.get_cell_value("Summary", "A1");
    assert_eq!(computed, Value::Number(60.0));

    let dbg = engine.debug_evaluate("Summary", "A1").unwrap();
    assert_eq!(dbg.value, computed);
    assert_eq!(
        slice(&dbg.formula, dbg.trace.span),
        "SUM([Book.xlsx]Sheet1!Table1[Col2])"
    );

    let arg = &dbg.trace.children[0];
    assert_eq!(
        slice(&dbg.formula, arg.span),
        "[Book.xlsx]Sheet1!Table1[Col2]"
    );
    assert!(matches!(arg.kind, TraceKind::StructuredRef));
    assert_eq!(
        arg.reference,
        Some(TraceRef::Range {
            sheet: formula_engine::functions::SheetId::External("[Book.xlsx]Sheet1".to_string()),
            start: CellAddr { row: 1, col: 1 },
            end: CellAddr { row: 3, col: 1 }
        })
    );
}

#[test]
fn trace_preserves_reference_context_for_sum_over_external_ranges() {
    let provider = Arc::new(TestExternalProvider::default());
    provider.set("[Book.xlsx]Sheet1", CellAddr { row: 0, col: 0 }, 1.0);
    provider.set("[Book.xlsx]Sheet1", CellAddr { row: 1, col: 0 }, 2.0);

    let mut engine = Engine::new();
    engine.set_external_value_provider(Some(provider));
    engine
        .set_cell_formula("Sheet1", "B1", "=SUM([Book.xlsx]Sheet1!A1:A2)")
        .unwrap();
    engine.recalculate();

    let dbg = engine.debug_evaluate("Sheet1", "B1").unwrap();
    assert_eq!(dbg.value, Value::Number(3.0));

    let range_node = &dbg.trace.children[0];
    assert_eq!(
        slice(&dbg.formula, range_node.span),
        "[Book.xlsx]Sheet1!A1:A2"
    );
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
    provider.set("[Book.xlsx]Sheet1", CellAddr { row: 0, col: 0 }, 7.0);

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

#[test]
fn debug_trace_rejects_external_3d_sheet_spans_without_sheet_order() {
    let provider = Arc::new(TestExternalProvider::default());
    provider.set("[Book.xlsx]Sheet1", CellAddr { row: 0, col: 0 }, 1.0);
    provider.set("[Book.xlsx]Sheet3", CellAddr { row: 0, col: 0 }, 3.0);

    let mut engine = Engine::new();
    engine.set_external_value_provider(Some(provider));
    engine
        .set_cell_formula("Sheet1", "A1", "=[Book.xlsx]Sheet1:Sheet3!A1")
        .unwrap();
    engine.recalculate();

    let computed = engine.get_cell_value("Sheet1", "A1");
    assert_eq!(computed, Value::Error(formula_engine::ErrorKind::Ref));

    let dbg = engine.debug_evaluate("Sheet1", "A1").unwrap();
    assert_eq!(dbg.value, computed);
    assert_eq!(
        slice(&dbg.formula, dbg.trace.span),
        "[Book.xlsx]Sheet1:Sheet3!A1"
    );
    assert!(matches!(dbg.trace.kind, TraceKind::CellRef));
    assert!(dbg.trace.reference.is_none());
}

#[test]
fn debug_trace_supports_external_3d_sheet_spans_with_sheet_order() {
    let provider = Arc::new(TestExternalProvider::default());
    provider.set_sheet_order("Book.xlsx", &["Sheet1", "Sheet2", "Sheet3"]);
    provider.set("[Book.xlsx]Sheet1", CellAddr { row: 0, col: 0 }, 1.0);
    provider.set("[Book.xlsx]Sheet2", CellAddr { row: 0, col: 0 }, 2.0);
    provider.set("[Book.xlsx]Sheet3", CellAddr { row: 0, col: 0 }, 3.0);

    let mut engine = Engine::new();
    engine.set_external_value_provider(Some(provider));
    engine
        .set_cell_formula("Sheet1", "A1", "=SUM([Book.xlsx]Sheet1:Sheet3!A1)")
        .unwrap();
    engine.recalculate();

    let computed = engine.get_cell_value("Sheet1", "A1");
    assert_eq!(computed, Value::Number(6.0));

    let dbg = engine.debug_evaluate("Sheet1", "A1").unwrap();
    assert_eq!(dbg.value, computed);
    assert_eq!(
        slice(&dbg.formula, dbg.trace.span),
        "SUM([Book.xlsx]Sheet1:Sheet3!A1)"
    );

    assert!(matches!(
        dbg.trace.kind,
        TraceKind::FunctionCall { ref name } if name == "SUM"
    ));
    assert_eq!(dbg.trace.children.len(), 1);

    let arg = &dbg.trace.children[0];
    assert_eq!(slice(&dbg.formula, arg.span), "[Book.xlsx]Sheet1:Sheet3!A1");
    assert!(matches!(arg.kind, TraceKind::CellRef));
    assert_eq!(arg.value, Value::Blank);
    assert!(
        arg.reference.is_none(),
        "3D sheet-range refs resolve to multiple sheets"
    );
}

#[test]
fn debug_trace_supports_field_access_on_record_values() {
    use formula_engine::value::Record;

    let mut engine = Engine::new();

    let mut record = Record::new("Widget");
    record
        .fields
        .insert("Price".to_string(), Value::Number(19.99));
    engine
        .set_cell_value("Sheet1", "A1", Value::Record(record))
        .unwrap();

    engine
        .set_cell_formula("Sheet1", "B1", "=A1.Price")
        .unwrap();
    engine.recalculate();

    let computed = engine.get_cell_value("Sheet1", "B1");
    assert_eq!(computed, Value::Number(19.99));

    let dbg = engine.debug_evaluate("Sheet1", "B1").unwrap();
    assert_eq!(dbg.value, computed);
    assert_eq!(slice(&dbg.formula, dbg.trace.span), "A1.Price");
    assert!(matches!(dbg.trace.kind, TraceKind::FieldAccess { .. }));
}

#[test]
fn debug_trace_propagates_field_error_for_missing_record_fields() {
    use formula_engine::value::Record;

    let mut engine = Engine::new();

    let mut record = Record::new("Widget");
    record
        .fields
        .insert("Other".to_string(), Value::Number(1.0));
    engine
        .set_cell_value("Sheet1", "A1", Value::Record(record))
        .unwrap();

    engine
        .set_cell_formula("Sheet1", "B1", "=A1.Price")
        .unwrap();
    engine.recalculate();

    let computed = engine.get_cell_value("Sheet1", "B1");
    assert_eq!(computed, Value::Error(formula_engine::ErrorKind::Field));

    let dbg = engine.debug_evaluate("Sheet1", "B1").unwrap();
    assert_eq!(dbg.value, computed);
    assert_eq!(slice(&dbg.formula, dbg.trace.span), "A1.Price");
}
