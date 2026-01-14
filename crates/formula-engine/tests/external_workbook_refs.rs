use formula_engine::calc_settings::{CalcSettings, CalculationMode};
use formula_engine::eval::CellAddr;
use formula_engine::{Engine, ExternalValueProvider, PrecedentNode, Value};
use formula_model::table::TableColumn;
use formula_model::{Range, Table};
use std::collections::HashMap;
use std::sync::{Arc, Mutex};

#[derive(Default)]
struct TestExternalProvider {
    values: Mutex<HashMap<(String, CellAddr), Value>>,
    sheet_order: Mutex<HashMap<String, Vec<String>>>,
    tables: Mutex<HashMap<(String, String), (String, Table)>>,
}

impl TestExternalProvider {
    fn set(&self, sheet: &str, addr: CellAddr, value: impl Into<Value>) {
        self.values
            .lock()
            .expect("lock poisoned")
            .insert((sheet.to_string(), addr), value.into());
    }

    fn set_sheet_order(&self, workbook: &str, order: impl Into<Vec<String>>) {
        self.sheet_order
            .lock()
            .expect("lock poisoned")
            .insert(workbook.to_string(), order.into());
    }

    fn set_table(&self, workbook: &str, sheet: &str, table: Table) {
        self.tables.lock().expect("lock poisoned").insert(
            (workbook.to_string(), table.name.clone()),
            (sheet.to_string(), table),
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

    fn sheet_order(&self, workbook: &str) -> Option<Vec<String>> {
        self.sheet_order
            .lock()
            .expect("lock poisoned")
            .get(workbook)
            .cloned()
    }

    fn workbook_table(&self, workbook: &str, table_name: &str) -> Option<(String, Table)> {
        self.tables
            .lock()
            .expect("lock poisoned")
            .get(&(workbook.to_string(), table_name.to_string()))
            .cloned()
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

#[derive(Default)]
struct ProviderWithoutSheetOrder;

impl ExternalValueProvider for ProviderWithoutSheetOrder {
    fn get(&self, _sheet: &str, _addr: CellAddr) -> Option<Value> {
        None
    }
}

#[derive(Default)]
struct ProviderWithWorkbookSheetNamesOnly {
    values: Mutex<HashMap<(String, CellAddr), Value>>,
    sheet_names: Mutex<HashMap<String, Arc<[String]>>>,
}

impl ProviderWithWorkbookSheetNamesOnly {
    fn set(&self, sheet: &str, addr: CellAddr, value: impl Into<Value>) {
        self.values
            .lock()
            .expect("lock poisoned")
            .insert((sheet.to_string(), addr), value.into());
    }

    fn set_sheet_names(&self, workbook: &str, names: impl Into<Vec<String>>) {
        self.sheet_names
            .lock()
            .expect("lock poisoned")
            .insert(workbook.to_string(), Arc::from(names.into()));
    }
}

impl ExternalValueProvider for ProviderWithWorkbookSheetNamesOnly {
    fn get(&self, sheet: &str, addr: CellAddr) -> Option<Value> {
        self.values
            .lock()
            .expect("lock poisoned")
            .get(&(sheet.to_string(), addr))
            .cloned()
    }

    fn workbook_sheet_names(&self, workbook: &str) -> Option<Arc<[String]>> {
        self.sheet_names
            .lock()
            .expect("lock poisoned")
            .get(workbook)
            .cloned()
    }
}

#[test]
fn external_cell_ref_resolves_via_provider() {
    let provider = Arc::new(TestExternalProvider::default());
    provider.set("[Book.xlsx]Sheet1", CellAddr { row: 0, col: 0 }, 41.0);

    let mut engine = Engine::new();
    engine.set_external_value_provider(Some(provider));
    engine
        .set_cell_formula("Sheet1", "A1", "=[Book.xlsx]Sheet1!A1")
        .unwrap();
    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(41.0));
}

#[test]
fn external_cell_ref_with_workbook_name_containing_lbracket_resolves_via_provider() {
    let provider = Arc::new(TestExternalProvider::default());
    provider.set("[A1[Name.xlsx]Sheet1", CellAddr { row: 0, col: 0 }, 41.0);

    let mut engine = Engine::new();
    engine.set_external_value_provider(Some(provider));
    engine
        .set_cell_formula("Sheet1", "A1", "=[A1[Name.xlsx]Sheet1!A1")
        .unwrap();
    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(41.0));
}

#[test]
fn external_cell_ref_with_workbook_name_containing_lbracket_and_escaped_rbracket_resolves_via_provider(
) {
    let provider = Arc::new(TestExternalProvider::default());
    provider.set(
        "[Book[Name]].xlsx]Sheet1",
        CellAddr { row: 0, col: 0 },
        41.0,
    );

    let mut engine = Engine::new();
    engine.set_external_value_provider(Some(provider));
    engine
        .set_cell_formula("Sheet1", "A1", "=[Book[Name]].xlsx]Sheet1!A1")
        .unwrap();
    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(41.0));
}

#[test]
fn quoted_external_cell_ref_with_workbook_name_containing_lbracket_resolves_via_provider() {
    let provider = Arc::new(TestExternalProvider::default());
    provider.set("[A1[Name.xlsx]Sheet1", CellAddr { row: 0, col: 0 }, 41.0);

    let mut engine = Engine::new();
    engine.set_external_value_provider(Some(provider));
    engine
        .set_cell_formula("Sheet1", "A1", "='[A1[Name.xlsx]Sheet1'!A1")
        .unwrap();
    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(41.0));
}

#[test]
fn quoted_external_cell_ref_with_workbook_name_containing_lbracket_and_escaped_rbracket_resolves_via_provider(
) {
    let provider = Arc::new(TestExternalProvider::default());
    provider.set(
        "[Book[Name]].xlsx]Sheet1",
        CellAddr { row: 0, col: 0 },
        41.0,
    );

    let mut engine = Engine::new();
    engine.set_external_value_provider(Some(provider));
    engine
        .set_cell_formula("Sheet1", "A1", "='[Book[Name]].xlsx]Sheet1'!A1")
        .unwrap();
    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(41.0));
}

#[test]
fn indirect_external_cell_ref_resolves_via_provider() {
    let provider = Arc::new(TestExternalProvider::default());
    provider.set("[Book.xlsx]Sheet1", CellAddr { row: 0, col: 0 }, 41.0);

    let mut engine = Engine::new();
    engine.set_external_value_provider(Some(provider));
    engine
        .set_cell_formula("Sheet1", "A1", r#"=INDIRECT("[Book.xlsx]Sheet1!A1")"#)
        .unwrap();
    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(41.0));
}

#[test]
fn external_cell_ref_participates_in_arithmetic() {
    let provider = Arc::new(TestExternalProvider::default());
    provider.set("[Book.xlsx]Sheet1", CellAddr { row: 0, col: 0 }, 41.0);

    let mut engine = Engine::new();
    engine.set_external_value_provider(Some(provider));
    engine
        .set_cell_formula("Sheet1", "A1", "=[Book.xlsx]Sheet1!A1+1")
        .unwrap();
    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(42.0));
}

#[test]
fn external_cell_ref_with_path_qualified_workbook_resolves_via_provider() {
    let provider = Arc::new(TestExternalProvider::default());
    provider.set(
        "[C:\\path\\Book.xlsx]Sheet1",
        CellAddr { row: 0, col: 0 },
        41.0,
    );

    let mut engine = Engine::new();
    engine.set_external_value_provider(Some(provider));
    engine
        .set_cell_formula("Sheet1", "A1", "='C:\\path\\[Book.xlsx]Sheet1'!A1")
        .unwrap();
    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(41.0));
}

#[test]
fn sum_over_external_range_uses_reference_semantics() {
    // Excel quirk: SUM over references ignores logicals/text stored in cells.
    let provider = Arc::new(TestExternalProvider::default());
    provider.set("[Book.xlsx]Sheet1", CellAddr { row: 0, col: 0 }, 1.0);
    provider.set(
        "[Book.xlsx]Sheet1",
        CellAddr { row: 1, col: 0 },
        Value::Text("2".to_string()),
    );
    provider.set(
        "[Book.xlsx]Sheet1",
        CellAddr { row: 2, col: 0 },
        Value::Bool(true),
    );

    let mut engine = Engine::new();
    engine.set_external_value_provider(Some(provider));
    engine
        .set_cell_formula("Sheet1", "A1", "=SUM([Book.xlsx]Sheet1!A1:A3)")
        .unwrap();
    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(1.0));
}

#[test]
fn indirect_external_range_ref_resolves_via_provider() {
    let provider = Arc::new(TestExternalProvider::default());
    provider.set("[Book.xlsx]Sheet1", CellAddr { row: 0, col: 0 }, 1.0);
    provider.set("[Book.xlsx]Sheet1", CellAddr { row: 1, col: 0 }, 2.0);

    let mut engine = Engine::new();
    engine.set_external_value_provider(Some(provider));
    engine
        .set_cell_formula(
            "Sheet1",
            "A1",
            r#"=SUM(INDIRECT("[Book.xlsx]Sheet1!A1:A2"))"#,
        )
        .unwrap();
    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(3.0));
}

#[test]
fn external_range_spills_to_grid() {
    let provider = Arc::new(TestExternalProvider::default());
    provider.set("[Book.xlsx]Sheet1", CellAddr { row: 0, col: 0 }, 1.0);
    provider.set("[Book.xlsx]Sheet1", CellAddr { row: 0, col: 1 }, 2.0);
    provider.set("[Book.xlsx]Sheet1", CellAddr { row: 1, col: 0 }, 3.0);
    provider.set("[Book.xlsx]Sheet1", CellAddr { row: 1, col: 1 }, 4.0);

    let mut engine = Engine::new();
    engine.set_external_value_provider(Some(provider));
    engine
        .set_cell_formula("Sheet1", "C1", "=[Book.xlsx]Sheet1!A1:B2")
        .unwrap();
    engine.recalculate();

    assert_eq!(
        engine.spill_range("Sheet1", "C1"),
        Some((CellAddr { row: 0, col: 2 }, CellAddr { row: 1, col: 3 }))
    );
    assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "D1"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "C2"), Value::Number(3.0));
    assert_eq!(engine.get_cell_value("Sheet1", "D2"), Value::Number(4.0));
}

#[test]
fn missing_external_value_is_ref_error() {
    let provider = Arc::new(TestExternalProvider::default());

    let mut engine = Engine::new();
    engine.set_external_value_provider(Some(provider));
    engine
        .set_cell_formula("Sheet1", "A1", "=[Book.xlsx]Sheet1!A1")
        .unwrap();
    engine.recalculate();

    assert_eq!(
        engine.get_cell_value("Sheet1", "A1"),
        Value::Error(formula_engine::ErrorKind::Ref)
    );
}

#[test]
fn precedents_include_external_refs() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "B1", "=[Book.xlsx]Sheet1!A1")
        .unwrap();

    assert_eq!(
        engine.precedents("Sheet1", "B1").unwrap(),
        vec![PrecedentNode::ExternalCell {
            sheet: "[Book.xlsx]Sheet1".to_string(),
            addr: CellAddr { row: 0, col: 0 },
        }]
    );
}

#[test]
fn precedents_include_external_table_structured_refs_when_metadata_available() {
    let provider = Arc::new(TestExternalProvider::default());
    provider.set_table("Book.xlsx", "Sheet1", table_fixture_multi_col());

    let mut engine = Engine::new();
    engine.set_external_value_provider(Some(provider));
    engine
        .set_cell_formula("Sheet1", "A1", "=SUM([Book.xlsx]Sheet1!Table1[Col2])")
        .unwrap();

    assert_eq!(
        engine.precedents("Sheet1", "A1").unwrap(),
        vec![PrecedentNode::ExternalRange {
            sheet: "[Book.xlsx]Sheet1".to_string(),
            // Table1[Col2] selects the Col2 data column (B2:B4) in the fixture table.
            start: CellAddr { row: 1, col: 1 },
            end: CellAddr { row: 3, col: 1 },
        }]
    );
}

#[test]
fn precedents_include_dynamic_external_precedents_from_indirect() {
    let provider = Arc::new(TestExternalProvider::default());
    provider.set("[Book.xlsx]Sheet1", CellAddr { row: 0, col: 0 }, 41.0);

    let mut engine = Engine::new();
    engine.set_external_value_provider(Some(provider));
    engine
        .set_cell_formula("Sheet1", "A1", r#"=INDIRECT("[Book.xlsx]Sheet1!A1")"#)
        .unwrap();
    // Ensure we compile to bytecode (no AST fallback) so we cover dynamic external precedent
    // tracking in the bytecode backend as well.
    assert!(
        engine.bytecode_compile_report(10).is_empty(),
        "{:?}",
        engine.bytecode_compile_report(10)
    );
    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(41.0));
    assert_eq!(
        engine.precedents("Sheet1", "A1").unwrap(),
        vec![PrecedentNode::ExternalCell {
            sheet: "[Book.xlsx]Sheet1".to_string(),
            addr: CellAddr { row: 0, col: 0 },
        }]
    );
}

#[test]
fn precedents_include_dynamic_external_precedents_from_indirect_ref_text_cell() {
    let provider = Arc::new(TestExternalProvider::default());
    provider.set("[Book.xlsx]Sheet1", CellAddr { row: 0, col: 0 }, 41.0);

    let mut engine = Engine::new();
    engine.set_external_value_provider(Some(provider));
    engine
        .set_cell_value("Sheet1", "B1", "[Book.xlsx]Sheet1!A1")
        .unwrap();
    engine.set_cell_formula("Sheet1", "A1", "=INDIRECT(B1)").unwrap();
    assert!(
        engine.bytecode_compile_report(10).is_empty(),
        "{:?}",
        engine.bytecode_compile_report(10)
    );
    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(41.0));
    assert_eq!(
        engine.precedents("Sheet1", "A1").unwrap(),
        vec![
            PrecedentNode::Cell {
                sheet: 0,
                addr: CellAddr { row: 0, col: 1 }, // B1
            },
            PrecedentNode::ExternalCell {
                sheet: "[Book.xlsx]Sheet1".to_string(),
                addr: CellAddr { row: 0, col: 0 },
            },
        ]
    );
}

#[test]
fn precedents_include_dynamic_external_precedents_from_offset() {
    let provider = Arc::new(TestExternalProvider::default());
    provider.set("[Book.xlsx]Sheet1", CellAddr { row: 1, col: 0 }, 41.0);

    let mut engine = Engine::new();
    engine.set_external_value_provider(Some(provider));
    engine
        .set_cell_formula("Sheet1", "A1", "=OFFSET([Book.xlsx]Sheet1!A1,1,0)")
        .unwrap();
    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(41.0));
    assert_eq!(
        engine.precedents("Sheet1", "A1").unwrap(),
        vec![
            PrecedentNode::ExternalCell {
                sheet: "[Book.xlsx]Sheet1".to_string(),
                addr: CellAddr { row: 0, col: 0 },
            },
            PrecedentNode::ExternalCell {
                sheet: "[Book.xlsx]Sheet1".to_string(),
                addr: CellAddr { row: 1, col: 0 },
            },
        ]
    );
}

#[test]
fn precedents_include_dynamic_external_range_from_offset() {
    let provider = Arc::new(TestExternalProvider::default());
    provider.set("[Book.xlsx]Sheet1", CellAddr { row: 1, col: 0 }, 1.0);
    provider.set("[Book.xlsx]Sheet1", CellAddr { row: 2, col: 0 }, 2.0);
    provider.set("[Book.xlsx]Sheet1", CellAddr { row: 3, col: 0 }, 3.0);

    let mut engine = Engine::new();
    engine.set_external_value_provider(Some(provider));
    engine
        .set_cell_formula("Sheet1", "A1", "=SUM(OFFSET([Book.xlsx]Sheet1!A1,1,0,3,1))")
        .unwrap();
    assert!(
        engine.bytecode_compile_report(10).is_empty(),
        "{:?}",
        engine.bytecode_compile_report(10)
    );
    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(6.0));
    assert_eq!(
        engine.precedents("Sheet1", "A1").unwrap(),
        vec![
            PrecedentNode::ExternalCell {
                sheet: "[Book.xlsx]Sheet1".to_string(),
                addr: CellAddr { row: 0, col: 0 },
            },
            PrecedentNode::ExternalRange {
                sheet: "[Book.xlsx]Sheet1".to_string(),
                start: CellAddr { row: 1, col: 0 },
                end: CellAddr { row: 3, col: 0 },
            },
        ]
    );
}

#[test]
fn precedents_transitive_include_dynamic_external_precedents() {
    let provider = Arc::new(TestExternalProvider::default());
    provider.set("[Book.xlsx]Sheet1", CellAddr { row: 1, col: 0 }, 1.0);
    provider.set("[Book.xlsx]Sheet1", CellAddr { row: 2, col: 0 }, 2.0);
    provider.set("[Book.xlsx]Sheet1", CellAddr { row: 3, col: 0 }, 3.0);

    let mut engine = Engine::new();
    engine.set_external_value_provider(Some(provider));
    engine
        .set_cell_formula("Sheet1", "A1", "=SUM(OFFSET([Book.xlsx]Sheet1!A1,1,0,3,1))")
        .unwrap();
    engine.set_cell_formula("Sheet1", "B1", "=A1+1").unwrap();
    assert!(
        engine.bytecode_compile_report(10).is_empty(),
        "{:?}",
        engine.bytecode_compile_report(10)
    );
    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(7.0));
    assert_eq!(
        engine.precedents_transitive("Sheet1", "B1").unwrap(),
        vec![
            PrecedentNode::Cell {
                sheet: 0,
                addr: CellAddr { row: 0, col: 0 },
            },
            PrecedentNode::ExternalCell {
                sheet: "[Book.xlsx]Sheet1".to_string(),
                addr: CellAddr { row: 0, col: 0 },
            },
            PrecedentNode::ExternalRange {
                sheet: "[Book.xlsx]Sheet1".to_string(),
                start: CellAddr { row: 1, col: 0 },
                end: CellAddr { row: 3, col: 0 },
            },
        ]
    );
}

#[test]
fn precedents_expand_external_3d_sheet_spans_when_sheet_order_is_available() {
    let provider = Arc::new(TestExternalProvider::default());
    provider.set_sheet_order(
        "Book.xlsx",
        vec![
            "Sheet1".to_string(),
            "Sheet2".to_string(),
            "Sheet3".to_string(),
        ],
    );
    for (sheet, value) in [("Sheet1", 1.0), ("Sheet2", 2.0), ("Sheet3", 3.0)] {
        provider.set(
            &format!("[Book.xlsx]{sheet}"),
            CellAddr { row: 0, col: 0 },
            value,
        );
    }

    let mut engine = Engine::new();
    engine.set_external_value_provider(Some(provider));
    engine
        .set_cell_formula("Sheet1", "A1", "=SUM([Book.xlsx]sheet1:sheet3!A1)")
        .unwrap();

    assert_eq!(
        engine.precedents("Sheet1", "A1").unwrap(),
        vec![
            PrecedentNode::ExternalCell {
                sheet: "[Book.xlsx]Sheet1".to_string(),
                addr: CellAddr { row: 0, col: 0 },
            },
            PrecedentNode::ExternalCell {
                sheet: "[Book.xlsx]Sheet2".to_string(),
                addr: CellAddr { row: 0, col: 0 },
            },
            PrecedentNode::ExternalCell {
                sheet: "[Book.xlsx]Sheet3".to_string(),
                addr: CellAddr { row: 0, col: 0 },
            },
        ]
    );
}

#[test]
fn precedents_expand_external_3d_sheet_span_matches_endpoints_nfkc_case_insensitively() {
    let provider = Arc::new(TestExternalProvider::default());
    provider.set_sheet_order(
        "Book.xlsx",
        vec![
            // U+212A KELVIN SIGN (K) is NFKC-equivalent to ASCII 'K'.
            "Kelvin".to_string(),
            "Sheet2".to_string(),
            "Sheet3".to_string(),
        ],
    );

    let mut engine = Engine::new();
    engine.set_external_value_provider(Some(provider));
    engine
        .set_cell_formula("Sheet1", "A1", "=[Book.xlsx]Kelvin:Sheet3!A1")
        .unwrap();

    assert_eq!(
        engine.precedents("Sheet1", "A1").unwrap(),
        vec![
            PrecedentNode::ExternalCell {
                sheet: "[Book.xlsx]Sheet2".to_string(),
                addr: CellAddr { row: 0, col: 0 },
            },
            PrecedentNode::ExternalCell {
                sheet: "[Book.xlsx]Sheet3".to_string(),
                addr: CellAddr { row: 0, col: 0 },
            },
            PrecedentNode::ExternalCell {
                sheet: "[Book.xlsx]Kelvin".to_string(),
                addr: CellAddr { row: 0, col: 0 },
            },
        ]
    );
}

#[test]
fn precedents_omit_external_3d_sheet_spans_when_sheet_order_unavailable() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=[Book.xlsx]Sheet1:Sheet3!A1")
        .unwrap();

    assert_eq!(engine.precedents("Sheet1", "A1").unwrap(), Vec::new());
}

#[test]
fn external_refs_are_volatile() {
    let provider = Arc::new(TestExternalProvider::default());
    provider.set("[Book.xlsx]Sheet1", CellAddr { row: 0, col: 0 }, 1.0);

    let mut engine = Engine::new();
    engine.set_external_value_provider(Some(provider.clone()));
    engine
        .set_cell_formula("Sheet1", "A1", "=[Book.xlsx]Sheet1!A1")
        .unwrap();
    engine.recalculate();
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(1.0));

    // Mutate the provider without marking any cells dirty. Volatile external references should be
    // included in subsequent recalculation passes.
    provider.set("[Book.xlsx]Sheet1", CellAddr { row: 0, col: 0 }, 2.0);
    engine.recalculate();
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(2.0));
}

#[test]
fn external_refs_can_be_non_volatile_with_explicit_invalidation() {
    let provider = Arc::new(TestExternalProvider::default());
    provider.set("[Book.xlsx]Sheet1", CellAddr { row: 0, col: 0 }, 1.0);

    let mut engine = Engine::new();
    engine.set_external_value_provider(Some(provider.clone()));
    engine.set_external_refs_volatile(false);
    engine
        .set_cell_formula("Sheet1", "A1", "=[Book.xlsx]Sheet1!A1")
        .unwrap();
    engine.recalculate();
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(1.0));

    // Mutate the provider without invalidation: the cell should not change because external refs
    // are treated as non-volatile.
    provider.set("[Book.xlsx]Sheet1", CellAddr { row: 0, col: 0 }, 2.0);
    engine.recalculate();
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(1.0));

    engine.mark_external_sheet_dirty("[Book.xlsx]Sheet1");
    engine.recalculate();
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(2.0));
}

#[test]
fn external_structured_refs_can_be_non_volatile_with_explicit_invalidation() {
    let provider = Arc::new(TestExternalProvider::default());
    provider.set_table("Book.xlsx", "Sheet1", table_fixture_multi_col());

    provider.set("[Book.xlsx]Sheet1", CellAddr { row: 1, col: 1 }, 10.0);
    provider.set("[Book.xlsx]Sheet1", CellAddr { row: 2, col: 1 }, 20.0);
    provider.set("[Book.xlsx]Sheet1", CellAddr { row: 3, col: 1 }, 30.0);

    let mut engine = Engine::new();
    engine.set_external_value_provider(Some(provider.clone()));
    engine.set_external_refs_volatile(false);
    engine
        .set_cell_formula("Sheet1", "A1", "=SUM([Book.xlsx]Sheet1!Table1[Col2])")
        .unwrap();
    engine.recalculate();
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(60.0));

    // Mutate the provider without invalidation: the cell should not change because external refs
    // are treated as non-volatile.
    provider.set("[Book.xlsx]Sheet1", CellAddr { row: 1, col: 1 }, 100.0);
    engine.recalculate();
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(60.0));

    engine.mark_external_sheet_dirty("[Book.xlsx]Sheet1");
    engine.recalculate();
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(150.0));
}

#[test]
fn external_sheet_invalidation_dirties_dynamic_external_indirect_dependents() {
    let provider = Arc::new(TestExternalProvider::default());
    provider.set("[Book.xlsx]Sheet1", CellAddr { row: 0, col: 0 }, 1.0);

    let mut engine = Engine::new();
    engine.set_external_value_provider(Some(provider));
    engine.set_external_refs_volatile(false);
    engine
        .set_cell_formula("Sheet1", "A1", r#"=INDIRECT("[Book.xlsx]Sheet1!A1")"#)
        .unwrap();
    // Ensure we're exercising the bytecode backend + bytecode dependency tracing.
    assert!(
        engine.bytecode_compile_report(10).is_empty(),
        "{:?}",
        engine.bytecode_compile_report(10)
    );
    engine.recalculate();
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(1.0));
    assert!(!engine.is_dirty("Sheet1", "A1"));

    engine.mark_external_sheet_dirty("[Book.xlsx]Sheet1");
    assert!(engine.is_dirty("Sheet1", "A1"));
}

#[test]
fn external_sheet_invalidation_dirties_dynamic_external_indirect_dependents_from_ref_text_cell() {
    let provider = Arc::new(TestExternalProvider::default());
    provider.set("[Book.xlsx]Sheet1", CellAddr { row: 0, col: 0 }, 1.0);

    let mut engine = Engine::new();
    engine.set_external_value_provider(Some(provider));
    engine.set_external_refs_volatile(false);
    engine
        .set_cell_value("Sheet1", "B1", "[Book.xlsx]Sheet1!A1")
        .unwrap();
    engine.set_cell_formula("Sheet1", "A1", "=INDIRECT(B1)").unwrap();
    assert!(
        engine.bytecode_compile_report(10).is_empty(),
        "{:?}",
        engine.bytecode_compile_report(10)
    );
    engine.recalculate();
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(1.0));
    assert!(!engine.is_dirty("Sheet1", "A1"));

    engine.mark_external_sheet_dirty("[Book.xlsx]Sheet1");
    assert!(engine.is_dirty("Sheet1", "A1"));
}

#[test]
fn external_workbook_invalidation_handles_workbook_ids_with_literal_brackets() {
    // Workbook names can contain literal `[` characters. Literal `]` characters are escaped as
    // `]]` in workbook ids. For a workbook name like `[Book]`, the canonical workbook id is:
    //   `[Book]]` (leading `[` is literal, trailing `]` is escaped).
    //
    // This workbook id appears in canonical external sheet keys as:
    //   `[[Book]]]Sheet1` (outer `[...]` plus the workbook id, then the sheet name).
    let provider = Arc::new(TestExternalProvider::default());
    provider.set("[[Book]]]Sheet1", CellAddr { row: 0, col: 0 }, 1.0);

    let mut engine = Engine::new();
    engine.set_external_value_provider(Some(provider.clone()));
    engine.set_external_refs_volatile(false);
    engine
        .set_cell_formula("Sheet1", "A1", "=[[Book]]]Sheet1!A1")
        .unwrap();
    engine.recalculate();
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(1.0));

    provider.set("[[Book]]]Sheet1", CellAddr { row: 0, col: 0 }, 2.0);
    engine.recalculate();
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(1.0));

    engine.mark_external_workbook_dirty("[Book]]");
    engine.recalculate();
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(2.0));
}

#[test]
fn external_structured_refs_respect_non_volatile_external_invalidation() {
    let provider = Arc::new(TestExternalProvider::default());
    provider.set_table("Book.xlsx", "Sheet1", table_fixture_multi_col());
    provider.set("[Book.xlsx]Sheet1", CellAddr { row: 1, col: 1 }, 10.0);
    provider.set("[Book.xlsx]Sheet1", CellAddr { row: 2, col: 1 }, 20.0);
    provider.set("[Book.xlsx]Sheet1", CellAddr { row: 3, col: 1 }, 30.0);

    let mut engine = Engine::new();
    engine.set_external_value_provider(Some(provider.clone()));
    engine.set_external_refs_volatile(false);
    engine
        .set_cell_formula("Sheet1", "A1", "=SUM([Book.xlsx]Sheet1!Table1[Col2])")
        .unwrap();
    engine.recalculate();
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(60.0));

    // Mutate the provider without invalidation: the result should not change because external refs
    // are treated as non-volatile.
    provider.set("[Book.xlsx]Sheet1", CellAddr { row: 2, col: 1 }, 25.0);
    engine.recalculate();
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(60.0));

    engine.mark_external_sheet_dirty("[Book.xlsx]Sheet1");
    engine.recalculate();
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(65.0));
}

#[test]
fn precedents_include_external_table_structured_ref_ranges() {
    let provider = Arc::new(TestExternalProvider::default());
    provider.set_table("Book.xlsx", "Sheet1", table_fixture_multi_col());

    let mut engine = Engine::new();
    engine.set_external_value_provider(Some(provider));
    engine
        .set_cell_formula("Sheet1", "A1", "=SUM([Book.xlsx]Sheet1!Table1[Col2])")
        .unwrap();

    assert_eq!(
        engine.precedents("Sheet1", "A1").unwrap(),
        vec![PrecedentNode::ExternalRange {
            sheet: "[Book.xlsx]Sheet1".to_string(),
            start: CellAddr { row: 1, col: 1 },
            end: CellAddr { row: 3, col: 1 },
        }]
    );
}

#[test]
fn external_sheet_invalidation_only_dirties_dependents_of_that_sheet() {
    let provider = Arc::new(TestExternalProvider::default());
    provider.set("[Book.xlsx]Sheet1", CellAddr { row: 0, col: 0 }, 1.0);
    provider.set("[Book.xlsx]Sheet2", CellAddr { row: 0, col: 0 }, 10.0);

    let mut engine = Engine::new();
    engine.set_external_value_provider(Some(provider.clone()));
    engine.set_external_refs_volatile(false);
    engine
        .set_cell_formula("Sheet1", "A1", "=[Book.xlsx]Sheet1!A1")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A2", "=[Book.xlsx]Sheet2!A1")
        .unwrap();
    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "A2"), Value::Number(10.0));

    provider.set("[Book.xlsx]Sheet1", CellAddr { row: 0, col: 0 }, 2.0);
    provider.set("[Book.xlsx]Sheet2", CellAddr { row: 0, col: 0 }, 20.0);

    engine.mark_external_sheet_dirty("[Book.xlsx]Sheet1");
    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(2.0));
    // A2 should not be refreshed because Sheet2 was not invalidated.
    assert_eq!(engine.get_cell_value("Sheet1", "A2"), Value::Number(10.0));

    engine.mark_external_sheet_dirty("[Book.xlsx]Sheet2");
    engine.recalculate();
    assert_eq!(engine.get_cell_value("Sheet1", "A2"), Value::Number(20.0));
}

#[test]
fn external_sheet_invalidation_dirties_external_3d_span_dependents() {
    let provider = Arc::new(TestExternalProvider::default());
    provider.set_sheet_order(
        "Book.xlsx",
        vec![
            "Sheet1".to_string(),
            "Sheet2".to_string(),
            "Sheet3".to_string(),
        ],
    );
    provider.set("[Book.xlsx]Sheet1", CellAddr { row: 0, col: 0 }, 1.0);
    provider.set("[Book.xlsx]Sheet2", CellAddr { row: 0, col: 0 }, 10.0);
    provider.set("[Book.xlsx]Sheet3", CellAddr { row: 0, col: 0 }, 100.0);

    let mut engine = Engine::new();
    engine.set_external_value_provider(Some(provider.clone()));
    engine.set_external_refs_volatile(false);
    engine
        .set_cell_formula("Sheet1", "A1", "=SUM([Book.xlsx]Sheet1:Sheet3!A1)")
        .unwrap();
    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(111.0));

    // Mutate one sheet inside the span without invalidation; the value should not change because
    // external refs are non-volatile.
    provider.set("[Book.xlsx]Sheet2", CellAddr { row: 0, col: 0 }, 20.0);
    engine.recalculate();
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(111.0));

    engine.mark_external_sheet_dirty("[Book.xlsx]Sheet2");
    engine.recalculate();
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(121.0));
}

#[test]
fn external_sheet_invalidation_dirties_dynamic_external_dependents_from_indirect() {
    let provider = Arc::new(TestExternalProvider::default());
    provider.set("[Book.xlsx]Sheet1", CellAddr { row: 0, col: 0 }, 1.0);

    let mut engine = Engine::new();
    engine.set_external_value_provider(Some(provider.clone()));
    engine.set_calc_settings(CalcSettings {
        calculation_mode: CalculationMode::Automatic,
        ..CalcSettings::default()
    });
    engine
        .set_cell_formula("Sheet1", "A1", r#"=INDIRECT("[Book.xlsx]Sheet1!A1")"#)
        .unwrap();
    // Ensure we're exercising the bytecode backend + bytecode dependency tracing.
    assert!(
        engine.bytecode_compile_report(10).is_empty(),
        "{:?}",
        engine.bytecode_compile_report(10)
    );
    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(1.0));

    provider.set("[Book.xlsx]Sheet1", CellAddr { row: 0, col: 0 }, 2.0);
    engine.mark_external_sheet_dirty("[Book.xlsx]Sheet1");
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(2.0));

    provider.set("[Book.xlsx]Sheet1", CellAddr { row: 0, col: 0 }, 3.0);
    engine.mark_external_workbook_dirty("Book.xlsx");
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(3.0));
}

#[test]
fn degenerate_external_3d_sheet_range_ref_resolves_via_provider() {
    let provider = Arc::new(TestExternalProvider::default());
    provider.set("[Book.xlsx]Sheet1", CellAddr { row: 0, col: 0 }, 41.0);

    let mut engine = Engine::new();
    engine.set_external_value_provider(Some(provider));
    engine
        .set_cell_formula("Sheet1", "A1", "=[Book.xlsx]Sheet1:Sheet1!A1")
        .unwrap();
    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(41.0));
}

#[test]
fn degenerate_external_3d_sheet_range_ref_matches_endpoints_nfkc_case_insensitively_for_bytecode() {
    let provider = Arc::new(TestExternalProvider::default());
    provider.set("[Book.xlsx]Kelvin", CellAddr { row: 0, col: 0 }, 41.0);

    let mut engine = Engine::new();
    engine.set_external_value_provider(Some(provider));
    engine
        .set_cell_formula("Sheet1", "A1", "=[Book.xlsx]'Kelvin':'Kelvin'!A1")
        .unwrap();

    // Degenerate external 3D spans (`Sheet1:Sheet1`) are representable in the bytecode backend.
    // Ensure we treat NFKC-equivalent endpoints (`Kelvin` / `Kelvin`) as degenerate too so the
    // reference can be lowered as a single external sheet key (not rejected as an external span).
    assert!(
        engine.bytecode_compile_report(10).is_empty(),
        "{:?}",
        engine.bytecode_compile_report(10)
    );

    engine.recalculate();
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(41.0));
}

#[test]
fn sheet_function_matches_external_sheet_order_nfkc_case_insensitively() {
    let provider = Arc::new(TestExternalProvider::default());
    provider.set_sheet_order(
        "Book.xlsx",
        vec!["Kelvin".to_string(), "Sheet2".to_string()],
    );

    let mut engine = Engine::new();
    engine.set_external_value_provider(Some(provider));
    engine
        .set_cell_formula("Sheet1", "A1", "=SHEET([Book.xlsx]Kelvin!A1)")
        .unwrap();
    engine.recalculate();

    // "Kelvin" is sheet 1 in the external workbook sheet order.
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(1.0));
}

#[test]
fn external_3d_sheet_span_with_quoted_sheet_names_expands_via_provider_sheet_order() {
    let provider = Arc::new(TestExternalProvider::default());
    provider.set_sheet_order(
        "Book.xlsx",
        vec![
            "Sheet 1".to_string(),
            "Sheet 2".to_string(),
            "Sheet 3".to_string(),
        ],
    );
    for (sheet, value) in [("Sheet 1", 1.0), ("Sheet 2", 2.0), ("Sheet 3", 3.0)] {
        provider.set(
            &format!("[Book.xlsx]{sheet}"),
            CellAddr { row: 0, col: 0 },
            value,
        );
    }

    let mut engine = Engine::new();
    engine.set_external_value_provider(Some(provider));
    engine
        .set_cell_formula("Sheet1", "A1", "=SUM([Book.xlsx]'Sheet 1':'Sheet 3'!A1)")
        .unwrap();
    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(6.0));
}

#[test]
fn external_3d_sheet_span_matches_endpoints_case_insensitively() {
    let provider = Arc::new(TestExternalProvider::default());
    provider.set_sheet_order(
        "Book.xlsx",
        vec![
            "Sheet1".to_string(),
            "Sheet2".to_string(),
            "Sheet3".to_string(),
        ],
    );
    for (sheet, value) in [("Sheet1", 1.0), ("Sheet2", 2.0), ("Sheet3", 3.0)] {
        provider.set(
            &format!("[Book.xlsx]{sheet}"),
            CellAddr { row: 0, col: 0 },
            value,
        );
    }

    let mut engine = Engine::new();
    engine.set_external_value_provider(Some(provider));
    engine
        .set_cell_formula("Sheet1", "A1", "=SUM([Book.xlsx]sheet1:sheet3!A1)")
        .unwrap();
    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(6.0));
}

#[test]
fn external_3d_sheet_span_allows_reversed_endpoints() {
    let provider = Arc::new(TestExternalProvider::default());
    provider.set_sheet_order(
        "Book.xlsx",
        vec![
            "Sheet1".to_string(),
            "Sheet2".to_string(),
            "Sheet3".to_string(),
        ],
    );
    for (sheet, value) in [("Sheet1", 1.0), ("Sheet2", 2.0), ("Sheet3", 3.0)] {
        provider.set(
            &format!("[Book.xlsx]{sheet}"),
            CellAddr { row: 0, col: 0 },
            value,
        );
    }

    let mut engine = Engine::new();
    engine.set_external_value_provider(Some(provider));
    engine
        .set_cell_formula("Sheet1", "A1", "=SUM([Book.xlsx]Sheet3:Sheet1!A1)")
        .unwrap();
    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(6.0));
}

#[test]
fn sheet_and_sheets_over_external_3d_span_use_provider_tab_order() {
    let provider = Arc::new(TestExternalProvider::default());
    provider.set_sheet_order(
        "Book.xlsx",
        vec![
            "Sheet1".to_string(),
            "Sheet2".to_string(),
            "Sheet3".to_string(),
        ],
    );
    for (sheet, value) in [("Sheet1", 1.0), ("Sheet2", 2.0), ("Sheet3", 3.0)] {
        provider.set(
            &format!("[Book.xlsx]{sheet}"),
            CellAddr { row: 0, col: 0 },
            value,
        );
    }

    let mut engine = Engine::new();
    engine.set_external_value_provider(Some(provider));
    engine
        .set_cell_formula("Sheet1", "A1", "=SHEET([Book.xlsx]Sheet3:Sheet1!A1)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A2", "=SHEETS([Book.xlsx]Sheet3:Sheet1!A1)")
        .unwrap();
    engine.recalculate();

    // SHEET returns the first sheet in tab order for a multi-area reference.
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "A2"), Value::Number(3.0));
}

#[test]
fn index_over_external_3d_span_uses_provider_tab_order_when_lexicographic_order_differs() {
    // When sorting a multi-area reference union produced from an external workbook 3D span, we
    // must preserve the provider's sheet tab order (not lexicographic sheet name order).
    //
    // This matters for INDEX(..., area_num) semantics.
    let provider = Arc::new(TestExternalProvider::default());
    provider.set_sheet_order(
        "Book.xlsx",
        vec!["Sheet2".to_string(), "Sheet10".to_string()],
    );
    provider.set("[Book.xlsx]Sheet2", CellAddr { row: 0, col: 0 }, 2.0);
    provider.set("[Book.xlsx]Sheet10", CellAddr { row: 0, col: 0 }, 10.0);

    let mut engine = Engine::new();
    engine.set_external_value_provider(Some(provider));
    engine
        .set_cell_formula("Sheet1", "A1", "=INDEX([Book.xlsx]Sheet2:Sheet10!A1,1,1,1)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A2", "=INDEX([Book.xlsx]Sheet2:Sheet10!A1,1,1,2)")
        .unwrap();
    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "A2"), Value::Number(10.0));
}

#[test]
fn external_3d_sheet_span_matches_endpoints_nfkc_case_insensitively() {
    let provider = Arc::new(TestExternalProvider::default());

    // U+212A KELVIN SIGN (K) is compatibility-equivalent (NFKC) to ASCII 'K'.
    // Excel applies NFKC + case-insensitive matching when resolving sheet names.
    provider.set_sheet_order(
        "Book.xlsx",
        vec![
            "Kelvin".to_string(),
            "Sheet2".to_string(),
            "Sheet3".to_string(),
        ],
    );
    for (sheet, value) in [("Kelvin", 1.0), ("Sheet2", 2.0), ("Sheet3", 3.0)] {
        provider.set(
            &format!("[Book.xlsx]{sheet}"),
            CellAddr { row: 0, col: 0 },
            value,
        );
    }

    let mut engine = Engine::new();
    engine.set_external_value_provider(Some(provider));
    engine
        .set_cell_formula("Sheet1", "A1", "=SUM([Book.xlsx]Kelvin:Sheet3!A1)")
        .unwrap();
    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(6.0));
}

#[test]
fn external_3d_sheet_span_with_missing_endpoint_is_ref_error() {
    let provider = Arc::new(TestExternalProvider::default());
    provider.set_sheet_order(
        "Book.xlsx",
        vec![
            "Sheet1".to_string(),
            "Sheet2".to_string(),
            "Sheet3".to_string(),
        ],
    );

    let mut engine = Engine::new();
    engine.set_external_value_provider(Some(provider));
    engine
        .set_cell_formula("Sheet1", "A1", "=SUM([Book.xlsx]Sheet1:Sheet4!A1)")
        .unwrap();
    engine.recalculate();

    assert_eq!(
        engine.get_cell_value("Sheet1", "A1"),
        Value::Error(formula_engine::ErrorKind::Ref)
    );
}

#[test]
fn external_3d_sheet_range_refs_are_ref_error_even_if_provider_has_value() {
    let provider = Arc::new(TestExternalProvider::default());
    provider.set(
        "[Book.xlsx]Sheet1:Sheet3",
        CellAddr { row: 0, col: 0 },
        41.0,
    );

    let mut engine = Engine::new();
    engine.set_external_value_provider(Some(provider));
    engine
        .set_cell_formula("Sheet1", "A1", "=[Book.xlsx]Sheet1:Sheet3!A1")
        .unwrap();
    engine.recalculate();

    assert_eq!(
        engine.get_cell_value("Sheet1", "A1"),
        Value::Error(formula_engine::ErrorKind::Ref)
    );
}

#[test]
fn sum_over_external_table_structured_ref_resolves_via_provider_metadata() {
    let provider = Arc::new(TestExternalProvider::default());
    provider.set_table("Book.xlsx", "Sheet1", table_fixture_multi_col());

    provider.set("[Book.xlsx]Sheet1", CellAddr { row: 1, col: 1 }, 10.0);
    provider.set("[Book.xlsx]Sheet1", CellAddr { row: 2, col: 1 }, 20.0);
    provider.set("[Book.xlsx]Sheet1", CellAddr { row: 3, col: 1 }, 30.0);

    let mut engine = Engine::new();
    engine.set_external_value_provider(Some(provider));
    engine
        .set_cell_formula("Sheet1", "A1", "=SUM([Book.xlsx]Sheet1!Table1[Col2])")
        .unwrap();
    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(60.0));
}

#[test]
fn external_table_structured_ref_with_workbook_name_containing_escaped_rbracket_resolves() {
    let provider = Arc::new(TestExternalProvider::default());
    provider.set_table("Book[Name]].xlsx", "Sheet1", table_fixture_multi_col());

    provider.set(
        "[Book[Name]].xlsx]Sheet1",
        CellAddr { row: 1, col: 1 },
        10.0,
    );
    provider.set(
        "[Book[Name]].xlsx]Sheet1",
        CellAddr { row: 2, col: 1 },
        20.0,
    );
    provider.set(
        "[Book[Name]].xlsx]Sheet1",
        CellAddr { row: 3, col: 1 },
        30.0,
    );

    let mut engine = Engine::new();
    engine.set_external_value_provider(Some(provider));
    engine
        .set_cell_formula("Sheet1", "A1", "=SUM([Book[Name]].xlsx]Table1[Col2])")
        .unwrap();
    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(60.0));
}

#[test]
fn external_table_this_row_structured_ref_is_ref_error() {
    let provider = Arc::new(TestExternalProvider::default());
    provider.set_table("Book.xlsx", "Sheet1", table_fixture_multi_col());
    provider.set("[Book.xlsx]Sheet1", CellAddr { row: 1, col: 1 }, 10.0);

    let mut engine = Engine::new();
    engine.set_external_value_provider(Some(provider));
    engine
        .set_cell_formula("Sheet1", "A1", "=[Book.xlsx]Sheet1!Table1[@Col2]")
        .unwrap();
    engine.recalculate();

    assert_eq!(
        engine.get_cell_value("Sheet1", "A1"),
        Value::Error(formula_engine::ErrorKind::Ref)
    );
}

#[test]
fn indirect_external_3d_span_is_ref_error() {
    let provider = Arc::new(TestExternalProvider::default());

    let mut engine = Engine::new();
    engine.set_external_value_provider(Some(provider));
    engine
        .set_cell_formula(
            "Sheet1",
            "A1",
            r#"=INDIRECT("[Book.xlsx]Sheet1:Sheet3!A1")"#,
        )
        .unwrap();
    engine.recalculate();

    assert_eq!(
        engine.get_cell_value("Sheet1", "A1"),
        Value::Error(formula_engine::ErrorKind::Ref)
    );
}

#[test]
fn database_functions_reject_external_3d_sheet_spans_as_database_range() {
    let provider = Arc::new(TestExternalProvider::default());
    provider.set_sheet_order(
        "Book.xlsx",
        vec![
            "Sheet1".to_string(),
            "Sheet2".to_string(),
            "Sheet3".to_string(),
        ],
    );

    let mut engine = Engine::new();
    engine.set_external_value_provider(Some(provider));
    engine
        .set_cell_formula(
            "Sheet1",
            "A1",
            r#"=DSUM([Book.xlsx]Sheet1:Sheet3!A1:D4,"Salary",F1:F2)"#,
        )
        .unwrap();
    engine.recalculate();

    assert_eq!(
        engine.get_cell_value("Sheet1", "A1"),
        Value::Error(formula_engine::ErrorKind::Value)
    );
}

#[test]
fn database_functions_support_computed_criteria_over_external_database() {
    let provider = Arc::new(TestExternalProvider::default());

    // External database: [Book.xlsx]Sheet1!A1:D4 (header + 3 records).
    provider.set("[Book.xlsx]Sheet1", CellAddr { row: 0, col: 0 }, "Name");
    provider.set("[Book.xlsx]Sheet1", CellAddr { row: 0, col: 1 }, "Dept");
    provider.set("[Book.xlsx]Sheet1", CellAddr { row: 0, col: 2 }, "Age");
    provider.set("[Book.xlsx]Sheet1", CellAddr { row: 0, col: 3 }, "Salary");

    provider.set("[Book.xlsx]Sheet1", CellAddr { row: 1, col: 0 }, "Alice");
    provider.set("[Book.xlsx]Sheet1", CellAddr { row: 1, col: 1 }, "Sales");
    provider.set("[Book.xlsx]Sheet1", CellAddr { row: 1, col: 2 }, 30.0);
    provider.set("[Book.xlsx]Sheet1", CellAddr { row: 1, col: 3 }, 1000.0);

    provider.set("[Book.xlsx]Sheet1", CellAddr { row: 2, col: 0 }, "Bob");
    provider.set("[Book.xlsx]Sheet1", CellAddr { row: 2, col: 1 }, "Sales");
    provider.set("[Book.xlsx]Sheet1", CellAddr { row: 2, col: 2 }, 35.0);
    provider.set("[Book.xlsx]Sheet1", CellAddr { row: 2, col: 3 }, 1500.0);

    provider.set("[Book.xlsx]Sheet1", CellAddr { row: 3, col: 0 }, "Carol");
    provider.set("[Book.xlsx]Sheet1", CellAddr { row: 3, col: 1 }, "HR");
    provider.set("[Book.xlsx]Sheet1", CellAddr { row: 3, col: 2 }, 28.0);
    provider.set("[Book.xlsx]Sheet1", CellAddr { row: 3, col: 3 }, 1200.0);

    let mut engine = Engine::new();
    engine.set_external_value_provider(Some(provider));

    // Computed criteria: blank header + formula referencing the first database record row.
    engine.set_cell_formula("Sheet1", "F2", "=C2>30").unwrap();

    // Evaluate a representative set of D* functions over the external database range.
    engine
        .set_cell_formula(
            "Sheet1",
            "A1",
            "=DSUM([Book.xlsx]Sheet1!A1:D4,\"Salary\",F1:F2)",
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "A2",
            "=DGET([Book.xlsx]Sheet1!A1:D4,\"Salary\",F1:F2)",
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "A3",
            "=DCOUNT([Book.xlsx]Sheet1!A1:D4,\"Salary\",F1:F2)",
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "A4",
            "=DVARP([Book.xlsx]Sheet1!A1:D4,\"Salary\",F1:F2)",
        )
        .unwrap();
    engine.recalculate();

    // Age > 30 matches only Bob => sum salary = 1500.
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(1500.0));
    assert_eq!(engine.get_cell_value("Sheet1", "A2"), Value::Number(1500.0));
    assert_eq!(engine.get_cell_value("Sheet1", "A3"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "A4"), Value::Number(0.0));
}

#[test]
fn sheet_returns_external_sheet_index_when_provider_exposes_sheet_order() {
    let provider = Arc::new(TestExternalProvider::default());
    provider.set_sheet_order(
        "Book.xlsx",
        vec!["Sheet1".to_string(), "Sheet2".to_string()],
    );

    let mut engine = Engine::new();
    engine.set_external_value_provider(Some(provider));
    engine
        .set_cell_formula("Sheet1", "A1", "=SHEET([Book.xlsx]Sheet2!A1)")
        .unwrap();
    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(2.0));
}

#[test]
fn external_workbook_sheet_names_api_drives_sheet_and_span_expansion() {
    // Ensure hosts can implement only `ExternalValueProvider::workbook_sheet_names` (Arc-based)
    // and still get correct Excel semantics for:
    // - SHEET(...) external sheet index mapping
    // - external workbook 3D span expansion
    let provider = Arc::new(ProviderWithWorkbookSheetNamesOnly::default());
    provider.set_sheet_names(
        "Book.xlsx",
        vec![
            "Sheet1".to_string(),
            "Sheet2".to_string(),
            "Sheet3".to_string(),
        ],
    );
    for (sheet, value) in [("Sheet1", 1.0), ("Sheet2", 2.0), ("Sheet3", 3.0)] {
        provider.set(
            &format!("[Book.xlsx]{sheet}"),
            CellAddr { row: 0, col: 0 },
            value,
        );
    }

    let mut engine = Engine::new();
    engine.set_external_value_provider(Some(provider));
    engine
        .set_cell_formula("Sheet1", "A1", "=SHEET([Book.xlsx]Sheet2!A1)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A2", "=SUM([Book.xlsx]Sheet1:Sheet3!A1)")
        .unwrap();
    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "A2"), Value::Number(6.0));
}

#[test]
fn sheet_returns_na_for_external_refs_when_provider_lacks_sheet_order() {
    let provider = Arc::new(ProviderWithoutSheetOrder::default());

    let mut engine = Engine::new();
    engine.set_external_value_provider(Some(provider));
    engine
        .set_cell_formula("Sheet1", "A1", "=SHEET([Book.xlsx]Sheet2!A1)")
        .unwrap();
    engine.recalculate();

    assert_eq!(
        engine.get_cell_value("Sheet1", "A1"),
        Value::Error(formula_engine::ErrorKind::NA)
    );
}

#[test]
fn sheet_returns_na_for_external_refs_when_sheet_missing_from_order() {
    let provider = Arc::new(TestExternalProvider::default());
    provider.set_sheet_order("Book.xlsx", vec!["Sheet1".to_string()]);

    let mut engine = Engine::new();
    engine.set_external_value_provider(Some(provider));
    engine
        .set_cell_formula("Sheet1", "A1", "=SHEET([Book.xlsx]Sheet2!A1)")
        .unwrap();
    engine.recalculate();

    assert_eq!(
        engine.get_cell_value("Sheet1", "A1"),
        Value::Error(formula_engine::ErrorKind::NA)
    );
}
