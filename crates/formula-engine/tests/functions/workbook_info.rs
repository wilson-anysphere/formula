use std::collections::{HashMap, HashSet};
use std::sync::{Arc, Mutex};

use formula_engine::eval::{
    compile_canonical_expr, CellAddr, EvalContext, Evaluator, RecalcContext, ValueResolver,
};
use formula_engine::{Engine, ErrorKind, ExternalValueProvider, ParseOptions, Value};

#[derive(Default)]
struct TestExternalProvider {
    sheet_order: Mutex<HashMap<String, Vec<String>>>,
}

impl TestExternalProvider {
    fn set_sheet_order(&self, workbook: &str, order: impl Into<Vec<String>>) {
        self.sheet_order
            .lock()
            .expect("lock poisoned")
            .insert(workbook.to_string(), order.into());
    }
}

impl ExternalValueProvider for TestExternalProvider {
    fn get(&self, _sheet: &str, _addr: CellAddr) -> Option<Value> {
        None
    }

    fn sheet_order(&self, workbook: &str) -> Option<Vec<String>> {
        self.sheet_order
            .lock()
            .expect("lock poisoned")
            .get(workbook)
            .cloned()
    }
}

#[test]
fn sheet_reports_current_and_referenced_sheet_numbers() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet2", "A1", 2.0).unwrap();

    engine.set_cell_formula("Sheet1", "B1", "=SHEET()").unwrap();
    engine.set_cell_formula("Sheet2", "B1", "=SHEET()").unwrap();
    engine
        .set_cell_formula("Sheet2", "B2", "=SHEET(A1)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B2", "=SHEET(Sheet2!A1)")
        .unwrap();

    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet2", "B1"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet2", "B2"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "B2"), Value::Number(2.0));
}

#[test]
fn sheets_reports_workbook_sheet_count_and_3d_reference_span() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet2", "A1", 2.0).unwrap();
    engine.set_cell_value("Sheet3", "A1", 3.0).unwrap();

    engine
        .set_cell_formula("Sheet1", "B1", "=SHEETS()")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B2", "=SHEETS(Sheet1:Sheet3!A1)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B3", "=SHEET(Sheet1:Sheet3!A1)")
        .unwrap();

    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(3.0));
    assert_eq!(engine.get_cell_value("Sheet1", "B2"), Value::Number(3.0));
    assert_eq!(engine.get_cell_value("Sheet1", "B3"), Value::Number(1.0));
}

#[test]
fn sheets_3d_span_expands_by_tab_order_after_reorder() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet2", "A1", 2.0).unwrap();
    engine.set_cell_value("Sheet3", "A1", 3.0).unwrap();
    engine.set_cell_value("Sheet4", "A1", 4.0).unwrap();

    engine
        .set_cell_formula("Sheet1", "B1", "=SHEETS(Sheet1:Sheet3!A1)")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(3.0));

    // Move Sheet4 into the middle of the span so the set of referenced sheets changes.
    assert!(engine.reorder_sheet("Sheet4", 1));
    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(4.0));
}

#[test]
fn sheets_3d_span_excludes_deleted_intermediate_sheets() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet2", "A1", 2.0).unwrap();
    engine.set_cell_value("Sheet3", "A1", 3.0).unwrap();

    engine
        .set_cell_formula("Sheet1", "B1", "=SHEETS(Sheet1:Sheet3!A1)")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(3.0));

    // Deleting an intermediate sheet should shrink the span without producing #REF!.
    engine.delete_sheet("Sheet2").unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(2.0));
}

#[test]
fn sheet_uses_tab_order_after_reorder() {
    let mut engine = Engine::new();
    engine.set_cell_formula("Sheet1", "A1", "=SHEET()").unwrap();
    engine.set_cell_formula("Sheet2", "A1", "=SHEET()").unwrap();
    engine.set_cell_formula("Sheet3", "A1", "=SHEET()").unwrap();

    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet2", "A1"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet3", "A1"), Value::Number(3.0));

    assert!(engine.reorder_sheet("Sheet3", 0));
    engine.recalculate_single_threaded();

    // After reordering, the sheet ids are stable but SHEET() should reflect tab order.
    assert_eq!(engine.get_cell_value("Sheet3", "A1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet2", "A1"), Value::Number(3.0));
}

#[test]
fn sheet_3d_span_uses_tab_order_after_reorder() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet2", "A1", 2.0).unwrap();
    engine.set_cell_value("Sheet3", "A1", 3.0).unwrap();

    engine
        .set_cell_formula("Sheet1", "B1", "=SHEET(Sheet1:Sheet3!A1)")
        .unwrap();
    // This span does *not* include the reordered Sheet3 tab, so the result is sensitive to
    // whether we choose `min(sheet_id)` vs `min(tab_index)`.
    engine
        .set_cell_formula("Sheet1", "B2", "=SHEET(Sheet1:Sheet2!A1)")
        .unwrap();

    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "B2"), Value::Number(1.0));

    assert!(engine.reorder_sheet("Sheet3", 0));
    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "B2"), Value::Number(2.0));
}

#[test]
fn sheet_3d_span_reversed_uses_tab_order_after_reorder() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet2", "A1", 2.0).unwrap();
    engine.set_cell_value("Sheet3", "A1", 3.0).unwrap();

    // This is a reversed span (`Sheet3:Sheet2`) which should behave the same as `Sheet2:Sheet3`
    // (i.e. the included sheets are determined by workbook tab order, not textual direction).
    engine
        .set_cell_formula("Sheet1", "B1", "=SHEET(Sheet3:Sheet2!A1)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B2", "=SHEETS(Sheet3:Sheet2!A1)")
        .unwrap();

    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "B2"), Value::Number(2.0));

    assert!(engine.reorder_sheet("Sheet3", 0));
    engine.recalculate_single_threaded();

    // Tab order is now: Sheet3, Sheet1, Sheet2.
    // So the span `Sheet3:Sheet2` includes all three sheets and starts on Sheet3.
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "B2"), Value::Number(3.0));
}

#[test]
fn sheet_string_name_uses_tab_order_after_reorder() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet2", "A1", 2.0).unwrap();
    engine.set_cell_value("Sheet3", "A1", 3.0).unwrap();

    engine
        .set_cell_formula("Sheet1", "B1", "=SHEET(\"Sheet1\")")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B2", "=SHEET(\"Sheet3\")")
        .unwrap();

    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "B2"), Value::Number(3.0));

    assert!(engine.reorder_sheet("Sheet3", 0));
    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "B2"), Value::Number(1.0));
}

#[test]
fn sheet_reports_external_sheet_number_when_order_available() {
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
        .set_cell_formula("Sheet1", "A1", "=SHEET([Book.xlsx]Sheet2!A1)")
        .unwrap();
    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(2.0));
}

#[test]
fn sheet_reports_external_sheet_number_using_nfkc_case_insensitive_matching() {
    let provider = Arc::new(TestExternalProvider::default());
    // Use the Kelvin sign (U+212A) to ensure we match Excel's NFKC sheet name semantics.
    provider.set_sheet_order("Book.xlsx", vec!["Kelvin".to_string()]);

    let mut engine = Engine::new();
    engine.set_external_value_provider(Some(provider));
    engine
        .set_cell_formula("Sheet1", "A1", "=SHEET([Book.xlsx]KELVIN!A1)")
        .unwrap();
    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(1.0));
}

#[test]
fn sheet_reports_external_sheet_number_for_path_qualified_workbook_with_brackets() {
    let provider = Arc::new(TestExternalProvider::default());
    provider.set_sheet_order(
        r"C:\[foo]\Book.xlsx",
        vec!["Sheet1".to_string(), "Sheet2".to_string()],
    );

    let mut engine = Engine::new();
    engine.set_external_value_provider(Some(provider));
    engine
        .set_cell_formula("Sheet1", "A1", r"=SHEET('C:\[foo]\[Book.xlsx]Sheet2'!A1)")
        .unwrap();
    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(2.0));
}

#[test]
fn sheet_returns_na_for_external_sheet_when_order_unavailable() {
    let provider = Arc::new(TestExternalProvider::default());

    let mut engine = Engine::new();
    engine.set_external_value_provider(Some(provider));
    engine
        .set_cell_formula("Sheet1", "A1", "=SHEET([Book.xlsx]Sheet2!A1)")
        .unwrap();
    engine.recalculate_single_threaded();

    assert_eq!(
        engine.get_cell_value("Sheet1", "A1"),
        Value::Error(ErrorKind::NA)
    );
}

#[test]
fn sheet_reports_external_sheet_number_for_3d_span_argument() {
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
        .set_cell_formula("Sheet1", "A1", "=SHEET([Book.xlsx]Sheet2:Sheet3!A1)")
        .unwrap();
    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(2.0));
}

#[test]
fn formulatext_and_isformula_reflect_cell_formula_presence() {
    let mut engine = Engine::new();
    engine.set_cell_formula("Sheet1", "A1", "=1+1").unwrap();
    engine.set_cell_value("Sheet1", "A2", 5.0).unwrap();

    engine
        .set_cell_formula("Sheet1", "B1", "=FORMULATEXT(A1)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B2", "=FORMULATEXT(A2)")
        .unwrap();

    engine
        .set_cell_formula("Sheet1", "C1", "=ISFORMULA(A1)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "C2", "=ISFORMULA(A2)")
        .unwrap();

    engine.recalculate_single_threaded();

    assert_eq!(
        engine.get_cell_value("Sheet1", "B1"),
        Value::Text("=1+1".to_string())
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "B2"),
        Value::Error(ErrorKind::NA)
    );

    assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::Bool(true));
    assert_eq!(engine.get_cell_value("Sheet1", "C2"), Value::Bool(false));
}

#[test]
fn normalize_formula_text_does_not_duplicate_equals_for_leading_whitespace_formulas() {
    assert_eq!(
        formula_engine::functions::information::workbook::normalize_formula_text(" =1+1"),
        " =1+1".to_string()
    )
}

#[derive(Debug, Clone)]
struct TabOrderResolver {
    tab_order: Vec<usize>,
    existing: HashSet<usize>,
    sheet_names: HashMap<usize, String>,
    sheet_ids: HashMap<String, usize>,
}

impl TabOrderResolver {
    fn new(sheets: Vec<(usize, &str)>) -> Self {
        let mut tab_order = Vec::with_capacity(sheets.len());
        let mut existing = HashSet::with_capacity(sheets.len());
        let mut sheet_names = HashMap::with_capacity(sheets.len());
        let mut sheet_ids = HashMap::with_capacity(sheets.len());
        for (id, name) in sheets {
            tab_order.push(id);
            existing.insert(id);
            sheet_names.insert(id, name.to_string());
            sheet_ids.insert(name.to_string(), id);
        }
        Self {
            tab_order,
            existing,
            sheet_names,
            sheet_ids,
        }
    }

    fn reorder(&mut self, tab_order: Vec<usize>) {
        debug_assert_eq!(
            tab_order.iter().copied().collect::<HashSet<_>>(),
            self.existing
        );
        self.tab_order = tab_order;
    }

    fn delete_sheet(&mut self, sheet_id: usize) {
        self.existing.remove(&sheet_id);
        self.tab_order.retain(|id| *id != sheet_id);
        if let Some(name) = self.sheet_names.remove(&sheet_id) {
            self.sheet_ids.remove(&name);
        }
    }
}

impl ValueResolver for TabOrderResolver {
    fn sheet_exists(&self, sheet_id: usize) -> bool {
        self.existing.contains(&sheet_id)
    }

    fn sheet_count(&self) -> usize {
        self.tab_order.len()
    }

    fn sheet_order_index(&self, sheet_id: usize) -> Option<usize> {
        self.tab_order.iter().position(|id| *id == sheet_id)
    }

    fn get_cell_value(&self, _sheet_id: usize, _addr: CellAddr) -> Value {
        Value::Blank
    }

    fn sheet_name(&self, sheet_id: usize) -> Option<&str> {
        self.sheet_names.get(&sheet_id).map(|s| s.as_str())
    }

    fn sheet_id(&self, name: &str) -> Option<usize> {
        self.sheet_ids.get(name).copied()
    }

    fn resolve_structured_ref(
        &self,
        _ctx: EvalContext,
        _sref: &formula_engine::structured_refs::StructuredRef,
    ) -> Result<Vec<(usize, CellAddr, CellAddr)>, ErrorKind> {
        Err(ErrorKind::Name)
    }

    fn resolve_name(
        &self,
        _sheet_id: usize,
        _name: &str,
    ) -> Option<formula_engine::eval::ResolvedName> {
        None
    }
}

fn compile(formula: &str, resolver: &TabOrderResolver) -> formula_engine::eval::CompiledExpr {
    let ast = formula_engine::parse_formula(formula, ParseOptions::default()).unwrap();
    let current_cell = CellAddr { row: 0, col: 0 };
    let current_sheet = *resolver.tab_order.first().expect("at least one sheet");
    let mut resolve_sheet = |name: &str| resolver.sheet_id(name);
    let mut sheet_dimensions =
        |_sheet_id: usize| (formula_model::EXCEL_MAX_ROWS, formula_model::EXCEL_MAX_COLS);
    compile_canonical_expr(
        &ast.expr,
        current_sheet,
        current_cell,
        &mut resolve_sheet,
        &mut sheet_dimensions,
    )
}

fn eval(
    resolver: &TabOrderResolver,
    sheet: usize,
    expr: &formula_engine::eval::CompiledExpr,
) -> Value {
    let recalc_ctx = RecalcContext::new(0);
    let ctx = EvalContext {
        current_sheet: sheet,
        current_cell: CellAddr { row: 0, col: 0 },
    };
    Evaluator::new(resolver, ctx, &recalc_ctx).eval_formula(expr)
}

#[test]
fn sheet_uses_tab_order_and_updates_after_reorder() {
    // Use stable sheet ids that are *not* tab indices to ensure the implementation uses the tab
    // mapping rather than `sheet_id + 1`.
    let mut resolver = TabOrderResolver::new(vec![(10, "Sheet1"), (20, "Sheet2"), (30, "Sheet3")]);

    let sheet_expr = compile("=SHEET()", &resolver);
    let sheet1_expr = compile("=SHEET(\"Sheet1\")", &resolver);
    let sheet3_expr = compile("=SHEET(\"Sheet3\")", &resolver);

    // Initial tab order: Sheet1, Sheet2, Sheet3.
    assert_eq!(eval(&resolver, 10, &sheet_expr), Value::Number(1.0));
    assert_eq!(eval(&resolver, 10, &sheet1_expr), Value::Number(1.0));
    assert_eq!(eval(&resolver, 10, &sheet3_expr), Value::Number(3.0));

    // Reorder: Sheet3, Sheet1, Sheet2.
    resolver.reorder(vec![30, 10, 20]);
    assert_eq!(eval(&resolver, 10, &sheet_expr), Value::Number(2.0));
    assert_eq!(eval(&resolver, 10, &sheet1_expr), Value::Number(2.0));
    assert_eq!(eval(&resolver, 10, &sheet3_expr), Value::Number(1.0));
}

#[test]
fn sheets_decreases_after_sheet_deletion() {
    let mut resolver = TabOrderResolver::new(vec![(10, "Sheet1"), (20, "Sheet2"), (30, "Sheet3")]);
    let sheets_expr = compile("=SHEETS()", &resolver);
    assert_eq!(eval(&resolver, 10, &sheets_expr), Value::Number(3.0));

    resolver.delete_sheet(20);
    assert_eq!(eval(&resolver, 10, &sheets_expr), Value::Number(2.0));
}
