use formula_engine::eval::{
    compile_canonical_expr, CellAddr, EvalContext, Evaluator, RecalcContext,
};
use formula_engine::{parse_formula, ParseOptions, Value};

#[derive(Debug)]
struct TestResolver {
    /// Sheet id -> tab order index (0 = leftmost tab).
    tab_order_index: Vec<usize>,
}

impl TestResolver {
    fn eval(&self, formula: &str) -> Value {
        let ast = parse_formula(formula, ParseOptions::default()).unwrap();
        let ctx = EvalContext {
            current_sheet: 0,
            current_cell: CellAddr { row: 0, col: 0 },
        };

        let mut resolve_sheet = |name: &str| match name {
            "Sheet1" => Some(0),
            "Sheet2" => Some(1),
            "Sheet3" => Some(2),
            _ => None,
        };
        let mut sheet_dimensions =
            |_sheet_id: usize| (formula_model::EXCEL_MAX_ROWS, formula_model::EXCEL_MAX_COLS);
        let compiled = compile_canonical_expr(
            &ast.expr,
            ctx.current_sheet,
            ctx.current_cell,
            &mut resolve_sheet,
            &mut sheet_dimensions,
        );

        let recalc_ctx = RecalcContext::new(0);
        let evaluator = Evaluator::new(self, ctx, &recalc_ctx);
        evaluator.eval_formula(&compiled)
    }
}

impl formula_engine::eval::ValueResolver for TestResolver {
    fn sheet_exists(&self, sheet_id: usize) -> bool {
        sheet_id < 3
    }

    fn sheet_count(&self) -> usize {
        3
    }

    fn sheet_order_index(&self, sheet_id: usize) -> Option<usize> {
        self.tab_order_index.get(sheet_id).copied()
    }

    fn get_cell_value(&self, sheet_id: usize, addr: CellAddr) -> Value {
        if addr.row == 0 && addr.col == 0 {
            // A1 values: Sheet1=1, Sheet2=2, Sheet3=3.
            return Value::Number((sheet_id + 1) as f64);
        }
        Value::Blank
    }

    fn resolve_structured_ref(
        &self,
        _ctx: EvalContext,
        _sref: &formula_engine::structured_refs::StructuredRef,
    ) -> Option<Vec<(usize, CellAddr, CellAddr)>> {
        None
    }
}

#[test]
fn reference_union_uses_sheet_tab_order_for_index_area_num() {
    // Simulate sheets Sheet1..Sheet3 with stable ids [0, 1, 2], but tab order reversed:
    // [Sheet3, Sheet2, Sheet1].
    //
    // Excel orders multi-area references by sheet tab order, so INDEX(..., area_num) should follow
    // the reversed ordering.
    let resolver = TestResolver {
        tab_order_index: vec![2, 1, 0],
    };

    assert_eq!(
        resolver.eval("=SUM(INDEX(Sheet1:Sheet3!A1,1,1,1))"),
        Value::Number(3.0)
    );
    assert_eq!(
        resolver.eval("=SUM(INDEX(Sheet1:Sheet3!A1,1,1,2))"),
        Value::Number(2.0)
    );
    assert_eq!(
        resolver.eval("=SUM(INDEX(Sheet1:Sheet3!A1,1,1,3))"),
        Value::Number(1.0)
    );
}
