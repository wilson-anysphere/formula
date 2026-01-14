use formula_engine::eval::{
    compile_canonical_expr, CellAddr, EvalContext, Evaluator, RecalcContext, ValueResolver,
};
use formula_engine::{parse_formula, ErrorKind, ParseOptions, Value};

#[derive(Debug, Clone, Copy)]
struct GapResolver;

impl GapResolver {
    fn eval(&self, formula: &str) -> Value {
        let ast = parse_formula(formula, ParseOptions::default()).unwrap();
        let ctx = EvalContext {
            current_sheet: 0,
            current_cell: CellAddr { row: 0, col: 0 },
        };

        let mut resolve_sheet = |name: &str| match name {
            "Sheet1" => Some(0),
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

impl ValueResolver for GapResolver {
    fn sheet_exists(&self, sheet_id: usize) -> bool {
        matches!(sheet_id, 0 | 2)
    }

    fn get_cell_value(&self, _sheet_id: usize, _addr: CellAddr) -> Value {
        Value::Blank
    }

    fn resolve_structured_ref(
        &self,
        _ctx: EvalContext,
        _sref: &formula_engine::structured_refs::StructuredRef,
    ) -> Result<Vec<(usize, CellAddr, CellAddr)>, ErrorKind> {
        Err(ErrorKind::Name)
    }
}

#[test]
fn default_sheet_order_index_returns_none_for_missing_sheets() {
    let resolver = GapResolver;
    assert_eq!(resolver.sheet_order_index(1), None);
}

#[test]
fn default_expand_sheet_span_skips_missing_sheets_in_3d_spans() {
    let resolver = GapResolver;
    assert_eq!(
        resolver.eval("=SHEETS(Sheet1:Sheet3!A1)"),
        Value::Number(2.0)
    );
}
