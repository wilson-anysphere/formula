use formula_engine::eval::{
    CellAddr, EvalContext, Evaluator, Expr as EvalExpr, NameRef, RecalcContext, SheetReference,
};
use formula_engine::functions::{ArgValue, FunctionContext};
use formula_engine::{
    parse_formula, parse_formula_partial, ErrorKind, ParseOptions, Value, EXCEL_MAX_ARGS,
};

#[test]
fn strict_parse_rejects_function_call_with_256_args() {
    let args = std::iter::repeat("1")
        .take(EXCEL_MAX_ARGS + 1)
        .collect::<Vec<_>>()
        .join(",");
    let formula = format!("=SUM({args})");

    let err = parse_formula(&formula, ParseOptions::default()).unwrap_err();
    assert!(
        err.message.contains("Too many arguments"),
        "unexpected error: {err}"
    );
}

#[test]
fn partial_parse_records_error_for_function_call_with_256_args() {
    let args = std::iter::repeat("1")
        .take(EXCEL_MAX_ARGS + 1)
        .collect::<Vec<_>>()
        .join(",");
    let formula = format!("=SUM({args})");

    let partial = parse_formula_partial(&formula, ParseOptions::default());
    let err = partial.error.expect("expected partial parse error");
    assert!(
        err.message.contains("Too many arguments"),
        "unexpected error: {err}"
    );
}

#[derive(Default)]
struct TestResolver;

impl formula_engine::eval::ValueResolver for TestResolver {
    fn sheet_exists(&self, sheet_id: usize) -> bool {
        sheet_id == 0
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
fn lambda_invocation_with_256_args_returns_value_error() {
    let resolver = TestResolver::default();
    let ctx = EvalContext {
        current_sheet: 0,
        current_cell: CellAddr { row: 0, col: 0 },
    };
    let recalc_ctx = RecalcContext::new(0);
    let evaluator = Evaluator::new(&resolver, ctx, &recalc_ctx);

    // Construct a lambda with 256 parameters (this is not possible via normal parsing due to
    // Excel's 255-arg limit, but can be constructed programmatically).
    let params = (0..(EXCEL_MAX_ARGS + 1))
        .map(|i| format!("p{i}"))
        .collect::<Vec<_>>();
    let lambda_value = evaluator.make_lambda(params, EvalExpr::Number(1.0));
    evaluator.set_local("F", ArgValue::Scalar(lambda_value));

    let call_expr = EvalExpr::Call {
        callee: Box::new(EvalExpr::NameRef(NameRef {
            sheet: SheetReference::Current,
            name: "F".to_string(),
        })),
        args: (0..(EXCEL_MAX_ARGS + 1))
            .map(|_| EvalExpr::Number(0.0))
            .collect(),
    };

    assert_eq!(
        evaluator.eval_formula(&call_expr),
        Value::Error(ErrorKind::Value)
    );
}

#[test]
fn strict_parse_rejects_postfix_call_with_256_args() {
    let args = std::iter::repeat("1")
        .take(EXCEL_MAX_ARGS + 1)
        .collect::<Vec<_>>()
        .join(",");
    let formula = format!("=LAMBDA(x,x)({args})");

    let err = parse_formula(&formula, ParseOptions::default()).unwrap_err();
    assert!(
        err.message.contains("Too many arguments"),
        "unexpected error: {err}"
    );
}

#[test]
fn partial_parse_records_error_for_postfix_call_with_256_args() {
    let args = std::iter::repeat("1")
        .take(EXCEL_MAX_ARGS + 1)
        .collect::<Vec<_>>()
        .join(",");
    let formula = format!("=LAMBDA(x,x)({args})");

    let partial = parse_formula_partial(&formula, ParseOptions::default());
    let err = partial.error.expect("expected partial parse error");
    assert!(
        err.message.contains("Too many arguments"),
        "unexpected error: {err}"
    );
}
