use std::collections::HashMap;

use formula_engine::eval::{
    CellAddr, EvalContext, Evaluator, NameRef as EvalNameRef, RecalcContext,
};
use formula_engine::eval::{Expr as EvalExpr, ResolvedName, SheetReference, ValueResolver};
use formula_engine::functions::{ArgValue, FunctionContext};
use formula_engine::value::{ErrorKind, Value};
use formula_engine::{parse_formula, parse_formula_partial, Expr, ParseOptions};

fn sum_formula(arg_count: usize) -> String {
    let mut out = String::from("=SUM(");
    for i in 0..arg_count {
        if i > 0 {
            out.push(',');
        }
        out.push('1');
    }
    out.push(')');
    out
}

#[test]
fn strict_parse_rejects_more_than_255_args() {
    let formula = sum_formula(256);
    let err = parse_formula(&formula, ParseOptions::default()).unwrap_err();
    assert!(
        err.message.contains("max 255"),
        "unexpected error message: {}",
        err.message
    );
}

#[test]
fn partial_parse_records_error_and_caps_args_at_255() {
    let formula = sum_formula(256);
    let partial = parse_formula_partial(&formula, ParseOptions::default());
    let err = partial.error.expect("expected partial parse error");
    assert!(
        err.message
            .to_ascii_lowercase()
            .contains("too many arguments")
            || err.message.contains("max 255"),
        "unexpected error message: {}",
        err.message
    );

    match partial.ast.expr {
        Expr::FunctionCall(call) => {
            assert!(
                call.args.len() <= 255,
                "expected args <= 255, got {}",
                call.args.len()
            );
        }
        other => panic!("expected FunctionCall expr, got {other:?}"),
    }
}

#[derive(Default)]
struct TestResolver {
    #[allow(dead_code)]
    names: HashMap<String, ResolvedName>,
}

impl ValueResolver for TestResolver {
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

    fn resolve_name(&self, sheet_id: usize, name: &str) -> Option<ResolvedName> {
        if sheet_id != 0 {
            return None;
        }
        let key = name.trim().to_ascii_uppercase();
        self.names.get(&key).cloned()
    }
}

#[test]
fn evaluator_rejects_calls_with_more_than_255_args() {
    let resolver = TestResolver::default();
    let ctx = EvalContext {
        current_sheet: 0,
        current_cell: CellAddr { row: 0, col: 0 },
    };
    let recalc_ctx = RecalcContext::new(0);
    let evaluator = Evaluator::new(&resolver, ctx, &recalc_ctx);

    // Create a lambda with >255 params so the existing params-length check wouldn't reject
    // a 256-arg call.
    let params: Vec<String> = (0..300).map(|i| format!("p{i}")).collect();
    let lambda_value = evaluator.make_lambda(params, EvalExpr::Number(0.0));
    evaluator.set_local("F", ArgValue::Scalar(lambda_value));

    let callee_ref = EvalExpr::NameRef(EvalNameRef {
        sheet: SheetReference::Current,
        name: "F".to_string(),
    });

    let args = (0..256).map(|_| EvalExpr::Number(1.0)).collect::<Vec<_>>();
    let call_expr = EvalExpr::Call {
        callee: Box::new(callee_ref),
        args,
    };

    assert_eq!(
        evaluator.eval_formula(&call_expr),
        Value::Error(ErrorKind::Value)
    );
}
