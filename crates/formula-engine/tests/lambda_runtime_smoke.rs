use std::collections::HashMap;
use std::sync::Arc;

use formula_engine::eval::{BinaryOp, Expr, NameRef, RangeRef, Ref};
use formula_engine::eval::{
    CellAddr, EvalContext, Evaluator, RecalcContext, ResolvedName, SheetReference, ValueResolver,
};
use formula_engine::functions::{ArgValue, FunctionContext, Reference, SheetId};
use formula_engine::value::{ErrorKind, Lambda, Value};
use formula_engine::Engine;

#[derive(Default)]
struct TestResolver {
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
fn lambda_value_captures_locals_and_supports_call_syntax() {
    let resolver = TestResolver::default();
    let ctx = EvalContext {
        current_sheet: 0,
        current_cell: CellAddr { row: 0, col: 0 },
    };
    let recalc_ctx = RecalcContext::new(0);
    let evaluator = Evaluator::new(&resolver, ctx, &recalc_ctx);

    // Capture `a` from the local LET scope.
    evaluator.push_local_scope();
    evaluator.set_local("a", ArgValue::Scalar(Value::Number(1.0)));

    let body = Expr::Binary {
        op: BinaryOp::Add,
        left: Box::new(Expr::NameRef(NameRef {
            sheet: SheetReference::Current,
            name: "x".to_string(),
        })),
        right: Box::new(Expr::NameRef(NameRef {
            sheet: SheetReference::Current,
            name: "a".to_string(),
        })),
    };

    let lambda_value = evaluator.make_lambda(vec!["x".to_string()], body);

    // Drop the defining scope to ensure the lambda really captured it.
    evaluator.pop_local_scope();

    // Store the lambda in a local binding so we can invoke it.
    evaluator.set_local("F", ArgValue::Scalar(lambda_value));

    let callee_ref = Expr::NameRef(NameRef {
        sheet: SheetReference::Current,
        name: "F".to_string(),
    });

    // Postfix-call evaluation: `expr(args)`.
    let postfix_call = Expr::Call {
        callee: Box::new(callee_ref.clone()),
        args: vec![Expr::Number(3.0)],
    };
    assert_eq!(evaluator.eval_formula(&postfix_call), Value::Number(4.0));

    // Name-call fallback: `Foo(args)` where Foo resolves to a lambda name/local.
    let named_call = Expr::FunctionCall {
        name: "F".to_string(),
        original_name: "F".to_string(),
        args: vec![Expr::Number(3.0)],
    };
    assert_eq!(evaluator.eval_formula(&named_call), Value::Number(4.0));

    assert!(matches!(
        evaluator.eval_formula(&callee_ref),
        Value::Lambda(_)
    ));
}

#[test]
fn local_name_refs_shadow_defined_names() {
    let mut resolver = TestResolver::default();
    resolver
        .names
        .insert("X".to_string(), ResolvedName::Constant(Value::Number(2.0)));

    let ctx = EvalContext {
        current_sheet: 0,
        current_cell: CellAddr { row: 0, col: 0 },
    };
    let recalc_ctx = RecalcContext::new(0);
    let evaluator = Evaluator::new(&resolver, ctx, &recalc_ctx);

    evaluator.set_local("x", ArgValue::Scalar(Value::Number(1.0)));

    let x_ref = Expr::NameRef(NameRef {
        sheet: SheetReference::Current,
        name: "x".to_string(),
    });
    assert_eq!(evaluator.eval_formula(&x_ref), Value::Number(1.0));
}

#[test]
fn local_bindings_preserve_references() {
    let resolver = TestResolver::default();
    let ctx = EvalContext {
        current_sheet: 0,
        current_cell: CellAddr { row: 0, col: 0 },
    };
    let recalc_ctx = RecalcContext::new(0);
    let evaluator = Evaluator::new(&resolver, ctx, &recalc_ctx);

    let reference = Reference {
        sheet_id: SheetId::Local(0),
        start: CellAddr { row: 0, col: 0 },
        end: CellAddr { row: 2, col: 0 },
    };
    evaluator.set_local("r", ArgValue::Reference(reference.clone()));

    let r_ref = Expr::NameRef(NameRef {
        sheet: SheetReference::Current,
        name: "r".to_string(),
    });
    assert_eq!(evaluator.eval_arg(&r_ref), ArgValue::Reference(reference));
}

#[test]
fn lambda_recursion_is_depth_limited() {
    let mut resolver = TestResolver::default();

    let recursive_body = Expr::FunctionCall {
        name: "F".to_string(),
        original_name: "F".to_string(),
        args: vec![Expr::NameRef(NameRef {
            sheet: SheetReference::Current,
            name: "x".to_string(),
        })],
    };

    let recursive_lambda = Lambda {
        params: vec!["x".to_string()].into(),
        body: Arc::new(recursive_body),
        env: Arc::new(HashMap::new()),
    };

    resolver.names.insert(
        "F".to_string(),
        ResolvedName::Constant(Value::Lambda(recursive_lambda)),
    );

    let ctx = EvalContext {
        current_sheet: 0,
        current_cell: CellAddr { row: 0, col: 0 },
    };
    let recalc_ctx = RecalcContext::new(0);
    let evaluator = Evaluator::new(&resolver, ctx, &recalc_ctx);

    let expr = Expr::FunctionCall {
        name: "F".to_string(),
        original_name: "F".to_string(),
        args: vec![Expr::Number(1.0)],
    };

    assert_eq!(evaluator.eval_formula(&expr), Value::Error(ErrorKind::Calc));
}

#[test]
fn lambda_calls_can_return_reference_values() {
    let resolver = TestResolver::default();
    let ctx = EvalContext {
        current_sheet: 0,
        current_cell: CellAddr { row: 0, col: 0 },
    };
    let recalc_ctx = RecalcContext::new(0);
    let evaluator = Evaluator::new(&resolver, ctx, &recalc_ctx);

    // LAMBDA(r, r) - identity function for references.
    let body = Expr::NameRef(NameRef {
        sheet: SheetReference::Current,
        name: "r".to_string(),
    });
    let lambda_value = evaluator.make_lambda(vec!["r".to_string()], body);
    evaluator.set_local("F", ArgValue::Scalar(lambda_value));

    let range_expr = Expr::RangeRef(RangeRef {
        sheet: SheetReference::Current,
        start: Ref::from_abs_cell_addr(CellAddr { row: 0, col: 0 }).unwrap(),
        end: Ref::from_abs_cell_addr(CellAddr { row: 2, col: 0 }).unwrap(),
    });

    let call_expr = Expr::FunctionCall {
        name: "F".to_string(),
        original_name: "F".to_string(),
        args: vec![range_expr],
    };

    let expected = Reference {
        sheet_id: SheetId::Local(0),
        start: CellAddr { row: 0, col: 0 },
        end: CellAddr { row: 2, col: 0 },
    };

    assert_eq!(
        evaluator.eval_arg(&call_expr),
        ArgValue::Reference(expected)
    );
}

#[test]
fn engine_coerces_top_level_lambda_results_to_calc_error() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=LAMBDA(x,x)")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Sheet1", "A1"),
        Value::Error(ErrorKind::Calc)
    );
}

#[test]
fn engine_coerces_lambda_values_inside_spilled_arrays_to_calc_error() {
    let mut engine = Engine::new();

    // IF is array-enabled and can produce a spilled array result. Ensure any lambda element is
    // stored as #CALC! rather than a callable Value::Lambda.
    engine
        // Use numeric truthy/falsy values inside the array literal to avoid locale/parser
        // differences in boolean literal handling.
        .set_cell_formula("Sheet1", "A1", "=IF({1,0}, LAMBDA(x,x), 0)")
        .unwrap();

    engine.recalculate_single_threaded();

    // Spill origin: lambda coerced to #CALC!.
    assert_eq!(
        engine.get_cell_value("Sheet1", "A1"),
        Value::Error(ErrorKind::Calc)
    );
    // Spill output: scalar branch preserved.
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(0.0));
}
