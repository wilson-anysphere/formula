#![cfg(not(target_arch = "wasm32"))]

use formula_engine::bytecode::{
    self, ast::Function, CellCoord, ColumnarGrid, ErrorKind, Expr, Ref, RangeRef, Value, Vm,
};
use std::sync::Arc;

fn identity_lambda() -> Expr {
    let param = Arc::<str>::from("X");
    Expr::Lambda {
        params: Arc::from(vec![param.clone()].into_boxed_slice()),
        body: Box::new(Expr::NameRef(param)),
    }
}

fn range_a1_a2() -> Expr {
    Expr::RangeRef(RangeRef::new(
        Ref::new(0, 0, false, false),
        Ref::new(1, 0, false, false),
    ))
}

fn range_b1_b2() -> Expr {
    Expr::RangeRef(RangeRef::new(
        Ref::new(0, 1, false, false),
        Ref::new(1, 1, false, false),
    ))
}

#[test]
fn bytecode_match_rejects_lambda_lookup_value() {
    let origin = CellCoord::new(0, 0);
    let expr = Expr::FuncCall {
        func: Function::Match,
        args: vec![
            identity_lambda(),
            range_a1_a2(),
            Expr::Literal(Value::Number(0.0)),
        ],
    };
    let program = bytecode::Compiler::compile(Arc::from("match_lambda_lookup_value"), &expr);

    let grid = ColumnarGrid::new(10, 10);
    let locale = formula_engine::LocaleConfig::en_us();
    let mut vm = Vm::with_capacity(32);
    let value = vm.eval(&program, &grid, 0, origin, &locale);
    assert_eq!(value, Value::Error(ErrorKind::Value));
}

#[test]
fn bytecode_xmatch_rejects_lambda_lookup_value() {
    let origin = CellCoord::new(0, 0);
    let expr = Expr::FuncCall {
        func: Function::XMatch,
        args: vec![identity_lambda(), range_a1_a2()],
    };
    let program = bytecode::Compiler::compile(Arc::from("xmatch_lambda_lookup_value"), &expr);

    let grid = ColumnarGrid::new(10, 10);
    let locale = formula_engine::LocaleConfig::en_us();
    let mut vm = Vm::with_capacity(32);
    let value = vm.eval(&program, &grid, 0, origin, &locale);
    assert_eq!(value, Value::Error(ErrorKind::Value));
}

#[test]
fn bytecode_xlookup_rejects_lambda_lookup_value() {
    let origin = CellCoord::new(0, 0);
    let expr = Expr::FuncCall {
        func: Function::XLookup,
        args: vec![identity_lambda(), range_a1_a2(), range_b1_b2()],
    };
    let program = bytecode::Compiler::compile(Arc::from("xlookup_lambda_lookup_value"), &expr);

    let grid = ColumnarGrid::new(10, 10);
    let locale = formula_engine::LocaleConfig::en_us();
    let mut vm = Vm::with_capacity(32);
    let value = vm.eval(&program, &grid, 0, origin, &locale);
    assert_eq!(value, Value::Error(ErrorKind::Value));
}
