#![cfg(not(target_arch = "wasm32"))]

use formula_engine::bytecode::{self, CellCoord, ColumnarGrid, Value, Vm};
use std::sync::Arc;

#[test]
fn bytecode_parser_supports_lambda_invocation_syntax() {
    let origin = CellCoord::new(0, 0);
    let locale = formula_engine::LocaleConfig::en_us();
    let grid = ColumnarGrid::new(1, 1);

    let expr = bytecode::parse_formula("=LAMBDA(x, x+1)(3)", origin).expect("parse");
    let program = bytecode::Compiler::compile(Arc::from("lambda_call"), &expr);

    let mut vm = Vm::with_capacity(32);
    let value = vm.eval(&program, &grid, 0, origin, &locale);
    assert_eq!(value, Value::Number(4.0));
}

#[test]
fn bytecode_parser_supports_let_bound_lambda_invocation() {
    let origin = CellCoord::new(0, 0);
    let locale = formula_engine::LocaleConfig::en_us();
    let grid = ColumnarGrid::new(1, 1);

    let expr = bytecode::parse_formula("=LET(f, LAMBDA(x, x+1), f(3))", origin).expect("parse");
    let program = bytecode::Compiler::compile(Arc::from("let_lambda_call"), &expr);

    let mut vm = Vm::with_capacity(32);
    let value = vm.eval(&program, &grid, 0, origin, &locale);
    assert_eq!(value, Value::Number(4.0));
}

