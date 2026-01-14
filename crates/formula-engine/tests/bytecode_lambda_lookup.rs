#![cfg(not(target_arch = "wasm32"))]

use formula_engine::bytecode::{
    self, ast::Function, Array, CellCoord, ColumnarGrid, ErrorKind, Expr, Value, Vm,
};
use std::sync::Arc;

fn identity_lambda() -> Expr {
    let param = Arc::<str>::from("X");
    Expr::Lambda {
        params: Arc::from(vec![param.clone()].into_boxed_slice()),
        body: Box::new(Expr::NameRef(param)),
    }
}

fn array_a1_b2() -> Expr {
    Expr::Literal(Value::Array(Array::new(
        2,
        2,
        vec![
            Value::Number(1.0),
            Value::Number(10.0),
            Value::Number(2.0),
            Value::Number(20.0),
        ],
    )))
}

#[test]
fn bytecode_match_rejects_lambda_lookup_value() {
    let origin = CellCoord::new(0, 0);
    let expr = bytecode::parse_formula("=MATCH(LAMBDA(x,x),A1:A2,0)", origin).expect("parse");
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
    let expr = bytecode::parse_formula("=XMATCH(LAMBDA(x,x),A1:A2)", origin).expect("parse");
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
    let expr = bytecode::parse_formula("=XLOOKUP(LAMBDA(x,x),A1:A2,B1:B2)", origin).expect("parse");
    let program = bytecode::Compiler::compile(Arc::from("xlookup_lambda_lookup_value"), &expr);

    let grid = ColumnarGrid::new(10, 10);
    let locale = formula_engine::LocaleConfig::en_us();
    let mut vm = Vm::with_capacity(32);
    let value = vm.eval(&program, &grid, 0, origin, &locale);
    assert_eq!(value, Value::Error(ErrorKind::Value));
}

#[test]
fn bytecode_vlookup_rejects_lambda_lookup_value() {
    let origin = CellCoord::new(0, 0);
    let expr =
        bytecode::parse_formula("=VLOOKUP(LAMBDA(x,x),A1:B2,2,FALSE)", origin).expect("parse");
    let program = bytecode::Compiler::compile(Arc::from("vlookup_lambda_lookup_value"), &expr);

    let grid = ColumnarGrid::new(10, 10);
    let locale = formula_engine::LocaleConfig::en_us();
    let mut vm = Vm::with_capacity(32);
    let value = vm.eval(&program, &grid, 0, origin, &locale);
    assert_eq!(value, Value::Error(ErrorKind::Value));
}

#[test]
fn bytecode_hlookup_rejects_lambda_lookup_value() {
    let origin = CellCoord::new(0, 0);
    let expr =
        bytecode::parse_formula("=HLOOKUP(LAMBDA(x,x),A1:B2,2,FALSE)", origin).expect("parse");
    let program = bytecode::Compiler::compile(Arc::from("hlookup_lambda_lookup_value"), &expr);

    let grid = ColumnarGrid::new(10, 10);
    let locale = formula_engine::LocaleConfig::en_us();
    let mut vm = Vm::with_capacity(32);
    let value = vm.eval(&program, &grid, 0, origin, &locale);
    assert_eq!(value, Value::Error(ErrorKind::Value));
}

#[test]
fn bytecode_vlookup_rejects_lambda_lookup_value_array_table() {
    let origin = CellCoord::new(0, 0);
    let expr = Expr::FuncCall {
        func: Function::VLookup,
        args: vec![
            identity_lambda(),
            array_a1_b2(),
            Expr::Literal(Value::Number(1.0)),
            Expr::Literal(Value::Bool(false)),
        ],
    };
    let program =
        bytecode::Compiler::compile(Arc::from("vlookup_lambda_lookup_value_array"), &expr);

    let grid = ColumnarGrid::new(10, 10);
    let locale = formula_engine::LocaleConfig::en_us();
    let mut vm = Vm::with_capacity(32);
    let value = vm.eval(&program, &grid, 0, origin, &locale);
    assert_eq!(value, Value::Error(ErrorKind::Value));
}

#[test]
fn bytecode_hlookup_rejects_lambda_lookup_value_array_table() {
    let origin = CellCoord::new(0, 0);
    let expr = Expr::FuncCall {
        func: Function::HLookup,
        args: vec![
            identity_lambda(),
            array_a1_b2(),
            Expr::Literal(Value::Number(1.0)),
            Expr::Literal(Value::Bool(false)),
        ],
    };
    let program =
        bytecode::Compiler::compile(Arc::from("hlookup_lambda_lookup_value_array"), &expr);

    let grid = ColumnarGrid::new(10, 10);
    let locale = formula_engine::LocaleConfig::en_us();
    let mut vm = Vm::with_capacity(32);
    let value = vm.eval(&program, &grid, 0, origin, &locale);
    assert_eq!(value, Value::Error(ErrorKind::Value));
}
