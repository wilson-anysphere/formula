#![cfg(not(target_arch = "wasm32"))]

use formula_engine::bytecode::ast::{Expr, Function};
use formula_engine::bytecode::{self, CellCoord, Value};
use std::sync::Arc;

#[derive(Clone, Copy)]
struct EmptyGrid;

impl bytecode::Grid for EmptyGrid {
    fn get_value(&self, _coord: CellCoord) -> Value {
        Value::Empty
    }

    fn column_slice(&self, _col: i32, _row_start: i32, _row_end: i32) -> Option<&[f64]> {
        None
    }

    fn bounds(&self) -> (i32, i32) {
        (10, 10)
    }
}

#[test]
fn bytecode_if_distinguishes_omitted_false_branch_from_explicit_missing_false_branch() {
    let origin = CellCoord::new(0, 0);
    let locale = formula_engine::LocaleConfig::en_us();
    let grid = EmptyGrid;

    // IF(FALSE, 1) => defaults to FALSE.
    let omitted = Expr::FuncCall {
        func: Function::If,
        args: vec![Expr::Literal(Value::Bool(false)), Expr::Literal(Value::Number(1.0))],
    };
    let omitted_program = bytecode::Compiler::compile(Arc::from("if_omitted_false_branch"), &omitted);
    let mut vm = bytecode::Vm::with_capacity(32);
    let omitted_vm = vm.eval(&omitted_program, &grid, 0, origin, &locale);
    let omitted_ast = bytecode::eval_ast(&omitted, &grid, 0, origin, &locale);
    assert_eq!(omitted_vm, Value::Bool(false));
    assert_eq!(omitted_ast, Value::Bool(false));

    // IF(FALSE, 1, ) => explicit empty argument yields a blank result.
    let explicit_missing = Expr::FuncCall {
        func: Function::If,
        args: vec![
            Expr::Literal(Value::Bool(false)),
            Expr::Literal(Value::Number(1.0)),
            Expr::Literal(Value::Missing),
        ],
    };
    let missing_program =
        bytecode::Compiler::compile(Arc::from("if_explicit_missing_false_branch"), &explicit_missing);
    let mut vm = bytecode::Vm::with_capacity(32);
    let missing_vm = vm.eval(&missing_program, &grid, 0, origin, &locale);
    let missing_ast = bytecode::eval_ast(&explicit_missing, &grid, 0, origin, &locale);
    assert_eq!(missing_vm, Value::Empty);
    assert_eq!(missing_ast, Value::Empty);
}

#[test]
fn bytecode_choose_returns_blank_for_selected_missing_choice() {
    let origin = CellCoord::new(0, 0);
    let locale = formula_engine::LocaleConfig::en_us();
    let grid = EmptyGrid;

    // CHOOSE(1, , 7) => blank.
    let expr = Expr::FuncCall {
        func: Function::Choose,
        args: vec![
            Expr::Literal(Value::Number(1.0)),
            Expr::Literal(Value::Missing),
            Expr::Literal(Value::Number(7.0)),
        ],
    };
    let program = bytecode::Compiler::compile(Arc::from("choose_selected_missing"), &expr);
    let mut vm = bytecode::Vm::with_capacity(32);
    let vm_value = vm.eval(&program, &grid, 0, origin, &locale);
    let ast_value = bytecode::eval_ast(&expr, &grid, 0, origin, &locale);

    assert_eq!(vm_value, Value::Empty);
    assert_eq!(ast_value, Value::Empty);
}

#[test]
fn bytecode_preserves_missing_for_address_optional_args() {
    // ADDRESS(1,1,,FALSE) should apply the default abs_num (1) even though arg 3 is present but
    // syntactically missing.
    let origin = CellCoord::new(0, 0);
    let locale = formula_engine::LocaleConfig::en_us();
    let grid = EmptyGrid;

    let expr = Expr::FuncCall {
        func: Function::Address,
        args: vec![
            Expr::Literal(Value::Number(1.0)),
            Expr::Literal(Value::Number(1.0)),
            Expr::Literal(Value::Missing),
            Expr::Literal(Value::Bool(false)),
        ],
    };
    let program = bytecode::Compiler::compile(Arc::from("address_missing_abs_num"), &expr);
    let mut vm = bytecode::Vm::with_capacity(32);
    let vm_value = vm.eval(&program, &grid, 0, origin, &locale);
    let ast_value = bytecode::eval_ast(&expr, &grid, 0, origin, &locale);

    assert_eq!(vm_value, Value::Text(Arc::from("R1C1")));
    assert_eq!(ast_value, Value::Text(Arc::from("R1C1")));
}

