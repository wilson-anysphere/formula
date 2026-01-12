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
    let omitted_formula = "=IF(FALSE,1)";
    let omitted = bytecode::parse_formula(omitted_formula, origin).expect("parse");
    if let Expr::FuncCall { func, args } = &omitted {
        assert_eq!(func, &Function::If);
        assert_eq!(args.len(), 2, "omitted IF should have 2 args");
    } else {
        panic!("expected IF func call: {omitted_formula}");
    }
    let omitted_program = bytecode::Compiler::compile(Arc::from(omitted_formula), &omitted);
    let mut vm = bytecode::Vm::with_capacity(32);
    let omitted_vm = vm.eval(&omitted_program, &grid, 0, origin, &locale);
    let omitted_ast = bytecode::eval_ast(&omitted, &grid, 0, origin, &locale);
    assert_eq!(omitted_vm, Value::Bool(false));
    assert_eq!(omitted_ast, Value::Bool(false));

    // IF(FALSE, 1, ) => explicit empty argument yields a blank result.
    let explicit_missing_formula = "=IF(FALSE,1,)";
    let explicit_missing =
        bytecode::parse_formula(explicit_missing_formula, origin).expect("parse");
    if let Expr::FuncCall { func, args } = &explicit_missing {
        assert_eq!(func, &Function::If);
        assert_eq!(args.len(), 3, "explicit-missing IF should have 3 args");
        assert_eq!(args[2], Expr::Literal(Value::Missing));
    } else {
        panic!("expected IF func call: {explicit_missing_formula}");
    }
    let missing_program =
        bytecode::Compiler::compile(Arc::from(explicit_missing_formula), &explicit_missing);
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
    let formula = "=CHOOSE(1,,7)";
    let expr = bytecode::parse_formula(formula, origin).expect("parse");
    if let Expr::FuncCall { func, args } = &expr {
        assert_eq!(func, &Function::Choose);
        assert_eq!(args.get(1), Some(&Expr::Literal(Value::Missing)));
    } else {
        panic!("expected CHOOSE func call: {formula}");
    }
    let program = bytecode::Compiler::compile(Arc::from(formula), &expr);
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

    let formula = "=ADDRESS(1,1,,FALSE)";
    let expr = bytecode::parse_formula(formula, origin).expect("parse");
    if let Expr::FuncCall { func, args } = &expr {
        assert_eq!(func, &Function::Address);
        assert_eq!(args.get(2), Some(&Expr::Literal(Value::Missing)));
    } else {
        panic!("expected ADDRESS func call: {formula}");
    }
    let program = bytecode::Compiler::compile(Arc::from(formula), &expr);
    let mut vm = bytecode::Vm::with_capacity(32);
    let vm_value = vm.eval(&program, &grid, 0, origin, &locale);
    let ast_value = bytecode::eval_ast(&expr, &grid, 0, origin, &locale);

    assert_eq!(vm_value, Value::Text(Arc::from("R1C1")));
    assert_eq!(ast_value, Value::Text(Arc::from("R1C1")));
}
