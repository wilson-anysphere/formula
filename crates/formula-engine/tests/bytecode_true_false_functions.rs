#![cfg(not(target_arch = "wasm32"))]

use formula_engine::bytecode::{self, CellCoord, ColumnarGrid, Value, Vm};
use std::sync::Arc;

#[test]
fn bytecode_true_false_zero_arg_functions_match_eval_ast() {
    let origin = CellCoord::new(0, 0);
    let locale = formula_engine::LocaleConfig::en_us();
    let grid = ColumnarGrid::new(1, 1);

    for (formula, expected) in [("=TRUE()", Value::Bool(true)), ("=FALSE()", Value::Bool(false))] {
        let expr = bytecode::parse_formula(formula, origin).expect("parse");
        let program = bytecode::Compiler::compile(Arc::from(formula), &expr);

        let mut vm = Vm::with_capacity(16);
        let vm_value = vm.eval(&program, &grid, 0, origin, &locale);
        let ast_value = bytecode::eval_ast(&expr, &grid, 0, origin, &locale);

        assert_eq!(vm_value, expected, "vm ({formula})");
        assert_eq!(ast_value, expected, "eval_ast ({formula})");
    }
}

#[test]
fn bytecode_true_false_reject_extra_args() {
    let origin = CellCoord::new(0, 0);
    let locale = formula_engine::LocaleConfig::en_us();
    let grid = ColumnarGrid::new(1, 1);

    for formula in ["=TRUE(1)", "=FALSE(1)"] {
        let expr = bytecode::parse_formula(formula, origin).expect("parse");
        let program = bytecode::Compiler::compile(Arc::from(formula), &expr);

        let mut vm = Vm::with_capacity(16);
        let vm_value = vm.eval(&program, &grid, 0, origin, &locale);
        let ast_value = bytecode::eval_ast(&expr, &grid, 0, origin, &locale);

        assert_eq!(vm_value, Value::Error(bytecode::ErrorKind::Value));
        assert_eq!(ast_value, Value::Error(bytecode::ErrorKind::Value));
    }
}

