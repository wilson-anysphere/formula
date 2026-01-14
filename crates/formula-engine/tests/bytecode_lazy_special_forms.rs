#![cfg(not(target_arch = "wasm32"))]

use formula_engine::bytecode::{self, CellCoord, Value};
use std::sync::Arc;

#[derive(Clone, Copy)]
struct PanicGrid {
    panic_coord: CellCoord,
}

impl bytecode::Grid for PanicGrid {
    fn get_value(&self, coord: CellCoord) -> Value {
        if coord == self.panic_coord {
            panic!("attempted to evaluate forbidden cell reference at {coord:?}");
        }
        Value::Empty
    }

    fn column_slice(&self, _col: i32, _row_start: i32, _row_end: i32) -> Option<&[f64]> {
        None
    }

    fn bounds(&self) -> (i32, i32) {
        (10, 10)
    }
}

#[derive(Clone, Copy)]
struct PanicGridWithNumber {
    number_coord: CellCoord,
    number: f64,
    panic_coord: CellCoord,
}

impl bytecode::Grid for PanicGridWithNumber {
    fn get_value(&self, coord: CellCoord) -> Value {
        if coord == self.panic_coord {
            panic!("attempted to evaluate forbidden cell reference at {coord:?}");
        }
        if coord == self.number_coord {
            return Value::Number(self.number);
        }
        Value::Empty
    }

    fn column_slice(&self, _col: i32, _row_start: i32, _row_end: i32) -> Option<&[f64]> {
        None
    }

    fn bounds(&self) -> (i32, i32) {
        (10, 10)
    }
}

#[derive(Clone)]
struct TextGrid {
    coord: CellCoord,
    value: Value,
}

impl bytecode::Grid for TextGrid {
    fn get_value(&self, coord: CellCoord) -> Value {
        if coord == self.coord {
            return self.value.clone();
        }
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
fn bytecode_choose_is_lazy() {
    // CHOOSE(2, <unused>, 7) must not evaluate the first choice expression.
    let origin = CellCoord::new(0, 0);
    let expr = bytecode::parse_formula("=CHOOSE(2, A2, 7)", origin).expect("parse");
    let program = bytecode::Compiler::compile(Arc::from("choose_lazy"), &expr);

    let grid = PanicGrid {
        // A2 relative to origin (A1) => (row=1, col=0)
        panic_coord: CellCoord::new(1, 0),
    };

    let mut vm = bytecode::Vm::with_capacity(32);
    let value = vm.eval(
        &program,
        &grid,
        0,
        origin,
        &formula_engine::LocaleConfig::en_us(),
    );
    assert_eq!(value, Value::Number(7.0));
}

#[test]
fn bytecode_eval_ast_choose_is_lazy() {
    // `bytecode::eval_ast` should match VM semantics and avoid evaluating unused CHOOSE branches.
    let origin = CellCoord::new(0, 0);
    let expr = bytecode::parse_formula("=CHOOSE(2, A2, 7)", origin).expect("parse");

    let grid = PanicGrid {
        // A2 relative to origin (A1) => (row=1, col=0)
        panic_coord: CellCoord::new(1, 0),
    };

    let value = bytecode::eval_ast(
        &expr,
        &grid,
        0,
        origin,
        &formula_engine::LocaleConfig::en_us(),
    );
    assert_eq!(value, Value::Number(7.0));
}

#[test]
fn bytecode_eval_ast_choose_matches_vm_reference_semantics() {
    // CHOOSE should evaluate its selected argument in "argument mode" (preserving direct cell
    // references as references), matching the VM compiler which lowers selected cell refs to
    // Range values.
    //
    // This matters when the CHOOSE result is consumed by a function with range/reference semantics
    // like SUM: if CHOOSE returns the *value* of a text cell, SUM will parse it, but if CHOOSE
    // returns a reference, SUM will ignore the text (Excel/engine semantics).
    let origin = CellCoord::new(0, 0);
    let expr = bytecode::parse_formula("=SUM(CHOOSE(1, A1, 0))", origin).expect("parse");
    let program = bytecode::Compiler::compile(Arc::from("choose_sum_ref_semantics"), &expr);

    let grid = TextGrid {
        coord: CellCoord::new(0, 0),
        value: Value::Text(Arc::from("5")),
    };
    let locale = formula_engine::LocaleConfig::en_us();

    let mut vm = bytecode::Vm::with_capacity(64);
    let vm_value = vm.eval(&program, &grid, 0, origin, &locale);

    let ast_value = bytecode::eval_ast(&expr, &grid, 0, origin, &locale);

    assert_eq!(ast_value, vm_value);
    assert_eq!(vm_value, Value::Number(0.0));
}

#[test]
fn bytecode_choose_selected_cell_ref_coerces_in_if_condition() {
    // CHOOSE can return a reference value (even for single cells). When used as a logical
    // condition, that single-cell reference should behave like a scalar (matching the evaluator's
    // `eval_arg` behavior for IF conditions).
    let origin = CellCoord::new(0, 0);
    let expr = bytecode::parse_formula("=IF(CHOOSE(1, A1, FALSE), 1, 0)", origin).expect("parse");
    let program = bytecode::Compiler::compile(Arc::from("choose_if_condition"), &expr);
    let locale = formula_engine::LocaleConfig::en_us();

    for (label, a1, expected) in [
        ("false", Value::Bool(false), Value::Number(0.0)),
        ("true", Value::Bool(true), Value::Number(1.0)),
    ] {
        let grid = TextGrid {
            coord: CellCoord::new(0, 0),
            value: a1,
        };

        let mut vm = bytecode::Vm::with_capacity(64);
        let vm_value = vm.eval(&program, &grid, 0, origin, &locale);
        assert_eq!(vm_value, expected, "vm ({label})");

        let ast_value = bytecode::eval_ast(&expr, &grid, 0, origin, &locale);
        assert_eq!(ast_value, expected, "eval_ast ({label})");
    }
}

#[test]
fn bytecode_choose_selected_cell_ref_coerces_in_not() {
    // NOT uses boolean coercion on its argument. When CHOOSE selects a single-cell reference,
    // that reference should be dereferenced before coercion so NOT behaves like `NOT(A1)`.
    let origin = CellCoord::new(0, 0);
    let expr = bytecode::parse_formula("=NOT(CHOOSE(1, A1, TRUE))", origin).expect("parse");
    let program = bytecode::Compiler::compile(Arc::from("choose_not"), &expr);
    let locale = formula_engine::LocaleConfig::en_us();

    let grid = TextGrid {
        coord: CellCoord::new(0, 0),
        value: Value::Bool(false),
    };

    let mut vm = bytecode::Vm::with_capacity(64);
    let vm_value = vm.eval(&program, &grid, 0, origin, &locale);
    assert_eq!(vm_value, Value::Bool(true));

    let ast_value = bytecode::eval_ast(&expr, &grid, 0, origin, &locale);
    assert_eq!(ast_value, Value::Bool(true));
}

#[test]
fn bytecode_choose_nan_index_is_value_error_and_does_not_evaluate_choices() {
    // NaN should coerce to an invalid CHOOSE index (0) and yield #VALUE!, without evaluating any
    // choice expressions.
    //
    // This is a regression test for implementations that normalize via INT(index) + comparisons:
    // Excel-style comparisons treat NaN as "equal" for ordering purposes, which can incorrectly
    // select a branch.
    let origin = CellCoord::new(0, 0);
    // Use a non-finite cell value for the index so this test doesn't depend on the engine's
    // arithmetic overflow/NaN behavior.
    let expr = bytecode::parse_formula("=CHOOSE(B1, A2, 7)", origin).expect("parse");
    let program = bytecode::Compiler::compile(Arc::from("choose_nan_index"), &expr);

    let grid = PanicGridWithNumber {
        // B1 relative to origin (A1) => (row=0, col=1)
        number_coord: CellCoord::new(0, 1),
        number: f64::NAN,
        // A2 relative to origin (A1) => (row=1, col=0)
        panic_coord: CellCoord::new(1, 0),
    };

    let mut vm = bytecode::Vm::with_capacity(32);
    let value = vm.eval(
        &program,
        &grid,
        0,
        origin,
        &formula_engine::LocaleConfig::en_us(),
    );
    assert_eq!(value, Value::Error(bytecode::ErrorKind::Value));
}

#[test]
fn bytecode_eval_ast_choose_nan_index_is_value_error_and_does_not_evaluate_choices() {
    // Ensure `bytecode::eval_ast` matches the VM semantics when the CHOOSE index is NaN: the
    // index should coerce to 0, yielding #VALUE!, and CHOOSE must not evaluate any choice
    // expressions.
    let origin = CellCoord::new(0, 0);
    // Use a non-finite cell value for the index so this test doesn't depend on the engine's
    // arithmetic overflow/NaN behavior.
    let expr = bytecode::parse_formula("=CHOOSE(B1, A2, 7)", origin).expect("parse");

    let grid = PanicGridWithNumber {
        // B1 relative to origin (A1) => (row=0, col=1)
        number_coord: CellCoord::new(0, 1),
        number: f64::NAN,
        // A2 relative to origin (A1) => (row=1, col=0)
        panic_coord: CellCoord::new(1, 0),
    };

    let value = bytecode::eval_ast(
        &expr,
        &grid,
        0,
        origin,
        &formula_engine::LocaleConfig::en_us(),
    );
    assert_eq!(value, Value::Error(bytecode::ErrorKind::Value));
}

#[test]
fn bytecode_choose_in_scalar_context_returns_scalar_value() {
    // Regression test: when CHOOSE is used in a scalar-function context (like ABS), it must not
    // return a reference value for a selected cell reference argument, otherwise scalar bytecode
    // functions will treat it as a spill attempt.
    let origin = CellCoord::new(0, 0);
    let expr = bytecode::parse_formula("=ABS(CHOOSE(1, A2, 2))", origin).expect("parse");
    let program = bytecode::Compiler::compile(Arc::from("choose_scalar_context"), &expr);

    #[derive(Clone, Copy)]
    struct ValueGrid {
        coord: CellCoord,
        value: f64,
    }

    impl bytecode::Grid for ValueGrid {
        fn get_value(&self, coord: CellCoord) -> Value {
            if coord == self.coord {
                return Value::Number(self.value);
            }
            Value::Empty
        }

        fn column_slice(&self, _col: i32, _row_start: i32, _row_end: i32) -> Option<&[f64]> {
            None
        }

        fn bounds(&self) -> (i32, i32) {
            (10, 10)
        }
    }

    let grid = ValueGrid {
        // A2 relative to origin (A1) => (row=1, col=0)
        coord: CellCoord::new(1, 0),
        value: -5.0,
    };

    let mut vm = bytecode::Vm::with_capacity(32);
    let value = vm.eval(
        &program,
        &grid,
        0,
        origin,
        &formula_engine::LocaleConfig::en_us(),
    );
    assert_eq!(value, Value::Number(5.0));
}

#[test]
fn bytecode_eval_ast_choose_in_scalar_context_returns_scalar_value() {
    // `bytecode::eval_ast` should match VM semantics for CHOOSE in scalar contexts.
    let origin = CellCoord::new(0, 0);
    let expr = bytecode::parse_formula("=ABS(CHOOSE(1, A2, 2))", origin).expect("parse");

    #[derive(Clone, Copy)]
    struct ValueGrid {
        coord: CellCoord,
        value: f64,
    }

    impl bytecode::Grid for ValueGrid {
        fn get_value(&self, coord: CellCoord) -> Value {
            if coord == self.coord {
                return Value::Number(self.value);
            }
            Value::Empty
        }

        fn column_slice(&self, _col: i32, _row_start: i32, _row_end: i32) -> Option<&[f64]> {
            None
        }

        fn bounds(&self) -> (i32, i32) {
            (10, 10)
        }
    }

    let grid = ValueGrid {
        // A2 relative to origin (A1) => (row=1, col=0)
        coord: CellCoord::new(1, 0),
        value: -5.0,
    };

    let value = bytecode::eval_ast(
        &expr,
        &grid,
        0,
        origin,
        &formula_engine::LocaleConfig::en_us(),
    );
    assert_eq!(value, Value::Number(5.0));
}

#[test]
fn bytecode_ifs_is_lazy() {
    // IFS(TRUE, 7, <unused_cond>, 8) must not evaluate the second condition/value pair.
    let origin = CellCoord::new(0, 0);
    let expr = bytecode::parse_formula("=IFS(TRUE, 7, A2, 8)", origin).expect("parse");
    let program = bytecode::Compiler::compile(Arc::from("ifs_lazy"), &expr);

    let grid = PanicGrid {
        panic_coord: CellCoord::new(1, 0),
    };

    let mut vm = bytecode::Vm::with_capacity(32);
    let value = vm.eval(
        &program,
        &grid,
        0,
        origin,
        &formula_engine::LocaleConfig::en_us(),
    );
    assert_eq!(value, Value::Number(7.0));
}

#[test]
fn bytecode_eval_ast_ifs_is_lazy() {
    // `bytecode::eval_ast` should match VM semantics and avoid evaluating unused IFS branches.
    let origin = CellCoord::new(0, 0);
    let expr = bytecode::parse_formula("=IFS(TRUE, 7, A2, 8)", origin).expect("parse");

    let grid = PanicGrid {
        // A2 relative to origin (A1) => (row=1, col=0)
        panic_coord: CellCoord::new(1, 0),
    };

    let value = bytecode::eval_ast(
        &expr,
        &grid,
        0,
        origin,
        &formula_engine::LocaleConfig::en_us(),
    );
    assert_eq!(value, Value::Number(7.0));
}

#[test]
fn bytecode_eval_ast_ifs_error_short_circuits() {
    // If an IFS condition is an error, IFS should return that error without evaluating any value
    // expressions.
    let origin = CellCoord::new(0, 0);
    let expr = bytecode::parse_formula("=IFS(1/0, A2, TRUE, 7)", origin).expect("parse");

    let grid = PanicGrid {
        // A2 relative to origin (A1) => (row=1, col=0)
        panic_coord: CellCoord::new(1, 0),
    };

    let value = bytecode::eval_ast(
        &expr,
        &grid,
        0,
        origin,
        &formula_engine::LocaleConfig::en_us(),
    );
    assert_eq!(value, Value::Error(bytecode::ErrorKind::Div0));
}

#[test]
fn bytecode_switch_is_lazy_for_unmatched_values() {
    // SWITCH(1, 2, <unused_value>, 1, 7) must not evaluate the value for the non-matching case.
    let origin = CellCoord::new(0, 0);
    let expr = bytecode::parse_formula("=SWITCH(1, 2, A2, 1, 7)", origin).expect("parse");
    let program = bytecode::Compiler::compile(Arc::from("switch_lazy"), &expr);

    let grid = PanicGrid {
        panic_coord: CellCoord::new(1, 0),
    };

    let mut vm = bytecode::Vm::with_capacity(32);
    let value = vm.eval(
        &program,
        &grid,
        0,
        origin,
        &formula_engine::LocaleConfig::en_us(),
    );
    assert_eq!(value, Value::Number(7.0));
}

#[test]
fn bytecode_eval_ast_switch_is_lazy_for_unmatched_values() {
    // `bytecode::eval_ast` should match VM semantics and avoid evaluating the result expression for
    // a non-matching SWITCH case.
    let origin = CellCoord::new(0, 0);
    let expr = bytecode::parse_formula("=SWITCH(1, 2, A2, 1, 7)", origin).expect("parse");

    let grid = PanicGrid {
        // A2 relative to origin (A1) => (row=1, col=0)
        panic_coord: CellCoord::new(1, 0),
    };

    let value = bytecode::eval_ast(
        &expr,
        &grid,
        0,
        origin,
        &formula_engine::LocaleConfig::en_us(),
    );
    assert_eq!(value, Value::Number(7.0));
}

#[test]
fn bytecode_eval_ast_switch_discriminant_error_short_circuits() {
    // If the SWITCH discriminant expression is an error, SWITCH should return it without
    // evaluating any case values/results or the default branch.
    let origin = CellCoord::new(0, 0);
    let expr = bytecode::parse_formula("=SWITCH(1/0, 1, A2, A2)", origin).expect("parse");

    let grid = PanicGrid {
        // A2 relative to origin (A1) => (row=1, col=0)
        panic_coord: CellCoord::new(1, 0),
    };

    let value = bytecode::eval_ast(
        &expr,
        &grid,
        0,
        origin,
        &formula_engine::LocaleConfig::en_us(),
    );
    assert_eq!(value, Value::Error(bytecode::ErrorKind::Div0));
}

#[test]
fn bytecode_if_is_lazy_for_unused_false_branch() {
    // IF(TRUE, 7, <unused>) must not evaluate the false branch.
    let origin = CellCoord::new(0, 0);
    let expr = bytecode::parse_formula("=IF(TRUE, 7, A2)", origin).expect("parse");
    let program = bytecode::Compiler::compile(Arc::from("if_lazy_false_branch"), &expr);

    let grid = PanicGrid {
        // A2 relative to origin (A1) => (row=1, col=0)
        panic_coord: CellCoord::new(1, 0),
    };

    let mut vm = bytecode::Vm::with_capacity(32);
    let value = vm.eval(
        &program,
        &grid,
        0,
        origin,
        &formula_engine::LocaleConfig::en_us(),
    );
    assert_eq!(value, Value::Number(7.0));
}

#[test]
fn bytecode_eval_ast_if_is_lazy_for_unused_false_branch() {
    // `bytecode::eval_ast` should match VM semantics and avoid evaluating unused IF branches.
    let origin = CellCoord::new(0, 0);
    let expr = bytecode::parse_formula("=IF(TRUE, 7, A2)", origin).expect("parse");

    let grid = PanicGrid {
        // A2 relative to origin (A1) => (row=1, col=0)
        panic_coord: CellCoord::new(1, 0),
    };

    let value = bytecode::eval_ast(
        &expr,
        &grid,
        0,
        origin,
        &formula_engine::LocaleConfig::en_us(),
    );
    assert_eq!(value, Value::Number(7.0));
}

#[test]
fn bytecode_eval_ast_if_condition_error_short_circuits() {
    // If the IF condition evaluation errors, IF should return that error without evaluating either
    // branch.
    let origin = CellCoord::new(0, 0);
    let expr = bytecode::parse_formula("=IF(1/0, A2, A2)", origin).expect("parse");

    let grid = PanicGrid {
        // A2 relative to origin (A1) => (row=1, col=0)
        panic_coord: CellCoord::new(1, 0),
    };

    let value = bytecode::eval_ast(
        &expr,
        &grid,
        0,
        origin,
        &formula_engine::LocaleConfig::en_us(),
    );
    assert_eq!(value, Value::Error(bytecode::ErrorKind::Div0));
}

#[test]
fn bytecode_iferror_is_lazy_for_unused_fallback() {
    // IFERROR(<non-error>, <unused>) must not evaluate the fallback.
    let origin = CellCoord::new(0, 0);
    let expr = bytecode::parse_formula("=IFERROR(1, A2)", origin).expect("parse");
    let program = bytecode::Compiler::compile(Arc::from("iferror_lazy_fallback"), &expr);

    let grid = PanicGrid {
        // A2 relative to origin (A1) => (row=1, col=0)
        panic_coord: CellCoord::new(1, 0),
    };

    let mut vm = bytecode::Vm::with_capacity(32);
    let value = vm.eval(
        &program,
        &grid,
        0,
        origin,
        &formula_engine::LocaleConfig::en_us(),
    );
    assert_eq!(value, Value::Number(1.0));
}

#[test]
fn bytecode_eval_ast_iferror_is_lazy_for_unused_fallback() {
    // `bytecode::eval_ast` should match VM semantics and avoid evaluating unused IFERROR fallbacks.
    let origin = CellCoord::new(0, 0);
    let expr = bytecode::parse_formula("=IFERROR(1, A2)", origin).expect("parse");

    let grid = PanicGrid {
        // A2 relative to origin (A1) => (row=1, col=0)
        panic_coord: CellCoord::new(1, 0),
    };

    let value = bytecode::eval_ast(
        &expr,
        &grid,
        0,
        origin,
        &formula_engine::LocaleConfig::en_us(),
    );
    assert_eq!(value, Value::Number(1.0));
}

#[test]
fn bytecode_ifna_is_lazy_for_unused_fallback() {
    // IFNA(<non-#N/A>, <unused>) must not evaluate the fallback.
    let origin = CellCoord::new(0, 0);
    let expr = bytecode::parse_formula("=IFNA(1, A2)", origin).expect("parse");
    let program = bytecode::Compiler::compile(Arc::from("ifna_lazy_fallback"), &expr);

    let grid = PanicGrid {
        // A2 relative to origin (A1) => (row=1, col=0)
        panic_coord: CellCoord::new(1, 0),
    };

    let mut vm = bytecode::Vm::with_capacity(32);
    let value = vm.eval(
        &program,
        &grid,
        0,
        origin,
        &formula_engine::LocaleConfig::en_us(),
    );
    assert_eq!(value, Value::Number(1.0));
}

#[test]
fn bytecode_eval_ast_ifna_is_lazy_for_unused_fallback() {
    // `bytecode::eval_ast` should match VM semantics and avoid evaluating unused IFNA fallbacks.
    let origin = CellCoord::new(0, 0);
    let expr = bytecode::parse_formula("=IFNA(1, A2)", origin).expect("parse");

    let grid = PanicGrid {
        // A2 relative to origin (A1) => (row=1, col=0)
        panic_coord: CellCoord::new(1, 0),
    };

    let value = bytecode::eval_ast(
        &expr,
        &grid,
        0,
        origin,
        &formula_engine::LocaleConfig::en_us(),
    );
    assert_eq!(value, Value::Number(1.0));
}

#[test]
fn bytecode_ifna_does_not_eval_fallback_for_non_na_errors() {
    // IFNA should only treat #N/A as a recoverable error. For other error types, it should return
    // the error without evaluating the fallback.
    let origin = CellCoord::new(0, 0);
    let expr = bytecode::parse_formula("=IFNA(1/0, A2)", origin).expect("parse");
    let program = bytecode::Compiler::compile(Arc::from("ifna_non_na_error_lazy"), &expr);

    let grid = PanicGrid {
        // A2 relative to origin (A1) => (row=1, col=0)
        panic_coord: CellCoord::new(1, 0),
    };

    let mut vm = bytecode::Vm::with_capacity(32);
    let value = vm.eval(
        &program,
        &grid,
        0,
        origin,
        &formula_engine::LocaleConfig::en_us(),
    );
    assert_eq!(value, Value::Error(bytecode::ErrorKind::Div0));
}

#[test]
fn bytecode_eval_ast_ifna_does_not_eval_fallback_for_non_na_errors() {
    // Ensure `bytecode::eval_ast` matches VM semantics for IFNA when the first argument is a
    // non-#N/A error.
    let origin = CellCoord::new(0, 0);
    let expr = bytecode::parse_formula("=IFNA(1/0, A2)", origin).expect("parse");

    let grid = PanicGrid {
        // A2 relative to origin (A1) => (row=1, col=0)
        panic_coord: CellCoord::new(1, 0),
    };

    let value = bytecode::eval_ast(
        &expr,
        &grid,
        0,
        origin,
        &formula_engine::LocaleConfig::en_us(),
    );
    assert_eq!(value, Value::Error(bytecode::ErrorKind::Div0));
}

#[test]
fn bytecode_iferror_evaluates_fallback_for_errors() {
    // IFERROR should evaluate its fallback when the first argument is an error.
    let origin = CellCoord::new(0, 0);
    let expr = bytecode::parse_formula("=IFERROR(1/0, A2)", origin).expect("parse");
    let program = bytecode::Compiler::compile(Arc::from("iferror_error_fallback"), &expr);

    let grid = TextGrid {
        // A2 relative to origin (A1) => (row=1, col=0)
        coord: CellCoord::new(1, 0),
        value: Value::Number(7.0),
    };

    let mut vm = bytecode::Vm::with_capacity(32);
    let value = vm.eval(
        &program,
        &grid,
        0,
        origin,
        &formula_engine::LocaleConfig::en_us(),
    );
    assert_eq!(value, Value::Number(7.0));
}

#[test]
fn bytecode_eval_ast_iferror_evaluates_fallback_for_errors() {
    // Ensure `bytecode::eval_ast` matches VM semantics for IFERROR when the first argument errors.
    let origin = CellCoord::new(0, 0);
    let expr = bytecode::parse_formula("=IFERROR(1/0, A2)", origin).expect("parse");

    let grid = TextGrid {
        // A2 relative to origin (A1) => (row=1, col=0)
        coord: CellCoord::new(1, 0),
        value: Value::Number(7.0),
    };

    let value = bytecode::eval_ast(
        &expr,
        &grid,
        0,
        origin,
        &formula_engine::LocaleConfig::en_us(),
    );
    assert_eq!(value, Value::Number(7.0));
}

#[test]
fn bytecode_iferror_evaluates_fallback_for_na_error() {
    // IFERROR should treat #N/A like any other error and evaluate the fallback.
    let origin = CellCoord::new(0, 0);
    let expr = bytecode::parse_formula("=IFERROR(NA(), A2)", origin).expect("parse");
    let program = bytecode::Compiler::compile(Arc::from("iferror_na_fallback"), &expr);

    let grid = TextGrid {
        // A2 relative to origin (A1) => (row=1, col=0)
        coord: CellCoord::new(1, 0),
        value: Value::Number(7.0),
    };

    let mut vm = bytecode::Vm::with_capacity(32);
    let value = vm.eval(
        &program,
        &grid,
        0,
        origin,
        &formula_engine::LocaleConfig::en_us(),
    );
    assert_eq!(value, Value::Number(7.0));
}

#[test]
fn bytecode_eval_ast_iferror_evaluates_fallback_for_na_error() {
    // `bytecode::eval_ast` should match VM semantics for IFERROR and treat #N/A as an error.
    let origin = CellCoord::new(0, 0);
    let expr = bytecode::parse_formula("=IFERROR(NA(), A2)", origin).expect("parse");

    let grid = TextGrid {
        // A2 relative to origin (A1) => (row=1, col=0)
        coord: CellCoord::new(1, 0),
        value: Value::Number(7.0),
    };

    let value = bytecode::eval_ast(
        &expr,
        &grid,
        0,
        origin,
        &formula_engine::LocaleConfig::en_us(),
    );
    assert_eq!(value, Value::Number(7.0));
}

#[test]
fn bytecode_eval_ast_ifna_evaluates_fallback_for_na_error() {
    // Ensure `bytecode::eval_ast` matches VM semantics for IFNA and evaluates the fallback on #N/A.
    let origin = CellCoord::new(0, 0);
    let expr = bytecode::parse_formula("=IFNA(NA(), A2)", origin).expect("parse");

    let grid = TextGrid {
        // A2 relative to origin (A1) => (row=1, col=0)
        coord: CellCoord::new(1, 0),
        value: Value::Number(7.0),
    };

    let value = bytecode::eval_ast(
        &expr,
        &grid,
        0,
        origin,
        &formula_engine::LocaleConfig::en_us(),
    );
    assert_eq!(value, Value::Number(7.0));
}

#[test]
fn bytecode_ifna_evaluates_fallback_for_na_error() {
    // IFNA should evaluate its fallback when the first argument is #N/A.
    let origin = CellCoord::new(0, 0);
    let expr = bytecode::parse_formula("=IFNA(NA(), A2)", origin).expect("parse");
    let program = bytecode::Compiler::compile(Arc::from("ifna_na_fallback"), &expr);

    let grid = TextGrid {
        // A2 relative to origin (A1) => (row=1, col=0)
        coord: CellCoord::new(1, 0),
        value: Value::Number(7.0),
    };

    let mut vm = bytecode::Vm::with_capacity(32);
    let value = vm.eval(
        &program,
        &grid,
        0,
        origin,
        &formula_engine::LocaleConfig::en_us(),
    );
    assert_eq!(value, Value::Number(7.0));
}

#[test]
fn bytecode_switch_is_lazy_for_unused_default() {
    // SWITCH(1, 1, 7, <unused_default>) must not evaluate the default argument.
    let origin = CellCoord::new(0, 0);
    let expr = bytecode::parse_formula("=SWITCH(1, 1, 7, A2)", origin).expect("parse");
    let program = bytecode::Compiler::compile(Arc::from("switch_lazy_default"), &expr);

    let grid = PanicGrid {
        // A2 relative to origin (A1) => (row=1, col=0)
        panic_coord: CellCoord::new(1, 0),
    };

    let mut vm = bytecode::Vm::with_capacity(32);
    let value = vm.eval(
        &program,
        &grid,
        0,
        origin,
        &formula_engine::LocaleConfig::en_us(),
    );
    assert_eq!(value, Value::Number(7.0));
}

#[test]
fn bytecode_eval_ast_switch_is_lazy_for_unused_default() {
    // Ensure `bytecode::eval_ast` matches VM semantics for SWITCH default laziness.
    let origin = CellCoord::new(0, 0);
    let expr = bytecode::parse_formula("=SWITCH(1, 1, 7, A2)", origin).expect("parse");

    let grid = PanicGrid {
        // A2 relative to origin (A1) => (row=1, col=0)
        panic_coord: CellCoord::new(1, 0),
    };

    let value = bytecode::eval_ast(
        &expr,
        &grid,
        0,
        origin,
        &formula_engine::LocaleConfig::en_us(),
    );
    assert_eq!(value, Value::Number(7.0));
}

#[test]
fn bytecode_choose_does_not_eval_choices_when_index_is_error() {
    // If the index expression evaluates to an error, CHOOSE should return that error without
    // evaluating any choice expressions.
    let origin = CellCoord::new(0, 0);
    let expr = bytecode::parse_formula("=CHOOSE(1/0, A2, 7)", origin).expect("parse");
    let program = bytecode::Compiler::compile(Arc::from("choose_index_error_lazy"), &expr);

    let grid = PanicGrid {
        // A2 relative to origin (A1) => (row=1, col=0)
        panic_coord: CellCoord::new(1, 0),
    };

    let mut vm = bytecode::Vm::with_capacity(32);
    let value = vm.eval(
        &program,
        &grid,
        0,
        origin,
        &formula_engine::LocaleConfig::en_us(),
    );
    assert_eq!(value, Value::Error(bytecode::ErrorKind::Div0));
}

#[test]
fn bytecode_eval_ast_choose_does_not_eval_choices_when_index_is_error() {
    // Ensure `bytecode::eval_ast` matches VM semantics when the CHOOSE index expression errors.
    let origin = CellCoord::new(0, 0);
    let expr = bytecode::parse_formula("=CHOOSE(1/0, A2, 7)", origin).expect("parse");

    let grid = PanicGrid {
        // A2 relative to origin (A1) => (row=1, col=0)
        panic_coord: CellCoord::new(1, 0),
    };

    let value = bytecode::eval_ast(
        &expr,
        &grid,
        0,
        origin,
        &formula_engine::LocaleConfig::en_us(),
    );
    assert_eq!(value, Value::Error(bytecode::ErrorKind::Div0));
}

#[test]
fn bytecode_choose_range_index_returns_spill_without_evaluating_choices() {
    // Bytecode coercion treats ranges/arrays in scalar contexts as a spill attempt.
    // CHOOSE should surface that error without evaluating any branch expressions.
    let origin = CellCoord::new(0, 0);
    let expr = bytecode::parse_formula("=CHOOSE(A1:A2, A2, 7)", origin).expect("parse");
    let program = bytecode::Compiler::compile(Arc::from("choose_range_index_spill_lazy"), &expr);

    let grid = PanicGrid {
        // A2 relative to origin (A1) => (row=1, col=0)
        panic_coord: CellCoord::new(1, 0),
    };

    let mut vm = bytecode::Vm::with_capacity(32);
    let value = vm.eval(
        &program,
        &grid,
        0,
        origin,
        &formula_engine::LocaleConfig::en_us(),
    );
    assert_eq!(value, Value::Error(bytecode::ErrorKind::Spill));
}

#[test]
fn bytecode_choose_does_not_eval_choices_when_index_is_out_of_range() {
    // If the index is out of range, CHOOSE should return #VALUE! without evaluating any choices.
    let origin = CellCoord::new(0, 0);
    let expr = bytecode::parse_formula("=CHOOSE(3, A2, 7)", origin).expect("parse");
    let program = bytecode::Compiler::compile(Arc::from("choose_index_oob_lazy"), &expr);

    let grid = PanicGrid {
        // A2 relative to origin (A1) => (row=1, col=0)
        panic_coord: CellCoord::new(1, 0),
    };

    let mut vm = bytecode::Vm::with_capacity(32);
    let value = vm.eval(
        &program,
        &grid,
        0,
        origin,
        &formula_engine::LocaleConfig::en_us(),
    );
    assert_eq!(value, Value::Error(bytecode::ErrorKind::Value));
}

#[test]
fn bytecode_choose_does_not_eval_choices_when_index_is_out_of_range_from_cell() {
    // If the index is out of range (even when provided via a cell value), CHOOSE should return
    // #VALUE! without evaluating any choices.
    let origin = CellCoord::new(0, 0);
    let expr = bytecode::parse_formula("=CHOOSE(B1, A2, 7)", origin).expect("parse");
    let program = bytecode::Compiler::compile(Arc::from("choose_index_oob_from_cell_lazy"), &expr);

    let grid = PanicGridWithNumber {
        // B1 relative to origin (A1) => (row=0, col=1)
        number_coord: CellCoord::new(0, 1),
        number: 3.0,
        // A2 relative to origin (A1) => (row=1, col=0)
        panic_coord: CellCoord::new(1, 0),
    };

    let mut vm = bytecode::Vm::with_capacity(32);
    let value = vm.eval(
        &program,
        &grid,
        0,
        origin,
        &formula_engine::LocaleConfig::en_us(),
    );
    assert_eq!(value, Value::Error(bytecode::ErrorKind::Value));
}

#[test]
fn bytecode_eval_ast_choose_does_not_eval_choices_when_index_is_out_of_range_from_cell() {
    // Ensure `bytecode::eval_ast` matches VM semantics when the CHOOSE index is out-of-range.
    let origin = CellCoord::new(0, 0);
    let expr = bytecode::parse_formula("=CHOOSE(B1, A2, 7)", origin).expect("parse");

    let grid = PanicGridWithNumber {
        // B1 relative to origin (A1) => (row=0, col=1)
        number_coord: CellCoord::new(0, 1),
        number: 3.0,
        // A2 relative to origin (A1) => (row=1, col=0)
        panic_coord: CellCoord::new(1, 0),
    };

    let value = bytecode::eval_ast(
        &expr,
        &grid,
        0,
        origin,
        &formula_engine::LocaleConfig::en_us(),
    );
    assert_eq!(value, Value::Error(bytecode::ErrorKind::Value));
}

#[test]
fn bytecode_ifs_does_not_eval_values_for_false_conditions() {
    // IFS(FALSE, <unused_value>, TRUE, 7) must not evaluate the first value argument.
    let origin = CellCoord::new(0, 0);
    let expr = bytecode::parse_formula("=IFS(FALSE, A2, TRUE, 7)", origin).expect("parse");
    let program = bytecode::Compiler::compile(Arc::from("ifs_false_value_lazy"), &expr);

    let grid = PanicGrid {
        // A2 relative to origin (A1) => (row=1, col=0)
        panic_coord: CellCoord::new(1, 0),
    };

    let mut vm = bytecode::Vm::with_capacity(32);
    let value = vm.eval(
        &program,
        &grid,
        0,
        origin,
        &formula_engine::LocaleConfig::en_us(),
    );
    assert_eq!(value, Value::Number(7.0));
}

#[test]
fn bytecode_eval_ast_ifs_does_not_eval_values_for_false_conditions() {
    // Ensure `bytecode::eval_ast` matches VM semantics when skipping values for false IFS
    // conditions.
    let origin = CellCoord::new(0, 0);
    let expr = bytecode::parse_formula("=IFS(FALSE, A2, TRUE, 7)", origin).expect("parse");

    let grid = PanicGrid {
        // A2 relative to origin (A1) => (row=1, col=0)
        panic_coord: CellCoord::new(1, 0),
    };

    let value = bytecode::eval_ast(
        &expr,
        &grid,
        0,
        origin,
        &formula_engine::LocaleConfig::en_us(),
    );
    assert_eq!(value, Value::Number(7.0));
}

#[test]
fn bytecode_ifs_does_not_eval_values_when_no_condition_matches() {
    // IFS(FALSE, <unused>, FALSE, <unused>) must not evaluate any value arguments and should
    // return #N/A.
    let origin = CellCoord::new(0, 0);
    let expr = bytecode::parse_formula("=IFS(FALSE, A2, FALSE, A2)", origin).expect("parse");
    let program = bytecode::Compiler::compile(Arc::from("ifs_no_match_lazy"), &expr);

    let grid = PanicGrid {
        // A2 relative to origin (A1) => (row=1, col=0)
        panic_coord: CellCoord::new(1, 0),
    };

    let mut vm = bytecode::Vm::with_capacity(32);
    let value = vm.eval(
        &program,
        &grid,
        0,
        origin,
        &formula_engine::LocaleConfig::en_us(),
    );
    assert_eq!(value, Value::Error(bytecode::ErrorKind::NA));
}

#[test]
fn bytecode_eval_ast_ifs_does_not_eval_values_when_no_condition_matches() {
    // Ensure `bytecode::eval_ast` matches VM semantics for the IFS no-match path.
    let origin = CellCoord::new(0, 0);
    let expr = bytecode::parse_formula("=IFS(FALSE, A2, FALSE, A2)", origin).expect("parse");

    let grid = PanicGrid {
        // A2 relative to origin (A1) => (row=1, col=0)
        panic_coord: CellCoord::new(1, 0),
    };

    let value = bytecode::eval_ast(
        &expr,
        &grid,
        0,
        origin,
        &formula_engine::LocaleConfig::en_us(),
    );
    assert_eq!(value, Value::Error(bytecode::ErrorKind::NA));
}

#[test]
fn bytecode_switch_short_circuits_later_case_values() {
    // SWITCH(1, 1, 7, <unused_case_value>, 8) must not evaluate the later case value.
    let origin = CellCoord::new(0, 0);
    let expr = bytecode::parse_formula("=SWITCH(1, 1, 7, A2, 8)", origin).expect("parse");
    let program = bytecode::Compiler::compile(Arc::from("switch_later_case_value_lazy"), &expr);

    let grid = PanicGrid {
        // A2 relative to origin (A1) => (row=1, col=0)
        panic_coord: CellCoord::new(1, 0),
    };

    let mut vm = bytecode::Vm::with_capacity(32);
    let value = vm.eval(
        &program,
        &grid,
        0,
        origin,
        &formula_engine::LocaleConfig::en_us(),
    );
    assert_eq!(value, Value::Number(7.0));
}

#[test]
fn bytecode_eval_ast_switch_short_circuits_later_case_values() {
    // Ensure `bytecode::eval_ast` matches VM semantics and does not evaluate later case values.
    let origin = CellCoord::new(0, 0);
    let expr = bytecode::parse_formula("=SWITCH(1, 1, 7, A2, 8)", origin).expect("parse");

    let grid = PanicGrid {
        // A2 relative to origin (A1) => (row=1, col=0)
        panic_coord: CellCoord::new(1, 0),
    };

    let value = bytecode::eval_ast(
        &expr,
        &grid,
        0,
        origin,
        &formula_engine::LocaleConfig::en_us(),
    );
    assert_eq!(value, Value::Number(7.0));
}

#[test]
fn bytecode_switch_does_not_eval_results_when_no_case_matches() {
    // SWITCH(3, 1, <unused_result>, 2, <unused_result>) must not evaluate any case results when
    // no case matches and there is no default.
    let origin = CellCoord::new(0, 0);
    let expr = bytecode::parse_formula("=SWITCH(3, 1, A2, 2, A2)", origin).expect("parse");
    let program = bytecode::Compiler::compile(Arc::from("switch_no_match_lazy"), &expr);

    let grid = PanicGrid {
        // A2 relative to origin (A1) => (row=1, col=0)
        panic_coord: CellCoord::new(1, 0),
    };

    let mut vm = bytecode::Vm::with_capacity(32);
    let value = vm.eval(
        &program,
        &grid,
        0,
        origin,
        &formula_engine::LocaleConfig::en_us(),
    );
    assert_eq!(value, Value::Error(bytecode::ErrorKind::NA));
}

#[test]
fn bytecode_eval_ast_switch_does_not_eval_results_when_no_case_matches() {
    // Ensure `bytecode::eval_ast` matches VM semantics for the SWITCH no-match path.
    let origin = CellCoord::new(0, 0);
    let expr = bytecode::parse_formula("=SWITCH(3, 1, A2, 2, A2)", origin).expect("parse");

    let grid = PanicGrid {
        // A2 relative to origin (A1) => (row=1, col=0)
        panic_coord: CellCoord::new(1, 0),
    };

    let value = bytecode::eval_ast(
        &expr,
        &grid,
        0,
        origin,
        &formula_engine::LocaleConfig::en_us(),
    );
    assert_eq!(value, Value::Error(bytecode::ErrorKind::NA));
}

#[test]
fn bytecode_switch_does_not_eval_anything_after_case_value_error() {
    // If a case value evaluation errors, SWITCH should return that error and not evaluate any
    // results or later case values/results.
    let origin = CellCoord::new(0, 0);
    let expr = bytecode::parse_formula("=SWITCH(2, 1/0, A2, A2, A2)", origin).expect("parse");
    let program = bytecode::Compiler::compile(Arc::from("switch_case_value_error_lazy"), &expr);

    let grid = PanicGrid {
        // A2 relative to origin (A1) => (row=1, col=0)
        panic_coord: CellCoord::new(1, 0),
    };

    let mut vm = bytecode::Vm::with_capacity(32);
    let value = vm.eval(
        &program,
        &grid,
        0,
        origin,
        &formula_engine::LocaleConfig::en_us(),
    );
    assert_eq!(value, Value::Error(bytecode::ErrorKind::Div0));
}

#[test]
fn bytecode_eval_ast_switch_does_not_eval_anything_after_case_value_error() {
    // Ensure `bytecode::eval_ast` matches VM semantics for SWITCH when a case value errors.
    let origin = CellCoord::new(0, 0);
    let expr = bytecode::parse_formula("=SWITCH(2, 1/0, A2, A2, A2)", origin).expect("parse");

    let grid = PanicGrid {
        // A2 relative to origin (A1) => (row=1, col=0)
        panic_coord: CellCoord::new(1, 0),
    };

    let value = bytecode::eval_ast(
        &expr,
        &grid,
        0,
        origin,
        &formula_engine::LocaleConfig::en_us(),
    );
    assert_eq!(value, Value::Error(bytecode::ErrorKind::Div0));
}

#[test]
fn bytecode_switch_does_not_eval_case_values_when_discriminant_is_error() {
    // If the discriminant evaluation errors, SWITCH should return that error without evaluating
    // any case values/results (or the default).
    let origin = CellCoord::new(0, 0);
    let expr = bytecode::parse_formula("=SWITCH(1/0, A2, 7, 1, 8)", origin).expect("parse");
    let program = bytecode::Compiler::compile(Arc::from("switch_discriminant_error_lazy"), &expr);

    let grid = PanicGrid {
        // A2 relative to origin (A1) => (row=1, col=0)
        panic_coord: CellCoord::new(1, 0),
    };

    let mut vm = bytecode::Vm::with_capacity(32);
    let value = vm.eval(
        &program,
        &grid,
        0,
        origin,
        &formula_engine::LocaleConfig::en_us(),
    );
    assert_eq!(value, Value::Error(bytecode::ErrorKind::Div0));
}

#[test]
fn bytecode_ifs_does_not_eval_anything_when_condition_is_error() {
    // If a condition evaluation errors, IFS should return that error without evaluating any
    // values or later condition/value pairs.
    let origin = CellCoord::new(0, 0);
    let expr = bytecode::parse_formula("=IFS(1/0, 7, TRUE, A2)", origin).expect("parse");
    let program = bytecode::Compiler::compile(Arc::from("ifs_condition_error_lazy"), &expr);

    let grid = PanicGrid {
        // A2 relative to origin (A1) => (row=1, col=0)
        panic_coord: CellCoord::new(1, 0),
    };

    let mut vm = bytecode::Vm::with_capacity(32);
    let value = vm.eval(
        &program,
        &grid,
        0,
        origin,
        &formula_engine::LocaleConfig::en_us(),
    );
    assert_eq!(value, Value::Error(bytecode::ErrorKind::Div0));
}

#[test]
fn bytecode_if_does_not_eval_branches_when_condition_is_error() {
    // If the IF condition evaluation errors, IF should return that error without evaluating
    // either branch.
    let origin = CellCoord::new(0, 0);
    let expr = bytecode::parse_formula("=IF(1/0, A2, A2)", origin).expect("parse");
    let program = bytecode::Compiler::compile(Arc::from("if_condition_error_lazy"), &expr);

    let grid = PanicGrid {
        // A2 relative to origin (A1) => (row=1, col=0)
        panic_coord: CellCoord::new(1, 0),
    };

    let mut vm = bytecode::Vm::with_capacity(32);
    let value = vm.eval(
        &program,
        &grid,
        0,
        origin,
        &formula_engine::LocaleConfig::en_us(),
    );
    assert_eq!(value, Value::Error(bytecode::ErrorKind::Div0));
}
