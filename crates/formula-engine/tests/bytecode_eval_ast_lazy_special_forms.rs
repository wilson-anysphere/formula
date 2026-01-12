#![cfg(not(target_arch = "wasm32"))]

use formula_engine::bytecode::{self, CellCoord, Value};

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

#[derive(Clone)]
struct CellValueGrid {
    coord: CellCoord,
    value: Value,
}

impl bytecode::Grid for CellValueGrid {
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

#[derive(Clone)]
struct ValueAndPanicGrid {
    value_coord: CellCoord,
    value: Value,
    panic_coord: CellCoord,
}

impl bytecode::Grid for ValueAndPanicGrid {
    fn get_value(&self, coord: CellCoord) -> Value {
        if coord == self.panic_coord {
            panic!("attempted to evaluate forbidden cell reference at {coord:?}");
        }
        if coord == self.value_coord {
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
fn bytecode_eval_ast_choose_does_not_eval_choices_when_index_is_error() {
    // If the index expression is an error, CHOOSE should return it without evaluating any choice
    // expressions.
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
fn bytecode_eval_ast_choose_does_not_eval_choices_when_index_is_out_of_range() {
    // If the index is out of range, CHOOSE should return #VALUE! without evaluating any choices.
    let origin = CellCoord::new(0, 0);
    let expr = bytecode::parse_formula("=CHOOSE(3, A2, 7)", origin).expect("parse");

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
    assert_eq!(value, Value::Error(bytecode::ErrorKind::Value));
}

#[test]
fn bytecode_eval_ast_choose_does_not_eval_choices_when_index_is_out_of_range_from_cell() {
    // Same as `bytecode_eval_ast_choose_does_not_eval_choices_when_index_is_out_of_range`, but with
    // the out-of-range index provided via a cell reference.
    let origin = CellCoord::new(0, 0);
    let expr = bytecode::parse_formula("=CHOOSE(B1, A2, 7)", origin).expect("parse");

    let grid = ValueAndPanicGrid {
        // B1 relative to origin (A1) => (row=0, col=1)
        value_coord: CellCoord::new(0, 1),
        value: Value::Number(3.0),
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
fn bytecode_eval_ast_choose_range_index_returns_spill_without_evaluating_choices() {
    // Bytecode coercion treats ranges/arrays in scalar contexts as a spill attempt.
    // CHOOSE should surface that error without evaluating any branch expressions.
    let origin = CellCoord::new(0, 0);
    let expr = bytecode::parse_formula("=CHOOSE(A1:A2, A2, 7)", origin).expect("parse");

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
    assert_eq!(value, Value::Error(bytecode::ErrorKind::Spill));
}

#[test]
fn bytecode_eval_ast_ifs_does_not_eval_values_for_false_conditions() {
    // IFS(FALSE, <unused_value>, TRUE, 7) must not evaluate the first value argument.
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
fn bytecode_eval_ast_ifs_does_not_eval_values_when_no_condition_matches() {
    // IFS(FALSE, <unused>, FALSE, <unused>) must not evaluate any value arguments and should
    // return #N/A.
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
fn bytecode_eval_ast_ifs_does_not_eval_anything_when_condition_is_error() {
    // If a condition evaluation errors, IFS should return that error without evaluating any
    // values or later condition/value pairs.
    let origin = CellCoord::new(0, 0);
    let expr = bytecode::parse_formula("=IFS(1/0, 7, TRUE, A2)", origin).expect("parse");

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
fn bytecode_eval_ast_if_is_lazy_for_unused_false_branch() {
    // IF(TRUE, 7, <unused>) must not evaluate the false branch.
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
    // If the IF condition is an error, IF should return it without evaluating any branch
    // expressions.
    let origin = CellCoord::new(0, 0);
    let expr = bytecode::parse_formula("=IF(1/0, 7, A2)", origin).expect("parse");

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
fn bytecode_eval_ast_iferror_is_lazy_for_unused_fallback() {
    // IFERROR(<non-error>, <unused>) must not evaluate the fallback.
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
fn bytecode_eval_ast_iferror_evaluates_fallback_for_errors() {
    // IFERROR should evaluate its fallback when the first argument is an error.
    let origin = CellCoord::new(0, 0);
    let expr = bytecode::parse_formula("=IFERROR(1/0, A2)", origin).expect("parse");

    let grid = CellValueGrid {
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
fn bytecode_eval_ast_ifna_is_lazy_for_unused_fallback() {
    // IFNA(<non-#N/A>, <unused>) must not evaluate the fallback.
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
fn bytecode_eval_ast_ifna_does_not_eval_fallback_for_non_na_errors() {
    // IFNA should only treat #N/A as a recoverable error. For other error types, it should return
    // the error without evaluating the fallback.
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
fn bytecode_eval_ast_ifna_evaluates_fallback_for_na_error() {
    // IFNA should evaluate its fallback when the first argument is #N/A.
    let origin = CellCoord::new(0, 0);
    let expr = bytecode::parse_formula("=IFNA(NA(), A2)", origin).expect("parse");

    let grid = CellValueGrid {
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
fn bytecode_eval_ast_switch_is_lazy_for_unused_default() {
    // SWITCH(1, 1, 7, <unused_default>) must not evaluate the default argument.
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
fn bytecode_eval_ast_switch_short_circuits_later_case_values() {
    // SWITCH(1, 1, 7, <unused_case_value>, 8) must not evaluate the later case value.
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
fn bytecode_eval_ast_switch_does_not_eval_results_when_no_case_matches() {
    // SWITCH(3, 1, <unused_result>, 2, <unused_result>) must not evaluate any case results when
    // no case matches and there is no default.
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
fn bytecode_eval_ast_switch_does_not_eval_anything_after_case_value_error() {
    // If a case value evaluation errors, SWITCH should return that error and not evaluate any
    // results or later case values/results.
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
fn bytecode_eval_ast_switch_does_not_eval_case_values_when_discriminant_is_error() {
    // If the discriminant evaluation errors, SWITCH should return that error without evaluating
    // any case values/results (or the default).
    let origin = CellCoord::new(0, 0);
    let expr = bytecode::parse_formula("=SWITCH(1/0, A2, 7, 1, 8)", origin).expect("parse");

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
fn bytecode_eval_ast_switch_evaluates_default_when_no_case_matches() {
    // When no case matches and a default is provided, SWITCH should evaluate and return the
    // default expression.
    let origin = CellCoord::new(0, 0);
    let expr = bytecode::parse_formula("=SWITCH(3, 1, 10, 2, 20, A2)", origin).expect("parse");

    let grid = CellValueGrid {
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
