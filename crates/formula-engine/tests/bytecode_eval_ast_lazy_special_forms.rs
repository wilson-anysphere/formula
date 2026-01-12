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

