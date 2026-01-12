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
