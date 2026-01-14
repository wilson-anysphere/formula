#![cfg(not(target_arch = "wasm32"))]

use formula_engine::bytecode::ast::{BinaryOp, Expr, Function};
use formula_engine::bytecode::{self, CellCoord, ColumnarGrid, Value, Vm};
use std::sync::Arc;

#[test]
fn bytecode_parser_supports_reference_intersection_in_parentheses() {
    let origin = CellCoord::new(0, 0);
    let expr = bytecode::parse_formula("=SUM((A1:B2 B1:C2))", origin).expect("parse");

    match &expr {
        Expr::FuncCall { func, args } => {
            assert_eq!(func, &Function::Sum);
            assert_eq!(args.len(), 1);
            match &args[0] {
                Expr::Binary { op, .. } => assert_eq!(*op, BinaryOp::Intersect),
                other => panic!("expected intersection binary op, got {other:?}"),
            }
        }
        other => panic!("expected SUM func call, got {other:?}"),
    }

    let program = bytecode::Compiler::compile(Arc::from("sum_intersect_parenthesized"), &expr);
    let locale = formula_engine::LocaleConfig::en_us();

    // A1:C2 = 1..=6 (row-major).
    // Intersect(A1:B2, B1:C2) = B1:B2 => 2 + 5 = 7.
    let mut grid = ColumnarGrid::new(2, 3);
    grid.set_number(CellCoord::new(0, 0), 1.0);
    grid.set_number(CellCoord::new(0, 1), 2.0);
    grid.set_number(CellCoord::new(0, 2), 3.0);
    grid.set_number(CellCoord::new(1, 0), 4.0);
    grid.set_number(CellCoord::new(1, 1), 5.0);
    grid.set_number(CellCoord::new(1, 2), 6.0);

    let mut vm = Vm::with_capacity(32);
    let vm_value = vm.eval(&program, &grid, 0, origin, &locale);
    let ast_value = bytecode::eval_ast(&expr, &grid, 0, origin, &locale);

    assert_eq!(vm_value, Value::Number(7.0));
    assert_eq!(ast_value, Value::Number(7.0));
}

#[test]
fn bytecode_reference_intersection_empty_yields_null_error() {
    let origin = CellCoord::new(0, 0);
    let expr = bytecode::parse_formula("=(A1:A2 C1:C2)", origin).expect("parse");
    let program = bytecode::Compiler::compile(Arc::from("empty_intersection"), &expr);
    let locale = formula_engine::LocaleConfig::en_us();

    let grid = ColumnarGrid::new(2, 3);

    let mut vm = Vm::with_capacity(32);
    let vm_value = vm.eval(&program, &grid, 0, origin, &locale);
    let ast_value = bytecode::eval_ast(&expr, &grid, 0, origin, &locale);

    assert_eq!(vm_value, Value::Error(bytecode::ErrorKind::Null));
    assert_eq!(ast_value, Value::Error(bytecode::ErrorKind::Null));
}

#[test]
fn bytecode_reference_intersection_binds_tighter_than_union() {
    let origin = CellCoord::new(0, 0);
    let expr = bytecode::parse_formula("=(A1:B2 B1:C2, C1:C2)", origin).expect("parse");

    match &expr {
        Expr::Binary { op, left, right } => {
            assert_eq!(*op, BinaryOp::Union);
            assert!(matches!(
                left.as_ref(),
                Expr::Binary {
                    op: BinaryOp::Intersect,
                    ..
                }
            ));
            assert!(matches!(right.as_ref(), Expr::RangeRef(_)));
        }
        other => panic!("expected union binary expression, got {other:?}"),
    }
}
