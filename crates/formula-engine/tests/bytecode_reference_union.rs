#![cfg(not(target_arch = "wasm32"))]

use formula_engine::bytecode::ast::{BinaryOp, Expr, Function};
use formula_engine::bytecode::{self, CellCoord, ColumnarGrid, Value, Vm};
use std::sync::Arc;

#[test]
fn bytecode_parser_supports_reference_union_inside_parentheses() {
    let origin = CellCoord::new(0, 0);
    let expr = bytecode::parse_formula("=SUM((A1,B1))", origin).expect("parse");

    match &expr {
        Expr::FuncCall { func, args } => {
            assert_eq!(func, &Function::Sum);
            assert_eq!(args.len(), 1);
            match &args[0] {
                Expr::Binary { op, .. } => assert_eq!(*op, BinaryOp::Union),
                other => panic!("expected union binary op, got {other:?}"),
            }
        }
        other => panic!("expected SUM func call, got {other:?}"),
    }

    let program = bytecode::Compiler::compile(Arc::from("sum_union_parenthesized"), &expr);
    let locale = formula_engine::LocaleConfig::en_us();
    let mut grid = ColumnarGrid::new(1, 2);
    grid.set_number(CellCoord::new(0, 0), 1.0);
    grid.set_number(CellCoord::new(0, 1), 2.0);

    let mut vm = Vm::with_capacity(32);
    let vm_value = vm.eval(&program, &grid, 0, origin, &locale);
    let ast_value = bytecode::eval_ast(&expr, &grid, 0, origin, &locale);

    assert_eq!(vm_value, Value::Number(3.0));
    assert_eq!(ast_value, Value::Number(3.0));
}

#[test]
fn bytecode_reference_union_can_be_used_as_one_argument_among_many() {
    let origin = CellCoord::new(0, 0);
    let expr = bytecode::parse_formula("=SUM((A1,B1),1)", origin).expect("parse");

    match &expr {
        Expr::FuncCall { func, args } => {
            assert_eq!(func, &Function::Sum);
            assert_eq!(args.len(), 2);
            assert!(matches!(
                args[0],
                Expr::Binary {
                    op: BinaryOp::Union,
                    ..
                }
            ));
            assert_eq!(args[1], Expr::Literal(Value::Number(1.0)));
        }
        other => panic!("expected SUM func call, got {other:?}"),
    }

    let program = bytecode::Compiler::compile(Arc::from("sum_union_plus_arg"), &expr);
    let locale = formula_engine::LocaleConfig::en_us();
    let mut grid = ColumnarGrid::new(1, 2);
    grid.set_number(CellCoord::new(0, 0), 1.0);
    grid.set_number(CellCoord::new(0, 1), 2.0);

    let mut vm = Vm::with_capacity(32);
    let vm_value = vm.eval(&program, &grid, 0, origin, &locale);
    let ast_value = bytecode::eval_ast(&expr, &grid, 0, origin, &locale);

    assert_eq!(vm_value, Value::Number(4.0));
    assert_eq!(ast_value, Value::Number(4.0));
}

#[test]
fn bytecode_reference_union_supports_ranges() {
    let origin = CellCoord::new(0, 0);
    let expr = bytecode::parse_formula("=SUM((A1:A2,B1:B2))", origin).expect("parse");
    let program = bytecode::Compiler::compile(Arc::from("sum_union_ranges"), &expr);
    let locale = formula_engine::LocaleConfig::en_us();
    let mut grid = ColumnarGrid::new(2, 2);
    grid.set_number(CellCoord::new(0, 0), 1.0);
    grid.set_number(CellCoord::new(1, 0), 2.0);
    grid.set_number(CellCoord::new(0, 1), 3.0);
    grid.set_number(CellCoord::new(1, 1), 4.0);

    let mut vm = Vm::with_capacity(32);
    let vm_value = vm.eval(&program, &grid, 0, origin, &locale);
    let ast_value = bytecode::eval_ast(&expr, &grid, 0, origin, &locale);

    assert_eq!(vm_value, Value::Number(10.0));
    assert_eq!(ast_value, Value::Number(10.0));
}
