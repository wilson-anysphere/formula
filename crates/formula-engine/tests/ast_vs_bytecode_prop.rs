#![cfg(not(target_arch = "wasm32"))]

use formula_engine::bytecode::{
    ast::{BinaryOp, Function, UnaryOp},
    eval_ast, parse_formula, BytecodeCache, CellCoord, ColumnarGrid, Expr, RangeRef, Ref, Value,
    Vm,
};
use formula_engine::LocaleConfig;
use proptest::prelude::*;
use std::sync::Arc;

fn arb_cell_coord(rows: i32, cols: i32) -> impl Strategy<Value = CellCoord> {
    (0..rows, 0..cols).prop_map(|(r, c)| CellCoord::new(r, c))
}

fn arb_ref(base: CellCoord, rows: i32, cols: i32) -> impl Strategy<Value = Ref> {
    arb_cell_coord(rows, cols).prop_map(move |target| {
        Ref::new(target.row - base.row, target.col - base.col, false, false)
    })
}

fn arb_range_ref(base: CellCoord, rows: i32, cols: i32) -> impl Strategy<Value = RangeRef> {
    (arb_cell_coord(rows, cols), arb_cell_coord(rows, cols)).prop_map(move |(a, b)| {
        let ra = Ref::new(a.row - base.row, a.col - base.col, false, false);
        let rb = Ref::new(b.row - base.row, b.col - base.col, false, false);
        RangeRef::new(ra, rb)
    })
}

fn arb_rect_range_ref(base: CellCoord, rows: i32, cols: i32) -> impl Strategy<Value = RangeRef> {
    // Small rectangles keep the dependency expansion in graph building reasonable.
    (1i32..=3, 1i32..=3).prop_flat_map(move |(h, w)| {
        let max_r = (rows - h).max(0);
        let max_c = (cols - w).max(0);
        (0..=max_r, 0..=max_c).prop_map(move |(r0, c0)| {
            let a = CellCoord::new(r0, c0);
            let b = CellCoord::new(r0 + h - 1, c0 + w - 1);
            let ra = Ref::new(a.row - base.row, a.col - base.col, false, false);
            let rb = Ref::new(b.row - base.row, b.col - base.col, false, false);
            RangeRef::new(ra, rb)
        })
    })
}

fn arb_sumproduct_ranges(
    base: CellCoord,
    rows: i32,
    cols: i32,
) -> impl Strategy<Value = (RangeRef, RangeRef)> {
    (1i32..=3, 1i32..=3).prop_flat_map(move |(h, w)| {
        let max_r = (rows - h).max(0);
        let max_c = (cols - w).max(0);
        (0..=max_r, 0..=max_c, 0..=max_r, 0..=max_c).prop_map(move |(r0a, c0a, r0b, c0b)| {
            let a0 = CellCoord::new(r0a, c0a);
            let a1 = CellCoord::new(r0a + h - 1, c0a + w - 1);
            let b0 = CellCoord::new(r0b, c0b);
            let b1 = CellCoord::new(r0b + h - 1, c0b + w - 1);

            let ra0 = Ref::new(a0.row - base.row, a0.col - base.col, false, false);
            let ra1 = Ref::new(a1.row - base.row, a1.col - base.col, false, false);
            let rb0 = Ref::new(b0.row - base.row, b0.col - base.col, false, false);
            let rb1 = Ref::new(b1.row - base.row, b1.col - base.col, false, false);

            (RangeRef::new(ra0, ra1), RangeRef::new(rb0, rb1))
        })
    })
}

fn arb_literal() -> impl Strategy<Value = Expr> {
    prop_oneof![
        (-1000i32..=1000).prop_map(|v| Expr::Literal(Value::Number(v as f64))),
        any::<bool>().prop_map(|b| Expr::Literal(Value::Bool(b))),
        Just(Expr::FuncCall {
            func: Function::Na,
            args: vec![],
        }),
    ]
}

fn arb_expr(base: CellCoord, rows: i32, cols: i32) -> impl Strategy<Value = Expr> {
    let leaf = prop_oneof![
        arb_literal(),
        arb_ref(base, rows, cols).prop_map(Expr::CellRef),
        arb_range_ref(base, rows, cols).prop_map(Expr::RangeRef),
    ];

    leaf.prop_recursive(
        4,  // depth
        32, // size
        4,  // items per collection
        move |inner| {
            prop_oneof![
                inner.clone().prop_map(|e| Expr::Unary {
                    op: UnaryOp::Neg,
                    expr: Box::new(e),
                }),
                (inner.clone(), inner.clone()).prop_map(|(l, r)| Expr::Binary {
                    op: BinaryOp::Add,
                    left: Box::new(l),
                    right: Box::new(r),
                }),
                (inner.clone(), inner.clone()).prop_map(|(l, r)| Expr::Binary {
                    op: BinaryOp::Mul,
                    left: Box::new(l),
                    right: Box::new(r),
                }),
                // IF(cond, t, f)
                (inner.clone(), inner.clone(), inner.clone()).prop_map(|(c, t, f)| Expr::FuncCall {
                    func: Function::If,
                    args: vec![c, t, f],
                }),
                // IF(cond, t)
                (inner.clone(), inner.clone()).prop_map(|(c, t)| Expr::FuncCall {
                    func: Function::If,
                    args: vec![c, t],
                }),
                // AND(a, b)
                (inner.clone(), inner.clone()).prop_map(|(a, b)| Expr::FuncCall {
                    func: Function::And,
                    args: vec![a, b],
                }),
                // OR(a, b)
                (inner.clone(), inner.clone()).prop_map(|(a, b)| Expr::FuncCall {
                    func: Function::Or,
                    args: vec![a, b],
                }),
                // IFERROR(a, b)
                (inner.clone(), inner.clone()).prop_map(|(a, b)| Expr::FuncCall {
                    func: Function::IfError,
                    args: vec![a, b],
                }),
                // IFNA(a, b)
                (inner.clone(), inner.clone()).prop_map(|(a, b)| Expr::FuncCall {
                    func: Function::IfNa,
                    args: vec![a, b],
                }),
                // ISERROR(a)
                inner.clone().prop_map(|a| Expr::FuncCall {
                    func: Function::IsError,
                    args: vec![a],
                }),
                // ISNA(a)
                inner.clone().prop_map(|a| Expr::FuncCall {
                    func: Function::IsNa,
                    args: vec![a],
                }),
                // SUM(range)
                arb_rect_range_ref(base, rows, cols).prop_map(|r| Expr::FuncCall {
                    func: Function::Sum,
                    args: vec![Expr::RangeRef(r)],
                }),
                // AVERAGE(range)
                arb_rect_range_ref(base, rows, cols).prop_map(|r| Expr::FuncCall {
                    func: Function::Average,
                    args: vec![Expr::RangeRef(r)],
                }),
                // MIN(range)
                arb_rect_range_ref(base, rows, cols).prop_map(|r| Expr::FuncCall {
                    func: Function::Min,
                    args: vec![Expr::RangeRef(r)],
                }),
                // MAX(range)
                arb_rect_range_ref(base, rows, cols).prop_map(|r| Expr::FuncCall {
                    func: Function::Max,
                    args: vec![Expr::RangeRef(r)],
                }),
                // COUNT(range)
                arb_rect_range_ref(base, rows, cols).prop_map(|r| Expr::FuncCall {
                    func: Function::Count,
                    args: vec![Expr::RangeRef(r)],
                }),
                // COUNTIF(range, number)
                (arb_rect_range_ref(base, rows, cols), -10i32..=10).prop_map(|(r, n)| {
                    Expr::FuncCall {
                        func: Function::CountIf,
                        args: vec![Expr::RangeRef(r), Expr::Literal(Value::Number(n as f64))],
                    }
                }),
                // SUMPRODUCT(range_a, range_b)
                arb_sumproduct_ranges(base, rows, cols).prop_map(|(a, b)| Expr::FuncCall {
                    func: Function::SumProduct,
                    args: vec![Expr::RangeRef(a), Expr::RangeRef(b)],
                }),
            ]
        },
    )
}

proptest! {
    #[test]
    fn prop_ast_matches_bytecode(
        expr in arb_expr(CellCoord::new(5, 5), 10, 10),
        cells in prop::collection::vec(prop_oneof![Just(None), (-1000i32..=1000).prop_map(Some)], 100),
    ) {
        let base = CellCoord::new(5, 5);

        let mut grid = ColumnarGrid::new(10, 10);
        for (idx, cell) in cells.into_iter().enumerate() {
            if let Some(v) = cell {
                let row = (idx / 10) as i32;
                let col = (idx % 10) as i32;
                grid.set_number(CellCoord::new(row, col), v as f64);
            }
        }

        let cache = BytecodeCache::new();
        let program = cache.get_or_compile(&expr);
        let mut vm = Vm::with_capacity(32);
        let locale = LocaleConfig::en_us();
        let ast_val = eval_ast(&expr, &grid, base, &locale);
        let bc_val = vm.eval(&program, &grid, base, &locale);

        prop_assert_eq!(ast_val, bc_val);
    }
}

#[test]
fn cache_shares_filled_formula_patterns() {
    let cache = BytecodeCache::new();

    // C1: =A1+B1
    let expr_c1 = parse_formula("=A1+B1", CellCoord::new(0, 2)).unwrap();
    // C2 after fill-down: =A2+B2
    let expr_c2 = parse_formula("=A2+B2", CellCoord::new(1, 2)).unwrap();

    let p1 = cache.get_or_compile(&expr_c1);
    let p2 = cache.get_or_compile(&expr_c2);

    assert_eq!(p1.key(), p2.key());
    assert!(Arc::ptr_eq(&p1, &p2));
}
