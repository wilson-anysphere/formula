use formula_engine::{
    parse_formula, BinaryExpr, BinaryOp, CellAddr, Coord, Expr, ParseOptions, SerializeOptions,
};

#[test]
fn parses_column_ranges() {
    let ast = parse_formula("=A:C", ParseOptions::default()).unwrap();
    let Expr::Binary(BinaryExpr { op, left, right }) = ast.expr else {
        panic!("expected binary expr");
    };
    assert_eq!(op, BinaryOp::Range);
    assert!(matches!(
        left.as_ref(),
        Expr::ColRef(r) if r.col == Coord::A1 { index: 0, abs: false }
    ));
    assert!(matches!(
        right.as_ref(),
        Expr::ColRef(r) if r.col == Coord::A1 { index: 2, abs: false }
    ));
}

#[test]
fn parses_row_ranges_from_number_literals() {
    let ast = parse_formula("=1:3", ParseOptions::default()).unwrap();
    let Expr::Binary(BinaryExpr { op, left, right }) = ast.expr else {
        panic!("expected binary expr");
    };
    assert_eq!(op, BinaryOp::Range);
    assert!(matches!(
        left.as_ref(),
        Expr::RowRef(r) if r.row == Coord::A1 { index: 0, abs: false }
    ));
    assert!(matches!(
        right.as_ref(),
        Expr::RowRef(r) if r.row == Coord::A1 { index: 2, abs: false }
    ));
}

#[test]
fn normalizes_column_ranges_for_shared_formulas() {
    let mut opts = ParseOptions::default();
    opts.normalize_relative_to = Some(CellAddr::new(0, 2)); // origin C1
    let ast = parse_formula("=A:A", opts).unwrap();

    let mut ser = SerializeOptions::default();
    ser.origin = Some(CellAddr::new(0, 3)); // render from D1

    assert_eq!(ast.to_string(ser).unwrap(), "=B:B");
}

#[test]
fn normalizes_row_ranges_for_shared_formulas() {
    let mut opts = ParseOptions::default();
    opts.normalize_relative_to = Some(CellAddr::new(4, 0)); // origin A5
    let ast = parse_formula("=1:1", opts).unwrap();

    let mut ser = SerializeOptions::default();
    ser.origin = Some(CellAddr::new(5, 0)); // render from A6

    assert_eq!(ast.to_string(ser).unwrap(), "=2:2");
}
