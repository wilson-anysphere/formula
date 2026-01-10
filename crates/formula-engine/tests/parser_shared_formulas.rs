use formula_engine::{
    parse_formula, BinaryExpr, BinaryOp, CellAddr, Coord, Expr, ParseOptions, SerializeOptions,
};

#[test]
fn normalizes_relative_cell_refs_to_offsets() {
    let mut opts = ParseOptions::default();
    opts.normalize_relative_to = Some(CellAddr::new(0, 2)); // origin C1

    let ast = parse_formula("=A1+B1", opts).unwrap();

    let Expr::Binary(BinaryExpr { op, left, right }) = ast.expr else {
        panic!("expected binary expr");
    };
    assert_eq!(op, BinaryOp::Add);

    let Expr::CellRef(left) = left.as_ref() else {
        panic!("expected left cell ref");
    };
    let Expr::CellRef(right) = right.as_ref() else {
        panic!("expected right cell ref");
    };

    assert_eq!(left.col, Coord::Offset(-2));
    assert_eq!(left.row, Coord::Offset(0));
    assert_eq!(right.col, Coord::Offset(-1));
    assert_eq!(right.row, Coord::Offset(0));
}

#[test]
fn renders_shared_formula_for_new_origin() {
    let mut opts = ParseOptions::default();
    opts.normalize_relative_to = Some(CellAddr::new(0, 2)); // normalized from C1
    let ast = parse_formula("=A1+B1", opts).unwrap();

    let mut ser = SerializeOptions::default();
    ser.origin = Some(CellAddr::new(1, 2)); // render for C2

    assert_eq!(ast.to_string(ser).unwrap(), "=A2+B2");
}
