use formula_engine::{
    parse_formula, BinaryExpr, BinaryOp, CellAddr, Coord, ErrorKind, Expr, ParseOptions,
    ReferenceStyle, SerializeOptions,
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

#[test]
fn normalizes_relative_refs_without_i32_overflow_when_origin_exceeds_i32_max() {
    let origin = CellAddr::new(i32::MAX as u32 + 1, 0);

    let mut opts = ParseOptions::default();
    opts.normalize_relative_to = Some(origin);
    let ast = parse_formula("=A1", opts).unwrap();

    let Expr::CellRef(cell_ref) = &ast.expr else {
        panic!("expected cell ref");
    };
    assert_eq!(cell_ref.col, Coord::Offset(0));
    assert_eq!(cell_ref.row, Coord::Offset(i32::MIN));

    // Round-trip back to A1 should not overflow when adding the large origin to the offset.
    let mut ser = SerializeOptions::default();
    ser.origin = Some(origin);
    assert_eq!(ast.to_string(ser).unwrap(), "=A1");
}

#[test]
fn renders_offsets_to_a1_without_overflow_when_origin_exceeds_i32_max() {
    let origin = CellAddr::new(i32::MAX as u32 + 1, 0);

    let mut opts = ParseOptions::default();
    opts.reference_style = ReferenceStyle::R1C1;
    let ast = parse_formula("=RC", opts).unwrap();

    let mut ser = SerializeOptions::default();
    ser.reference_style = ReferenceStyle::A1;
    ser.origin = Some(origin);

    assert_eq!(ast.to_string(ser).unwrap(), "=A2147483649");
}

#[test]
fn normalize_relative_falls_back_when_offset_out_of_i32_range() {
    let mut opts = ParseOptions::default();
    opts.normalize_relative_to = Some(CellAddr::new(0, 0));

    // Row index is `i32::MAX + 1`, which cannot be represented as an i32 offset from row 0.
    let ast = parse_formula("=A2147483649", opts).unwrap();

    let Expr::CellRef(cell_ref) = &ast.expr else {
        panic!("expected cell ref");
    };
    assert_eq!(cell_ref.col, Coord::Offset(0));
    assert_eq!(
        cell_ref.row,
        Coord::A1 {
            index: i32::MAX as u32 + 1,
            abs: false
        }
    );

    // R1C1 rendering should fall back to an absolute row instead of overflowing/wrapping.
    let mut ser = SerializeOptions::default();
    ser.reference_style = ReferenceStyle::R1C1;
    ser.origin = Some(CellAddr::new(0, 0));
    assert_eq!(ast.to_string(ser).unwrap(), "=R2147483649C");
}

#[test]
fn eval_compiler_caps_rows_to_i32_max_minus_one() {
    let ast_ok = parse_formula("=A2147483647", ParseOptions::default()).unwrap();
    let lowered_ok = formula_engine::eval::lower_ast(&ast_ok, None);
    assert!(matches!(lowered_ok, formula_engine::eval::Expr::CellRef(_)));

    // Row 2147483648 (1-indexed) is out of range for the eval IR's i32 encoding.
    let ast_bad = parse_formula("=A2147483648", ParseOptions::default()).unwrap();
    let lowered_bad = formula_engine::eval::lower_ast(&ast_bad, None);
    assert_eq!(
        lowered_bad,
        formula_engine::eval::Expr::Error(ErrorKind::Ref)
    );
}
