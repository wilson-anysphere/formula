use formula_biff::decode_rgce;
use formula_engine::parse_formula;
use pretty_assertions::assert_eq;

fn assert_parses_and_roundtrips(src: &str) {
    let ast = parse_formula(src, Default::default()).expect("formula should parse");
    let back = ast
        .to_string(formula_engine::SerializeOptions {
            omit_equals: true,
            ..Default::default()
        })
        .expect("serialize");
    assert_eq!(back, src);
}

#[test]
fn decodes_optimized_sum_using_tattrsum() {
    // Excel can encode `SUM(A1:A3)` in optimized form as:
    //   PtgArea(A1:A3) + PtgAttr(tAttrSum)
    //
    // This stream intentionally omits any explicit `PtgFuncVar(SUM)` token.
    let rgce = [
        0x25, // PtgArea
        0x00, 0x00, 0x00, 0x00, // rowFirst = 0 (A1)
        0x02, 0x00, 0x00, 0x00, // rowLast  = 2 (A3)
        0x00, 0xC0, // colFirst = A, relative row/col
        0x00, 0xC0, // colLast  = A, relative row/col
        0x19, // PtgAttr
        0x10, // tAttrSum
        0x00, 0x00, // wAttr (unused for tAttrSum)
    ];

    let text = decode_rgce(&rgce).expect("decode");
    assert_eq!(text, "SUM(A1:A3)");
    assert_parses_and_roundtrips(&text);
}

#[test]
fn consumes_tattrchoose_jump_table_bytes_to_keep_offsets_aligned() {
    // `tAttrChoose` is followed by a `u16` jump table (wAttr entries).
    // The jump offsets are evaluator metadata, but we must still consume them so subsequent tokens
    // stay aligned.
    let rgce = [
        0x1E, 0x01, 0x00, // PtgInt(1)
        0x19, 0x04, 0x01, 0x00, // PtgAttr(tAttrChoose, wAttr=1)
        0xFF, 0xFF, // jump table entry (ignored)
        0x1E, 0x02, 0x00, // PtgInt(2)
        0x03, // PtgAdd
    ];

    let text = decode_rgce(&rgce).expect("decode");
    assert_eq!(text, "1+2");
    assert_parses_and_roundtrips(&text);
}

