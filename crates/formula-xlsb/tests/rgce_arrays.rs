use formula_engine::{parse_formula, ParseOptions};
use formula_xlsb::rgce::{decode_rgce_with_rgcb, encode_rgce_with_context, CellCoord};
use formula_xlsb::workbook_context::WorkbookContext;
use pretty_assertions::assert_eq;

#[test]
fn decodes_ptgarray_with_trailing_array_data() {
    // rgce = [PtgArray][unused7]
    let rgce = vec![0x20, 0, 0, 0, 0, 0, 0, 0];

    // rgcb = [cols_minus1: u16][rows_minus1: u16] + cells (row-major)
    let mut rgcb = Vec::new();
    rgcb.extend_from_slice(&1u16.to_le_bytes()); // 2 cols
    rgcb.extend_from_slice(&1u16.to_le_bytes()); // 2 rows

    for v in [1.0f64, 2.0, 3.0, 4.0] {
        rgcb.push(0x01); // xltypeNum
        rgcb.extend_from_slice(&v.to_le_bytes());
    }

    let text = decode_rgce_with_rgcb(&rgce, &rgcb).expect("decode rgce");
    assert_eq!(text, "{1,2;3,4}");

    parse_formula(&text, ParseOptions::default()).expect("formula-engine parses decoded array");
}

#[test]
fn encode_decode_roundtrip_array_constant() {
    let ctx = WorkbookContext::default();
    let encoded =
        encode_rgce_with_context("={1,2;3,4}", &ctx, CellCoord::new(0, 0)).expect("encode");
    assert!(!encoded.rgcb.is_empty());
    let text = decode_rgce_with_rgcb(&encoded.rgce, &encoded.rgcb).expect("decode");
    assert_eq!(text, "{1,2;3,4}");
}

#[test]
fn encode_decode_roundtrip_sum_over_array_constant() {
    let ctx = WorkbookContext::default();
    let encoded =
        encode_rgce_with_context("=SUM({1,2,3})", &ctx, CellCoord::new(0, 0)).expect("encode");
    assert!(!encoded.rgcb.is_empty());
    let text = decode_rgce_with_rgcb(&encoded.rgce, &encoded.rgcb).expect("decode");
    assert_eq!(text, "SUM({1,2,3})");

    parse_formula(&format!("={text}"), ParseOptions::default())
        .expect("formula-engine parses decoded formula");
}

#[test]
fn encode_accepts_na_bang_error_literal() {
    let ctx = WorkbookContext::default();
    let encoded = encode_rgce_with_context("=#N/A!", &ctx, CellCoord::new(0, 0)).expect("encode");
    assert!(encoded.rgcb.is_empty());
    let text = decode_rgce_with_rgcb(&encoded.rgce, &encoded.rgcb).expect("decode");
    assert_eq!(text, "#N/A");
}

#[test]
fn encode_accepts_na_bang_error_literal_in_array_constant() {
    let ctx = WorkbookContext::default();
    let encoded = encode_rgce_with_context("={#N/A!}", &ctx, CellCoord::new(0, 0)).expect("encode");
    assert!(!encoded.rgcb.is_empty());
    let text = decode_rgce_with_rgcb(&encoded.rgce, &encoded.rgcb).expect("decode");
    assert_eq!(text, "{#N/A}");
}

#[test]
fn encode_decode_roundtrip_array_constant_mixed_types_and_blanks() {
    // 2x3 array with mixed literal types and blanks.
    let ctx = WorkbookContext::default();
    let encoded =
        encode_rgce_with_context("={1,,\"hi\";TRUE,#DIV/0!,FALSE}", &ctx, CellCoord::new(0, 0))
            .expect("encode");
    assert!(!encoded.rgcb.is_empty());
    let text = decode_rgce_with_rgcb(&encoded.rgce, &encoded.rgcb).expect("decode");
    assert_eq!(text, "{1,,\"hi\";TRUE,#DIV/0!,FALSE}");

    parse_formula(&format!("={text}"), ParseOptions::default())
        .expect("formula-engine parses decoded formula");
}

#[test]
fn encode_decode_roundtrip_array_constant_string_with_quotes() {
    let ctx = WorkbookContext::default();
    let encoded =
        encode_rgce_with_context("={\"a\"\"b\"}", &ctx, CellCoord::new(0, 0)).expect("encode");
    assert!(!encoded.rgcb.is_empty());
    let text = decode_rgce_with_rgcb(&encoded.rgce, &encoded.rgcb).expect("decode");
    assert_eq!(text, "{\"a\"\"b\"}");
}

#[test]
fn encode_unary_plus_and_minus_in_array_constants() {
    // Unary `+` is valid syntax but is not preserved by BIFF encoding, which stores the literal
    // numeric value.
    let ctx = WorkbookContext::default();
    let encoded = encode_rgce_with_context("={+1,-2}", &ctx, CellCoord::new(0, 0)).expect("encode");
    assert!(!encoded.rgcb.is_empty());
    let text = decode_rgce_with_rgcb(&encoded.rgce, &encoded.rgcb).expect("decode");
    assert_eq!(text, "{1,-2}");
}

#[test]
fn encode_decode_roundtrip_multiple_array_constants_in_one_formula() {
    // BIFF12 stores array literals in `rgcb` as a sequence of blocks. Ensure we can roundtrip
    // multiple array literals by advancing the `rgcb` cursor correctly.
    let ctx = WorkbookContext::default();
    let encoded =
        encode_rgce_with_context("=SUM({1,2},{3,4})", &ctx, CellCoord::new(0, 0)).expect("encode");
    assert!(!encoded.rgcb.is_empty());
    assert!(
        encoded
            .rgce
            .iter()
            .filter(|&&b| matches!(b, 0x20 | 0x40 | 0x60))
            .count()
            >= 2,
        "expected at least two PtgArray tokens in rgce"
    );
    let text = decode_rgce_with_rgcb(&encoded.rgce, &encoded.rgcb).expect("decode");
    assert_eq!(text, "SUM({1,2},{3,4})");
}

#[test]
fn decode_ptgarray_inside_memfunc_advances_rgcb_cursor() {
    // `PtgMemFunc` contains a nested token stream that is not printed, but it can still contain
    // `PtgArray` tokens that consume `rgcb` blocks. Ensure we advance the rgcb cursor through the
    // mem payload so later visible `PtgArray` tokens decode correctly.
    let ptg_array = [0x20u8, 0, 0, 0, 0, 0, 0, 0]; // PtgArray + 7 unused bytes

    // rgce = [PtgMemFunc][cce][PtgArray (nested)][PtgArray (visible)]
    let mut rgce = vec![0x29];
    rgce.extend_from_slice(&u16::try_from(ptg_array.len()).unwrap().to_le_bytes());
    rgce.extend_from_slice(&ptg_array);
    rgce.extend_from_slice(&ptg_array);

    // rgcb contains two array constant blocks: first for the nested PtgArray, second for the
    // visible one.
    let mut rgcb = Vec::new();
    // {111}
    rgcb.extend_from_slice(&0u16.to_le_bytes()); // cols_minus1
    rgcb.extend_from_slice(&0u16.to_le_bytes()); // rows_minus1
    rgcb.push(0x01);
    rgcb.extend_from_slice(&111f64.to_le_bytes());
    // {222}
    rgcb.extend_from_slice(&0u16.to_le_bytes()); // cols_minus1
    rgcb.extend_from_slice(&0u16.to_le_bytes()); // rows_minus1
    rgcb.push(0x01);
    rgcb.extend_from_slice(&222f64.to_le_bytes());

    let text = decode_rgce_with_rgcb(&rgce, &rgcb).expect("decode");
    assert_eq!(text, "{222}");
}

#[test]
fn decode_ptgarray_inside_memfunc_with_ptgname_advances_rgcb_cursor() {
    // Like `decode_ptgarray_inside_memfunc_advances_rgcb_cursor`, but with a `PtgName` token
    // before the nested `PtgArray`. This ensures the mem-subexpression scanner skips the full
    // PtgName payload (u32 nameIndex + u16 unused) so we still consume the nested array constant.
    let ptg_array = [0x20u8, 0, 0, 0, 0, 0, 0, 0]; // PtgArray + 7 unused bytes

    // PtgName: [ptg=0x23][nameIndex: u32][unused: u16]
    let mut mem_subexpr = vec![0x23];
    mem_subexpr.extend_from_slice(&123u32.to_le_bytes());
    mem_subexpr.extend_from_slice(&0u16.to_le_bytes());
    mem_subexpr.extend_from_slice(&ptg_array);

    // rgce = [PtgMemFunc][cce][subexpr...][PtgArray (visible)]
    let mut rgce = vec![0x29];
    rgce.extend_from_slice(&u16::try_from(mem_subexpr.len()).unwrap().to_le_bytes());
    rgce.extend_from_slice(&mem_subexpr);
    rgce.extend_from_slice(&ptg_array);

    let mut rgcb = Vec::new();
    // {111}
    rgcb.extend_from_slice(&0u16.to_le_bytes()); // cols_minus1
    rgcb.extend_from_slice(&0u16.to_le_bytes()); // rows_minus1
    rgcb.push(0x01);
    rgcb.extend_from_slice(&111f64.to_le_bytes());
    // {222}
    rgcb.extend_from_slice(&0u16.to_le_bytes()); // cols_minus1
    rgcb.extend_from_slice(&0u16.to_le_bytes()); // rows_minus1
    rgcb.push(0x01);
    rgcb.extend_from_slice(&222f64.to_le_bytes());

    let text = decode_rgce_with_rgcb(&rgce, &rgcb).expect("decode");
    assert_eq!(text, "{222}");
}

#[test]
fn decode_ptgarray_inside_memfunc_with_ptgexp_advances_rgcb_cursor() {
    // Like `decode_ptgarray_inside_memfunc_advances_rgcb_cursor`, but with a `PtgExp` token
    // before the nested `PtgArray`. This ensures the mem-subexpression scanner skips the full
    // PtgExp payload (row + col) so we still consume the nested array constant.
    let ptg_array = [0x20u8, 0, 0, 0, 0, 0, 0, 0]; // PtgArray + 7 unused bytes

    // PtgExp: [ptg=0x01][row: u16][col: u16]
    let mut mem_subexpr = vec![0x01];
    mem_subexpr.extend_from_slice(&0u16.to_le_bytes());
    mem_subexpr.extend_from_slice(&0u16.to_le_bytes());
    mem_subexpr.extend_from_slice(&ptg_array);

    // rgce = [PtgMemFunc][cce][subexpr...][PtgArray (visible)]
    let mut rgce = vec![0x29];
    rgce.extend_from_slice(&u16::try_from(mem_subexpr.len()).unwrap().to_le_bytes());
    rgce.extend_from_slice(&mem_subexpr);
    rgce.extend_from_slice(&ptg_array);

    let mut rgcb = Vec::new();
    // {111}
    rgcb.extend_from_slice(&0u16.to_le_bytes()); // cols_minus1
    rgcb.extend_from_slice(&0u16.to_le_bytes()); // rows_minus1
    rgcb.push(0x01);
    rgcb.extend_from_slice(&111f64.to_le_bytes());
    // {222}
    rgcb.extend_from_slice(&0u16.to_le_bytes()); // cols_minus1
    rgcb.extend_from_slice(&0u16.to_le_bytes()); // rows_minus1
    rgcb.push(0x01);
    rgcb.extend_from_slice(&222f64.to_le_bytes());

    let text = decode_rgce_with_rgcb(&rgce, &rgcb).expect("decode");
    assert_eq!(text, "{222}");
}

#[test]
fn decode_ptgarray_inside_memfunc_with_ptgtbl_advances_rgcb_cursor() {
    // Like `decode_ptgarray_inside_memfunc_with_ptgexp_advances_rgcb_cursor`, but with a `PtgTbl`
    // token before the nested `PtgArray`.
    let ptg_array = [0x20u8, 0, 0, 0, 0, 0, 0, 0]; // PtgArray + 7 unused bytes

    // PtgTbl: [ptg=0x02][row: u16][col: u16]
    let mut mem_subexpr = vec![0x02];
    mem_subexpr.extend_from_slice(&0u16.to_le_bytes());
    mem_subexpr.extend_from_slice(&0u16.to_le_bytes());
    mem_subexpr.extend_from_slice(&ptg_array);

    // rgce = [PtgMemFunc][cce][subexpr...][PtgArray (visible)]
    let mut rgce = vec![0x29];
    rgce.extend_from_slice(&u16::try_from(mem_subexpr.len()).unwrap().to_le_bytes());
    rgce.extend_from_slice(&mem_subexpr);
    rgce.extend_from_slice(&ptg_array);

    let mut rgcb = Vec::new();
    // {111}
    rgcb.extend_from_slice(&0u16.to_le_bytes()); // cols_minus1
    rgcb.extend_from_slice(&0u16.to_le_bytes()); // rows_minus1
    rgcb.push(0x01);
    rgcb.extend_from_slice(&111f64.to_le_bytes());
    // {222}
    rgcb.extend_from_slice(&0u16.to_le_bytes()); // cols_minus1
    rgcb.extend_from_slice(&0u16.to_le_bytes()); // rows_minus1
    rgcb.push(0x01);
    rgcb.extend_from_slice(&222f64.to_le_bytes());

    let text = decode_rgce_with_rgcb(&rgce, &rgcb).expect("decode");
    assert_eq!(text, "{222}");
}

#[test]
fn decode_ptgarray_inside_memfunc_with_prefixed_ptgextend_list_advances_rgcb_cursor() {
    // Like `decode_ptgarray_inside_memfunc_advances_rgcb_cursor`, but with a prefixed structured
    // reference token (`PtgExtend` + `etpg=0x19`, a.k.a. `PtgList`) before the nested `PtgArray`.
    //
    // Some producers appear to insert 2/4 bytes of padding before the canonical 12-byte PtgList
    // payload. Ensure the mem-subexpression scanner can skip the full payload (prefix + core) so
    // it still finds and consumes the nested array constant.
    let ptg_array = [0x20u8, 0, 0, 0, 0, 0, 0, 0]; // PtgArray + 7 unused bytes

    // PtgExtend (structured ref): [ptg=0x18][etpg=0x19][payload...]
    // Include a 2-byte prefix before the 12-byte core payload.
    let mut mem_subexpr = vec![0x18, 0x19];
    mem_subexpr.extend_from_slice(&[0u8; 2]); // prefix padding
    mem_subexpr.extend_from_slice(&[
        0x01, 0x00, 0x00, 0x00, // table_id = 1
        0x00, 0x00, 0x10, 0x00, // col_first_raw = flags<<16 (flags=0x0010, col_first=0)
        0x00, 0x00, 0x00, 0x00, // col_last_raw = 0
    ]);
    mem_subexpr.extend_from_slice(&ptg_array);

    // rgce = [PtgMemFunc][cce][subexpr...][PtgArray (visible)]
    let mut rgce = vec![0x29];
    rgce.extend_from_slice(&u16::try_from(mem_subexpr.len()).unwrap().to_le_bytes());
    rgce.extend_from_slice(&mem_subexpr);
    rgce.extend_from_slice(&ptg_array);

    let mut rgcb = Vec::new();
    // {111}
    rgcb.extend_from_slice(&0u16.to_le_bytes()); // cols_minus1
    rgcb.extend_from_slice(&0u16.to_le_bytes()); // rows_minus1
    rgcb.push(0x01);
    rgcb.extend_from_slice(&111f64.to_le_bytes());
    // {222}
    rgcb.extend_from_slice(&0u16.to_le_bytes()); // cols_minus1
    rgcb.extend_from_slice(&0u16.to_le_bytes()); // rows_minus1
    rgcb.push(0x01);
    rgcb.extend_from_slice(&222f64.to_le_bytes());

    let text = decode_rgce_with_rgcb(&rgce, &rgcb).expect("decode");
    assert_eq!(text, "{222}");
}
