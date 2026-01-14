use formula_engine::parse_formula;
use formula_xlsb::rgce::{
    decode_formula_rgce, decode_formula_rgce_with_rgcb, decode_rgce, decode_rgce_with_base,
    decode_rgce_with_context, CellCoord, DecodeError, DecodeFailureKind, DecodeWarning,
};
use formula_xlsb::workbook_context::WorkbookContext;
use pretty_assertions::assert_eq;

fn rgce_area(ptg: u8) -> Vec<u8> {
    // A1:A10 in BIFF12 encoding:
    // - rows are 0-indexed u32
    // - cols are 0-indexed u14 in a u16 where:
    //   - bit 14 (0x4000): row relative
    //   - bit 15 (0x8000): col relative
    let mut out = vec![ptg];
    out.extend_from_slice(&0u32.to_le_bytes()); // rowFirst = 0 (A1)
    out.extend_from_slice(&9u32.to_le_bytes()); // rowLast  = 9 (A10)
    out.extend_from_slice(&0xC000u16.to_le_bytes()); // colFirst = A, relative row/col
    out.extend_from_slice(&0xC000u16.to_le_bytes()); // colLast  = A, relative row/col
    out
}

fn rgce_ref(ptg: u8) -> Vec<u8> {
    // A1 as a PtgRef* token: [ptg][row: u32][col: u16]
    let mut out = vec![ptg];
    out.extend_from_slice(&0u32.to_le_bytes()); // row = 0 (A1)
    out.extend_from_slice(&0xC000u16.to_le_bytes()); // col = A, relative row/col
    out
}

fn rgce_area_n(ptg: u8) -> Vec<u8> {
    // A1:A10 as a PtgAreaN* token, relative to the base cell A1:
    // [ptg][r1_off: i32][r2_off: i32][c1_off: i16][c2_off: i16]
    let mut out = vec![ptg];
    out.extend_from_slice(&0i32.to_le_bytes()); // rowFirst offset
    out.extend_from_slice(&9i32.to_le_bytes()); // rowLast offset
    out.extend_from_slice(&0i16.to_le_bytes()); // colFirst offset
    out.extend_from_slice(&0i16.to_le_bytes()); // colLast offset
    out
}

fn assert_parses_and_roundtrips(src: &str) {
    let ast = parse_formula(src, Default::default()).expect("formula should parse");
    let back = ast.to_string(Default::default()).expect("serialize");
    assert_eq!(back, src);
}

#[test]
fn decodes_ptg_areav_as_explicit_implicit_intersection() {
    // PtgAreaV (value class) should render as `@` to preserve legacy implicit intersection.
    let rgce = rgce_area(0x45);
    let text = decode_rgce(&rgce).expect("decode");
    assert_eq!(text, "@A1:A10");
    assert_parses_and_roundtrips(&text);
}

#[test]
fn decodes_ptg_area_ref_class_without_at() {
    // PtgArea (ref class) should not render `@`.
    let rgce = rgce_area(0x25);
    let text = decode_rgce(&rgce).expect("decode");
    assert_eq!(text, "A1:A10");
    assert_parses_and_roundtrips(&text);
}

#[test]
fn decodes_ptg_area3dv_with_sheet_prefix_and_at() {
    // PtgArea3dV: [ptg][ixti: u16][area...]
    let mut rgce = vec![0x5B];
    rgce.extend_from_slice(&1u16.to_le_bytes()); // Sheet2 (by index in our decode context)
    rgce.extend_from_slice(&0u32.to_le_bytes());
    rgce.extend_from_slice(&9u32.to_le_bytes());
    rgce.extend_from_slice(&0xC000u16.to_le_bytes());
    rgce.extend_from_slice(&0xC000u16.to_le_bytes());

    let mut ctx = WorkbookContext::default();
    ctx.add_extern_sheet("Sheet2", "Sheet2", 1);

    let text = decode_rgce_with_context(&rgce, &ctx).expect("decode");

    assert_eq!(text, "@Sheet2!A1:A10");
    assert_parses_and_roundtrips(&text);
}

#[test]
fn decodes_ptg_ref3d_with_external_workbook_prefix() {
    // PtgRef3d: [ptg][ixti: u16][row: u32][col: u16]
    let mut rgce = vec![0x3A];
    rgce.extend_from_slice(&0u16.to_le_bytes()); // extern sheet index
    rgce.extend_from_slice(&0u32.to_le_bytes()); // row = 0 (A1)
    rgce.extend_from_slice(&0xC000u16.to_le_bytes()); // col = A, relative row/col

    let mut ctx = WorkbookContext::default();
    ctx.add_extern_sheet_external_workbook("Book2.xlsb", "SheetA", "SheetB", 0);

    let text = decode_rgce_with_context(&rgce, &ctx).expect("decode");
    assert_eq!(text, "'[Book2.xlsb]SheetA:SheetB'!A1");
    parse_formula(&text, Default::default()).expect("formula should parse");
}

#[test]
fn does_not_emit_at_for_single_cell_ptg_refv() {
    let rgce = rgce_ref(0x44);
    let text = decode_rgce(&rgce).expect("decode");
    assert_eq!(text, "A1");
    assert_parses_and_roundtrips(&text);
}

#[test]
fn decodes_ptg_areanv_as_explicit_implicit_intersection() {
    // PtgAreaNV (value class) should render as `@` when it denotes a multi-cell range.
    let rgce = rgce_area_n(0x4D);
    let text = decode_rgce_with_base(&rgce, CellCoord::new(0, 0)).expect("decode");
    assert_eq!(text, "@A1:A10");
    assert_parses_and_roundtrips(&text);
}

#[test]
fn decodes_ptg_namev_as_explicit_implicit_intersection() {
    // PtgNameV (value class) should render with `@` to preserve legacy implicit intersection.
    let mut ctx = WorkbookContext::default();
    ctx.add_workbook_name("MyNamedRange", 1);

    // PtgNameV: [ptg][nameId: u32][reserved: u16]
    let mut rgce = vec![0x43];
    rgce.extend_from_slice(&1u32.to_le_bytes());
    rgce.extend_from_slice(&0u16.to_le_bytes());

    let text = decode_rgce_with_context(&rgce, &ctx).expect("decode");
    assert_eq!(text, "@MyNamedRange");
    assert_parses_and_roundtrips(&text);
}

#[test]
fn decodes_ptgerr_known_code() {
    let rgce = [0x1C, 0x07]; // PtgErr #DIV/0!
    let decoded = decode_formula_rgce(&rgce);
    assert_eq!(decoded.text.as_deref(), Some("#DIV/0!"));
    assert!(decoded.warnings.is_empty());
}

#[test]
fn decodes_ptgerr_modern_codes() {
    for (code, lit) in [
        (0x2C, "#SPILL!"),
        (0x2D, "#CALC!"),
        (0x2E, "#FIELD!"),
        (0x2F, "#CONNECT!"),
        (0x30, "#BLOCKED!"),
        (0x31, "#UNKNOWN!"),
    ] {
        let rgce = [0x1C, code]; // PtgErr
        let decoded = decode_formula_rgce(&rgce);
        assert_eq!(decoded.text.as_deref(), Some(lit), "code={code:#04x}");
        assert!(decoded.warnings.is_empty(), "code={code:#04x}");
    }
}

#[test]
fn decodes_ptgerr_unknown_code_without_aborting() {
    let rgce = [0x1C, 0xFF]; // PtgErr unknown/extended code
    let decoded = decode_formula_rgce(&rgce);
    assert_eq!(decoded.text.as_deref(), Some("#UNKNOWN!"));
    assert_eq!(
        decoded.warnings,
        vec![DecodeWarning::UnknownErrorCode {
            code: 0xFF,
            offset: 1
        }]
    );
}

#[test]
fn renders_unknown_ptgfuncvar_ids_as_parseable_function_calls_in_best_effort_decoder() {
    // Unknown `PtgFuncVar` iftab => emit a stable placeholder name rather than failing the entire
    // formula decode.
    let rgce = [0x22, 0x00, 0xFF, 0xFF]; // PtgFuncVar(argc=0, iftab=0xFFFF)
    let decoded = decode_formula_rgce(&rgce);
    assert_eq!(decoded.text.as_deref(), Some("_UNKNOWN_FUNC_0XFFFF()"));
    assert_eq!(
        decoded.warnings,
        vec![DecodeWarning::DecodeFailed {
            kind: DecodeFailureKind::UnknownPtg,
            offset: 0,
            ptg: 0x22
        }]
    );
    assert_parses_and_roundtrips(decoded.text.as_ref().expect("text"));

    // Strict decode APIs should still surface the unknown function id as an error.
    let err = decode_rgce(&rgce).expect_err("expected error");
    assert!(matches!(err, DecodeError::UnknownPtg { offset: 0, ptg: 0x22 }));
}

#[test]
fn renders_unknown_ptgfunc_ids_as_parseable_function_calls_best_effort_unary() {
    // Best-effort: unknown `PtgFunc` does not store argc, so the decoder assumes unary if
    // possible.
    //
    // _UNKNOWN_FUNC_0XFFFF(1):
    //   PtgInt(1)
    //   PtgFunc(iftab=0xFFFF)
    let rgce = [0x1E, 0x01, 0x00, 0x21, 0xFF, 0xFF];
    let decoded = decode_formula_rgce(&rgce);
    assert_eq!(decoded.text.as_deref(), Some("_UNKNOWN_FUNC_0XFFFF(1)"));
    assert_eq!(
        decoded.warnings,
        vec![DecodeWarning::DecodeFailed {
            kind: DecodeFailureKind::UnknownPtg,
            offset: 3,
            ptg: 0x21
        }]
    );
    assert_parses_and_roundtrips(decoded.text.as_ref().expect("text"));

    let err = decode_rgce(&rgce).expect_err("expected error");
    assert!(matches!(err, DecodeError::UnknownPtg { offset: 3, ptg: 0x21 }));
}

#[test]
fn renders_unknown_ptgfunc_ids_with_empty_stack_as_parseable_function_calls() {
    // Unknown `PtgFunc` with no arguments on the stack => `_UNKNOWN_FUNC_0XFFFF()`.
    let rgce = [0x21, 0xFF, 0xFF];
    let decoded = decode_formula_rgce(&rgce);
    assert_eq!(decoded.text.as_deref(), Some("_UNKNOWN_FUNC_0XFFFF()"));
    assert_eq!(
        decoded.warnings,
        vec![DecodeWarning::DecodeFailed {
            kind: DecodeFailureKind::UnknownPtg,
            offset: 0,
            ptg: 0x21
        }]
    );
    assert_parses_and_roundtrips(decoded.text.as_ref().expect("text"));

    let err = decode_rgce(&rgce).expect_err("expected error");
    assert!(matches!(err, DecodeError::UnknownPtg { offset: 0, ptg: 0x21 }));
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
fn ignores_tattrif_and_tattrskip_without_breaking_offsets() {
    // `tAttrIf` / `tAttrSkip` are control-flow metadata used by Excel's evaluator
    // for short-circuiting. They should not break best-effort formula printing.
    let rgce = [
        0x1E, 0x01, 0x00, // PtgInt(1)
        0x19, 0x02, 0x00, 0x00, // PtgAttr(tAttrIf, wAttr=0)
        0x1E, 0x02, 0x00, // PtgInt(2)
        0x19, 0x08, 0x00, 0x00, // PtgAttr(tAttrSkip, wAttr=0)
        0x03, // PtgAdd
    ];

    let text = decode_rgce(&rgce).expect("decode");
    assert_eq!(text, "1+2");
    assert_parses_and_roundtrips(&text);
}

#[test]
fn quotes_sheet_names_in_ptgref3d() {
    let mut ctx = WorkbookContext::default();
    ctx.add_extern_sheet("My Sheet", "My Sheet", 0);
    ctx.add_extern_sheet("A1", "A1", 1);

    let mut rgce = vec![0x3A]; // PtgRef3d
    rgce.extend_from_slice(&0u16.to_le_bytes()); // ixti=0 => "My Sheet"
    rgce.extend_from_slice(&0u32.to_le_bytes()); // row=0
    rgce.extend_from_slice(&0xC000u16.to_le_bytes()); // col=A, relative
    let text = decode_rgce_with_context(&rgce, &ctx).expect("decode");
    assert_eq!(text, "'My Sheet'!A1");
    assert_parses_and_roundtrips(&text);

    let mut rgce = vec![0x3A]; // PtgRef3d
    rgce.extend_from_slice(&1u16.to_le_bytes()); // ixti=1 => "A1"
    rgce.extend_from_slice(&0u32.to_le_bytes()); // row=0
    rgce.extend_from_slice(&0xC000u16.to_le_bytes()); // col=A, relative
    let text = decode_rgce_with_context(&rgce, &ctx).expect("decode");
    assert_eq!(text, "'A1'!A1");
    assert_parses_and_roundtrips(&text);
}

#[test]
fn wraps_union_operator_inside_function_args() {
    // Canonical Excel text disambiguates union inside args with parentheses.
    // This rgce stream intentionally omits `PtgParen` so the decoder must add them.
    //
    // Formula: IF(1,(A1,B1))
    let mut rgce = vec![0x1E, 0x01, 0x00]; // PtgInt(1)
    rgce.extend_from_slice(&rgce_ref(0x24)); // A1

    let mut b1 = vec![0x24]; // PtgRef(B1)
    b1.extend_from_slice(&0u32.to_le_bytes());
    b1.extend_from_slice(&0xC001u16.to_le_bytes());
    rgce.extend_from_slice(&b1);

    rgce.push(0x10); // PtgUnion
    rgce.extend_from_slice(&[0x22, 0x02, 0x01, 0x00]); // PtgFuncVar(argc=2, IF)

    let text = decode_rgce(&rgce).expect("decode");
    assert_eq!(text, "IF(1,(A1,B1))");
    assert_parses_and_roundtrips(&text);
}

#[test]
fn skips_ptgmem_tokens() {
    // Mem tokens are evaluation hints and should not affect printed formula text.
    let rgce = [
        0x1E, 0x01, 0x00, // PtgInt(1)
        0x29, 0x00, 0x00, // PtgMemFunc(cce=0)
        0x1E, 0x02, 0x00, // PtgInt(2)
        0x03, // PtgAdd
    ];

    let text = decode_rgce(&rgce).expect("decode");
    assert_eq!(text, "1+2");
    assert_parses_and_roundtrips(&text);
}

#[test]
fn decodes_unknown_array_error_codes_without_aborting() {
    let rgce = [0x20, 0, 0, 0, 0, 0, 0, 0]; // PtgArray + 7 unused bytes
    let rgcb = [
        0x00, 0x00, // cols_minus1
        0x00, 0x00, // rows_minus1
        0x10, // xltypeErr
        0xFF, // unknown error code
    ];

    let decoded = decode_formula_rgce_with_rgcb(&rgce, &rgcb);
    assert_eq!(decoded.text.as_deref(), Some("{#UNKNOWN!}"));
    assert_eq!(
        decoded.warnings,
        vec![DecodeWarning::UnknownArrayErrorCode {
            code: 0xFF,
            offset: 5
        }]
    );
    assert_parses_and_roundtrips(decoded.text.as_ref().expect("text"));
}

#[test]
fn surfaces_stack_underflow_as_structured_warning() {
    // A single binary operator token should fail with a stack underflow.
    let rgce = [0x03]; // PtgAdd
    let decoded = decode_formula_rgce(&rgce);
    assert_eq!(decoded.text, None);
    assert_eq!(
        decoded.warnings,
        vec![DecodeWarning::DecodeFailed {
            kind: DecodeFailureKind::StackUnderflow,
            offset: 0,
            ptg: 0x03
        }]
    );
}

#[test]
fn surfaces_unexpected_eof_as_structured_warning() {
    // Truncated PtgInt: expects 2 trailing bytes for the i16 value.
    let rgce = [0x1E]; // PtgInt (missing payload)
    let decoded = decode_formula_rgce(&rgce);
    assert_eq!(decoded.text, None);
    assert_eq!(
        decoded.warnings,
        vec![DecodeWarning::DecodeFailed {
            kind: DecodeFailureKind::UnexpectedEof,
            offset: 0,
            ptg: 0x1E
        }]
    );
}

#[test]
fn best_effort_returns_top_expression_when_stack_not_singular() {
    // Two ints in a row: "1 2" (RPN) leaves stack depth 2 at end.
    let rgce = [0x1E, 0x01, 0x00, 0x1E, 0x02, 0x00];
    let decoded = decode_formula_rgce(&rgce);
    assert_eq!(decoded.text.as_deref(), Some("2"));
    assert_eq!(
        decoded.warnings,
        vec![DecodeWarning::DecodeFailed {
            kind: DecodeFailureKind::StackNotSingular,
            offset: 3,
            ptg: 0x1E
        }]
    );

    let err = decode_rgce(&rgce).expect_err("expected error");
    assert!(matches!(
        err,
        DecodeError::StackNotSingular {
            offset: 3,
            ptg: 0x1E,
            stack_len: 2
        }
    ));
}
