use formula_biff::{decode_rgce, decode_rgce_with_base, decode_rgce_with_rgcb, DecodeRgceError};

#[test]
fn decode_rgce_reports_offset_for_truncated_ptgstr() {
    // PtgStr requires at least a 2-byte cch field.
    let err = decode_rgce(&[0x17]).expect_err("expected truncated PtgStr to fail");
    assert!(
        matches!(
            err,
            DecodeRgceError::UnexpectedEof {
                offset: 0,
                ptg: 0x17,
                needed: 2,
                remaining: 0
            }
        ),
        "expected UnexpectedEof at offset 0 for ptg=0x17, got {err:?}"
    );
}

#[test]
fn decode_rgce_reports_offset_for_truncated_token_after_prefix() {
    // Prefix the failing token with a 3-byte PtgInt so the error offset is non-zero.
    // PtgRef requires 6 bytes of payload after the ptg.
    let rgce = [0x1E, 0x01, 0x00, 0x24];
    let err = decode_rgce(&rgce).expect_err("expected truncated PtgRef to fail");
    assert!(
        matches!(
            err,
            DecodeRgceError::UnexpectedEof {
                offset: 3,
                ptg: 0x24,
                needed: 6,
                remaining: 0
            }
        ),
        "expected UnexpectedEof at offset 3 for ptg=0x24, got {err:?}"
    );
}

#[test]
fn decode_rgce_reports_offset_for_stack_underflow() {
    // PtgAdd requires two stack operands; with an empty stack this should be a stack underflow at
    // offset 0.
    let err = decode_rgce(&[0x03]).expect_err("expected stack underflow");
    assert!(
        matches!(
            err,
            DecodeRgceError::StackUnderflow {
                offset: 0,
                ptg: 0x03
            }
        ),
        "expected StackUnderflow at offset 0 for ptg=0x03, got {err:?}"
    );
}

#[test]
fn decode_rgce_reports_offset_for_stack_not_singular() {
    // Two consecutive PtgInt tokens without an operator will leave 2 items on the stack.
    let rgce = [0x1E, 0x01, 0x00, 0x1E, 0x02, 0x00];
    let err = decode_rgce(&rgce).expect_err("expected StackNotSingular");
    assert!(
        matches!(
            err,
            DecodeRgceError::StackNotSingular {
                offset: 3,
                ptg: 0x1E,
                stack_len: 2
            }
        ),
        "expected StackNotSingular at offset 3 for ptg=0x1E, got {err:?}"
    );
}

#[test]
fn decode_rgce_reports_offset_for_truncated_ptgarray_header() {
    // PtgArray requires 7 unused bytes in rgce; even with an rgcb stream present, a truncated
    // token should report EOF at the ptg offset.
    let err = decode_rgce_with_rgcb(&[0x20], &[]).expect_err("expected truncated PtgArray");
    assert!(
        matches!(
            err,
            DecodeRgceError::UnexpectedEof {
                offset: 0,
                ptg: 0x20,
                needed: 7,
                remaining: 0
            }
        ),
        "expected UnexpectedEof at offset 0 for ptg=0x20, got {err:?}"
    );
}

#[test]
fn decode_rgce_reports_offset_for_unknown_function_id() {
    // PtgFuncVar with an unknown function id should report its offset/ptg.
    let err = decode_rgce(&[0x22, 0x00, 0xFF, 0xFF]).expect_err("expected unknown function id");
    assert!(
        matches!(
            err,
            DecodeRgceError::UnknownFunctionId {
                offset: 0,
                ptg: 0x22,
                func_id: 0xFFFF
            }
        ),
        "expected UnknownFunctionId at offset 0 for ptg=0x22, got {err:?}"
    );
}

#[test]
fn decode_rgce_reports_offset_for_invalid_utf16() {
    // Prefix to make the failing token offset non-zero.
    //
    // Note: `PtgStr` decodes UTF-16 lossily (best-effort), so we trigger InvalidUtf16 via a
    // string within an array constant instead.
    let rgce = [
        0x1E, 0x01, 0x00, // PtgInt(1) prefix
        0x20, // PtgArray
        0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, // unused 7 bytes
    ];
    // 1x1 array containing a string with an unpaired surrogate (0xD800).
    let rgcb = [0x00, 0x00, 0x00, 0x00, 0x02, 0x01, 0x00, 0x00, 0xD8];
    let err = decode_rgce_with_rgcb(&rgce, &rgcb).expect_err("expected invalid utf16");
    assert!(
        matches!(
            err,
            DecodeRgceError::InvalidUtf16 {
                offset: 3,
                ptg: 0x20
            }
        ),
        "expected InvalidUtf16 at offset 3 for ptg=0x20, got {err:?}"
    );
}

#[test]
fn decode_rgce_reports_offset_for_unsupported_token_after_prefix() {
    // Prefix the failing token with a 3-byte PtgInt so the unsupported token offset is non-zero.
    let rgce = [0x1E, 0x01, 0x00, 0xFF];
    let err = decode_rgce(&rgce).expect_err("expected unsupported token");
    assert!(
        matches!(
            err,
            DecodeRgceError::UnsupportedToken {
                offset: 3,
                ptg: 0xFF
            }
        ),
        "expected UnsupportedToken at offset 3 for ptg=0xFF, got {err:?}"
    );
}

#[test]
fn decode_rgce_error_messages_include_ptg_and_offset() {
    // Keep this test intentionally broad: we just want to ensure the Display strings always
    // include enough context for diagnostics tooling (ptg + rgce offset).
    // 1x1 array containing a string with an unpaired surrogate (0xD800).
    let rgcb_invalid_utf16 = [0x00, 0x00, 0x00, 0x00, 0x02, 0x01, 0x00, 0x00, 0xD8];
    let cases = [
        decode_rgce(&[0xFF]).unwrap_err(),                      // UnsupportedToken
        decode_rgce(&[0x17]).unwrap_err(),                      // UnexpectedEof
        decode_rgce(&[0x03]).unwrap_err(),                      // StackUnderflow
        decode_rgce(&[0x22, 0x00, 0xFF, 0xFF]).unwrap_err(),    // UnknownFunctionId
        decode_rgce_with_rgcb(
            &[0x20, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00],
            &rgcb_invalid_utf16,
        )
        .unwrap_err(), // InvalidUtf16
        decode_rgce(&[0x1E, 0x01, 0x00, 0x1E, 0x02, 0x00]).unwrap_err(), // StackNotSingular
    ];

    for err in cases {
        let msg = err.to_string();
        assert!(
            msg.contains("ptg=0x") && msg.contains("offset"),
            "expected error message to contain ptg and offset, got: {msg}"
        );
    }
}

#[test]
fn decode_rgce_reports_offset_for_truncated_ptgattr_header() {
    // PtgAttr requires [grbit: u8][wAttr: u16] after the ptg byte.
    let err = decode_rgce(&[0x19]).expect_err("expected truncated PtgAttr");
    assert!(
        matches!(
            err,
            DecodeRgceError::UnexpectedEof {
                offset: 0,
                ptg: 0x19,
                needed: 3,
                remaining: 0
            }
        ),
        "expected UnexpectedEof at offset 0 for ptg=0x19, got {err:?}"
    );
}

#[test]
fn decode_rgce_reports_offset_for_truncated_tattrchoose_jump_table() {
    // PtgAttr(tAttrChoose, wAttr=2) requires 4 jump-table bytes after the 3-byte attr header.
    // Provide only 2.
    let rgce = [
        0x1E, 0x01, 0x00, // PtgInt(1) (prefix)
        0x19, 0x04, 0x02, 0x00, // PtgAttr(tAttrChoose, wAttr=2)
        0xFF, 0xFF, // truncated jump table (needs 4 bytes)
    ];
    let err = decode_rgce(&rgce).expect_err("expected truncated tAttrChoose jump table");
    assert!(
        matches!(
            err,
            DecodeRgceError::UnexpectedEof {
                offset: 3,
                ptg: 0x19,
                needed: 4,
                remaining: 2
            }
        ),
        "expected UnexpectedEof at offset 3 for ptg=0x19, got {err:?}"
    );
}

#[test]
fn decode_rgce_reports_offset_for_tattrsum_stack_underflow() {
    // PtgAttr(tAttrSum) rewrites the previous arg as SUM(arg). With an empty stack, this is a
    // stack underflow.
    let rgce = [0x19, 0x10, 0x00, 0x00];
    let err = decode_rgce(&rgce).expect_err("expected stack underflow");
    assert!(
        matches!(
            err,
            DecodeRgceError::StackUnderflow {
                offset: 0,
                ptg: 0x19
            }
        ),
        "expected StackUnderflow at offset 0 for ptg=0x19, got {err:?}"
    );
}

#[test]
fn decode_rgce_reports_offset_for_truncated_ptgrefn_payload() {
    // PtgRefN requires 6 bytes of payload after the ptg. Truncated payload should be EOF even
    // when using the base-aware decoder.
    let err = decode_rgce_with_base(&[0x2C], 0, 0).expect_err("expected truncated PtgRefN");
    assert!(
        matches!(
            err,
            DecodeRgceError::UnexpectedEof {
                offset: 0,
                ptg: 0x2C,
                needed: 6,
                remaining: 0
            }
        ),
        "expected UnexpectedEof at offset 0 for ptg=0x2C, got {err:?}"
    );
}
