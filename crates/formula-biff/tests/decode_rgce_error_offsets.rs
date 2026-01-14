use formula_biff::{decode_rgce, decode_rgce_with_rgcb, DecodeRgceError};

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
    // PtgStr(cch=1, unit=0xD800) is invalid (unpaired surrogate).
    let rgce = [0x1E, 0x01, 0x00, 0x17, 0x01, 0x00, 0x00, 0xD8];
    let err = decode_rgce(&rgce).expect_err("expected invalid utf16");
    assert!(
        matches!(
            err,
            DecodeRgceError::InvalidUtf16 {
                offset: 3,
                ptg: 0x17
            }
        ),
        "expected InvalidUtf16 at offset 3 for ptg=0x17, got {err:?}"
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
    let cases = [
        decode_rgce(&[0xFF]).unwrap_err(),                      // UnsupportedToken
        decode_rgce(&[0x17]).unwrap_err(),                      // UnexpectedEof
        decode_rgce(&[0x03]).unwrap_err(),                      // StackUnderflow
        decode_rgce(&[0x22, 0x00, 0xFF, 0xFF]).unwrap_err(),    // UnknownFunctionId
        decode_rgce(&[0x17, 0x01, 0x00, 0x00, 0xD8]).unwrap_err(), // InvalidUtf16
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
