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
