use formula_biff::{decode_rgce, DecodeRgceError};

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

