use formula_biff::{decode_rgce, DecodeRgceError};

#[test]
fn decode_rgce_rejects_legacy_structured_ref_placeholder_tokens() {
    for ptg in [0x30u8, 0x50u8, 0x70u8] {
        let err = decode_rgce(&[ptg]).expect_err("expected unsupported token");
        assert!(
            matches!(err, DecodeRgceError::UnsupportedToken { offset: 0, ptg: got } if got == ptg),
            "expected UnsupportedToken {{ offset: 0, ptg: 0x{ptg:02x} }}, got {err:?}"
        );
    }
}
