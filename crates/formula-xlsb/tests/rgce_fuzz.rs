use formula_xlsb::rgce::{decode_rgce, DecodeError};
use proptest::prelude::*;

proptest! {
    #![proptest_config(ProptestConfig {
        failure_persistence: None,
        .. ProptestConfig::default()
    })]

    #[test]
    fn decode_rgce_is_robust(rgce in proptest::collection::vec(any::<u8>(), 0..=256)) {
        let res = decode_rgce(&rgce);

        match res {
            Ok(s) => {
                prop_assert!(s.len() <= rgce.len().saturating_mul(10));
            }
            Err(e) => {
                prop_assert!(!rgce.is_empty());

                let offset = e.offset();
                prop_assert!(offset < rgce.len());
                prop_assert_eq!(e.ptg(), Some(rgce[offset]));
            }
        }
    }
}

#[test]
fn truncated_ptgnum_reports_offset_and_ptg() {
    // PtgNum requires 8 bytes.
    let rgce = vec![0x1F, 0, 1, 2];
    let err = decode_rgce(&rgce).unwrap_err();
    assert!(matches!(err, DecodeError::UnexpectedEof { .. }));
    assert_eq!(err.offset(), 0);
    assert_eq!(err.ptg(), Some(0x1F));
}

#[test]
fn unterminated_ptgstr_reports_offset_and_ptg() {
    // PtgStr: [ptg=0x17][cch=2][utf16 bytes...]
    // Provide only 1 UTF-16 char (2 bytes) instead of 2 (4 bytes).
    let rgce = vec![0x17, 0x02, 0x00, b'A', 0x00];
    let err = decode_rgce(&rgce).unwrap_err();
    assert!(matches!(err, DecodeError::UnexpectedEof { .. }));
    assert_eq!(err.offset(), 0);
    assert_eq!(err.ptg(), Some(0x17));
}

#[test]
fn operator_stack_underflow_is_error() {
    // Binary '+' with no operands.
    let rgce = vec![0x03];
    let err = decode_rgce(&rgce).unwrap_err();
    assert!(matches!(err, DecodeError::StackUnderflow { .. }));
    assert_eq!(err.offset(), 0);
    assert_eq!(err.ptg(), Some(0x03));
}

#[test]
fn trailing_bytes_after_complete_expression_is_error() {
    // Two ints in a row: "1 2" (RPN) leaves stack depth 2 at end.
    let rgce = vec![0x1E, 1, 0, 0x1E, 2, 0];
    let err = decode_rgce(&rgce).unwrap_err();
    assert!(matches!(err, DecodeError::StackNotSingular { stack_len: 2, .. }));
    assert_eq!(err.ptg(), Some(0x1E));
}

#[test]
fn ptgstr_escapes_quotes_excel_style() {
    // PtgStr payload stores raw characters. Excel formula text escapes inner quotes by doubling them.
    // This token stores the raw string value `A"B`.
    let rgce = vec![0x17, 0x03, 0x00, b'A', 0x00, b'"', 0x00, b'B', 0x00];
    let decoded = decode_rgce(&rgce).expect("decode string literal");
    assert_eq!(decoded, "\"A\"\"B\"");
}
