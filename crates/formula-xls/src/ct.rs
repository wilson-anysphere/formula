use subtle::ConstantTimeEq;

/// Constant-time byte slice equality.
///
/// This is used for password verifier/hash comparisons during legacy `.xls` decryption to avoid
/// timing side channels from early-exit comparisons (`==` / `!=`).
pub(crate) fn ct_eq(a: &[u8], b: &[u8]) -> bool {
    bool::from(a.ct_eq(b))
}

#[cfg(test)]
mod tests {
    use super::ct_eq;

    #[test]
    fn ct_eq_true_for_equal_slices() {
        assert!(ct_eq(b"", b""));
        assert!(ct_eq(b"abc", b"abc"));
        assert!(ct_eq(&[0u8, 1, 2, 3], &[0u8, 1, 2, 3]));
    }

    #[test]
    fn ct_eq_false_for_mismatched_slices() {
        assert!(!ct_eq(b"abc", b"xbc"));
        assert!(!ct_eq(b"abc", b"axc"));
        assert!(!ct_eq(b"abc", b"abx"));
    }

    #[test]
    fn ct_eq_false_for_different_lengths() {
        assert!(!ct_eq(b"a", b""));
        assert!(!ct_eq(b"ab", b"abc"));
    }
}

