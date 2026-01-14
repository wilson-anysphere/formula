use subtle::{Choice, ConstantTimeEq};

#[cfg(test)]
use std::cell::Cell;

#[cfg(test)]
thread_local! {
    static CT_EQ_CALLS: Cell<usize> = Cell::new(0);
}

/// Constant-time byte slice equality.
///
/// This is used for password verifier/hash comparisons during legacy `.xls` decryption to avoid
/// timing side channels from early-exit comparisons (`==` / `!=`).
pub(crate) fn ct_eq(a: &[u8], b: &[u8]) -> bool {
    #[cfg(test)]
    CT_EQ_CALLS.with(|calls| calls.set(calls.get().saturating_add(1)));

    // Treat lengths as non-secret metadata, but still avoid early returns so callers don't
    // accidentally reintroduce short-circuit timing behavior.
    let max_len = a.len().max(b.len());
    let mut ok = Choice::from(1u8);
    for idx in 0..max_len {
        let av = a.get(idx).copied().unwrap_or(0);
        let bv = b.get(idx).copied().unwrap_or(0);
        ok &= av.ct_eq(&bv);
    }
    ok &= Choice::from((a.len() == b.len()) as u8);

    bool::from(ok)
}

#[cfg(test)]
pub(crate) fn reset_ct_eq_calls() {
    CT_EQ_CALLS.with(|calls| calls.set(0));
}

#[cfg(test)]
pub(crate) fn ct_eq_call_count() -> usize {
    CT_EQ_CALLS.with(|calls| calls.get())
}

#[cfg(test)]
mod tests {
    use super::{ct_eq, ct_eq_call_count, reset_ct_eq_calls};

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

    #[test]
    fn ct_eq_call_count_increments() {
        reset_ct_eq_calls();
        assert_eq!(ct_eq_call_count(), 0);

        ct_eq(b"abc", b"abc");
        assert_eq!(ct_eq_call_count(), 1);

        ct_eq(b"abc", b"abd");
        assert_eq!(ct_eq_call_count(), 2);
    }
}
