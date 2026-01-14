use subtle::{Choice, ConstantTimeEq};

#[cfg(test)]
use std::cell::Cell;

// Unit tests run in parallel by default. Use a thread-local counter so tests that reset/inspect
// the counter don't race each other.
#[cfg(test)]
thread_local! {
    // Keep test instrumentation thread-local to avoid cross-test interference when the test runner
    // executes unit tests in parallel.
    static CT_EQ_CALLS: Cell<usize> = Cell::new(0);
}

/// Compare two byte slices in constant time.
///
/// This should be used for comparing any password verifier digests (e.g.
/// `encryptedVerifierHashValue`) to avoid timing side channels.
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
