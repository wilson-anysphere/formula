use subtle::ConstantTimeEq;

#[cfg(test)]
use std::sync::atomic::{AtomicUsize, Ordering};

#[cfg(test)]
static CT_EQ_CALLS: AtomicUsize = AtomicUsize::new(0);

/// Compare two byte slices in constant time.
///
/// This should be used for comparing any password verifier digests (e.g.
/// `encryptedVerifierHashValue`) to avoid timing side channels.
pub(crate) fn ct_eq(a: &[u8], b: &[u8]) -> bool {
    #[cfg(test)]
    CT_EQ_CALLS.fetch_add(1, Ordering::Relaxed);
    bool::from(a.ct_eq(b))
}

#[cfg(test)]
pub(crate) fn reset_ct_eq_calls() {
    CT_EQ_CALLS.store(0, Ordering::Relaxed);
}

#[cfg(test)]
pub(crate) fn ct_eq_call_count() -> usize {
    CT_EQ_CALLS.load(Ordering::Relaxed)
}
