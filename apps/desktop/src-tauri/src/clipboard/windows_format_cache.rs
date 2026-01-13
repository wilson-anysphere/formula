use std::sync::OnceLock;

/// Cache for a registered Windows clipboard format ID.
///
/// Registering custom clipboard formats requires calling into Win32 via
/// `RegisterClipboardFormatW`, which can allocate and is relatively expensive on hot paths.
/// This helper memoizes the resulting format ID so reads/writes can reuse it.
///
/// The cached value is `Option<u32>` so callers can preserve best-effort semantics:
/// - `Some(id)` when registration succeeds.
/// - `None` when registration fails (the format is treated as unavailable).
pub(crate) struct CachedClipboardFormat {
    name: &'static str,
    id: OnceLock<Option<u32>>,
}

impl CachedClipboardFormat {
    pub(crate) const fn new(name: &'static str) -> Self {
        Self {
            name,
            id: OnceLock::new(),
        }
    }

    pub(crate) fn get_with(
        &self,
        register: impl FnOnce(&'static str) -> Option<u32>,
    ) -> Option<u32> {
        *self.id.get_or_init(|| register(self.name))
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::sync::atomic::{AtomicUsize, Ordering};

    #[test]
    fn get_with_only_runs_registration_once() {
        let calls = AtomicUsize::new(0);
        let fmt = CachedClipboardFormat::new("test-format");

        let id1 = fmt.get_with(|_| {
            calls.fetch_add(1, Ordering::SeqCst);
            Some(42)
        });
        let id2 = fmt.get_with(|_| {
            calls.fetch_add(1, Ordering::SeqCst);
            Some(43)
        });

        assert_eq!(id1, Some(42));
        assert_eq!(id2, Some(42));
        assert_eq!(calls.load(Ordering::SeqCst), 1);
    }

    #[test]
    fn get_with_caches_failures() {
        let calls = AtomicUsize::new(0);
        let fmt = CachedClipboardFormat::new("test-format-failure");

        let id1 = fmt.get_with(|_| {
            calls.fetch_add(1, Ordering::SeqCst);
            None
        });
        let id2 = fmt.get_with(|_| {
            calls.fetch_add(1, Ordering::SeqCst);
            Some(99)
        });

        assert_eq!(id1, None);
        assert_eq!(id2, None);
        assert_eq!(calls.load(Ordering::SeqCst), 1);
    }
}
