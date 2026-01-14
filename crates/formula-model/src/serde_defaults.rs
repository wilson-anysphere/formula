/// Serde helper for `#[serde(default = "crate::serde_defaults::default_true")]`.
///
/// Prefer using the fully-qualified path in serde attributes to avoid having to import this symbol
/// into individual modules (which can lead to merge-conflict reimports).
pub(crate) const fn default_true() -> bool {
    true
}
