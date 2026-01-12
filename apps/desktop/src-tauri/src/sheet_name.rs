use unicode_normalization::UnicodeNormalization as _;

/// Excel compares sheet names case-insensitively across Unicode, not ASCII.
///
/// We approximate Excel's behavior by normalizing both inputs with Unicode NFKC and then applying
/// Unicode uppercasing.
pub(crate) fn sheet_name_eq_case_insensitive(a: &str, b: &str) -> bool {
    a.nfkc()
        .flat_map(|c| c.to_uppercase())
        .eq(b.nfkc().flat_map(|c| c.to_uppercase()))
}

