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

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn compares_ascii_case_insensitively() {
        assert!(sheet_name_eq_case_insensitive("Sheet1", "sHeEt1"));
        assert!(!sheet_name_eq_case_insensitive("Sheet1", "Sheet2"));
    }

    #[test]
    fn compares_unicode_case_insensitively() {
        // German ß uppercases to "SS".
        assert!(sheet_name_eq_case_insensitive("straße", "STRASSE"));
    }

    #[test]
    fn compares_nfkc_compatibility_equivalents() {
        // Fullwidth Latin A normalizes to ASCII 'A' under NFKC.
        assert!(sheet_name_eq_case_insensitive("Ａ", "A"));
    }
}
