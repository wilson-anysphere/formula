/// Helpers for working with formula strings across layers (UI, file formats, engine).
///
/// # Invariant
///
/// The canonical formula representation stored in [`crate::Cell::formula`] is:
/// - trimmed
/// - **without** a leading `'='`
///
/// SpreadsheetML stores formulas in `<f>` elements without a leading `'='`, while
/// most UIs display formulas with one. These helpers provide a single
/// implementation of that policy to avoid ad-hoc conversions across the
/// codebase.
///
/// Note: these helpers intentionally do not validate formula syntax.

/// Normalize formula text into the canonical `formula-model` representation.
///
/// - Trims leading/trailing whitespace.
/// - Strips a single leading `'='` if present.
/// - Returns `None` for empty/whitespace-only strings (and bare `"="`).
pub fn normalize_formula_text(s: &str) -> Option<String> {
    let trimmed = s.trim();
    let stripped = trimmed.strip_prefix('=').unwrap_or(trimmed);
    let stripped = stripped.trim();

    if stripped.is_empty() {
        None
    } else {
        Some(stripped.to_string())
    }
}

/// Convert formula text into a UI-friendly display form.
///
/// The returned string is either empty (if the input is empty/whitespace) or
/// starts with a leading `'='`.
///
/// This accepts either canonical text (no `'='`) or display text (leading `'='`)
/// and produces the display form.
pub fn display_formula_text(s: &str) -> String {
    match normalize_formula_text(s) {
        Some(normalized) => format!("={normalized}"),
        None => String::new(),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn normalize_returns_none_for_empty_or_whitespace() {
        assert_eq!(normalize_formula_text(""), None);
        assert_eq!(normalize_formula_text("   "), None);
        assert_eq!(normalize_formula_text("\n\t"), None);
        assert_eq!(normalize_formula_text("="), None);
        assert_eq!(normalize_formula_text("   =   "), None);
    }

    #[test]
    fn normalize_strips_single_leading_equals_and_trims() {
        assert_eq!(normalize_formula_text("=1+1"), Some("1+1".to_string()));
        assert_eq!(
            normalize_formula_text("  =  SUM(A1:A3)  "),
            Some("SUM(A1:A3)".to_string())
        );
        assert_eq!(
            normalize_formula_text("==1+1"),
            Some("=1+1".to_string()),
            "only a single leading '=' is stripped"
        );
    }

    #[test]
    fn display_ensures_leading_equals() {
        assert_eq!(display_formula_text("1+1"), "=1+1");
        assert_eq!(display_formula_text("=1+1"), "=1+1");
        assert_eq!(display_formula_text("   = 1+1 "), "=1+1");
        assert_eq!(display_formula_text(""), "");
        assert_eq!(display_formula_text("   "), "");
    }
}
