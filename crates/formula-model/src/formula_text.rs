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

/// Normalize formula text into the canonical `formula-model` representation.
///
/// - Trims leading/trailing whitespace.
/// - Strips a single leading `'='` if present.
///
/// This function is intentionally conservative: it does not attempt to validate
/// formula syntax.
pub fn normalize_formula_text(s: &str) -> String {
    let mut trimmed = s.trim();
    if let Some(rest) = trimmed.strip_prefix('=') {
        trimmed = rest.trim();
    }
    trimmed.to_string()
}

/// Convert formula text into a UI-friendly display form.
///
/// The returned string is either empty (if the input is empty/whitespace) or
/// starts with a leading `'='`.
///
/// This accepts either canonical text (no `'='`) or display text (leading `'='`)
/// and produces the display form.
pub fn display_formula_text(s: &str) -> String {
    let normalized = normalize_formula_text(s);
    if normalized.is_empty() {
        String::new()
    } else {
        format!("={normalized}")
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn normalize_strips_equals_and_trims() {
        assert_eq!(normalize_formula_text("=1+1"), "1+1");
        assert_eq!(normalize_formula_text("  =  SUM(A1:A3)  "), "SUM(A1:A3)");
        assert_eq!(normalize_formula_text("SUM(A1:A3)"), "SUM(A1:A3)");
        assert_eq!(normalize_formula_text("   "), "");
        assert_eq!(normalize_formula_text("="), "");
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
