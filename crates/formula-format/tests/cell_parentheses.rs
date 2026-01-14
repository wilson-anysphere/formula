use formula_format::cell_parentheses_flag;

#[test]
fn cell_parentheses_ignores_underscore_and_fill_operands() {
    // Underscore alignment tokens reserve the width of the next character, but do not render it.
    // Parentheses used as the underscore operand must not count toward CELL("parentheses").
    assert_eq!(cell_parentheses_flag(Some("0;0_)")), 0);
    assert_eq!(cell_parentheses_flag(Some("0;0_(")), 0);

    // Fill tokens repeat the next character to fill cell width; the operand is a layout hint and
    // should not be treated as a literal parenthesis for CELL("parentheses") classification.
    assert_eq!(cell_parentheses_flag(Some("0;0*)")), 0);
}

#[test]
fn cell_parentheses_detects_accounting_parentheses() {
    // Canonical accounting-style format: negative section wraps the number in parentheses.
    assert_eq!(cell_parentheses_flag(Some("#,##0_);(#,##0)")), 1);
}
