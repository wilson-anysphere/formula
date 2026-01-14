use formula_format::{cell_format_info, FormatOptions};

#[test]
fn cell_parentheses_ignores_underscore_and_fill_operands() {
    let options = FormatOptions::default();

    // Underscore alignment tokens reserve the width of the next character, but do not render it.
    // Parentheses used as the underscore operand must not count toward CELL("parentheses").
    assert_eq!(cell_format_info(Some("0;0_)"), &options).parentheses, 0);
    assert_eq!(cell_format_info(Some("0;0_("), &options).parentheses, 0);

    // Fill tokens repeat the next character to fill cell width; the operand is a layout hint and
    // should not be treated as a literal parenthesis for CELL("parentheses") classification.
    assert_eq!(cell_format_info(Some("0;0*)"), &options).parentheses, 0);

    // Regression: when both '(' and ')' appear only as underscore/fill operands, they should not
    // trigger the negative-parentheses flag.
    assert_eq!(cell_format_info(Some("0;0_(_)"), &options).parentheses, 0);
    assert_eq!(cell_format_info(Some("0;0*(*)"), &options).parentheses, 0);
}

#[test]
fn cell_parentheses_detects_accounting_parentheses() {
    let options = FormatOptions::default();
    // Canonical accounting-style format: negative section wraps the number in parentheses.
    assert_eq!(
        cell_format_info(Some("#,##0_);(#,##0)"), &options).parentheses,
        1
    );
}
