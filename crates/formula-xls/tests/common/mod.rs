pub mod xls_fixture_builder;

use formula_engine::{parse_formula, ParseOptions};

/// Assert that a formula expression string is parseable by `formula-engine`.
///
/// Note: `parse_formula` accepts formulas both with and without a leading `=`, so callers should
/// pass whatever representation they have (canonical `formula-model` strings omit it).
#[allow(dead_code)]
pub fn assert_parseable_formula(expr: &str) {
    let expr = expr.trim();
    assert!(!expr.is_empty(), "expected formula text to be non-empty");
    parse_formula(expr, ParseOptions::default()).unwrap_or_else(|err| {
        panic!("expected formula to be parseable, expr={expr:?}, err={err:?}")
    });
}
