use formula_engine::{parse_formula, ParseOptions};

fn nested_sum_calls(depth: usize) -> String {
    let mut out = String::from("=");
    for _ in 0..depth {
        out.push_str("SUM(");
    }
    out.push_str("1");
    for _ in 0..depth {
        out.push(')');
    }
    out
}

fn nested_parens(depth: usize) -> String {
    let mut out = String::from("=");
    out.extend(std::iter::repeat('(').take(depth));
    out.push('1');
    out.extend(std::iter::repeat(')').take(depth));
    out
}

fn pow_chain(ops: usize) -> String {
    let mut out = String::from("=1");
    for _ in 0..ops {
        out.push('^');
        out.push('1');
    }
    out
}

#[test]
fn parse_rejects_formula_over_8192_chars() {
    // `="aaaa..."` (includes the leading `=`).
    let formula = format!("=\"{}\"", "a".repeat(8190));
    assert_eq!(formula.chars().count(), 8193);
    assert!(parse_formula(&formula, ParseOptions::default()).is_err());
}

#[test]
fn parse_allows_formula_at_8192_chars() {
    let formula = format!("=\"{}\"", "a".repeat(8189));
    assert_eq!(formula.chars().count(), 8192);
    assert!(parse_formula(&formula, ParseOptions::default()).is_ok());
}

#[test]
fn parse_rejects_function_nesting_over_64() {
    let formula = nested_sum_calls(65);
    assert!(parse_formula(&formula, ParseOptions::default()).is_err());
}

#[test]
fn parse_allows_function_nesting_at_64() {
    let formula = nested_sum_calls(64);
    assert!(parse_formula(&formula, ParseOptions::default()).is_ok());
}

#[test]
fn parse_rejects_parenthesis_nesting_over_64() {
    let formula = nested_parens(65);
    assert!(parse_formula(&formula, ParseOptions::default()).is_err());
}

#[test]
fn parse_allows_parenthesis_nesting_at_64() {
    let formula = nested_parens(64);
    assert!(parse_formula(&formula, ParseOptions::default()).is_ok());
}

#[test]
fn parse_rejects_pow_nesting_over_64() {
    let formula = pow_chain(65);
    assert!(parse_formula(&formula, ParseOptions::default()).is_err());
}

#[test]
fn parse_allows_pow_nesting_at_64() {
    let formula = pow_chain(64);
    assert!(parse_formula(&formula, ParseOptions::default()).is_ok());
}

#[test]
fn parse_rejects_formula_over_tokenized_size_limit() {
    // Build a formula that is well under 8,192 characters, but has enough numeric literals that
    // the (estimated) tokenized byte size exceeds Excel's 16,384-byte limit.
    //
    // We use a non-integer numeric literal (`1.1`) so it is represented as an 8-byte float token.
    let terms = 1639;
    let mut formula = String::from("=");
    for i in 0..terms {
        if i > 0 {
            formula.push('+');
        }
        formula.push_str("1.1");
    }
    assert!(formula.chars().count() < 8192);
    assert!(parse_formula(&formula, ParseOptions::default()).is_err());
}
