use std::collections::BTreeSet;
use std::path::PathBuf;

use formula_engine::{eval, functions};

#[test]
fn excel_oracle_function_calls_are_registered() {
    // Keep `tests/compatibility/excel-oracle/cases.json` aligned with the function registry so
    // new oracle cases don't silently regress to `#NAME?`.
    let mut path = PathBuf::from(env!("CARGO_MANIFEST_DIR"));
    path.push("../../tests/compatibility/excel-oracle/cases.json");
    let corpus_bytes =
        std::fs::read_to_string(&path).unwrap_or_else(|e| panic!("read {}: {e}", path.display()));

    let corpus: serde_json::Value =
        serde_json::from_str(&corpus_bytes).expect("parse excel oracle corpus JSON");
    let cases = corpus
        .get("cases")
        .and_then(|v| v.as_array())
        .expect("cases[]");

    let mut unknown = BTreeSet::new();
    for case in cases {
        let formula = case
            .get("formula")
            .and_then(|v| v.as_str())
            .expect("case.formula");
        let parsed = eval::Parser::parse(formula)
            .unwrap_or_else(|e| panic!("parse formula {formula:?}: {e}"));
        collect_unknown_function_calls(&parsed, &mut unknown);
    }

    // The corpus intentionally includes `NO_SUCH_FUNCTION` to validate that unknown functions
    // still evaluate to `#NAME?`.
    assert_eq!(unknown, BTreeSet::from(["NO_SUCH_FUNCTION".to_string()]));
}

fn collect_unknown_function_calls(expr: &eval::Expr<String>, unknown: &mut BTreeSet<String>) {
    match expr {
        eval::Expr::FunctionCall { name, args, .. } => {
            if functions::lookup_function(name).is_none() {
                unknown.insert(name.clone());
            }
            for arg in args {
                collect_unknown_function_calls(arg, unknown);
            }
        }
        eval::Expr::Call { callee, args } => {
            collect_unknown_function_calls(callee, unknown);
            for arg in args {
                collect_unknown_function_calls(arg, unknown);
            }
        }
        eval::Expr::Unary { expr, .. } => collect_unknown_function_calls(expr, unknown),
        eval::Expr::Postfix { expr, .. } => collect_unknown_function_calls(expr, unknown),
        eval::Expr::Binary { left, right, .. } | eval::Expr::Compare { left, right, .. } => {
            collect_unknown_function_calls(left, unknown);
            collect_unknown_function_calls(right, unknown);
        }
        eval::Expr::ArrayLiteral { values, .. } => {
            for el in values.iter() {
                collect_unknown_function_calls(el, unknown);
            }
        }
        eval::Expr::ImplicitIntersection(inner) => collect_unknown_function_calls(inner, unknown),
        eval::Expr::SpillRange(inner) => collect_unknown_function_calls(inner, unknown),
        eval::Expr::Number(_)
        | eval::Expr::Text(_)
        | eval::Expr::Bool(_)
        | eval::Expr::Blank
        | eval::Expr::Error(_)
        | eval::Expr::NameRef(_)
        | eval::Expr::CellRef(_)
        | eval::Expr::RangeRef(_)
        | eval::Expr::StructuredRef(_) => {}
    }
}
