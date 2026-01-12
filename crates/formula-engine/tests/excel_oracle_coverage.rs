use std::collections::BTreeSet;
use std::path::PathBuf;

use formula_engine::{eval, functions};
use serde::Deserialize;

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

#[derive(Debug, Deserialize)]
struct FunctionCatalog {
    functions: Vec<CatalogFunction>,
}

#[derive(Debug, Deserialize)]
struct CatalogFunction {
    name: String,
    volatility: String,
}

#[test]
fn excel_oracle_corpus_covers_nonvolatile_function_catalog() {
    // Keep `tests/compatibility/excel-oracle/cases.json` aligned with `shared/functionCatalog.json`
    // so we have at least one deterministic oracle case for every implemented non-volatile
    // function. Volatile functions are intentionally excluded from the oracle corpus.
    let mut cases_path = PathBuf::from(env!("CARGO_MANIFEST_DIR"));
    cases_path.push("../../tests/compatibility/excel-oracle/cases.json");
    let corpus_bytes = std::fs::read_to_string(&cases_path)
        .unwrap_or_else(|e| panic!("read {}: {e}", cases_path.display()));
    let corpus: serde_json::Value =
        serde_json::from_str(&corpus_bytes).expect("parse excel oracle corpus JSON");
    let cases = corpus
        .get("cases")
        .and_then(|v| v.as_array())
        .expect("cases[]");

    let mut called = BTreeSet::new();
    for case in cases {
        let formula = case
            .get("formula")
            .and_then(|v| v.as_str())
            .expect("case.formula");
        let parsed = eval::Parser::parse(formula)
            .unwrap_or_else(|e| panic!("parse formula {formula:?}: {e}"));
        collect_function_calls(&parsed, &mut called);
    }

    let mut catalog_path = PathBuf::from(env!("CARGO_MANIFEST_DIR"));
    catalog_path.push("../../shared/functionCatalog.json");
    let raw_catalog = std::fs::read_to_string(&catalog_path)
        .unwrap_or_else(|e| panic!("read {}: {e}", catalog_path.display()));
    let catalog: FunctionCatalog =
        serde_json::from_str(&raw_catalog).expect("parse shared/functionCatalog.json");

    let mut nonvolatile = BTreeSet::new();
    let mut volatile = BTreeSet::new();
    for f in catalog.functions {
        let name = f.name.to_ascii_uppercase();
        match f.volatility.as_str() {
            "volatile" => {
                volatile.insert(name);
            }
            "non_volatile" => {
                nonvolatile.insert(name);
            }
            other => panic!(
                "unknown volatility in shared/functionCatalog.json for {}: {}",
                f.name, other
            ),
        }
    }

    // If a deterministic function cannot yet be represented in the oracle harness (e.g. it
    // depends on workbook-level state not modeled in `cases.json`), add it to this allow-list
    // with a justification comment. Keep this list small.
    const EXCEPTIONS: &[&str] = &[];

    for &exception in EXCEPTIONS {
        assert!(
            nonvolatile.contains(exception),
            "excel oracle completeness exception {exception:?} is not a non-volatile catalog function"
        );
    }

    // Volatile functions should not appear in the oracle corpus at all.
    let present_volatile: BTreeSet<_> = volatile.intersection(&called).cloned().collect();
    assert!(
        present_volatile.is_empty(),
        "oracle corpus includes volatile functions:\n{}",
        present_volatile
            .iter()
            .map(|name| format!("  - {name}"))
            .collect::<Vec<_>>()
            .join("\n")
    );

    let missing: BTreeSet<_> = nonvolatile
        .difference(&called)
        .filter(|name| !EXCEPTIONS.contains(&name.as_str()))
        .cloned()
        .collect();
    assert!(
        missing.is_empty(),
        "oracle corpus missing coverage for {} non-volatile functions from shared/functionCatalog.json:\n{}",
        missing.len(),
        missing
            .iter()
            .map(|name| format!("  - {name}"))
            .collect::<Vec<_>>()
            .join("\n")
    );
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

fn collect_function_calls(expr: &eval::Expr<String>, called: &mut BTreeSet<String>) {
    match expr {
        eval::Expr::FunctionCall { name, args, .. } => {
            let upper = name.to_ascii_uppercase();
            let normalized = if let Some(rest) = upper.strip_prefix("_XLFN.") {
                rest.to_string()
            } else {
                upper
            };
            called.insert(normalized);
            for arg in args {
                collect_function_calls(arg, called);
            }
        }
        eval::Expr::Unary { expr, .. } => collect_function_calls(expr, called),
        eval::Expr::Postfix { expr, .. } => collect_function_calls(expr, called),
        eval::Expr::Binary { left, right, .. } | eval::Expr::Compare { left, right, .. } => {
            collect_function_calls(left, called);
            collect_function_calls(right, called);
        }
        eval::Expr::Call { callee, args } => {
            collect_function_calls(callee, called);
            for arg in args {
                collect_function_calls(arg, called);
            }
        }
        eval::Expr::ArrayLiteral { values, .. } => {
            for el in values.iter() {
                collect_function_calls(el, called);
            }
        }
        eval::Expr::ImplicitIntersection(inner) => collect_function_calls(inner, called),
        eval::Expr::SpillRange(inner) => collect_function_calls(inner, called),
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
