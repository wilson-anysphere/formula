use std::collections::BTreeSet;
use std::path::PathBuf;

use formula_engine::{eval, functions};
use serde::Deserialize;

#[derive(Debug, Deserialize)]
struct OracleCorpus {
    cases: Vec<OracleCase>,
}

#[derive(Debug, Deserialize)]
struct OracleCase {
    id: String,
    formula: String,
    #[serde(default)]
    inputs: Vec<OracleCellInput>,
}

#[derive(Debug, Deserialize)]
struct OracleCellInput {
    cell: String,
    #[serde(default)]
    formula: Option<String>,
}

fn load_excel_oracle_cases() -> Vec<OracleCase> {
    let mut path = PathBuf::from(env!("CARGO_MANIFEST_DIR"));
    path.push("../../tests/compatibility/excel-oracle/cases.json");
    let corpus_bytes =
        std::fs::read_to_string(&path).unwrap_or_else(|e| panic!("read {}: {e}", path.display()));

    let corpus: OracleCorpus = serde_json::from_str(&corpus_bytes)
        .unwrap_or_else(|e| panic!("parse excel oracle corpus JSON ({}): {e}", path.display()));
    corpus.cases
}

#[test]
fn excel_oracle_function_calls_are_registered() {
    // Keep `tests/compatibility/excel-oracle/cases.json` aligned with the function registry so
    // new oracle cases don't silently regress to `#NAME?`.
    let cases = load_excel_oracle_cases();

    let mut unknown = BTreeSet::new();
    for case in cases {
        let parsed = eval::Parser::parse(&case.formula).unwrap_or_else(|e| {
            panic!(
                "parse excel oracle formula ({}) {:?}: {e}",
                case.id, case.formula
            )
        });
        collect_unknown_function_calls(&parsed, &mut unknown);

        // Input cells can also contain formulas (e.g. `=NA()`).
        for input in &case.inputs {
            let Some(input_formula) = input.formula.as_deref() else {
                continue;
            };
            let parsed = eval::Parser::parse(input_formula).unwrap_or_else(|e| {
                panic!(
                    "parse excel oracle input formula ({} {}) {input_formula:?}: {e}",
                    case.id, input.cell
                )
            });
            collect_unknown_function_calls(&parsed, &mut unknown);
        }
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

fn normalize_function_call_name(name: &str) -> String {
    // Mirror `functions::lookup_function` behavior: Excel stores newer functions with an
    // `_xlfn.` prefix, but they should count as coverage for the unprefixed built-in.
    let upper = name.to_ascii_uppercase();
    upper.strip_prefix("_XLFN.").unwrap_or(&upper).to_string()
}

#[test]
fn excel_oracle_corpus_covers_nonvolatile_function_catalog() {
    // Keep `tests/compatibility/excel-oracle/cases.json` aligned with `shared/functionCatalog.json`
    // so we have at least one deterministic oracle case for every implemented non-volatile
    // function. Volatile functions are intentionally excluded from the oracle corpus.
    let cases = load_excel_oracle_cases();

    let mut called_in_case_formulas = BTreeSet::new();
    let mut called_in_any_formula = BTreeSet::new();
    for case in cases {
        let parsed = eval::Parser::parse(&case.formula).unwrap_or_else(|e| {
            panic!(
                "parse excel oracle formula ({}) {:?}: {e}",
                case.id, case.formula
            )
        });
        collect_function_calls(&parsed, &mut called_in_case_formulas);
        collect_function_calls(&parsed, &mut called_in_any_formula);

        // Input cells can also contain formulas (e.g. `=NA()`); these should not contain volatile
        // functions, but they do not count towards coverage of `case.formula`.
        for input in &case.inputs {
            let Some(input_formula) = input.formula.as_deref() else {
                continue;
            };
            let parsed = eval::Parser::parse(input_formula).unwrap_or_else(|e| {
                panic!(
                    "parse excel oracle input formula ({} {}) {input_formula:?}: {e}",
                    case.id, input.cell
                )
            });
            collect_function_calls(&parsed, &mut called_in_any_formula);
        }
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
    let present_volatile: BTreeSet<_> = volatile
        .intersection(&called_in_any_formula)
        .cloned()
        .collect();
    assert!(
        present_volatile.is_empty(),
        "oracle corpus includes volatile functions (non-deterministic).\n\
         Remove them from tests/compatibility/excel-oracle/cases.json (the oracle corpus is intended to stay deterministic).\n\
         Volatile functions:\n{}",
        present_volatile
            .iter()
            .map(|name| format!("  - {name}"))
            .collect::<Vec<_>>()
            .join("\n")
    );

    let missing: BTreeSet<_> = nonvolatile
        .difference(&called_in_case_formulas)
        .filter(|name| !EXCEPTIONS.contains(&name.as_str()))
        .cloned()
        .collect();
    assert!(
        missing.is_empty(),
        "oracle corpus is missing coverage (case.formula) for {} non-volatile functions from shared/functionCatalog.json.\n\
         Add at least one case in tests/compatibility/excel-oracle/cases.json for each missing function.\n\
         (Tip: regenerate the corpus with `python tools/excel-oracle/generate_cases.py --out tests/compatibility/excel-oracle/cases.json`.)\n\
         Missing functions:\n{}",
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
        eval::Expr::FieldAccess { base, .. } => collect_unknown_function_calls(base, unknown),
        eval::Expr::FunctionCall { name, args, .. } => {
            if functions::lookup_function(name).is_none() {
                unknown.insert(normalize_function_call_name(name));
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
        eval::Expr::FieldAccess { base, .. } => collect_unknown_function_calls(base, unknown),
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
        eval::Expr::FieldAccess { base, .. } => collect_function_calls(base, called),
        eval::Expr::FunctionCall { name, args, .. } => {
            called.insert(normalize_function_call_name(name));
            for arg in args {
                collect_function_calls(arg, called);
            }
        }
        eval::Expr::Unary { expr, .. } => collect_function_calls(expr, called),
        eval::Expr::Postfix { expr, .. } => collect_function_calls(expr, called),
        eval::Expr::FieldAccess { base, .. } => collect_function_calls(base, called),
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
