use std::collections::{BTreeMap, BTreeSet};
use std::path::PathBuf;

use formula_engine::functions::{self, Volatility};
use serde::Deserialize;

#[derive(Debug, Deserialize)]
struct FunctionCatalog {
    functions: Vec<CatalogFunction>,
}

#[derive(Debug, Deserialize)]
struct CatalogFunction {
    name: String,
    volatility: String,
}

fn load_function_catalog() -> FunctionCatalog {
    let catalog_path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("..")
        .join("..")
        .join("shared")
        .join("functionCatalog.json");
    let raw_catalog = std::fs::read_to_string(&catalog_path)
        .unwrap_or_else(|e| panic!("read {}: {e}", catalog_path.display()));
    serde_json::from_str(&raw_catalog).unwrap_or_else(|e| {
        panic!(
            "parse shared/functionCatalog.json ({}): {e}",
            catalog_path.display()
        )
    })
}

fn normalize_function_name(name: &str) -> String {
    // Mirror `functions::lookup_function` behavior: Excel stores newer functions with an
    // `_xlfn.` prefix, but they should resolve to the unprefixed built-in.
    let upper = name.to_ascii_uppercase();
    upper.strip_prefix("_XLFN.").unwrap_or(&upper).to_string()
}

fn parse_catalog_volatility(name: &str, volatility: &str) -> Volatility {
    match volatility {
        "volatile" => Volatility::Volatile,
        "non_volatile" => Volatility::NonVolatile,
        other => panic!("unknown volatility in shared/functionCatalog.json for {name}: {other}"),
    }
}

fn format_volatility(volatility: Volatility) -> &'static str {
    match volatility {
        Volatility::Volatile => "volatile",
        Volatility::NonVolatile => "non_volatile",
    }
}

#[test]
fn function_catalog_functions_are_registered_and_volatility_matches() {
    let catalog = load_function_catalog();

    // If a catalog entry is intentionally not exposed via `functions::lookup_function`, add it
    // here with a justification comment. Keep this list small.
    const EXCEPTIONS: &[&str] = &[];

    let mut missing_in_registry = BTreeSet::new();
    let mut volatility_mismatches = BTreeMap::new();

    for entry in &catalog.functions {
        let normalized = normalize_function_name(&entry.name);
        if EXCEPTIONS.contains(&normalized.as_str()) {
            continue;
        }

        let Some(spec) = functions::lookup_function(&entry.name) else {
            missing_in_registry.insert(normalized);
            continue;
        };

        let expected = parse_catalog_volatility(&entry.name, &entry.volatility);
        if spec.volatility != expected {
            volatility_mismatches.insert(
                normalized,
                (entry.volatility.clone(), format_volatility(spec.volatility)),
            );
        }
    }

    if missing_in_registry.is_empty() && volatility_mismatches.is_empty() {
        return;
    }

    let mut report =
        String::from("shared/functionCatalog.json is out of sync with formula-engine registry\n");
    report.push_str(
        "\nTo regenerate function catalog artifacts, run:\n  pnpm -w run generate:function-catalog\n",
    );

    if !missing_in_registry.is_empty() {
        report.push_str("\nMissing in registry (present in functionCatalog.json):\n");
        for name in &missing_in_registry {
            report.push_str(&format!("  - {name}\n"));
        }
    }

    if !volatility_mismatches.is_empty() {
        report.push_str("\nVolatility mismatches:\n");
        for (name, (catalog, registry)) in &volatility_mismatches {
            report.push_str(&format!(
                "  - {name}: catalog={catalog} registry={registry}\n"
            ));
        }
    }

    panic!("{report}");
}

#[test]
fn all_registered_functions_exist_in_function_catalog() {
    let catalog = load_function_catalog();
    let catalog_names: BTreeSet<String> = catalog
        .functions
        .iter()
        .map(|entry| normalize_function_name(&entry.name))
        .collect();

    // The engine may expose internal-only synthetic functions via its runtime registry.
    // These should not be part of the shared function catalog consumed by JS tooling.
    const ALLOWLIST: &[&str] = &[
        // Synthetic function used by expression lowering for the field-access operator.
        "_FIELDACCESS",
    ];

    let mut missing_in_catalog = BTreeSet::new();
    for spec in functions::iter_function_specs() {
        let name = normalize_function_name(spec.name);
        if ALLOWLIST.contains(&name.as_str()) {
            continue;
        }
        if !catalog_names.contains(&name) {
            missing_in_catalog.insert(name);
        }
    }

    assert!(
        missing_in_catalog.is_empty(),
        "formula-engine registry contains functions missing from shared/functionCatalog.json.\n\
         Add them to shared/functionCatalog.json (or regenerate via `pnpm -w run generate:function-catalog`).\n\
         Missing functions:\n{}",
        missing_in_catalog
            .iter()
            .map(|name| format!("  - {name}"))
            .collect::<Vec<_>>()
            .join("\n")
    );
}
