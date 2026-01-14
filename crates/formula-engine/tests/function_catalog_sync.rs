use std::collections::{BTreeMap, BTreeSet};
use std::fs;
use std::path::PathBuf;

use formula_engine::functions::{iter_function_specs, ValueType, Volatility};
use serde::Deserialize;
use serde_json::Value as JsonValue;

#[derive(Debug, Deserialize)]
struct FunctionCatalog {
    functions: Vec<CatalogFunction>,
}

#[derive(Debug, Deserialize)]
struct CatalogFunction {
    name: String,
    min_args: usize,
    max_args: usize,
    volatility: String,
    return_type: String,
    arg_types: Vec<String>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct ComparableFunctionSpec {
    min_args: usize,
    max_args: usize,
    volatility: Volatility,
    return_type: ValueType,
    arg_types: Vec<ValueType>,
}

fn parse_value_type(value_type: &str) -> ValueType {
    match value_type {
        "any" => ValueType::Any,
        "number" => ValueType::Number,
        "text" => ValueType::Text,
        "bool" => ValueType::Bool,
        other => panic!("unknown value type in functionCatalog.json: {other}"),
    }
}

fn format_value_type(value_type: ValueType) -> &'static str {
    match value_type {
        ValueType::Any => "any",
        ValueType::Number => "number",
        ValueType::Text => "text",
        ValueType::Bool => "bool",
    }
}

fn parse_volatility(volatility: &str) -> Volatility {
    match volatility {
        "volatile" => Volatility::Volatile,
        "non_volatile" => Volatility::NonVolatile,
        other => panic!("unknown volatility in functionCatalog.json: {other}"),
    }
}

fn format_volatility(volatility: Volatility) -> &'static str {
    match volatility {
        Volatility::Volatile => "volatile",
        Volatility::NonVolatile => "non_volatile",
    }
}

fn format_value_types(types: &[ValueType]) -> String {
    let mut out = String::from("[");
    for (idx, ty) in types.iter().copied().enumerate() {
        if idx > 0 {
            out.push_str(", ");
        }
        out.push_str(format_value_type(ty));
    }
    out.push(']');
    out
}

#[test]
fn function_catalog_sync() {
    let catalog_path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("..")
        .join("..")
        .join("shared")
        .join("functionCatalog.json");

    let raw_catalog = fs::read_to_string(&catalog_path).unwrap_or_else(|err| {
        panic!(
            "failed to read function catalog at {}: {err}",
            catalog_path.display()
        )
    });

    let catalog: FunctionCatalog = serde_json::from_str(&raw_catalog).unwrap_or_else(|err| {
        panic!(
            "failed to parse function catalog JSON at {}: {err}",
            catalog_path.display()
        )
    });

    // Keep the catalog stable for deterministic diffs and downstream tooling.
    let catalog_names: Vec<String> = catalog
        .functions
        .iter()
        .map(|entry| entry.name.clone())
        .collect();
    for name in &catalog_names {
        assert_eq!(
            name,
            &name.to_ascii_uppercase(),
            "expected functionCatalog.json function names to be uppercase, got {name}"
        );
    }
    let mut sorted_names = catalog_names.clone();
    sorted_names.sort();
    assert_eq!(
        catalog_names, sorted_names,
        "expected functionCatalog.json to be sorted by name for deterministic diffs"
    );
    assert_eq!(
        catalog_names.len(),
        catalog_names.iter().collect::<BTreeSet<_>>().len(),
        "expected functionCatalog.json to contain unique names"
    );

    let mut catalog_specs: BTreeMap<String, ComparableFunctionSpec> = BTreeMap::new();
    for entry in catalog.functions {
        let key = entry.name.to_ascii_uppercase();
        let spec = ComparableFunctionSpec {
            min_args: entry.min_args,
            max_args: entry.max_args,
            volatility: parse_volatility(&entry.volatility),
            return_type: parse_value_type(&entry.return_type),
            arg_types: entry
                .arg_types
                .iter()
                .map(|t| parse_value_type(t))
                .collect(),
        };
        if catalog_specs.insert(key.clone(), spec).is_some() {
            panic!(
                "duplicate function name in functionCatalog.json (case-insensitive): {}",
                entry.name
            );
        }
    }

    let mut registry_specs: BTreeMap<String, ComparableFunctionSpec> = BTreeMap::new();
    for spec in iter_function_specs() {
        let key = spec.name.to_ascii_uppercase();
        let comparable = ComparableFunctionSpec {
            min_args: spec.min_args,
            max_args: spec.max_args,
            volatility: spec.volatility,
            return_type: spec.return_type,
            arg_types: spec.arg_types.to_vec(),
        };

        if registry_specs.insert(key.clone(), comparable).is_some() {
            panic!(
                "duplicate function name in registry (case-insensitive): {}",
                spec.name
            );
        }
    }

    let catalog_names: BTreeSet<String> = catalog_specs.keys().cloned().collect();
    let registry_names: BTreeSet<String> = registry_specs.keys().cloned().collect();

    let missing_in_registry: Vec<String> =
        catalog_names.difference(&registry_names).cloned().collect();
    let missing_in_catalog: Vec<String> =
        registry_names.difference(&catalog_names).cloned().collect();

    let mut mismatches: BTreeMap<String, Vec<String>> = BTreeMap::new();
    for name in catalog_names.intersection(&registry_names) {
        let catalog_spec = &catalog_specs[name];
        let registry_spec = &registry_specs[name];

        let mut diffs = Vec::new();
        if catalog_spec.min_args != registry_spec.min_args {
            diffs.push(format!(
                "min_args catalog={} registry={}",
                catalog_spec.min_args, registry_spec.min_args
            ));
        }
        if catalog_spec.max_args != registry_spec.max_args {
            diffs.push(format!(
                "max_args catalog={} registry={}",
                catalog_spec.max_args, registry_spec.max_args
            ));
        }
        if catalog_spec.volatility != registry_spec.volatility {
            diffs.push(format!(
                "volatility catalog={} registry={}",
                format_volatility(catalog_spec.volatility),
                format_volatility(registry_spec.volatility)
            ));
        }
        if catalog_spec.return_type != registry_spec.return_type {
            diffs.push(format!(
                "return_type catalog={} registry={}",
                format_value_type(catalog_spec.return_type),
                format_value_type(registry_spec.return_type)
            ));
        }
        if catalog_spec.arg_types != registry_spec.arg_types {
            diffs.push(format!(
                "arg_types catalog={} registry={}",
                format_value_types(&catalog_spec.arg_types),
                format_value_types(&registry_spec.arg_types)
            ));
        }

        if !diffs.is_empty() {
            mismatches.insert(name.clone(), diffs);
        }
    }

    if missing_in_registry.is_empty() && missing_in_catalog.is_empty() && mismatches.is_empty() {
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

    if !missing_in_catalog.is_empty() {
        report.push_str("\nMissing in functionCatalog.json (present in registry):\n");
        for name in &missing_in_catalog {
            report.push_str(&format!("  - {name}\n"));
        }
    }

    if !mismatches.is_empty() {
        report.push_str("\nField mismatches:\n");
        for (name, diffs) in &mismatches {
            report.push_str(&format!("  - {name}:\n"));
            for diff in diffs {
                report.push_str(&format!("      {diff}\n"));
            }
        }
    }

    panic!("{report}");
}

#[test]
fn function_catalog_mjs_matches_committed_json() {
    let shared_dir = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("..")
        .join("..")
        .join("shared");
    let json_path = shared_dir.join("functionCatalog.json");
    let mjs_path = shared_dir.join("functionCatalog.mjs");

    let raw_json = fs::read_to_string(&json_path)
        .unwrap_or_else(|err| panic!("failed to read {}: {err}", json_path.display()));
    let json_value: JsonValue = serde_json::from_str(&raw_json)
        .unwrap_or_else(|err| panic!("failed to parse {}: {err}", json_path.display()));

    let raw_mjs = fs::read_to_string(&mjs_path)
        .unwrap_or_else(|err| panic!("failed to read {}: {err}", mjs_path.display()));
    let raw_mjs = raw_mjs
        .strip_prefix(
            "// This file is generated by scripts/generate-function-catalog.js. Do not edit.\n",
        )
        .unwrap_or_else(|| {
            panic!(
                "{} did not start with expected generated header",
                mjs_path.display()
            )
        });

    let raw_mjs = raw_mjs
        .strip_prefix("export default ")
        .unwrap_or_else(|| panic!("{} did not start with `export default`", mjs_path.display()));
    let raw_mjs = raw_mjs
        .trim_end()
        .strip_suffix(';')
        .unwrap_or_else(|| panic!("{} did not end with a semicolon", mjs_path.display()));

    let mjs_value: JsonValue = serde_json::from_str(raw_mjs).unwrap_or_else(|err| {
        panic!(
            "failed to parse embedded JSON in {}: {err}",
            mjs_path.display()
        )
    });

    assert_eq!(
        json_value,
        mjs_value,
        "{} is out of sync with {}",
        mjs_path.display(),
        json_path.display()
    );
}
