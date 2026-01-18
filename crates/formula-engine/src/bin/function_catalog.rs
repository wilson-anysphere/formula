use std::collections::BTreeMap;
use std::io::{self, Write};

use formula_engine::functions::{FunctionSpec, ValueType, Volatility};
use serde::Serialize;

#[derive(Serialize)]
struct FunctionCatalog {
    functions: Vec<FunctionCatalogEntry>,
}

#[derive(Serialize)]
struct FunctionCatalogEntry {
    name: String,
    min_args: usize,
    max_args: usize,
    volatility: String,
    return_type: String,
    arg_types: Vec<String>,
}

fn volatility_to_string(volatility: Volatility) -> String {
    match volatility {
        Volatility::NonVolatile => "non_volatile".to_string(),
        Volatility::Volatile => "volatile".to_string(),
    }
}

fn value_type_to_string(value_type: ValueType) -> String {
    match value_type {
        ValueType::Any => "any".to_string(),
        ValueType::Number => "number".to_string(),
        ValueType::Text => "text".to_string(),
        ValueType::Bool => "bool".to_string(),
    }
}

fn main() {
    if let Err(err) = main_inner() {
        // Allow piping output to tools like `head` without panicking.
        if let Some(io_err) = err.downcast_ref::<io::Error>() {
            if io_err.kind() == io::ErrorKind::BrokenPipe {
                return;
            }
        }
        if let Some(json_err) = err.downcast_ref::<serde_json::Error>() {
            if json_err.io_error_kind() == Some(io::ErrorKind::BrokenPipe) {
                return;
            }
        }
        eprintln!("error: {err}");
        std::process::exit(1);
    }
}

fn main_inner() -> Result<(), Box<dyn std::error::Error>> {
    let mut functions = BTreeMap::<String, FunctionCatalogEntry>::new();

    for spec in inventory::iter::<FunctionSpec> {
        let name = spec.name.to_ascii_uppercase();
        let mut arg_types: Vec<String> = Vec::new();
        if arg_types.try_reserve_exact(spec.arg_types.len()).is_err() {
            return Err(format!(
                "allocation failed (function catalog arg_types, len={})",
                spec.arg_types.len()
            )
            .into());
        }
        for &arg_type in spec.arg_types {
            arg_types.push(value_type_to_string(arg_type));
        }
        let entry = FunctionCatalogEntry {
            name: name.clone(),
            min_args: spec.min_args,
            max_args: spec.max_args,
            volatility: volatility_to_string(spec.volatility),
            return_type: value_type_to_string(spec.return_type),
            arg_types,
        };

        if functions.insert(name.clone(), entry).is_some() {
            return Err(format!(
                "duplicate function name registered in formula-engine inventory: {name}"
            )
            .into());
        }
    }

    let mut entries: Vec<FunctionCatalogEntry> = Vec::new();
    if entries.try_reserve_exact(functions.len()).is_err() {
        return Err(format!(
            "allocation failed (function catalog entries, len={})",
            functions.len()
        )
        .into());
    }
    for entry in functions.into_values() {
        entries.push(entry);
    }
    let catalog = FunctionCatalog {
        functions: entries,
    };

    let stdout = io::stdout();
    let mut out = io::BufWriter::new(stdout.lock());
    serde_json::to_writer_pretty(&mut out, &catalog)?;
    out.write_all(b"\n")?;
    out.flush()?;
    Ok(())
}
