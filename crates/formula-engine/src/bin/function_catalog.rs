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
        let entry = FunctionCatalogEntry {
            name: name.clone(),
            min_args: spec.min_args,
            max_args: spec.max_args,
            volatility: volatility_to_string(spec.volatility),
            return_type: value_type_to_string(spec.return_type),
            arg_types: spec
                .arg_types
                .iter()
                .copied()
                .map(value_type_to_string)
                .collect(),
        };

        if functions.insert(name.clone(), entry).is_some() {
            return Err(format!(
                "duplicate function name registered in formula-engine inventory: {name}"
            )
            .into());
        }
    }

    let catalog = FunctionCatalog {
        functions: functions.into_values().collect(),
    };

    let stdout = io::stdout();
    let mut out = io::BufWriter::new(stdout.lock());
    serde_json::to_writer_pretty(&mut out, &catalog)?;
    out.write_all(b"\n")?;
    out.flush()?;
    Ok(())
}
