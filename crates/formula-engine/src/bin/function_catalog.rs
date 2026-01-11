use std::collections::BTreeMap;

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
    let mut functions = BTreeMap::<String, FunctionCatalogEntry>::new();

    for spec in inventory::iter::<FunctionSpec> {
        let name = spec.name.to_ascii_uppercase();
        let entry = FunctionCatalogEntry {
            name: name.clone(),
            min_args: spec.min_args,
            max_args: spec.max_args,
            volatility: volatility_to_string(spec.volatility),
            return_type: value_type_to_string(spec.return_type),
        };

        if functions.insert(name.clone(), entry).is_some() {
            panic!("Duplicate function name registered in formula-engine inventory: {name}");
        }
    }

    let catalog = FunctionCatalog {
        functions: functions.into_values().collect(),
    };

    let json = serde_json::to_string_pretty(&catalog).expect("serialize function catalog");
    println!("{json}");
}

