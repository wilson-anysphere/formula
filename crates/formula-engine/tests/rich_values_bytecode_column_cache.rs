#![cfg(not(target_arch = "wasm32"))]

use formula_engine::eval::CellAddr;
use formula_engine::value::{EntityValue, RecordValue};
use formula_engine::{Engine, ExternalValueProvider, Value};
use std::sync::Arc;

#[derive(Debug)]
struct TwoCellProvider {
    a1: Value,
    a2: Value,
}

impl TwoCellProvider {
    fn new(a1: Value, a2: Value) -> Self {
        Self { a1, a2 }
    }
}

impl ExternalValueProvider for TwoCellProvider {
    fn get(&self, sheet: &str, addr: CellAddr) -> Option<Value> {
        if sheet != "Sheet1" {
            return None;
        }

        match (addr.row, addr.col) {
            (0, 0) => Some(self.a1.clone()),
            (1, 0) => Some(self.a2.clone()),
            _ => None,
        }
    }
}

#[test]
fn bytecode_column_cache_ignores_entity_values_in_ranges() {
    let provider = Arc::new(TwoCellProvider::new(
        Value::Number(1.0),
        Value::Entity(EntityValue::new("Entity#1")),
    ));

    let mut engine = Engine::new();
    engine.set_external_value_provider(Some(provider.clone()));
    engine
        .set_cell_formula("Sheet1", "B1", "=SUM(A1:A2)")
        .unwrap();

    engine.recalculate_single_threaded();

    // Entity values in references should be ignored (like text) by SUM ranges.
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(1.0));

    // SUM should be compiled to bytecode (and thus exercise the column cache build path).
    let stats = engine.bytecode_compile_stats();
    assert_eq!(stats.total_formula_cells, 1);
    assert_eq!(stats.compiled, 1);
}

#[test]
fn bytecode_column_cache_ignores_record_values_in_ranges() {
    let provider = Arc::new(TwoCellProvider::new(
        Value::Number(1.0),
        Value::Record(RecordValue::new("Record#1")),
    ));

    let mut engine = Engine::new();
    engine.set_external_value_provider(Some(provider.clone()));
    engine
        .set_cell_formula("Sheet1", "B1", "=SUM(A1:A2)")
        .unwrap();

    engine.recalculate_single_threaded();

    // Record values in references should be ignored (like text) by SUM ranges.
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(1.0));

    let stats = engine.bytecode_compile_stats();
    assert_eq!(stats.total_formula_cells, 1);
    assert_eq!(stats.compiled, 1);
}
