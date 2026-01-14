use formula_engine::eval::CellAddr;
use formula_engine::{Engine, ErrorKind, ExternalValueProvider, Value};
use std::sync::Arc;

struct Provider;

impl ExternalValueProvider for Provider {
    fn get(&self, sheet: &str, addr: CellAddr) -> Option<Value> {
        if sheet == "[Book.xlsx]Sheet1" && addr.row == 0 && addr.col == 0 {
            return Some(Value::Number(41.0));
        }
        None
    }
}

#[test]
fn bytecode_external_cell_ref_evaluates_via_provider() {
    let mut engine = Engine::new();
    engine.set_external_value_provider(Some(Arc::new(Provider)));
    engine.set_bytecode_enabled(true);
    engine
        .set_cell_formula("Sheet1", "A1", "=[Book.xlsx]Sheet1!A1+1")
        .unwrap();

    // Ensure we compile to bytecode (no AST fallback).
    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(42.0));
}

#[test]
fn bytecode_missing_external_cell_ref_is_ref_error() {
    struct EmptyProvider;

    impl ExternalValueProvider for EmptyProvider {
        fn get(&self, _sheet: &str, _addr: CellAddr) -> Option<Value> {
            None
        }
    }

    let mut engine = Engine::new();
    engine.set_external_value_provider(Some(Arc::new(EmptyProvider)));
    engine.set_bytecode_enabled(true);
    engine
        .set_cell_formula("Sheet1", "A1", "=[Book.xlsx]Sheet1!A1")
        .unwrap();

    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Sheet1", "A1"),
        Value::Error(ErrorKind::Ref)
    );
}
