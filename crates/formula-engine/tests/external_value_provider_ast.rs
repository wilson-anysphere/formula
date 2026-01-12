use std::sync::Arc;

use formula_engine::{Engine, ExternalValueProvider, Value};

#[test]
fn ast_backend_respects_external_value_provider_for_range_iteration() {
    struct Provider;

    impl ExternalValueProvider for Provider {
        fn get(&self, sheet: &str, addr: formula_engine::eval::CellAddr) -> Option<Value> {
            if sheet != "Sheet1" {
                return None;
            }
            match (addr.row, addr.col) {
                (0, 0) => Some(Value::Number(1.0)),
                (1, 0) => Some(Value::Number(2.0)),
                (2, 0) => Some(Value::Number(3.0)),
                _ => None,
            }
        }
    }

    let mut engine = Engine::new();
    engine.set_external_value_provider(Some(Arc::new(Provider)));
    engine.set_bytecode_enabled(false);
    engine
        .set_cell_formula("Sheet1", "B1", "=SUM(A1:A3)")
        .unwrap();

    engine.recalculate_single_threaded();
    assert_eq!(engine.bytecode_program_count(), 0);
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(6.0));
}

