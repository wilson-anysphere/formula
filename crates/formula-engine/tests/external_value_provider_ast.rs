use std::sync::Arc;

use formula_engine::{Engine, ExternalValueProvider, Value};
use formula_model::Style;

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

#[test]
fn ast_backend_style_only_cells_do_not_override_external_value_provider() {
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

    // Create a "style-only" cell record at A1 (blank + no formula + non-default style).
    let style_id = engine.intern_style(Style {
        number_format: Some("0.00".to_string()),
        ..Default::default()
    });
    engine
        .set_cell_style_id("Sheet1", "A1", style_id)
        .expect("set style id");

    engine
        .set_cell_formula("Sheet1", "B1", "=SUM(A1:A3)")
        .unwrap();

    engine.recalculate_single_threaded();
    assert_eq!(engine.bytecode_program_count(), 0);

    // Provider values should flow through style-only cells.
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(6.0));
}
