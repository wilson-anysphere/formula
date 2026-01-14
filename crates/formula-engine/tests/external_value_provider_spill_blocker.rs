use std::sync::Arc;

use formula_engine::{Engine, ErrorKind, ExternalValueProvider, Value};
use formula_model::Style;

#[test]
fn spill_is_blocked_by_provider_values_even_with_style_only_cells() {
    struct Provider;

    impl ExternalValueProvider for Provider {
        fn get(&self, sheet: &str, addr: formula_engine::eval::CellAddr) -> Option<Value> {
            if sheet != "Sheet1" {
                return None;
            }
            // Provide a non-blank value at D2 to block a spill from D1.
            if addr.row == 1 && addr.col == 3 {
                return Some(Value::Number(999.0));
            }
            None
        }
    }

    let mut engine = Engine::new();
    engine.set_external_value_provider(Some(Arc::new(Provider)));

    // Create a style-only record at the blocking cell (D2). Formatting should not mask the
    // provider-backed value when deciding whether a spill is blocked.
    let style_id = engine.intern_style(Style {
        number_format: Some("0.00".to_string()),
        ..Default::default()
    });
    engine
        .set_cell_style_id("Sheet1", "D2", style_id)
        .expect("set style id");

    engine
        .set_cell_formula("Sheet1", "D1", "=SEQUENCE(2,1)")
        .unwrap();
    engine.recalculate_single_threaded();

    assert_eq!(
        engine.get_cell_value("Sheet1", "D1"),
        Value::Error(ErrorKind::Spill)
    );
    assert_eq!(engine.get_cell_value("Sheet1", "D2"), Value::Number(999.0));
    assert!(engine.spill_range("Sheet1", "D1").is_none());
}
