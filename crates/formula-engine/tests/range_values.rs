use formula_engine::eval::CellAddr;
use formula_engine::{Engine, ErrorKind, ExternalValueProvider, Value};
use formula_model::Range;
use std::sync::atomic::{AtomicUsize, Ordering};
use std::sync::Arc;

#[test]
fn get_range_values_includes_blanks_for_unset_cells() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).expect("set A1");
    engine
        .set_cell_value("Sheet1", "C2", "hello")
        .expect("set C2");

    let range = Range::from_a1("A1:C2").expect("range");
    let values = engine
        .get_range_values("Sheet1", range)
        .expect("get_range_values");

    assert_eq!(
        values,
        vec![
            vec![Value::Number(1.0), Value::Blank, Value::Blank],
            vec![Value::Blank, Value::Blank, Value::Text("hello".to_string())],
        ]
    );
}

#[test]
fn get_range_values_returns_ref_for_out_of_bounds_cells() {
    let mut engine = Engine::new();
    engine.set_sheet_dimensions("Sheet1", 2, 2).unwrap(); // A1:B2
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "B2", 2.0).unwrap();

    let range = Range::from_a1("A1:C3").unwrap();
    let values = engine.get_range_values("Sheet1", range).unwrap();

    assert_eq!(
        values,
        vec![
            vec![
                Value::Number(1.0),
                Value::Blank,
                Value::Error(ErrorKind::Ref)
            ],
            vec![
                Value::Blank,
                Value::Number(2.0),
                Value::Error(ErrorKind::Ref)
            ],
            vec![
                Value::Error(ErrorKind::Ref),
                Value::Error(ErrorKind::Ref),
                Value::Error(ErrorKind::Ref)
            ],
        ]
    );
}

#[test]
fn get_range_values_includes_spilled_array_outputs() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=SEQUENCE(2,2)")
        .unwrap();
    engine.recalculate();

    let range = Range::from_a1("A1:B2").unwrap();
    let values = engine.get_range_values("Sheet1", range).unwrap();

    assert_eq!(
        values,
        vec![
            vec![Value::Number(1.0), Value::Number(2.0)],
            vec![Value::Number(3.0), Value::Number(4.0)],
        ]
    );
}

#[test]
fn get_range_values_queries_external_provider_for_missing_cells() {
    struct Provider {
        calls: AtomicUsize,
    }

    impl ExternalValueProvider for Provider {
        fn get(&self, _sheet: &str, addr: CellAddr) -> Option<Value> {
            self.calls.fetch_add(1, Ordering::SeqCst);
            if addr.row == 1 && addr.col == 1 {
                Some(Value::Number(9.0))
            } else {
                None
            }
        }
    }

    let provider = Arc::new(Provider {
        calls: AtomicUsize::new(0),
    });

    let mut engine = Engine::new();
    engine.ensure_sheet("Sheet1");
    engine.set_external_value_provider(Some(provider.clone()));

    let range = Range::from_a1("A1:B2").unwrap();
    let values = engine.get_range_values("Sheet1", range).unwrap();

    assert_eq!(
        values,
        vec![
            vec![Value::Blank, Value::Blank],
            vec![Value::Blank, Value::Number(9.0)],
        ]
    );
    assert_eq!(provider.calls.load(Ordering::SeqCst), 4);
}
