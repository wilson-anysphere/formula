use formula_engine::eval::CellAddr;
use formula_engine::{Engine, ErrorKind, ExternalValueProvider, PrecedentNode, Value};
use std::sync::atomic::{AtomicUsize, Ordering};
use std::sync::Arc;

#[test]
fn indirect_distinguishes_a1_vs_r1c1_modes() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 123.0).unwrap();
    engine
        .set_cell_formula("Sheet1", "B1", r#"=INDIRECT("R1C1")"#)
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B2", r#"=INDIRECT("R1C1",FALSE)"#)
        .unwrap();

    engine.recalculate();

    // In A1 mode (default), `R1C1` parses as a name rather than an R1C1 reference.
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Error(ErrorKind::Ref));

    // In R1C1 mode, `R1C1` is an absolute ref to `A1`.
    assert_eq!(engine.get_cell_value("Sheet1", "B2"), Value::Number(123.0));
}

#[test]
fn indirect_r1c1_relative_is_resolved_against_formula_cell() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "B2", 42.0).unwrap();
    engine
        .set_cell_formula("Sheet1", "C3", r#"=INDIRECT("R[-1]C[-1]",FALSE)"#)
        .unwrap();

    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "C3"), Value::Number(42.0));
    assert_eq!(
        engine.precedents("Sheet1", "C3").unwrap(),
        vec![PrecedentNode::Cell {
            sheet: 0,
            addr: CellAddr { row: 1, col: 1 } // B2
        }]
    );
}

#[test]
fn indirect_rejects_external_workbook_refs_even_with_external_provider() {
    struct CountingExternalProvider {
        calls: AtomicUsize,
    }

    impl CountingExternalProvider {
        fn calls(&self) -> usize {
            self.calls.load(Ordering::SeqCst)
        }
    }

    impl ExternalValueProvider for CountingExternalProvider {
        fn get(&self, _sheet: &str, _addr: CellAddr) -> Option<Value> {
            self.calls.fetch_add(1, Ordering::SeqCst);
            Some(Value::Number(999.0))
        }
    }

    let provider = Arc::new(CountingExternalProvider {
        calls: AtomicUsize::new(0),
    });

    let mut engine = Engine::new();
    engine.set_external_value_provider(Some(provider.clone()));
    engine
        .set_cell_formula("Sheet1", "A1", r#"=INDIRECT("[Book.xlsx]Sheet1!A1")"#)
        .unwrap();
    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Error(ErrorKind::Ref));
    assert_eq!(provider.calls(), 0, "INDIRECT should not query external providers");
    assert!(engine.precedents("Sheet1", "A1").unwrap().is_empty());
}

