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
fn indirect_external_workbook_refs_resolve_via_provider_with_bytecode() {
    struct CountingExternalProvider {
        calls: AtomicUsize,
    }

    impl CountingExternalProvider {
        fn calls(&self) -> usize {
            self.calls.load(Ordering::SeqCst)
        }
    }

    impl ExternalValueProvider for CountingExternalProvider {
        fn get(&self, sheet: &str, addr: CellAddr) -> Option<Value> {
            self.calls.fetch_add(1, Ordering::SeqCst);
            assert_eq!(sheet, "[Book.xlsx]Sheet1");
            assert_eq!(addr, CellAddr { row: 0, col: 0 });
            Some(Value::Number(999.0))
        }
    }

    let provider = Arc::new(CountingExternalProvider {
        calls: AtomicUsize::new(0),
    });

    let mut engine = Engine::new();
    engine.set_external_value_provider(Some(provider.clone()));
    engine.set_bytecode_enabled(true);
    engine
        .set_cell_formula("Sheet1", "A1", r#"=INDIRECT("[Book.xlsx]Sheet1!A1")"#)
        .unwrap();
    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(999.0));
    assert!(
        provider.calls() > 0,
        "expected INDIRECT to consult the external provider when dereferencing external workbook refs"
    );
    assert_eq!(
        engine.precedents("Sheet1", "A1").unwrap(),
        vec![PrecedentNode::ExternalCell {
            sheet: "[Book.xlsx]Sheet1".to_string(),
            addr: CellAddr { row: 0, col: 0 },
        }]
    );
}

#[test]
fn indirect_path_qualified_external_workbook_refs_resolve_via_provider_with_bytecode() {
    struct CountingExternalProvider {
        calls: AtomicUsize,
    }

    impl CountingExternalProvider {
        fn calls(&self) -> usize {
            self.calls.load(Ordering::SeqCst)
        }
    }

    impl ExternalValueProvider for CountingExternalProvider {
        fn get(&self, sheet: &str, addr: CellAddr) -> Option<Value> {
            self.calls.fetch_add(1, Ordering::SeqCst);
            assert_eq!(sheet, "[C:\\path\\Book.xlsx]Sheet1");
            assert_eq!(addr, CellAddr { row: 0, col: 0 });
            Some(Value::Number(999.0))
        }
    }

    let provider = Arc::new(CountingExternalProvider {
        calls: AtomicUsize::new(0),
    });

    let mut engine = Engine::new();
    engine.set_external_value_provider(Some(provider.clone()));
    engine.set_bytecode_enabled(true);
    engine
        .set_cell_formula(
            "Sheet1",
            "A1",
            // External workbook reference with a path-qualified workbook.
            r#"=INDIRECT("'C:\path\[Book.xlsx]Sheet1'!A1")"#,
        )
        .unwrap();

    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(999.0));
    assert!(
        provider.calls() > 0,
        "expected INDIRECT to consult the external provider when dereferencing external workbook refs"
    );
    assert_eq!(
        engine.precedents("Sheet1", "A1").unwrap(),
        vec![PrecedentNode::ExternalCell {
            sheet: "[C:\\path\\Book.xlsx]Sheet1".to_string(),
            addr: CellAddr { row: 0, col: 0 },
        }]
    );
}

#[test]
fn indirect_external_workbook_refs_are_ref_error_without_bytecode() {
    struct CountingExternalProvider {
        calls: AtomicUsize,
    }

    impl CountingExternalProvider {
        fn calls(&self) -> usize {
            self.calls.load(Ordering::SeqCst)
        }
    }

    impl ExternalValueProvider for CountingExternalProvider {
        fn get(&self, sheet: &str, addr: CellAddr) -> Option<Value> {
            self.calls.fetch_add(1, Ordering::SeqCst);
            assert_eq!(sheet, "[Book.xlsx]Sheet1");
            assert_eq!(addr, CellAddr { row: 0, col: 0 });
            Some(Value::Number(999.0))
        }
    }

    let provider = Arc::new(CountingExternalProvider {
        calls: AtomicUsize::new(0),
    });

    let mut engine = Engine::new();
    engine.set_bytecode_enabled(false);
    engine.set_external_value_provider(Some(provider.clone()));
    engine
        .set_cell_formula("Sheet1", "A1", r#"=INDIRECT("[Book.xlsx]Sheet1!A1")"#)
        .unwrap();

    engine.recalculate();

    assert_eq!(
        engine.get_cell_value("Sheet1", "A1"),
        Value::Error(ErrorKind::Ref)
    );
    assert_eq!(
        provider.calls(),
        0,
        "expected INDIRECT to reject external workbook refs without consulting the provider"
    );
    assert!(engine.precedents("Sheet1", "A1").unwrap().is_empty());
}

#[test]
fn indirect_dynamic_external_workbook_refs_are_ref_error_without_bytecode() {
    struct CountingExternalProvider {
        calls: AtomicUsize,
    }

    impl CountingExternalProvider {
        fn calls(&self) -> usize {
            self.calls.load(Ordering::SeqCst)
        }
    }

    impl ExternalValueProvider for CountingExternalProvider {
        fn get(&self, sheet: &str, addr: CellAddr) -> Option<Value> {
            self.calls.fetch_add(1, Ordering::SeqCst);
            assert_eq!(sheet, "[Book.xlsx]Sheet1");
            assert_eq!(addr, CellAddr { row: 0, col: 0 });
            Some(Value::Number(999.0))
        }
    }

    let provider = Arc::new(CountingExternalProvider {
        calls: AtomicUsize::new(0),
    });

    let mut engine = Engine::new();
    engine.set_bytecode_enabled(false);
    engine.set_external_value_provider(Some(provider.clone()));
    engine
        .set_cell_value("Sheet1", "B1", "[Book.xlsx]Sheet1!A1")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A1", "=INDIRECT(B1)")
        .unwrap();

    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Error(ErrorKind::Ref));
    assert_eq!(
        provider.calls(),
        0,
        "expected INDIRECT to reject external workbook refs without consulting the provider"
    );
    // The dynamic ref text is sourced from `B1`, so that cell is a static precedent even though
    // the external workbook reference is rejected.
    assert_eq!(
        engine.precedents("Sheet1", "A1").unwrap(),
        vec![PrecedentNode::Cell {
            sheet: 0,
            addr: CellAddr { row: 0, col: 1 } // B1
        }]
    );
}

#[test]
fn indirect_external_workbook_refs_are_ref_error_in_r1c1_mode_without_bytecode() {
    struct CountingExternalProvider {
        calls: AtomicUsize,
    }

    impl CountingExternalProvider {
        fn calls(&self) -> usize {
            self.calls.load(Ordering::SeqCst)
        }
    }

    impl ExternalValueProvider for CountingExternalProvider {
        fn get(&self, sheet: &str, addr: CellAddr) -> Option<Value> {
            self.calls.fetch_add(1, Ordering::SeqCst);
            assert_eq!(sheet, "[Book.xlsx]Sheet1");
            assert_eq!(addr, CellAddr { row: 0, col: 0 });
            Some(Value::Number(999.0))
        }
    }

    let provider = Arc::new(CountingExternalProvider {
        calls: AtomicUsize::new(0),
    });

    let mut engine = Engine::new();
    engine.set_bytecode_enabled(false);
    engine.set_external_value_provider(Some(provider.clone()));
    engine
        .set_cell_formula(
            "Sheet1",
            "A1",
            r#"=INDIRECT("[Book.xlsx]Sheet1!R1C1",FALSE)"#,
        )
        .unwrap();

    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Error(ErrorKind::Ref));
    assert_eq!(
        provider.calls(),
        0,
        "expected INDIRECT to reject external workbook refs without consulting the provider"
    );
    assert!(engine.precedents("Sheet1", "A1").unwrap().is_empty());
}
