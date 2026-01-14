use formula_engine::eval::CellAddr;
use formula_engine::{Engine, ErrorKind, ExternalValueProvider, Value};
use std::sync::atomic::{AtomicUsize, Ordering};
use std::sync::Arc;

struct Provider {
    calls: AtomicUsize,
}

impl Provider {
    fn new() -> Self {
        Self {
            calls: AtomicUsize::new(0),
        }
    }

    fn calls(&self) -> usize {
        self.calls.load(Ordering::SeqCst)
    }
}

impl ExternalValueProvider for Provider {
    fn get(&self, sheet: &str, addr: CellAddr) -> Option<Value> {
        self.calls.fetch_add(1, Ordering::SeqCst);
        if sheet == "[Book.xlsx]Sheet1" && addr.row == 0 && addr.col == 0 {
            return Some(Value::Number(41.0));
        }
        None
    }
}

#[test]
fn bytecode_external_cell_ref_evaluates_via_provider() {
    let mut engine = Engine::new();
    let provider = Arc::new(Provider::new());
    engine.set_external_value_provider(Some(provider));
    engine.set_bytecode_enabled(true);
    engine
        .set_cell_formula("Sheet1", "A1", "=[Book.xlsx]Sheet1!A1+1")
        .unwrap();

    // Ensure we compile to bytecode (no AST fallback).
    assert_eq!(
        engine.bytecode_program_count(),
        1,
        "expected external workbook refs to compile to bytecode (stats={:?}, report={:?})",
        engine.bytecode_compile_stats(),
        engine.bytecode_compile_report(32)
    );

    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(42.0));
}

#[test]
fn bytecode_external_cell_ref_with_path_qualified_workbook_evaluates_via_provider() {
    struct Provider;

    impl ExternalValueProvider for Provider {
        fn get(&self, sheet: &str, addr: CellAddr) -> Option<Value> {
            assert_eq!(sheet, "[C:\\path\\Book.xlsx]Sheet1");
            assert_eq!(addr, CellAddr { row: 0, col: 0 });
            Some(Value::Number(41.0))
        }
    }

    let mut engine = Engine::new();
    engine.set_external_value_provider(Some(Arc::new(Provider)));
    engine.set_bytecode_enabled(true);
    engine
        .set_cell_formula("Sheet1", "A1", r#"='C:\path\[Book.xlsx]Sheet1'!A1+1"#)
        .unwrap();

    // Ensure we compile to bytecode (no AST fallback).
    assert_eq!(
        engine.bytecode_program_count(),
        1,
        "expected path-qualified external workbook refs to compile to bytecode (stats={:?}, report={:?})",
        engine.bytecode_compile_stats(),
        engine.bytecode_compile_report(32)
    );

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

#[test]
fn bytecode_indirect_external_cell_ref_is_ref_error() {
    let mut engine = Engine::new();
    let provider = Arc::new(Provider::new());
    engine.set_external_value_provider(Some(provider.clone()));
    engine.set_bytecode_enabled(true);
    engine
        .set_cell_formula("Sheet1", "A1", r#"=INDIRECT("[Book.xlsx]Sheet1!A1")+1"#)
        .unwrap();

    // INDIRECT rejects external workbook references; ensure bytecode compilation falls back so
    // diagnostics are consistent.
    assert_eq!(
        engine.bytecode_program_count(),
        0,
        "expected INDIRECT external workbook refs to fall back from bytecode (stats={:?}, report={:?})",
        engine.bytecode_compile_stats(),
        engine.bytecode_compile_report(32)
    );

    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Sheet1", "A1"),
        Value::Error(ErrorKind::Ref)
    );
    assert_eq!(
        provider.calls(),
        0,
        "bytecode INDIRECT should not consult the external provider for unsupported external workbook refs"
    );
}
