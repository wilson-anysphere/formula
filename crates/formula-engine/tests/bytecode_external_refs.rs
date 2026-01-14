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
fn bytecode_indirect_external_cell_ref_compiles_and_evaluates_via_provider() {
    let mut engine = Engine::new();
    let provider = Arc::new(Provider::new());
    engine.set_external_value_provider(Some(provider.clone()));
    engine.set_bytecode_enabled(true);
    engine
        .set_cell_formula("Sheet1", "A1", r#"=INDIRECT("[Book.xlsx]Sheet1!A1")"#)
        .unwrap();

    // Ensure we compile to bytecode (no AST fallback).
    assert_eq!(
        engine.bytecode_program_count(),
        1,
        "expected INDIRECT external workbook refs to compile to bytecode (stats={:?}, report={:?})",
        engine.bytecode_compile_stats(),
        engine.bytecode_compile_report(32)
    );

    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(41.0));
    assert!(
        provider.calls() > 0,
        "expected INDIRECT to consult the external provider when dereferencing external workbook refs"
    );
}

#[test]
fn bytecode_indirect_dynamic_external_cell_ref_compiles_and_evaluates_via_provider() {
    let mut engine = Engine::new();
    let provider = Arc::new(Provider::new());
    engine.set_external_value_provider(Some(provider.clone()));
    engine.set_bytecode_enabled(true);

    engine
        .set_cell_value("Sheet1", "B1", "[Book.xlsx]Sheet1!A1")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A1", "=INDIRECT(B1)+1")
        .unwrap();

    // The external workbook reference is produced at runtime (from `B1`), so the bytecode backend
    // must support dynamic external dereferencing without requiring AST fallback.
    assert_eq!(
        engine.bytecode_program_count(),
        1,
        "expected dynamic INDIRECT formulas to compile to bytecode (stats={:?}, report={:?})",
        engine.bytecode_compile_stats(),
        engine.bytecode_compile_report(32)
    );

    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(42.0));
    assert!(
        provider.calls() > 0,
        "expected INDIRECT to consult the external provider when dereferencing external workbook refs"
    );
}

#[test]
fn bytecode_indirect_dynamic_external_range_ref_compiles_and_evaluates_via_provider() {
    struct Provider {
        calls: AtomicUsize,
    }

    impl Provider {
        fn calls(&self) -> usize {
            self.calls.load(Ordering::SeqCst)
        }
    }

    impl ExternalValueProvider for Provider {
        fn get(&self, sheet: &str, addr: CellAddr) -> Option<Value> {
            self.calls.fetch_add(1, Ordering::SeqCst);
            if sheet != "[Book.xlsx]Sheet1" || addr.col != 0 {
                return None;
            }
            match addr.row {
                0 => Some(Value::Number(1.0)),
                1 => Some(Value::Number(2.0)),
                _ => None,
            }
        }
    }

    let mut engine = Engine::new();
    let provider = Arc::new(Provider {
        calls: AtomicUsize::new(0),
    });
    engine.set_external_value_provider(Some(provider.clone()));
    engine.set_bytecode_enabled(true);

    engine
        .set_cell_value("Sheet1", "B1", "[Book.xlsx]Sheet1!A1:A2")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A1", "=SUM(INDIRECT(B1))")
        .unwrap();

    // Ensure we compile to bytecode (no AST fallback).
    assert_eq!(
        engine.bytecode_program_count(),
        1,
        "expected dynamic INDIRECT formulas to compile to bytecode (stats={:?}, report={:?})",
        engine.bytecode_compile_stats(),
        engine.bytecode_compile_report(32)
    );

    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(3.0));
    assert!(
        provider.calls() > 0,
        "expected INDIRECT to consult the external provider when dereferencing external workbook refs"
    );
}

#[test]
fn bytecode_sum_over_external_range_compiles_and_uses_reference_semantics() {
    // Excel quirk: SUM over references ignores logicals/text stored in cells.
    struct Provider;

    impl ExternalValueProvider for Provider {
        fn get(&self, sheet: &str, addr: CellAddr) -> Option<Value> {
            if sheet != "[Book.xlsx]Sheet1" || addr.col != 0 {
                return None;
            }
            match addr.row {
                0 => Some(Value::Number(1.0)),
                1 => Some(Value::Text("2".to_string())),
                2 => Some(Value::Bool(true)),
                _ => None,
            }
        }
    }

    let mut engine = Engine::new();
    engine.set_external_value_provider(Some(Arc::new(Provider)));
    engine.set_bytecode_enabled(true);
    engine
        .set_cell_formula("Sheet1", "A1", "=SUM([Book.xlsx]Sheet1!A1:A3)")
        .unwrap();

    // Ensure we compile to bytecode (no AST fallback).
    assert_eq!(
        engine.bytecode_program_count(),
        1,
        "expected external workbook range refs to compile to bytecode (stats={:?}, report={:?})",
        engine.bytecode_compile_stats(),
        engine.bytecode_compile_report(32)
    );

    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(1.0));
}

#[test]
fn bytecode_sum_over_external_reference_union_uses_provider_tab_order_for_error_precedence() {
    // When bytecode evaluates a multi-area reference union spanning external sheets, the
    // deterministic area ordering should follow the external workbook tab order (when supplied by
    // the provider), not lexicographic sheet name order.
    struct Provider;

    impl ExternalValueProvider for Provider {
        fn get(&self, sheet: &str, addr: CellAddr) -> Option<Value> {
            if addr != (CellAddr { row: 0, col: 0 }) {
                return None;
            }
            match sheet {
                "[Book.xlsx]Sheet2" => Some(Value::Error(ErrorKind::Ref)),
                "[Book.xlsx]Sheet10" => Some(Value::Error(ErrorKind::Div0)),
                _ => None,
            }
        }

        fn sheet_order(&self, workbook: &str) -> Option<Vec<String>> {
            (workbook == "Book.xlsx").then(|| vec!["Sheet2".to_string(), "Sheet10".to_string()])
        }
    }

    let mut engine = Engine::new();
    engine.set_external_value_provider(Some(Arc::new(Provider)));
    engine.set_bytecode_enabled(true);
    engine
        .set_cell_formula(
            "Sheet1",
            "A1",
            "=SUM(([Book.xlsx]Sheet2!A1,[Book.xlsx]Sheet10!A1))",
        )
        .unwrap();

    assert_eq!(
        engine.bytecode_program_count(),
        1,
        "expected union over external refs to compile to bytecode (stats={:?}, report={:?})",
        engine.bytecode_compile_stats(),
        engine.bytecode_compile_report(32)
    );

    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Sheet1", "A1"),
        Value::Error(ErrorKind::Ref)
    );
}

#[test]
fn bytecode_degenerate_external_3d_sheet_range_matches_endpoints_nfkc_case_insensitively() {
    struct NfkcProvider;

    impl ExternalValueProvider for NfkcProvider {
        fn get(&self, sheet: &str, addr: CellAddr) -> Option<Value> {
            if sheet == "[Book.xlsx]Kelvin" && addr.row == 0 && addr.col == 0 {
                return Some(Value::Number(41.0));
            }
            None
        }
    }

    let mut engine = Engine::new();
    engine.set_external_value_provider(Some(Arc::new(NfkcProvider)));
    engine.set_bytecode_enabled(true);
    engine
        .set_cell_formula("Sheet1", "A1", "=[Book.xlsx]'Kelvin':'KELVIN'!A1")
        .unwrap();

    // Ensure we compile to bytecode (no AST fallback).
    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(41.0));
}
