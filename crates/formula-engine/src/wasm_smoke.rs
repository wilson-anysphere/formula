use crate::{Engine, Value};

/// Compile-only smoke test for `wasm32-unknown-unknown`.
///
/// `cargo check -p formula-engine --target wasm32-unknown-unknown` should typecheck
/// the core Engine APIs used by the web worker calc backend.
#[allow(dead_code)]
pub(crate) fn wasm_compile_smoke() {
    let mut engine = Engine::new();
    let _ = engine.set_cell_value("Sheet1", "A1", 41.0);
    let _ = engine.set_cell_formula("Sheet1", "A2", "=A1+1");
    engine.recalculate_single_threaded();

    let _ = engine.get_cell_value("Sheet1", "A2");
    let _ = Value::Blank;

    engine.recalculate();
}
