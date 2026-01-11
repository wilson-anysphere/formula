use formula_engine::date::ExcelDateSystem;
use formula_engine::{Engine, Value};

pub struct TestSheet {
    engine: Engine,
    sheet: &'static str,
    scratch_cell: &'static str,
}

impl TestSheet {
    pub fn new() -> Self {
        Self {
            engine: Engine::new(),
            sheet: "Sheet1",
            scratch_cell: "Z1",
        }
    }

    pub fn set_date_system(&mut self, system: ExcelDateSystem) {
        self.engine.set_date_system(system);
    }

    pub fn set(&mut self, addr: &str, value: impl Into<Value>) {
        self.engine
            .set_cell_value(self.sheet, addr, value)
            .expect("set cell value");
    }

    pub fn set_formula(&mut self, addr: &str, formula: &str) {
        self.engine
            .set_cell_formula(self.sheet, addr, formula)
            .expect("set cell formula");
    }

    pub fn recalc(&mut self) {
        self.engine.recalculate();
    }

    pub fn get(&self, addr: &str) -> Value {
        self.engine.get_cell_value(self.sheet, addr)
    }

    pub fn eval(&mut self, formula: &str) -> Value {
        self.set_formula(self.scratch_cell, formula);
        self.recalc();
        self.get(self.scratch_cell)
    }
}

pub fn assert_number(value: &Value, expected: f64) {
    match value {
        Value::Number(n) => {
            assert!(
                (*n - expected).abs() < 1e-9,
                "expected {expected}, got {n}"
            );
        }
        other => panic!("expected number {expected}, got {other:?}"),
    }
}
