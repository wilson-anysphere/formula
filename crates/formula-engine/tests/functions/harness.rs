#![allow(dead_code)]

use formula_engine::date::ExcelDateSystem;
use formula_engine::locale::ValueLocaleConfig;
use formula_engine::pivot::PivotTable;
use formula_engine::{EditOp, EditResult, Engine, Value};
use formula_model::{Range, Style};

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

    pub fn set_value_locale(&mut self, locale: ValueLocaleConfig) {
        self.engine.set_value_locale(locale);
    }

    pub fn set_text_codepage(&mut self, codepage: u16) {
        self.engine.set_text_codepage(codepage);
    }

    pub fn set(&mut self, addr: &str, value: impl Into<Value>) {
        self.engine
            .set_cell_value(self.sheet, addr, value)
            .expect("set cell value");
    }

    pub fn set_on(&mut self, sheet: &str, addr: &str, value: impl Into<Value>) {
        self.engine
            .set_cell_value(sheet, addr, value)
            .expect("set cell value");
    }

    pub fn set_range_values(&mut self, range_a1: &str, values: &[Vec<Value>]) {
        let range = Range::from_a1(range_a1).expect("range");
        self.engine
            .set_range_values(self.sheet, range, values, true)
            .expect("set range values");
    }

    pub fn clear_cell(&mut self, addr: &str) {
        self.engine
            .clear_cell(self.sheet, addr)
            .expect("clear cell");
    }

    pub fn clear_range(&mut self, range_a1: &str) {
        let range = Range::from_a1(range_a1).expect("range");
        self.engine
            .clear_range(self.sheet, range, true)
            .expect("clear range");
    }

    pub fn set_phonetic(&mut self, addr: &str, phonetic: Option<&str>) {
        self.engine
            .set_cell_phonetic(self.sheet, addr, phonetic.map(|s| s.to_string()))
            .expect("set cell phonetic");
    }

    pub fn set_formula(&mut self, addr: &str, formula: &str) {
        self.engine
            .set_cell_formula(self.sheet, addr, formula)
            .expect("set cell formula");
    }

    pub fn set_formula_on(&mut self, sheet: &str, addr: &str, formula: &str) {
        self.engine
            .set_cell_formula(sheet, addr, formula)
            .expect("set cell formula");
    }

    pub fn register_pivot_table(&mut self, destination: Range, pivot: PivotTable) {
        self.engine
            .register_pivot_table(self.sheet, destination, pivot)
            .expect("register pivot table");
    }

    pub fn apply_operation(&mut self, op: EditOp) -> EditResult {
        self.engine.apply_operation(op).expect("apply operation")
    }

    pub fn set_default_col_width(&mut self, width: Option<f32>) {
        self.engine.set_sheet_default_col_width(self.sheet, width);
    }

    pub fn set_col_width(&mut self, col: u32, width: Option<f32>) {
        self.engine.set_col_width(self.sheet, col, width);
    }

    pub fn intern_style(&mut self, style: Style) -> u32 {
        self.engine.intern_style(style)
    }

    pub fn set_col_style_id(&mut self, col_0based: u32, style_id: Option<u32>) {
        self.engine
            .set_col_style_id(self.sheet, col_0based, style_id);
    }

    pub fn set_row_style_id(&mut self, row_0based: u32, style_id: Option<u32>) {
        self.engine
            .set_row_style_id(self.sheet, row_0based, style_id);
    }

    pub fn set_cell_style_id(&mut self, addr: &str, style_id: u32) {
        self.engine
            .set_cell_style_id(self.sheet, addr, style_id)
            .expect("set cell style id");
    }

    pub fn set_sheet_protection_enabled(&mut self, enabled: bool) {
        self.engine
            .set_sheet_protection_enabled(self.sheet, enabled);
    }

    pub fn recalc(&mut self) {
        // Use the single-threaded recalc path in tests to avoid initializing a global Rayon pool
        // (which can fail on shared CI/agent hosts due to OS thread limits).
        self.engine.recalculate_single_threaded();
    }

    pub fn recalculate(&mut self) {
        self.recalc();
    }

    pub fn get(&self, addr: &str) -> Value {
        self.engine.get_cell_value(self.sheet, addr)
    }

    pub fn get_on(&self, sheet: &str, addr: &str) -> Value {
        self.engine.get_cell_value(sheet, addr)
    }

    pub fn bytecode_program_count(&self) -> usize {
        self.engine.bytecode_program_count()
    }

    pub fn circular_reference_count(&self) -> usize {
        self.engine.circular_reference_count()
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
            assert!((*n - expected).abs() < 1e-9, "expected {expected}, got {n}");
        }
        other => panic!("expected number {expected}, got {other:?}"),
    }
}
