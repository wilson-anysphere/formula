use formula_engine::eval::{
    parse_a1, EvalContext, Evaluator, RecalcContext, SheetReference, ValueResolver,
};
use formula_engine::date::{ymd_to_serial, ExcelDate, ExcelDateSystem};
use formula_engine::{Engine, ErrorKind, ExternalValueProvider, Value};
use proptest::prelude::*;
use std::sync::Arc;

fn cell_addr_to_a1(addr: formula_engine::eval::CellAddr) -> String {
    fn col_to_name(mut col: u32) -> String {
        col += 1;
        let mut out = Vec::<u8>::new();
        while col > 0 {
            let rem = (col - 1) % 26;
            out.push(b'A' + rem as u8);
            col = (col - 1) / 26;
        }
        out.reverse();
        String::from_utf8(out).expect("column letters are ASCII")
    }

    format!("{}{}", col_to_name(addr.col), addr.row + 1)
}

struct EngineResolver<'a> {
    engine: &'a Engine,
}

impl ValueResolver for EngineResolver<'_> {
    fn sheet_exists(&self, sheet_id: usize) -> bool {
        sheet_id == 0
    }

    fn get_cell_value(&self, sheet_id: usize, addr: formula_engine::eval::CellAddr) -> Value {
        let sheet = match sheet_id {
            0 => "Sheet1",
            _ => return Value::Blank,
        };
        self.engine.get_cell_value(sheet, &cell_addr_to_a1(addr))
    }

    fn resolve_structured_ref(
        &self,
        _ctx: EvalContext,
        _sref: &formula_engine::structured_refs::StructuredRef,
    ) -> Option<Vec<(usize, formula_engine::eval::CellAddr, formula_engine::eval::CellAddr)>> {
        None
    }
}

fn eval_via_ast(engine: &Engine, formula: &str, current_cell: &str) -> Value {
    let resolver = EngineResolver { engine };
    let recalc_ctx = RecalcContext::new(0);

    let parsed = formula_engine::eval::Parser::parse(formula).unwrap();
    let compiled = {
        let mut map = |sref: &SheetReference<String>| match sref {
            SheetReference::Current => SheetReference::Current,
            SheetReference::Sheet(_name) => SheetReference::Sheet(0),
            SheetReference::SheetRange(_start, _end) => SheetReference::SheetRange(0, 0),
            SheetReference::External(wb) => SheetReference::External(wb.clone()),
        };
        parsed.map_sheets(&mut map)
    };

    let ctx = EvalContext {
        current_sheet: 0,
        current_cell: parse_a1(current_cell).unwrap(),
    };
    Evaluator::new_with_date_system_and_locale(
        &resolver,
        ctx,
        &recalc_ctx,
        engine.date_system(),
        engine.value_locale(),
    )
    .eval_formula(&compiled)
}

fn assert_engine_matches_ast(engine: &Engine, formula: &str, cell: &str) {
    let expected = eval_via_ast(engine, formula, cell);
    assert_eq!(engine.get_cell_value("Sheet1", cell), expected);
}

#[test]
fn bytecode_backend_matches_ast_for_sum_and_countif() {
    let mut engine = Engine::new();

    for row in 1..=1000 {
        engine
            .set_cell_value("Sheet1", &format!("A{row}"), row as f64)
            .unwrap();
    }

    engine
        .set_cell_formula("Sheet1", "B1", "=SUM(A1:A1001)")
        .unwrap();
    engine
        // Include a trailing blank in the range and ensure COUNTIF's numeric comparisons treat
        // blanks as zero.
        .set_cell_formula("Sheet1", "B2", "=COUNTIF(A1:A1001, \"<1\")")
        .unwrap();

    engine.recalculate_single_threaded();

    assert_engine_matches_ast(&engine, "=SUM(A1:A1001)", "B1");
    assert_engine_matches_ast(&engine, "=COUNTIF(A1:A1001, \"<1\")", "B2");

    // SUM + COUNTIF should both be compiled to bytecode.
    assert_eq!(engine.bytecode_program_count(), 2);
}

#[test]
fn bytecode_backend_matches_ast_for_countif_grouped_numeric_criteria() {
    let mut engine = Engine::new();
    engine
        .set_cell_value("Sheet1", "A1", 1000.0)
        .expect("set A1");
    engine
        .set_cell_value("Sheet1", "A2", 999.0)
        .expect("set A2");

    engine
        .set_cell_formula("Sheet1", "B1", "=COUNTIF(A1:A2, \"1,000\")")
        .unwrap();
    engine.recalculate_single_threaded();

    assert_engine_matches_ast(&engine, "=COUNTIF(A1:A2, \"1,000\")", "B1");
    assert_eq!(engine.bytecode_program_count(), 1);
}

#[test]
fn bytecode_cache_reuses_filled_formula_patterns_in_engine() {
    let mut engine = Engine::new();

    engine.set_cell_formula("Sheet1", "C1", "=A1+B1").unwrap();
    engine.set_cell_formula("Sheet1", "C2", "=A2+B2").unwrap();
    engine.set_cell_formula("Sheet1", "C3", "=A3+B3").unwrap();

    assert_eq!(engine.bytecode_program_count(), 1);
}

#[test]
fn bytecode_backend_respects_external_value_provider() {
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
    engine
        .set_cell_formula("Sheet1", "B1", "=SUM(A1:A3)")
        .unwrap();

    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(6.0));
    assert!(engine.bytecode_program_count() > 0);
}

#[test]
fn sumproduct_coerces_bools_in_ranges() {
    let mut engine = Engine::new();

    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", true).unwrap();
    engine.set_cell_value("Sheet1", "A3", 3.0).unwrap();

    engine
        .set_cell_formula("Sheet1", "B1", "=SUMPRODUCT(A1:A3, A1:A3)")
        .unwrap();
    engine.recalculate_single_threaded();

    // [1, TRUE, 3] is coerced to [1, 1, 3] for SUMPRODUCT.
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(11.0));
    assert!(engine.bytecode_program_count() > 0);
}

#[test]
fn bytecode_backend_matches_ast_for_scalar_math_and_comparisons() {
    let mut engine = Engine::new();

    engine.set_cell_value("Sheet1", "A1", -1.5).unwrap();
    engine.set_cell_value("Sheet1", "A2", 1.9).unwrap();
    engine.set_cell_value("Sheet1", "A3", 3.0).unwrap();

    // Scalar-only math.
    engine.set_cell_formula("Sheet1", "B1", "=ABS(A1)").unwrap();
    engine.set_cell_formula("Sheet1", "B2", "=INT(A2)").unwrap();
    engine.set_cell_formula("Sheet1", "B3", "=ROUND(A2, 0)").unwrap();
    engine.set_cell_formula("Sheet1", "B4", "=ROUNDUP(A1, 0)").unwrap();
    engine.set_cell_formula("Sheet1", "B5", "=ROUNDDOWN(A1, 0)").unwrap();
    engine.set_cell_formula("Sheet1", "B6", "=MOD(7, 4)").unwrap();
    engine.set_cell_formula("Sheet1", "B7", "=SIGN(A1)").unwrap();

    // CONCAT (scalar-only fast path).
    engine
        .set_cell_formula("Sheet1", "B8", "=CONCAT(\"foo\", A3, TRUE)")
        .unwrap();

    // Pow + comparisons (new bytecode ops).
    engine.set_cell_formula("Sheet1", "C1", "=2^3").unwrap();
    engine.set_cell_formula("Sheet1", "C2", "=\"a\"=\"A\"").unwrap();
    engine.set_cell_formula("Sheet1", "C3", "=(-1)^0.5").unwrap();

    assert_eq!(engine.bytecode_program_count(), 11);

    engine.recalculate_single_threaded();

    for (formula, cell) in [
        ("=ABS(A1)", "B1"),
        ("=INT(A2)", "B2"),
        ("=ROUND(A2, 0)", "B3"),
        ("=ROUNDUP(A1, 0)", "B4"),
        ("=ROUNDDOWN(A1, 0)", "B5"),
        ("=MOD(7, 4)", "B6"),
        ("=SIGN(A1)", "B7"),
        ("=CONCAT(\"foo\", A3, TRUE)", "B8"),
        ("=2^3", "C1"),
        ("=\"a\"=\"A\"", "C2"),
        ("=(-1)^0.5", "C3"),
    ] {
        assert_engine_matches_ast(&engine, formula, cell);
    }
}

proptest! {
    #![proptest_config(ProptestConfig { cases: 32, .. ProptestConfig::default() })]
    #[test]
    fn bytecode_backend_matches_ast_for_random_supported_formulas(
        a in -1000f64..1000f64,
        b in -1000f64..1000f64,
        digits in -6i32..6i32,
        choice in 0u8..12u8,
    ) {
        let formula = match choice {
            0 => "=A1+B1".to_string(),
            1 => "=A1-B1".to_string(),
            2 => "=A1*B1".to_string(),
            3 => "=A1/B1".to_string(),
            4 => "=A1^B1".to_string(),
            5 => "=A1=B1".to_string(),
            6 => "=A1<>B1".to_string(),
            7 => "=A1<B1".to_string(),
            8 => "=ABS(A1)".to_string(),
            9 => format!("=ROUND(A1, {digits})"),
            10 => "=MOD(A1, B1)".to_string(),
            11 => "=SIGN(A1)".to_string(),
            _ => unreachable!(),
        };

        let mut engine = Engine::new();
        engine.set_cell_value("Sheet1", "A1", a).unwrap();
        engine.set_cell_value("Sheet1", "B1", b).unwrap();
        engine.set_cell_formula("Sheet1", "C1", &formula).unwrap();

        // Ensure we're exercising the bytecode path.
        prop_assert_eq!(engine.bytecode_program_count(), 1);

        engine.recalculate_single_threaded();
        let expected = eval_via_ast(&engine, &formula, "C1");
        prop_assert_eq!(engine.get_cell_value("Sheet1", "C1"), expected);
    }
}

#[test]
fn bytecode_backend_matches_ast_for_countif_full_criteria_semantics() {
    let mut engine = Engine::new();

    // Text + wildcards.
    engine.set_cell_value("Sheet1", "A1", "apple").unwrap();
    engine.set_cell_value("Sheet1", "A2", "apricot").unwrap();
    engine.set_cell_value("Sheet1", "A3", "banana").unwrap();
    engine.set_cell_value("Sheet1", "A4", "*").unwrap();
    engine.set_cell_value("Sheet1", "A5", "ab").unwrap();
    engine.set_cell_value("Sheet1", "A6", "abc").unwrap();
    engine.set_cell_value("Sheet1", "A7", "").unwrap(); // empty string counts as blank
    // A8 left blank
    engine
        .set_cell_value("Sheet1", "A9", Value::Error(ErrorKind::Div0))
        .unwrap();
    engine.set_cell_value("Sheet1", "A10", true).unwrap();

    // Date criteria strings are parsed in the default (1900) date system.
    let d1 = ymd_to_serial(ExcelDate::new(2020, 1, 1), ExcelDateSystem::EXCEL_1900).unwrap();
    let d2 = ymd_to_serial(ExcelDate::new(2020, 1, 2), ExcelDateSystem::EXCEL_1900).unwrap();
    let d3 = ymd_to_serial(ExcelDate::new(2020, 1, 3), ExcelDateSystem::EXCEL_1900).unwrap();
    engine.set_cell_value("Sheet1", "B1", d1 as f64).unwrap();
    engine.set_cell_value("Sheet1", "B2", d2 as f64).unwrap();
    engine.set_cell_value("Sheet1", "B3", d3 as f64).unwrap();

    engine
        .set_cell_formula("Sheet1", "C1", "=COUNTIF(A1:A10, D1)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "C2", "=COUNTIF(B1:B3, D2)")
        .unwrap();

    // Ensure we're actually exercising the bytecode COUNTIF implementation (criteria args are
    // cell references, not numeric literals).
    assert_eq!(engine.bytecode_program_count(), 2);

    for (criteria, expected) in [
        (Value::Text("ap*".to_string()), Value::Number(2.0)),
        (Value::Text("~*".to_string()), Value::Number(1.0)),
        (Value::Text("??".to_string()), Value::Number(1.0)),
        (Value::Text("".to_string()), Value::Number(2.0)),  // blanks
        (Value::Text("<>".to_string()), Value::Number(7.0)), // non-blanks (errors excluded)
        (Value::Text("#DIV/0!".to_string()), Value::Number(1.0)),
        (Value::Bool(true), Value::Number(1.0)),
        (Value::Error(ErrorKind::Div0), Value::Error(ErrorKind::Div0)),
    ] {
        engine.set_cell_value("Sheet1", "D1", criteria).unwrap();
        engine.recalculate_single_threaded();

        assert_eq!(engine.get_cell_value("Sheet1", "C1"), expected);
        assert_engine_matches_ast(&engine, "=COUNTIF(A1:A10, D1)", "C1");
    }

    engine
        .set_cell_value("Sheet1", "D2", Value::Text(">=1/2/2020".to_string()))
        .unwrap();
    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Sheet1", "C2"), Value::Number(2.0));
    assert_engine_matches_ast(&engine, "=COUNTIF(B1:B3, D2)", "C2");
}

#[test]
fn bytecode_backend_countif_date_criteria_respects_engine_date_system() {
    let mut engine = Engine::new();
    engine.set_date_system(ExcelDateSystem::Excel1904);

    let system = ExcelDateSystem::Excel1904;
    let d2019 = ymd_to_serial(ExcelDate::new(2019, 12, 31), system).unwrap();
    let d2020 = ymd_to_serial(ExcelDate::new(2020, 1, 1), system).unwrap();
    let d2020_next = ymd_to_serial(ExcelDate::new(2020, 1, 2), system).unwrap();

    engine.set_cell_value("Sheet1", "A1", d2019 as f64).unwrap();
    engine.set_cell_value("Sheet1", "A2", d2020 as f64).unwrap();
    engine
        .set_cell_value("Sheet1", "A3", d2020_next as f64)
        .unwrap();

    engine
        .set_cell_formula("Sheet1", "B1", r#"=COUNTIF(A1:A3, ">1/1/2020")"#)
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B2", r#"=COUNTIF(A1:A3, "=1/1/2020")"#)
        .unwrap();

    // Ensure we're exercising the bytecode path.
    assert_eq!(engine.bytecode_program_count(), 2);

    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "B2"), Value::Number(1.0));
    assert_engine_matches_ast(&engine, r#"=COUNTIF(A1:A3, ">1/1/2020")"#, "B1");
    assert_engine_matches_ast(&engine, r#"=COUNTIF(A1:A3, "=1/1/2020")"#, "B2");
}
