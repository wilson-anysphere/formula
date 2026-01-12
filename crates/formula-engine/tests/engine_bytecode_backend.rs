use formula_engine::date::{ymd_to_serial, ExcelDate, ExcelDateSystem};
use formula_engine::eval::{
    parse_a1, EvalContext, Evaluator, RecalcContext, SheetReference, ValueResolver,
};
use formula_engine::locale::ValueLocaleConfig;
use formula_engine::value::NumberLocale;
use formula_engine::{
    bytecode, BytecodeCompileReason, Engine, ErrorKind, ExternalValueProvider, NameDefinition,
    NameScope, ParseOptions, ReferenceStyle, Value,
};
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
    ) -> Option<
        Vec<(
            usize,
            formula_engine::eval::CellAddr,
            formula_engine::eval::CellAddr,
        )>,
    > {
        None
    }
}

fn eval_via_ast(engine: &Engine, formula: &str, current_cell: &str) -> Value {
    let resolver = EngineResolver { engine };
    let mut recalc_ctx = RecalcContext::new(0);
    let separators = engine.value_locale().separators;
    recalc_ctx.number_locale =
        NumberLocale::new(separators.decimal_sep, Some(separators.thousands_sep));

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

fn bytecode_value_to_engine(value: formula_engine::bytecode::Value) -> Value {
    use formula_engine::bytecode::{ErrorKind as ByteErrorKind, Value as ByteValue};
    match value {
        ByteValue::Number(n) => Value::Number(n),
        ByteValue::Bool(b) => Value::Bool(b),
        ByteValue::Text(s) => Value::Text(s.to_string()),
        ByteValue::Empty => Value::Blank,
        ByteValue::Error(e) => Value::Error(match e {
            ByteErrorKind::Null => ErrorKind::Null,
            ByteErrorKind::Div0 => ErrorKind::Div0,
            ByteErrorKind::Ref => ErrorKind::Ref,
            ByteErrorKind::Value => ErrorKind::Value,
            ByteErrorKind::Name => ErrorKind::Name,
            ByteErrorKind::Num => ErrorKind::Num,
            ByteErrorKind::NA => ErrorKind::NA,
            ByteErrorKind::Spill => ErrorKind::Spill,
            ByteErrorKind::Calc => ErrorKind::Calc,
        }),
        // Array/range values are not valid scalar results for the engine API; treat them as spills.
        ByteValue::Array(_) | ByteValue::Range(_) => Value::Error(ErrorKind::Spill),
    }
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
fn bytecode_backend_inlines_defined_name_static_ranges() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", 3.0).unwrap();
    engine
        .define_name(
            "MyRange",
            NameScope::Workbook,
            NameDefinition::Reference("Sheet1!$A$1:$A$3".to_string()),
        )
        .unwrap();

    engine
        .set_cell_formula("Sheet1", "B1", "=SUM(MyRange)")
        .unwrap();

    // Ensure the named range was inlined and compiled to bytecode.
    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();
    let via_bytecode = engine.get_cell_value("Sheet1", "B1");

    // Compare against the AST backend for the same workbook state.
    engine.set_bytecode_enabled(false);
    engine.recalculate_single_threaded();
    let via_ast = engine.get_cell_value("Sheet1", "B1");

    assert_eq!(via_bytecode, Value::Number(6.0));
    assert_eq!(via_bytecode, via_ast);
}

#[test]
fn bytecode_backend_does_not_inline_dynamic_defined_name_formulas() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine
        .define_name(
            "MyDyn",
            NameScope::Workbook,
            NameDefinition::Formula("=INDIRECT(\"A1\")".to_string()),
        )
        .unwrap();

    engine
        .set_cell_formula("Sheet1", "B1", "=SUM(MyDyn)")
        .unwrap();

    // The dynamic name definition should not be inlined, so the formula should stay on the AST backend.
    assert_eq!(engine.bytecode_program_count(), 0);

    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(1.0));
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
fn bytecode_backend_sumproduct_respects_engine_value_locale() {
    let mut engine = Engine::new();
    engine.set_value_locale(ValueLocaleConfig::de_de());

    // Text numeric values should be coerced using the workbook locale.
    engine.set_cell_value("Sheet1", "A1", "1,5").unwrap();
    engine.set_cell_value("Sheet1", "A2", 2.0).unwrap();

    engine.set_cell_value("Sheet1", "B1", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "B2", 2.0).unwrap();

    engine
        .set_cell_formula("Sheet1", "C1", "=SUMPRODUCT(A1:A2,B1:B2)")
        .unwrap();
    engine.recalculate_single_threaded();

    assert_eq!(engine.bytecode_program_count(), 1);
    assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::Number(7.0));
    assert_engine_matches_ast(&engine, "=SUMPRODUCT(A1:A2,B1:B2)", "C1");
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
    engine
        .set_cell_formula("Sheet1", "B3", "=ROUND(A2, 0)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B4", "=ROUNDUP(A1, 0)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B5", "=ROUNDDOWN(A1, 0)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B6", "=MOD(7, 4)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B7", "=SIGN(A1)")
        .unwrap();

    // CONCAT (scalar-only fast path).
    engine
        .set_cell_formula("Sheet1", "B8", "=CONCAT(\"foo\", A3, TRUE)")
        .unwrap();

    // Pow + comparisons (new bytecode ops).
    engine.set_cell_formula("Sheet1", "C1", "=2^3").unwrap();
    engine
        .set_cell_formula("Sheet1", "C2", "=\"a\"=\"A\"")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "C3", "=(-1)^0.5")
        .unwrap();

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

#[test]
fn bytecode_backend_matches_ast_for_concat_operator() {
    let mut engine = Engine::new();

    engine
        .set_cell_formula("Sheet1", "A1", "=\"a\"&\"b\"")
        .unwrap();

    // Ensure number-to-text coercion uses "General" formatting (scientific notation for large
    // magnitudes).
    engine
        .set_cell_value("Sheet1", "A2", 100000000000.0)
        .unwrap();
    engine.set_cell_value("Sheet1", "B2", "x").unwrap();
    engine
        .set_cell_formula("Sheet1", "C2", "=A2&B2")
        .unwrap();

    // Ensure we're exercising the bytecode path for both formulas.
    assert_eq!(engine.bytecode_program_count(), 2);

    engine.recalculate_single_threaded();

    assert_engine_matches_ast(&engine, "=\"a\"&\"b\"", "A1");
    assert_engine_matches_ast(&engine, "=A2&B2", "C2");
}

#[test]
fn bytecode_cache_reuses_filled_formula_patterns_for_concat_operator() {
    let mut engine = Engine::new();

    engine.set_cell_formula("Sheet1", "C1", "=A1&B1").unwrap();
    engine.set_cell_formula("Sheet1", "C2", "=A2&B2").unwrap();
    engine.set_cell_formula("Sheet1", "C3", "=A3&B3").unwrap();

    assert_eq!(engine.bytecode_program_count(), 1);
}

#[test]
fn bytecode_backend_matches_ast_for_postfix_percent_operator() {
    let mut engine = Engine::new();

    engine.set_cell_formula("Sheet1", "A1", "=10%").unwrap();

    engine.set_cell_value("Sheet1", "A2", 50.0).unwrap();
    engine.set_cell_formula("Sheet1", "B2", "=A2%").unwrap();

    // Ensure we're exercising the bytecode path.
    assert_eq!(engine.bytecode_program_count(), 2);

    engine.recalculate_single_threaded();

    assert_engine_matches_ast(&engine, "=10%", "A1");
    assert_engine_matches_ast(&engine, "=A2%", "B2");
}

#[test]
fn bytecode_backend_matches_ast_for_postfix_percent_with_locale_sensitive_numeric_text() {
    // en-US: comma thousands separator.
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", r#"="1,234"%"#)
        .unwrap();
    assert_eq!(engine.bytecode_program_count(), 1);
    engine.recalculate_single_threaded();
    assert_engine_matches_ast(&engine, r#"="1,234"%"#, "A1");

    // de-DE: '.' thousands separator, ',' decimal separator.
    let mut engine = Engine::new();
    engine.set_value_locale(ValueLocaleConfig::de_de());
    engine
        .set_cell_formula("Sheet1", "A1", r#"="1.234,56"%"#)
        .unwrap();
    assert_eq!(engine.bytecode_program_count(), 1);
    engine.recalculate_single_threaded();
    assert_engine_matches_ast(&engine, r#"="1.234,56"%"#, "A1");
}

#[test]
fn bytecode_backend_matches_ast_for_explicit_implicit_intersection_operator() {
    let mut engine = Engine::new();

    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", 3.0).unwrap();

    engine.set_cell_value("Sheet1", "B1", 10.0).unwrap();
    engine.set_cell_value("Sheet1", "C1", 30.0).unwrap();
    engine.set_cell_value("Sheet1", "B2", 22.0).unwrap();

    engine.set_cell_formula("Sheet1", "D1", "=@A1").unwrap();
    engine.set_cell_formula("Sheet1", "D2", "=@A1:A3").unwrap();
    engine.set_cell_formula("Sheet1", "B3", "=@A1:C1").unwrap();
    engine.set_cell_formula("Sheet1", "C3", "=@A1:B2").unwrap();

    // Ensure these formulas are compiled to bytecode (and aren't forcing AST fallback).
    assert_eq!(engine.bytecode_program_count(), 4);

    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Sheet1", "D1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "D2"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "B3"), Value::Number(10.0));
    assert_eq!(
        engine.get_cell_value("Sheet1", "C3"),
        Value::Error(ErrorKind::Value)
    );

    assert_engine_matches_ast(&engine, "=@A1", "D1");
    assert_engine_matches_ast(&engine, "=@A1:A3", "D2");
    assert_engine_matches_ast(&engine, "=@A1:C1", "B3");
    assert_engine_matches_ast(&engine, "=@A1:B2", "C3");
}

#[test]
fn bytecode_implicit_intersection_matches_ast_for_2d_range_inside_rectangle() {
    use formula_engine::bytecode::{CellCoord, GridMut, SparseGrid, Vm};

    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "B1", 10.0).unwrap();
    engine.set_cell_value("Sheet1", "B2", 42.0).unwrap();

    let mut grid = SparseGrid::new(10, 10);
    grid.set_value(CellCoord::new(0, 0), formula_engine::bytecode::Value::Number(1.0));
    grid.set_value(CellCoord::new(1, 0), formula_engine::bytecode::Value::Number(2.0));
    grid.set_value(CellCoord::new(0, 1), formula_engine::bytecode::Value::Number(10.0));
    grid.set_value(CellCoord::new(1, 1), formula_engine::bytecode::Value::Number(42.0));

    let formula = "=@A1:B2";
    let current_cell = "B2";
    let expected = eval_via_ast(&engine, formula, current_cell);

    let origin = parse_a1(current_cell).expect("parse current cell");
    let origin = formula_engine::CellAddr::new(origin.row, origin.col);
    let ast = formula_engine::parse_formula(
        formula,
        ParseOptions {
            locale: formula_engine::LocaleConfig::en_us(),
            reference_style: ReferenceStyle::A1,
            normalize_relative_to: Some(origin),
        },
    )
    .expect("parse canonical formula");

    let mut resolve_sheet = |_name: &str| Some(0usize);
    let bc_expr = formula_engine::bytecode::lower_canonical_expr(&ast.expr, origin, 0, &mut resolve_sheet)
        .expect("lower to bytecode expr");

    let cache = formula_engine::bytecode::BytecodeCache::new();
    let program = cache.get_or_compile(&bc_expr);
    assert_eq!(cache.program_count(), 1);
    assert!(
        program
            .instrs()
            .iter()
            .any(|inst| inst.op() == formula_engine::bytecode::OpCode::ImplicitIntersection),
        "bytecode program should contain ImplicitIntersection opcode"
    );

    let mut vm = Vm::with_capacity(32);
    let base = CellCoord::new(origin.row as i32, origin.col as i32);
    let bc_value = vm.eval(&program, &grid, base, &formula_engine::LocaleConfig::en_us());

    assert_eq!(bytecode_value_to_engine(bc_value), expected);
}

#[test]
fn bytecode_backend_matches_ast_for_vlookup_and_hlookup() {
    let mut engine = Engine::new();

    // VLOOKUP table: ascending numeric keys.
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "B1", "a").unwrap();
    engine.set_cell_value("Sheet1", "A2", 3.0).unwrap();
    engine.set_cell_value("Sheet1", "B2", "b").unwrap();
    engine.set_cell_value("Sheet1", "A3", 5.0).unwrap();
    engine.set_cell_value("Sheet1", "B3", "c").unwrap();

    // Mixed numeric/text keys (valid Excel ascending order: numbers < text).
    engine.set_cell_value("Sheet1", "D1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "D2", 3.0).unwrap();
    engine.set_cell_value("Sheet1", "D3", "A").unwrap();
    engine.set_cell_value("Sheet1", "E1", 10.0).unwrap();
    engine.set_cell_value("Sheet1", "E2", 30.0).unwrap();
    engine.set_cell_value("Sheet1", "E3", 40.0).unwrap();

    // HLOOKUP table: ascending numeric keys.
    engine.set_cell_value("Sheet1", "A10", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "B10", 3.0).unwrap();
    engine.set_cell_value("Sheet1", "C10", 5.0).unwrap();
    engine.set_cell_value("Sheet1", "A11", "a").unwrap();
    engine.set_cell_value("Sheet1", "B11", "b").unwrap();
    engine.set_cell_value("Sheet1", "C11", "c").unwrap();

    // Error propagation from lookup_value.
    engine
        .set_cell_value("Sheet1", "F1", Value::Error(ErrorKind::Div0))
        .unwrap();

    // VLOOKUP: exact match.
    engine
        .set_cell_formula("Sheet1", "G1", "=VLOOKUP(3, A1:B3, 2, FALSE)")
        .unwrap();
    // VLOOKUP: approximate match (range_lookup omitted => TRUE).
    engine
        .set_cell_formula("Sheet1", "G2", "=VLOOKUP(4, A1:B3, 2)")
        .unwrap();
    // VLOOKUP: missing => #N/A.
    engine
        .set_cell_formula("Sheet1", "G3", "=VLOOKUP(0, A1:B3, 2)")
        .unwrap();
    // VLOOKUP: bad index (< 1) => #VALUE!.
    engine
        .set_cell_formula("Sheet1", "G4", "=VLOOKUP(3, A1:B3, 0, FALSE)")
        .unwrap();
    // VLOOKUP: bad index (out of range) => #REF!.
    engine
        .set_cell_formula("Sheet1", "G5", "=VLOOKUP(3, A1:B3, 3, FALSE)")
        .unwrap();
    // VLOOKUP: approximate match with mixed-type key column.
    engine
        .set_cell_formula("Sheet1", "G6", "=VLOOKUP(4, D1:E3, 2)")
        .unwrap();
    // VLOOKUP: propagate lookup_value error.
    engine
        .set_cell_formula("Sheet1", "G7", "=VLOOKUP(F1, A1:B3, 2, FALSE)")
        .unwrap();

    // HLOOKUP: exact match.
    engine
        .set_cell_formula("Sheet1", "H1", "=HLOOKUP(3, A10:C11, 2, FALSE)")
        .unwrap();
    // HLOOKUP: approximate match.
    engine
        .set_cell_formula("Sheet1", "H2", "=HLOOKUP(4, A10:C11, 2)")
        .unwrap();
    // HLOOKUP: missing => #N/A.
    engine
        .set_cell_formula("Sheet1", "H3", "=HLOOKUP(0, A10:C11, 2)")
        .unwrap();
    // HLOOKUP: bad index (< 1) => #VALUE!.
    engine
        .set_cell_formula("Sheet1", "H4", "=HLOOKUP(3, A10:C11, 0, FALSE)")
        .unwrap();
    // HLOOKUP: bad index (out of range) => #REF!.
    engine
        .set_cell_formula("Sheet1", "H5", "=HLOOKUP(3, A10:C11, 3, FALSE)")
        .unwrap();

    // Ensure VLOOKUP/HLOOKUP formulas compile to bytecode (i.e. don't fall back to AST).
    assert_eq!(engine.bytecode_program_count(), 12);

    engine.recalculate_single_threaded();

    for (formula, cell) in [
        ("=VLOOKUP(3, A1:B3, 2, FALSE)", "G1"),
        ("=VLOOKUP(4, A1:B3, 2)", "G2"),
        ("=VLOOKUP(0, A1:B3, 2)", "G3"),
        ("=VLOOKUP(3, A1:B3, 0, FALSE)", "G4"),
        ("=VLOOKUP(3, A1:B3, 3, FALSE)", "G5"),
        ("=VLOOKUP(4, D1:E3, 2)", "G6"),
        ("=VLOOKUP(F1, A1:B3, 2, FALSE)", "G7"),
        ("=HLOOKUP(3, A10:C11, 2, FALSE)", "H1"),
        ("=HLOOKUP(4, A10:C11, 2)", "H2"),
        ("=HLOOKUP(0, A10:C11, 2)", "H3"),
        ("=HLOOKUP(3, A10:C11, 0, FALSE)", "H4"),
        ("=HLOOKUP(3, A10:C11, 3, FALSE)", "H5"),
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
        choice in 0u8..14u8,
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
            12 => "=A1&B1".to_string(),
            13 => "=A1%".to_string(),
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
        (Value::Text("".to_string()), Value::Number(2.0)), // blanks
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

#[test]
fn bytecode_backend_countif_criteria_respects_engine_value_locale() {
    // Numeric parsing uses the workbook value locale (decimal/thousands separators).
    let mut engine = Engine::new();
    engine.set_value_locale(ValueLocaleConfig::de_de());

    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    // Text numbers should be coerced using the workbook locale ("," as decimal separator in de-DE).
    engine.set_cell_value("Sheet1", "A2", "1,5").unwrap();
    engine.set_cell_value("Sheet1", "A3", 2.0).unwrap();
    engine
        .set_cell_value("Sheet1", "D1", Value::Text(">1,5".to_string()))
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B1", "=COUNTIF(A1:A3, D1)")
        .unwrap();

    // Ensure we're exercising the bytecode path (criteria arg is a cell reference).
    assert_eq!(engine.bytecode_program_count(), 1);
    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(1.0));
    assert_engine_matches_ast(&engine, "=COUNTIF(A1:A3, D1)", "B1");

    // Date parsing uses the workbook value locale's date order (DMY for de-DE).
    let mut engine = Engine::new();
    engine.set_value_locale(ValueLocaleConfig::de_de());

    let system = engine.date_system();
    let jan_2 = ymd_to_serial(ExcelDate::new(2020, 1, 2), system).unwrap();
    let feb_1 = ymd_to_serial(ExcelDate::new(2020, 2, 1), system).unwrap();

    engine.set_cell_value("Sheet1", "A1", jan_2 as f64).unwrap();
    engine.set_cell_value("Sheet1", "A2", feb_1 as f64).unwrap();
    engine
        .set_cell_value("Sheet1", "D1", Value::Text(">=1/2/2020".to_string()))
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B1", "=COUNTIF(A1:A2, D1)")
        .unwrap();

    // Ensure we're exercising the bytecode path (criteria arg is a cell reference).
    assert_eq!(engine.bytecode_program_count(), 1);
    engine.recalculate_single_threaded();

    // In de-DE (DMY), "1/2/2020" is Feb 1, 2020.
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(1.0));
    assert_engine_matches_ast(&engine, "=COUNTIF(A1:A2, D1)", "B1");
}

#[test]
fn bytecode_backend_coerces_scalar_text_using_engine_value_locale() {
    let mut engine = Engine::new();
    engine.set_value_locale(ValueLocaleConfig::de_de());

    // In de-DE, dot-separated numeric strings with three components (like `1.5.2020`) should be
    // treated as dates rather than a number with stripped thousands separators.
    engine
        .set_cell_formula("Sheet1", "A1", r#"="1.5.2020"+0"#)
        .unwrap();

    // Ensure we're exercising the bytecode path.
    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();

    let system = engine.date_system();
    let expected = ymd_to_serial(ExcelDate::new(2020, 5, 1), system).unwrap();

    assert_eq!(
        engine.get_cell_value("Sheet1", "A1"),
        Value::Number(expected as f64)
    );
    assert_engine_matches_ast(&engine, r#"="1.5.2020"+0"#, "A1");
}

#[test]
fn bytecode_compile_diagnostics_reports_fallback_reasons() {
    let mut engine = Engine::new();

    // Supported + eligible.
    engine.set_cell_formula("Sheet1", "A1", "=1+1").unwrap();
    // Volatile.
    engine.set_cell_formula("Sheet1", "A2", "=RAND()").unwrap();
    // Cross-sheet reference.
    engine.set_cell_value("Sheet2", "A1", 42.0).unwrap();
    engine.set_cell_formula("Sheet1", "A3", "=Sheet2!A1").unwrap();
    // Unsupported operator (intersection).
    engine
        .set_cell_formula("Sheet1", "A4", "=A1:A2 B1:B2")
        .unwrap();

    let stats = engine.bytecode_compile_stats();
    assert_eq!(stats.total_formula_cells, 4);
    assert_eq!(stats.compiled, 1);
    assert_eq!(stats.fallback, 3);

    assert_eq!(
        stats
            .fallback_reasons
            .get(&BytecodeCompileReason::Volatile)
            .copied()
            .unwrap_or(0),
        1
    );
    assert_eq!(
        stats
            .fallback_reasons
            .get(&BytecodeCompileReason::LowerError(
                formula_engine::bytecode::LowerError::CrossSheetReference
            ))
            .copied()
            .unwrap_or(0),
        1
    );

    // Intersection can fail in lowering (Unsupported) or by failing the eligibility gate.
    let unsupported = stats
        .fallback_reasons
        .get(&BytecodeCompileReason::LowerError(
            formula_engine::bytecode::LowerError::Unsupported,
        ))
        .copied()
        .unwrap_or(0);
    let ineligible = stats
        .fallback_reasons
        .get(&BytecodeCompileReason::IneligibleExpr)
        .copied()
        .unwrap_or(0);
    assert_eq!(unsupported + ineligible, 1);

    let report = engine.bytecode_compile_report(usize::MAX);
    assert_eq!(report.len(), 3);

    let a2 = parse_a1("A2").unwrap();
    let a3 = parse_a1("A3").unwrap();
    let a4 = parse_a1("A4").unwrap();

    let reason_for = |addr| {
        report
            .iter()
            .find(|e| e.sheet == "Sheet1" && e.addr == addr)
            .map(|e| e.reason.clone())
    };

    assert_eq!(reason_for(a2), Some(BytecodeCompileReason::Volatile));
    assert_eq!(
        reason_for(a3),
        Some(BytecodeCompileReason::LowerError(
            formula_engine::bytecode::LowerError::CrossSheetReference
        ))
    );
    let a4_reason = reason_for(a4).expect("A4 should appear in fallback report");
    assert!(
        matches!(
            a4_reason,
            BytecodeCompileReason::IneligibleExpr
                | BytecodeCompileReason::LowerError(formula_engine::bytecode::LowerError::Unsupported)
        ),
        "unexpected A4 bytecode compile reason: {a4_reason:?}"
    );
}

#[test]
fn bytecode_compile_diagnostics_reports_disabled_reason() {
    let mut engine = Engine::new();
    engine.set_bytecode_enabled(false);

    engine.set_cell_formula("Sheet1", "A1", "=1+1").unwrap();

    let stats = engine.bytecode_compile_stats();
    assert_eq!(stats.total_formula_cells, 1);
    assert_eq!(stats.compiled, 0);
    assert_eq!(stats.fallback, 1);
    assert_eq!(
        stats
            .fallback_reasons
            .get(&BytecodeCompileReason::Disabled)
            .copied()
            .unwrap_or(0),
        1
    );

    let report = engine.bytecode_compile_report(10);
    assert_eq!(report.len(), 1);
    assert_eq!(report[0].sheet, "Sheet1");
    assert_eq!(report[0].addr, parse_a1("A1").unwrap());
    assert_eq!(report[0].reason, BytecodeCompileReason::Disabled);
}

#[test]
fn bytecode_compile_diagnostics_reports_grid_and_range_limits() {
    let mut engine = Engine::new();

    // Out-of-bounds cell reference (column exceeds XFD).
    engine.set_cell_formula("Sheet1", "A1", "=XFE1").unwrap();

    // Huge range that would require an enormous columnar buffer; bytecode compilation should skip it.
    engine
        .set_cell_formula("Sheet1", "B1", "=SUM(A1:XFD1048576)")
        .unwrap();

    let stats = engine.bytecode_compile_stats();
    assert_eq!(stats.total_formula_cells, 2);
    assert_eq!(stats.compiled, 0);
    assert_eq!(stats.fallback, 2);
    assert_eq!(
        stats
            .fallback_reasons
            .get(&BytecodeCompileReason::ExceedsGridLimits)
            .copied()
            .unwrap_or(0),
        1
    );
    assert_eq!(
        stats
            .fallback_reasons
            .get(&BytecodeCompileReason::ExceedsRangeCellLimit)
            .copied()
            .unwrap_or(0),
        1
    );

    let report = engine.bytecode_compile_report(usize::MAX);
    assert_eq!(report.len(), 2);

    let a1 = report
        .iter()
        .find(|e| e.sheet == "Sheet1" && e.addr == parse_a1("A1").unwrap())
        .map(|e| e.reason.clone());
    assert_eq!(a1, Some(BytecodeCompileReason::ExceedsGridLimits));

    let b1 = report
        .iter()
        .find(|e| e.sheet == "Sheet1" && e.addr == parse_a1("B1").unwrap())
        .map(|e| e.reason.clone());
    assert_eq!(b1, Some(BytecodeCompileReason::ExceedsRangeCellLimit));
}

#[test]
fn bytecode_backend_inlines_constant_defined_names() {
    let mut engine_bc = Engine::new();
    engine_bc
        .define_name(
            "RATE",
            NameScope::Workbook,
            NameDefinition::Constant(Value::Number(0.1)),
        )
        .unwrap();
    engine_bc.set_cell_formula("Sheet1", "A1", "=RATE*2").unwrap();
    engine_bc.recalculate_single_threaded();

    assert_eq!(engine_bc.bytecode_program_count(), 1);
    assert_eq!(engine_bc.get_cell_value("Sheet1", "A1"), Value::Number(0.2));

    let mut engine_ast = Engine::new();
    engine_ast.set_bytecode_enabled(false);
    engine_ast
        .define_name(
            "RATE",
            NameScope::Workbook,
            NameDefinition::Constant(Value::Number(0.1)),
        )
        .unwrap();
    engine_ast.set_cell_formula("Sheet1", "A1", "=RATE*2").unwrap();
    engine_ast.recalculate_single_threaded();

    assert_eq!(
        engine_bc.get_cell_value("Sheet1", "A1"),
        engine_ast.get_cell_value("Sheet1", "A1")
    );
}

#[test]
fn bytecode_backend_inlines_reference_defined_names_when_static_refs() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 0.1).unwrap();
    engine
        .define_name(
            "RATE",
            NameScope::Workbook,
            NameDefinition::Reference("Sheet1!A1".to_string()),
        )
        .unwrap();

    engine.set_cell_formula("Sheet1", "B1", "=RATE*2").unwrap();
    engine.recalculate_single_threaded();

    // RATE should be inlined to a cell reference and compile to bytecode.
    assert_eq!(engine.bytecode_program_count(), 1);
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(0.2));
}

#[test]
fn bytecode_backend_recompiles_when_constant_defined_name_changes() {
    let mut engine = Engine::new();
    engine
        .define_name(
            "RATE",
            NameScope::Workbook,
            NameDefinition::Constant(Value::Number(0.1)),
        )
        .unwrap();
    engine.set_cell_formula("Sheet1", "A1", "=RATE*2").unwrap();
    engine.recalculate_single_threaded();

    assert_eq!(engine.bytecode_program_count(), 1);
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(0.2));

    engine
        .define_name(
            "RATE",
            NameScope::Workbook,
            NameDefinition::Constant(Value::Number(0.2)),
        )
        .unwrap();
    assert!(engine.is_dirty("Sheet1", "A1"));
    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(0.4));
    assert!(engine.bytecode_program_count() >= 2);
}

#[test]
fn bytecode_backend_compiles_error_literals() {
    for (formula, expected) in [
        ("=#N/A", Value::Error(ErrorKind::NA)),
        ("=#DIV/0!", Value::Error(ErrorKind::Div0)),
    ] {
        let mut engine = Engine::new();
        engine.set_cell_formula("Sheet1", "A1", formula).unwrap();

        // Ensure we're exercising the bytecode path.
        assert_eq!(engine.bytecode_program_count(), 1);

        engine.recalculate_single_threaded();
        assert_eq!(engine.get_cell_value("Sheet1", "A1"), expected);
        assert_engine_matches_ast(&engine, formula, "A1");
    }
}

#[test]
fn bytecode_backend_matches_ast_for_conditional_aggregates_numeric_criteria() {
    let mut engine = Engine::new();

    // Criteria range: numbers + implicit blank.
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", 3.0).unwrap();
    engine.set_cell_value("Sheet1", "A4", 4.0).unwrap();
    // A5 left blank

    // Value range.
    engine.set_cell_value("Sheet1", "B1", 10.0).unwrap();
    engine.set_cell_value("Sheet1", "B2", 20.0).unwrap();
    engine.set_cell_value("Sheet1", "B3", 30.0).unwrap();
    engine.set_cell_value("Sheet1", "B4", 40.0).unwrap();
    engine.set_cell_value("Sheet1", "B5", 50.0).unwrap();

    // Second criteria range.
    engine.set_cell_value("Sheet1", "C1", 100.0).unwrap();
    engine.set_cell_value("Sheet1", "C2", 200.0).unwrap();
    engine.set_cell_value("Sheet1", "C3", 100.0).unwrap();
    engine.set_cell_value("Sheet1", "C4", 200.0).unwrap();
    engine.set_cell_value("Sheet1", "C5", 100.0).unwrap();

    // Each formula should compile to bytecode (new function support).
    engine
        .set_cell_formula("Sheet1", "D1", r#"=SUMIF(A1:A5,">2",B1:B5)"#)
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "D2", r#"=SUMIFS(B1:B5,A1:A5,">2",C1:C5,"=100")"#)
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "D3", r#"=COUNTIFS(A1:A5,">2",C1:C5,"=100")"#)
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "D4", r#"=AVERAGEIF(A1:A5,">1",B1:B5)"#)
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "D5",
            r#"=AVERAGEIFS(B1:B5,A1:A5,">1",C1:C5,"=200")"#,
        )
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "D6", r#"=MINIFS(B1:B5,A1:A5,">1",C1:C5,"=200")"#)
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "D7", r#"=MAXIFS(B1:B5,A1:A5,">1",C1:C5,"=200")"#)
        .unwrap();
    engine
        // Exercise `<>` numeric criteria (blanks are coerced to 0, so they satisfy `<>2`).
        .set_cell_formula("Sheet1", "D8", r#"=SUMIF(A1:A5,"<>2",B1:B5)"#)
        .unwrap();

    assert_eq!(engine.bytecode_program_count(), 8);
    engine.recalculate_single_threaded();

    for (formula, cell) in [
        (r#"=SUMIF(A1:A5,">2",B1:B5)"#, "D1"),
        (r#"=SUMIFS(B1:B5,A1:A5,">2",C1:C5,"=100")"#, "D2"),
        (r#"=COUNTIFS(A1:A5,">2",C1:C5,"=100")"#, "D3"),
        (r#"=AVERAGEIF(A1:A5,">1",B1:B5)"#, "D4"),
        (r#"=AVERAGEIFS(B1:B5,A1:A5,">1",C1:C5,"=200")"#, "D5"),
        (r#"=MINIFS(B1:B5,A1:A5,">1",C1:C5,"=200")"#, "D6"),
        (r#"=MAXIFS(B1:B5,A1:A5,">1",C1:C5,"=200")"#, "D7"),
        (r#"=SUMIF(A1:A5,"<>2",B1:B5)"#, "D8"),
    ] {
        assert_engine_matches_ast(&engine, formula, cell);
    }
}

#[test]
fn bytecode_backend_propagates_error_literals_through_supported_ops() {
    let mut engine = Engine::new();
    engine.set_cell_formula("Sheet1", "A1", "=#N/A+1").unwrap();

    // Ensure we're exercising the bytecode path.
    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Error(ErrorKind::NA));
    assert_engine_matches_ast(&engine, "=#N/A+1", "A1");
}

#[test]
fn bytecode_backend_propagates_error_literals_through_supported_functions() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=ABS(#DIV/0!)")
        .unwrap();

    // Ensure we're exercising the bytecode path.
    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();

    assert_eq!(
        engine.get_cell_value("Sheet1", "A1"),
        Value::Error(ErrorKind::Div0)
    );
    assert_engine_matches_ast(&engine, "=ABS(#DIV/0!)", "A1");
}

#[test]
fn bytecode_backend_matches_ast_for_conditional_aggregates_full_criteria_semantics() {
    let mut engine = Engine::new();

    engine.set_cell_value("Sheet1", "A1", "apple").unwrap();
    engine.set_cell_value("Sheet1", "A2", "apricot").unwrap();
    engine.set_cell_value("Sheet1", "A3", "banana").unwrap();
    engine.set_cell_value("Sheet1", "A4", "*").unwrap();
    engine.set_cell_value("Sheet1", "A5", "").unwrap(); // empty string counts as blank
                                                        // A6 left blank (implicit blank)
    engine
        .set_cell_value("Sheet1", "A7", Value::Error(ErrorKind::Div0))
        .unwrap();

    for (row, v) in (1..=7).zip([1.0, 2.0, 3.0, 4.0, 5.0, 6.0, 7.0]) {
        engine
            .set_cell_value("Sheet1", &format!("B{row}"), v)
            .unwrap();
    }

    engine
        .set_cell_formula("Sheet1", "C1", r#"=SUMIF(A1:A7,"ap*",B1:B7)"#)
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "C2", r#"=SUMIF(A1:A7,"~*",B1:B7)"#)
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "C3", r#"=COUNTIFS(A1:A7,"ap*")"#)
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "C4", r#"=COUNTIFS(A1:A7,"~*")"#)
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "C5", r#"=SUMIF(A1:A7,"",B1:B7)"#)
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "C6", r#"=SUMIF(A1:A7,"<>",B1:B7)"#)
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "C7", r##"=SUMIF(A1:A7,"#DIV/0!",B1:B7)"##)
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "C8", r##"=COUNTIFS(A1:A7,"#DIV/0!")"##)
        .unwrap();
    engine
        // Errors in the criteria argument itself must propagate.
        .set_cell_formula("Sheet1", "C9", r#"=SUMIF(A1:A7,1/0,B1:B7)"#)
        .unwrap();

    assert_eq!(engine.bytecode_program_count(), 9);
    engine.recalculate_single_threaded();

    for (formula, cell) in [
        (r#"=SUMIF(A1:A7,"ap*",B1:B7)"#, "C1"),
        (r#"=SUMIF(A1:A7,"~*",B1:B7)"#, "C2"),
        (r#"=COUNTIFS(A1:A7,"ap*")"#, "C3"),
        (r#"=COUNTIFS(A1:A7,"~*")"#, "C4"),
        (r#"=SUMIF(A1:A7,"",B1:B7)"#, "C5"),
        (r#"=SUMIF(A1:A7,"<>",B1:B7)"#, "C6"),
        (r##"=SUMIF(A1:A7,"#DIV/0!",B1:B7)"##, "C7"),
        (r##"=COUNTIFS(A1:A7,"#DIV/0!")"##, "C8"),
        (r#"=SUMIF(A1:A7,1/0,B1:B7)"#, "C9"),
    ] {
        assert_engine_matches_ast(&engine, formula, cell);
    }
}

#[test]
fn bytecode_backend_matches_ast_for_conditional_aggregates_locale_criteria_parsing() {
    let mut engine = Engine::new();
    engine.set_value_locale(ValueLocaleConfig::de_de());

    // Numeric criteria parsing (`1,5` in de-DE).
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 1.5).unwrap();
    engine.set_cell_value("Sheet1", "A3", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "B1", 10.0).unwrap();
    engine.set_cell_value("Sheet1", "B2", 20.0).unwrap();
    engine.set_cell_value("Sheet1", "B3", 30.0).unwrap();
    engine
        .set_cell_value("Sheet1", "D1", Value::Text(">1,5".to_string()))
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "C1", "=SUMIF(A1:A3, D1, B1:B3)")
        .unwrap();

    // Date criteria parsing uses the workbook value locale's date order (DMY for de-DE).
    let system = engine.date_system();
    let jan_2 = ymd_to_serial(ExcelDate::new(2020, 1, 2), system).unwrap();
    let feb_1 = ymd_to_serial(ExcelDate::new(2020, 2, 1), system).unwrap();
    engine.set_cell_value("Sheet1", "E1", jan_2 as f64).unwrap();
    engine.set_cell_value("Sheet1", "E2", feb_1 as f64).unwrap();
    engine.set_cell_value("Sheet1", "F1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "F2", 2.0).unwrap();
    engine
        .set_cell_value("Sheet1", "D2", Value::Text(">=1/2/2020".to_string()))
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "C2", "=SUMIF(E1:E2, D2, F1:F2)")
        .unwrap();

    assert_eq!(engine.bytecode_program_count(), 2);
    engine.recalculate_single_threaded();

    assert_engine_matches_ast(&engine, "=SUMIF(A1:A3, D1, B1:B3)", "C1");
    assert_engine_matches_ast(&engine, "=SUMIF(E1:E2, D2, F1:F2)", "C2");
}

#[test]
fn bytecode_backend_matches_ast_for_conditional_aggregates_shape_mismatch_and_out_of_bounds() {
    // Shape mismatch should surface `#VALUE!`.
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", 3.0).unwrap();
    engine.set_cell_value("Sheet1", "B1", 10.0).unwrap();
    engine.set_cell_value("Sheet1", "B2", 20.0).unwrap();
    engine.set_cell_value("Sheet1", "B3", 30.0).unwrap();
    engine
        .set_cell_formula("Sheet1", "C1", r#"=SUMIFS(B1:B3,A1:A2,">0")"#)
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "C2", r#"=SUMIF(A1:A3,">0",B1:B2)"#)
        .unwrap();

    assert_eq!(engine.bytecode_program_count(), 2);
    engine.recalculate_single_threaded();

    assert_eq!(
        engine.get_cell_value("Sheet1", "C1"),
        Value::Error(ErrorKind::Value)
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "C2"),
        Value::Error(ErrorKind::Value)
    );
    assert_engine_matches_ast(&engine, r#"=SUMIFS(B1:B3,A1:A2,">0")"#, "C1");
    assert_engine_matches_ast(&engine, r#"=SUMIF(A1:A3,">0",B1:B2)"#, "C2");

    // Out-of-bounds ranges should surface `#REF!` in the bytecode runtime (ColumnarGrid bounds).
    let grid = bytecode::ColumnarGrid::new(2, 2);
    let expr = bytecode::parse_formula(
        r#"=SUMIF(A1:A3,">0",B1:B3)"#,
        bytecode::CellCoord::new(0, 0),
    )
    .unwrap();
    let program = bytecode::Compiler::compile(Arc::from("out_of_bounds"), &expr);
    let mut vm = bytecode::Vm::with_capacity(16);
    let v = vm.eval(
        &program,
        &grid,
        bytecode::CellCoord::new(0, 0),
        &formula_engine::LocaleConfig::en_us(),
    );
    assert_eq!(v, bytecode::Value::Error(bytecode::ErrorKind::Ref));
}
