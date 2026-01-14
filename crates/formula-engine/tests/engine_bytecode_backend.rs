#![cfg(not(target_arch = "wasm32"))]

use chrono::{DateTime, Utc};
use formula_engine::date::{serial_to_ymd, ymd_to_serial, ExcelDate, ExcelDateSystem};
use formula_engine::eval::{
    parse_a1, EvalContext, Evaluator, RecalcContext, SheetReference, ValueResolver,
};
use formula_engine::functions::{
    ArraySupport, FunctionContext, FunctionSpec, ThreadSafety, ValueType, Volatility,
};
use formula_engine::locale::ValueLocaleConfig;
use formula_engine::value::{EntityValue, NumberLocale, RecordValue};
use formula_engine::{
    bytecode, BytecodeCompileReason, Engine, ErrorKind, ExternalValueProvider, NameDefinition,
    NameScope, ParseOptions, ReferenceStyle, Value,
};
use formula_model::table::TableColumn;
use formula_model::{Range, Style, Table};
use proptest::prelude::*;
use std::sync::Arc;

fn not_thread_safe_test(
    _ctx: &dyn FunctionContext,
    _args: &[formula_engine::eval::CompiledExpr],
) -> Value {
    Value::Number(1.0)
}

inventory::submit! {
    FunctionSpec {
        name: "NOT_THREAD_SAFE_TEST",
        min_args: 0,
        max_args: 0,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::NotThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[],
        implementation: not_thread_safe_test,
    }
}

fn bytecode_unsupported_test(
    _ctx: &dyn FunctionContext,
    _args: &[formula_engine::eval::CompiledExpr],
) -> Value {
    Value::Number(1.0)
}

inventory::submit! {
    FunctionSpec {
        name: "BYTECODE_UNSUPPORTED_TEST",
        min_args: 0,
        max_args: 0,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[],
        implementation: bytecode_unsupported_test,
    }
}

fn cell_addr_to_a1(addr: formula_engine::eval::CellAddr) -> String {
    addr.to_a1()
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
    ) -> Result<
        Vec<(
            usize,
            formula_engine::eval::CellAddr,
            formula_engine::eval::CellAddr,
        )>,
        ErrorKind,
    > {
        Err(ErrorKind::Name)
    }

    fn spill_origin(
        &self,
        sheet_id: usize,
        addr: formula_engine::eval::CellAddr,
    ) -> Option<formula_engine::eval::CellAddr> {
        let sheet = match sheet_id {
            0 => "Sheet1",
            _ => return None,
        };
        let (_origin_sheet, origin) = self.engine.spill_origin(sheet, &cell_addr_to_a1(addr))?;
        Some(origin)
    }

    fn spill_range(
        &self,
        sheet_id: usize,
        origin: formula_engine::eval::CellAddr,
    ) -> Option<(
        formula_engine::eval::CellAddr,
        formula_engine::eval::CellAddr,
    )> {
        let sheet = match sheet_id {
            0 => "Sheet1",
            _ => return None,
        };
        self.engine.spill_range(sheet, &cell_addr_to_a1(origin))
    }

    fn text_codepage(&self) -> u16 {
        self.engine.text_codepage()
    }
}

fn eval_via_ast(engine: &Engine, formula: &str, current_cell: &str) -> Value {
    let resolver = EngineResolver { engine };
    let mut recalc_ctx = RecalcContext::new(0);
    let separators = engine.value_locale().separators;
    recalc_ctx.number_locale =
        NumberLocale::new(separators.decimal_sep, Some(separators.thousands_sep));
    recalc_ctx.calculation_mode = engine.calc_settings().calculation_mode;

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

fn eval_via_ast_with_now_utc(
    engine: &Engine,
    formula: &str,
    current_cell: &str,
    now_utc: DateTime<Utc>,
) -> Value {
    let resolver = EngineResolver { engine };
    let separators = engine.value_locale().separators;
    let recalc_ctx = RecalcContext {
        now_utc,
        recalc_id: 0,
        number_locale: NumberLocale::new(separators.decimal_sep, Some(separators.thousands_sep)),
        calculation_mode: engine.calc_settings().calculation_mode,
    };

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

fn assert_engine_spill_matches_ast(engine: &Engine, formula: &str, origin_cell: &str) {
    let expected = eval_via_ast(engine, formula, origin_cell);
    let Value::Array(arr) = expected else {
        panic!("expected formula {formula} to spill an array, got {expected:?}");
    };
    let origin = parse_a1(origin_cell).expect("parse spill origin");

    let end = formula_engine::eval::CellAddr {
        row: origin.row + (arr.rows as u32).saturating_sub(1),
        col: origin.col + (arr.cols as u32).saturating_sub(1),
    };
    assert_eq!(
        engine.spill_range("Sheet1", origin_cell),
        Some((origin, end)),
        "spill footprint mismatch for {origin_cell} {formula}"
    );

    for r in 0..arr.rows {
        for c in 0..arr.cols {
            let addr = formula_engine::eval::CellAddr {
                row: origin.row + r as u32,
                col: origin.col + c as u32,
            };
            let a1 = addr.to_a1();
            let expected = arr.get(r, c).cloned().unwrap_or(Value::Blank);
            assert_eq!(
                engine.get_cell_value("Sheet1", &a1),
                expected,
                "spill value mismatch at {a1} for origin {origin_cell} {formula}"
            );
        }
    }
}

fn bytecode_value_to_engine(value: formula_engine::bytecode::Value) -> Value {
    use formula_engine::bytecode::Value as ByteValue;
    match value {
        ByteValue::Number(n) => Value::Number(n),
        ByteValue::Bool(b) => Value::Bool(b),
        ByteValue::Text(s) => Value::Text(s.to_string()),
        ByteValue::Entity(v) => Value::Entity(v.as_ref().clone()),
        ByteValue::Record(v) => Value::Record(v.as_ref().clone()),
        ByteValue::Empty => Value::Blank,
        ByteValue::Missing => Value::Blank,
        ByteValue::Error(e) => Value::Error(e.into()),
        ByteValue::Lambda(_) => Value::Error(ErrorKind::Calc),
        // Array/range values are not valid scalar results for the engine API; treat them as spills.
        ByteValue::Array(_) | ByteValue::Range(_) | ByteValue::MultiRange(_) => {
            Value::Error(ErrorKind::Spill)
        }
    }
}

fn table_fixture_multi_col(range_a1: &str) -> Table {
    Table {
        id: 1,
        name: "Table1".into(),
        display_name: "Table1".into(),
        range: Range::from_a1(range_a1).unwrap(),
        header_row_count: 1,
        totals_row_count: 0,
        columns: vec![
            TableColumn {
                id: 1,
                name: "Col1".into(),
                formula: None,
                totals_formula: None,
            },
            TableColumn {
                id: 2,
                name: "Col2".into(),
                formula: None,
                totals_formula: None,
            },
            TableColumn {
                id: 3,
                name: "Col3".into(),
                formula: None,
                totals_formula: None,
            },
            TableColumn {
                id: 4,
                name: "Col4".into(),
                formula: None,
                totals_formula: None,
            },
        ],
        style: None,
        auto_filter: None,
        relationship_id: None,
        part_path: None,
    }
}

#[test]
fn bytecode_backend_compiles_structured_refs_and_recompiles_on_table_changes() {
    let mut engine = Engine::new();

    // Create a table and some values in Col1.
    engine.set_sheet_tables("Sheet1", vec![table_fixture_multi_col("A1:D4")]);
    engine.set_cell_value("Sheet1", "A2", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "A4", 3.0).unwrap();

    engine
        .set_cell_formula("Sheet1", "E1", "=SUM(Table1[Col1])")
        .unwrap();
    engine.recalculate_single_threaded();

    // Structured ref should take the bytecode path.
    assert_eq!(engine.bytecode_program_count(), 1);
    assert_eq!(engine.get_cell_value("Sheet1", "E1"), Value::Number(6.0));

    // Compare to AST-only evaluation.
    let mut engine_ast = Engine::new();
    engine_ast.set_bytecode_enabled(false);
    engine_ast.set_sheet_tables("Sheet1", vec![table_fixture_multi_col("A1:D4")]);
    engine_ast.set_cell_value("Sheet1", "A2", 1.0).unwrap();
    engine_ast.set_cell_value("Sheet1", "A3", 2.0).unwrap();
    engine_ast.set_cell_value("Sheet1", "A4", 3.0).unwrap();
    engine_ast
        .set_cell_formula("Sheet1", "E1", "=SUM(Table1[Col1])")
        .unwrap();
    engine_ast.recalculate_single_threaded();
    assert_eq!(engine_ast.bytecode_program_count(), 0);
    assert_eq!(
        engine.get_cell_value("Sheet1", "E1"),
        engine_ast.get_cell_value("Sheet1", "E1")
    );

    // Expand the table range to include another data row; Col1 sum should grow accordingly.
    let before_programs = engine.bytecode_program_count();
    engine.set_sheet_tables("Sheet1", vec![table_fixture_multi_col("A1:D5")]);
    engine.set_cell_value("Sheet1", "A5", 4.0).unwrap();
    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Sheet1", "E1"), Value::Number(10.0));
    assert_eq!(engine.bytecode_program_count(), before_programs + 1);

    engine_ast.set_sheet_tables("Sheet1", vec![table_fixture_multi_col("A1:D5")]);
    engine_ast.set_cell_value("Sheet1", "A5", 4.0).unwrap();
    engine_ast.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Sheet1", "E1"),
        engine_ast.get_cell_value("Sheet1", "E1")
    );
}

#[test]
fn bytecode_backend_reuses_program_for_this_row_structured_refs() {
    let mut engine = Engine::new();

    engine.set_sheet_tables("Sheet1", vec![table_fixture_multi_col("A1:D4")]);
    engine.set_cell_value("Sheet1", "A2", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "A4", 3.0).unwrap();

    // `[@Col]` depends on the current row; the bytecode backend should still be able to compile it
    // and reuse the same program pattern across rows.
    engine.set_cell_formula("Sheet1", "D2", "=[@Col1]").unwrap();
    engine.set_cell_formula("Sheet1", "D3", "=[@Col1]").unwrap();
    engine.set_cell_formula("Sheet1", "D4", "=[@Col1]").unwrap();

    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "D2"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "D3"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "D4"), Value::Number(3.0));
}

#[test]
fn bytecode_backend_compiles_and_evaluates_today() {
    let mut engine = Engine::new();
    engine.set_date_system(ExcelDateSystem::Excel1904);
    engine
        .set_cell_formula("Sheet1", "A1", "=TODAY()")
        .expect("set TODAY()");

    // Volatile formulas should still compile to bytecode (when thread-safe and non-dynamic).
    assert_eq!(engine.bytecode_program_count(), 1);

    let before = Utc::now();
    engine.recalculate_single_threaded();
    let after = Utc::now();

    let got = engine.get_cell_value("Sheet1", "A1");
    // `TODAY()` only changes at the day boundary; if the test happens to cross midnight between
    // `before` and `after`, accept either result.
    let expected_before = eval_via_ast_with_now_utc(&engine, "=TODAY()", "A1", before);
    let expected_after = eval_via_ast_with_now_utc(&engine, "=TODAY()", "A1", after);
    assert!(
        got == expected_before || got == expected_after,
        "TODAY() mismatch: bytecode={got:?}, ast_before={expected_before:?}, ast_after={expected_after:?}"
    );
}

#[test]
fn bytecode_backend_compiles_and_evaluates_now() {
    let mut engine = Engine::new();
    engine.set_date_system(ExcelDateSystem::Excel1904);
    engine
        .set_cell_formula("Sheet1", "A1", "=NOW()")
        .expect("set NOW()");

    // Volatile formulas should still compile to bytecode (when thread-safe and non-dynamic).
    assert_eq!(engine.bytecode_program_count(), 1);

    let before = Utc::now();
    engine.recalculate_single_threaded();
    let after = Utc::now();

    let got = engine.get_cell_value("Sheet1", "A1");
    let expected_before = eval_via_ast_with_now_utc(&engine, "=NOW()", "A1", before);
    let expected_after = eval_via_ast_with_now_utc(&engine, "=NOW()", "A1", after);
    match (got, expected_before, expected_after) {
        (Value::Number(got), Value::Number(before), Value::Number(after)) => {
            let (low, high) = if before <= after {
                (before, after)
            } else {
                (after, before)
            };
            // Allow a small epsilon to account for runtime differences between `Utc::now()` and
            // formula evaluation.
            let eps = 1.0 / 86_400.0; // 1 second
            assert!(
                got >= low - eps && got <= high + eps,
                "NOW() mismatch: bytecode={got}, ast_low={low}, ast_high={high}"
            );
        }
        (got, expected_before, expected_after) => {
            // Should never happen for NOW(), but keep the failure mode clear.
            assert!(
                got == expected_before || got == expected_after,
                "NOW() mismatch: bytecode={got:?}, ast_before={expected_before:?}, ast_after={expected_after:?}"
            );
        }
    }
}

#[test]
fn bytecode_backend_matches_ast_for_let_simple() {
    let mut engine = Engine::new();

    engine
        .set_cell_formula("Sheet1", "A1", "=LET(x, 1, x+1)")
        .unwrap();
    engine.recalculate_single_threaded();

    assert_engine_matches_ast(&engine, "=LET(x, 1, x+1)", "A1");
    assert_eq!(engine.bytecode_program_count(), 1);
}

#[test]
fn bytecode_backend_matches_ast_for_let_multiple_bindings_and_case_insensitive_names() {
    let mut engine = Engine::new();

    // Use different cases for the binding name and its references to assert case-insensitive lookup.
    engine
        .set_cell_formula("Sheet1", "A1", "=LET(X, 1, y, x+1, y+X)")
        .unwrap();
    engine.recalculate_single_threaded();

    assert_engine_matches_ast(&engine, "=LET(X, 1, y, x+1, y+X)", "A1");
    assert_eq!(engine.bytecode_program_count(), 1);
}

#[test]
fn bytecode_backend_matches_ast_for_let_with_cell_refs_in_bindings() {
    let mut engine = Engine::new();

    engine.set_cell_value("Sheet1", "A1", 10.0).unwrap();
    engine
        .set_cell_formula("Sheet1", "B1", "=LET(x, A1, x+1)")
        .unwrap();
    engine.recalculate_single_threaded();

    assert_engine_matches_ast(&engine, "=LET(x, A1, x+1)", "B1");
    assert_eq!(engine.bytecode_program_count(), 1);
}

#[test]
fn bytecode_backend_matches_ast_for_let_shadowing() {
    let mut engine = Engine::new();

    engine
        .set_cell_formula("Sheet1", "A1", "=LET(x, 1, LET(x, 2, x+1)+x)")
        .unwrap();
    engine.recalculate_single_threaded();

    assert_engine_matches_ast(&engine, "=LET(x, 1, LET(x, 2, x+1)+x)", "A1");
    assert_eq!(engine.bytecode_program_count(), 1);
}

#[test]
fn bytecode_backend_let_shadows_defined_name_constants() {
    // LET locals should shadow workbook/sheet defined names, and bytecode compilation must not
    // inline constant defined names in a way that breaks lexical scoping.
    let mut engine = Engine::new();
    engine
        .define_name(
            "X",
            NameScope::Workbook,
            NameDefinition::Constant(Value::Number(10.0)),
        )
        .unwrap();

    // Use mixed-case references to assert case-insensitive shadowing of defined names.
    engine
        .set_cell_formula("Sheet1", "A1", "=LET(x, 1, X+1)")
        .unwrap();

    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(2.0));
    assert_engine_matches_ast(&engine, "=LET(x, 1, X+1)", "A1");
}

#[test]
fn bytecode_backend_let_shadows_defined_name_static_refs() {
    // Defined names that point at static references (cell/range) are inlined during bytecode
    // compilation to improve eligibility. LET locals must still shadow those defined names.
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 100.0).unwrap();
    engine
        .define_name(
            "X",
            NameScope::Workbook,
            NameDefinition::Reference("Sheet1!$A$1".to_string()),
        )
        .unwrap();

    engine
        .set_cell_formula("Sheet1", "B1", "=LET(x, 1, X+1)")
        .unwrap();

    // This formula should still be bytecode-eligible: `X` inside the LET body refers to the local
    // binding (not the defined name), so inlining the defined name must not rewrite it.
    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(2.0));
    assert_engine_matches_ast(&engine, "=LET(x, 1, X+1)", "B1");
}

#[test]
fn bytecode_backend_matches_ast_for_let_rebinding_same_name() {
    let mut engine = Engine::new();

    // Rebinding `x` in the same LET should see the previous `x` value in the RHS and then shadow it.
    engine
        .set_cell_formula("Sheet1", "A1", "=LET(x, 1, x, x+1, x)")
        .unwrap();
    engine.recalculate_single_threaded();

    assert_engine_matches_ast(&engine, "=LET(x, 1, x, x+1, x)", "A1");
    assert_eq!(engine.bytecode_program_count(), 1);
}

#[test]
fn bytecode_backend_matches_ast_for_let_shadowing_uses_outer_binding_in_inner_rhs() {
    let mut engine = Engine::new();

    // The inner binding RHS should resolve `x` from the outer LET because the inner `x` is not
    // visible until after its value expression has been evaluated.
    engine
        .set_cell_formula("Sheet1", "A1", "=LET(x, 1, LET(x, x+1, x))")
        .unwrap();
    engine.recalculate_single_threaded();

    assert_engine_matches_ast(&engine, "=LET(x, 1, LET(x, x+1, x))", "A1");
    assert_eq!(engine.bytecode_program_count(), 1);
}

#[test]
fn bytecode_backend_matches_ast_for_let_error_propagation() {
    let mut engine = Engine::new();

    // Errors inside binding expressions should propagate like the AST evaluator.
    engine
        .set_cell_formula("Sheet1", "A1", "=LET(x, 1/0, x+1)")
        .unwrap();
    engine.recalculate_single_threaded();

    assert_engine_matches_ast(&engine, "=LET(x, 1/0, x+1)", "A1");
    assert_eq!(engine.bytecode_program_count(), 1);
}

#[test]
fn bytecode_backend_matches_ast_for_let_unused_error_binding() {
    let mut engine = Engine::new();

    // LET evaluates binding expressions eagerly, but errors are values (not exceptions), so an
    // unused error binding should not force the overall LET result to be an error.
    engine
        .set_cell_formula("Sheet1", "A1", "=LET(x, 1/0, 1)")
        .unwrap();
    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(1.0));
    assert_engine_matches_ast(&engine, "=LET(x, 1/0, 1)", "A1");
    assert_eq!(engine.bytecode_program_count(), 1);
}

#[test]
fn bytecode_backend_matches_ast_for_let_with_iferror_short_circuit() {
    let mut engine = Engine::new();

    // IFERROR should short-circuit when the first argument is not an error, even when that argument
    // is a LET local.
    engine
        .set_cell_formula("Sheet1", "A1", "=LET(x, 1, IFERROR(x, 1/0))")
        .unwrap();
    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(1.0));
    assert_engine_matches_ast(&engine, "=LET(x, 1, IFERROR(x, 1/0))", "A1");
    assert_eq!(engine.bytecode_program_count(), 1);
}

#[test]
fn bytecode_backend_matches_ast_for_let_with_iferror_on_error_local() {
    let mut engine = Engine::new();

    engine
        .set_cell_formula("Sheet1", "A1", "=LET(x, 1/0, IFERROR(x, 0))")
        .unwrap();
    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(0.0));
    assert_engine_matches_ast(&engine, "=LET(x, 1/0, IFERROR(x, 0))", "A1");
    assert_eq!(engine.bytecode_program_count(), 1);
}

#[test]
fn bytecode_backend_short_circuits_let_in_if_branches() {
    let mut engine = Engine::new();

    engine
        .set_cell_formula("Sheet1", "A1", "=IF(FALSE, LET(x, 1/0, 1), 2)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A2", "=IF(TRUE, 2, LET(x, 1/0, 1))")
        .unwrap();

    // Both formulas should compile to bytecode; the LET inside the untaken IF branch should not be
    // evaluated (so the #DIV/0! is never observed).
    assert_eq!(engine.bytecode_program_count(), 2);

    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "A2"), Value::Number(2.0));

    assert_engine_matches_ast(&engine, "=IF(FALSE, LET(x, 1/0, 1), 2)", "A1");
    assert_engine_matches_ast(&engine, "=IF(TRUE, 2, LET(x, 1/0, 1))", "A2");
}

#[test]
fn bytecode_backend_reuses_program_for_filled_let_patterns() {
    let mut engine = Engine::new();

    engine.set_cell_value("Sheet1", "A1", 10.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 20.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", 30.0).unwrap();

    engine
        .set_cell_formula("Sheet1", "B1", "=LET(x, A1, x+1)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B2", "=LET(x, A2, x+1)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B3", "=LET(x, A3, x+1)")
        .unwrap();

    // Filled LET formulas should share a single normalized bytecode program.
    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();
    assert_engine_matches_ast(&engine, "=LET(x, A1, x+1)", "B1");
    assert_engine_matches_ast(&engine, "=LET(x, A2, x+1)", "B2");
    assert_engine_matches_ast(&engine, "=LET(x, A3, x+1)", "B3");
}

#[test]
fn bytecode_backend_rejects_invalid_let_name_arg() {
    // LET's binding "name" args must be identifiers; invalid name args should fall back to the AST
    // evaluator (and produce #VALUE!).
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=LET(1, 2, 3)")
        .unwrap();
    assert_eq!(engine.bytecode_program_count(), 0);

    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Sheet1", "A1"),
        Value::Error(ErrorKind::Value)
    );
    assert_engine_matches_ast(&engine, "=LET(1, 2, 3)", "A1");

    engine
        .set_cell_formula("Sheet1", "A2", "=LET(A1, 2, 3)")
        .unwrap();
    // Still not bytecode-eligible because the binding "name" is a cell ref, not a bare identifier.
    assert_eq!(engine.bytecode_program_count(), 0);

    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Sheet1", "A2"),
        Value::Error(ErrorKind::Value)
    );
    assert_engine_matches_ast(&engine, "=LET(A1, 2, 3)", "A2");
}

#[test]
fn bytecode_backend_matches_ast_for_lambda_invocation_call_expr() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=LAMBDA(x,x+1)(2)")
        .unwrap();
    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();
    assert_engine_matches_ast(&engine, "=LAMBDA(x,x+1)(2)", "A1");
}

#[test]
fn bytecode_backend_preserves_reference_results_from_lambda_calls() {
    // Lambdas can return references, which should remain as references so reference-only functions
    // can consume them (rather than forcing an eager dereference/spill inside the lambda body).
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=ROW(LAMBDA(r,r)(A10))")
        .unwrap();
    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();
    assert_engine_matches_ast(&engine, "=ROW(LAMBDA(r,r)(A10))", "A1");
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(10.0));
}

#[test]
fn bytecode_backend_matches_ast_for_lambda_call_with_array_literal_arg() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=LAMBDA(x,SUM(x))({1,2,3})")
        .unwrap();
    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();
    assert_engine_matches_ast(&engine, "=LAMBDA(x,SUM(x))({1,2,3})", "A1");
}

#[test]
fn bytecode_backend_matches_ast_for_let_captured_env_lambda_call() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=LET(a,10,LAMBDA(x,a+x)(2))")
        .unwrap();
    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();
    assert_engine_matches_ast(&engine, "=LET(a,10,LAMBDA(x,a+x)(2))", "A1");
}

#[test]
fn bytecode_backend_preserves_reference_semantics_for_let_single_cell_reference_passed_to_lambda() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=LET(r,A10,LAMBDA(x,ROW(x))(r))")
        .unwrap();
    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();
    assert_engine_matches_ast(&engine, "=LET(r,A10,LAMBDA(x,ROW(x))(r))", "A1");
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(10.0));
}

#[test]
fn bytecode_backend_enforces_lambda_recursion_limit() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=LET(f,LAMBDA(x,f(x)),f(1))")
        .unwrap();
    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();
    assert_engine_matches_ast(&engine, "=LET(f,LAMBDA(x,f(x)),f(1))", "A1");
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Error(ErrorKind::Calc));
}

#[test]
fn bytecode_backend_supports_isomitted_inside_lambdas() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula(
            "Sheet1",
            "A1",
            "=LET(f,LAMBDA(x,y,IF(ISOMITTED(y),x,x+y)),f(2))",
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "A2",
            "=LET(f,LAMBDA(x,y,IF(ISOMITTED(y),x,x+y)),f(2,3))",
        )
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A3", "=LET(f,LAMBDA(x,y,ISOMITTED(y)),f(1))")
        .unwrap();
    // A blank placeholder is not the same as an omitted argument.
    engine
        .set_cell_formula("Sheet1", "A4", "=LET(f,LAMBDA(x,y,ISOMITTED(y)),f(1,))")
        .unwrap();

    let stats = engine.bytecode_compile_stats();
    assert_eq!(stats.total_formula_cells, 4);
    assert_eq!(
        stats.compiled, 4,
        "expected ISOMITTED lambda formulas to compile to bytecode (report={:?})",
        engine.bytecode_compile_report(32)
    );
    assert_eq!(stats.fallback, 0);

    engine.recalculate_single_threaded();

    for (formula, cell) in [
        ("=LET(f,LAMBDA(x,y,IF(ISOMITTED(y),x,x+y)),f(2))", "A1"),
        ("=LET(f,LAMBDA(x,y,IF(ISOMITTED(y),x,x+y)),f(2,3))", "A2"),
        ("=LET(f,LAMBDA(x,y,ISOMITTED(y)),f(1))", "A3"),
        ("=LET(f,LAMBDA(x,y,ISOMITTED(y)),f(1,))", "A4"),
    ] {
        assert_engine_matches_ast(&engine, formula, cell);
    }

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "A2"), Value::Number(5.0));
    assert_eq!(engine.get_cell_value("Sheet1", "A3"), Value::Bool(true));
    assert_eq!(engine.get_cell_value("Sheet1", "A4"), Value::Bool(false));
}

#[test]
fn bytecode_backend_supports_isomitted_outside_lambda() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=ISOMITTED(x)")
        .unwrap();

    // `x` is not a LAMBDA parameter here, so ISOMITTED should return FALSE; this should still be
    // bytecode-eligible (the identifier argument is not evaluated as a name reference).
    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();
    assert_engine_matches_ast(&engine, "=ISOMITTED(x)", "A1");
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Bool(false));
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
fn bytecode_backend_supports_array_literal_arguments_for_sum() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=SUM({1,2;3,4})")
        .unwrap();

    // Ensure the formula was compiled to bytecode (array literals should not force an AST fallback
    // in supported contexts).
    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();
    assert_engine_matches_ast(&engine, "=SUM({1,2;3,4})", "A1");
}

#[test]
fn bytecode_backend_supports_array_literal_arguments_for_count() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=COUNT({1,\"x\";TRUE,2})")
        .unwrap();

    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();
    assert_engine_matches_ast(&engine, "=COUNT({1,\"x\";TRUE,2})", "A1");
}

#[test]
fn bytecode_backend_supports_countif_array_literal_ranges() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=COUNTIF({1,2,3}, \">1\")")
        .unwrap();

    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(2.0));
    assert_engine_matches_ast(&engine, "=COUNTIF({1,2,3}, \">1\")", "A1");
}

#[test]
fn bytecode_backend_supports_countif_array_literal_locals_via_let() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=LET(a, {1,2,3}, COUNTIF(a, \">1\"))")
        .unwrap();

    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(2.0));
    assert_engine_matches_ast(&engine, "=LET(a, {1,2,3}, COUNTIF(a, \">1\"))", "A1");
}

#[test]
fn bytecode_backend_supports_countifs_array_literal_ranges() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=COUNTIFS({1,2,3}, \">1\")")
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "A2",
            "=COUNTIFS({\"A\",\"A\",\"B\",\"B\"},\"A\",{1,2,3,4},\">1\")",
        )
        .unwrap();

    assert_eq!(engine.bytecode_program_count(), 2);

    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(2.0));
    assert_engine_matches_ast(&engine, "=COUNTIFS({1,2,3}, \">1\")", "A1");

    assert_eq!(engine.get_cell_value("Sheet1", "A2"), Value::Number(1.0));
    assert_engine_matches_ast(
        &engine,
        "=COUNTIFS({\"A\",\"A\",\"B\",\"B\"},\"A\",{1,2,3,4},\">1\")",
        "A2",
    );
}

#[test]
fn bytecode_backend_supports_countifs_array_literal_locals_via_let() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=LET(a, {1,2,3}, COUNTIFS(a, \">1\"))")
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "A2",
            "=LET(a,{\"A\",\"A\",\"B\",\"B\"},b,{1,2,3,4},COUNTIFS(a,\"A\",b,\">1\"))",
        )
        .unwrap();

    assert_eq!(engine.bytecode_program_count(), 2);

    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(2.0));
    assert_engine_matches_ast(&engine, "=LET(a, {1,2,3}, COUNTIFS(a, \">1\"))", "A1");

    assert_eq!(engine.get_cell_value("Sheet1", "A2"), Value::Number(1.0));
    assert_engine_matches_ast(
        &engine,
        "=LET(a,{\"A\",\"A\",\"B\",\"B\"},b,{1,2,3,4},COUNTIFS(a,\"A\",b,\">1\"))",
        "A2",
    );
}

#[test]
fn bytecode_backend_supports_sumifs_averageifs_minifs_maxifs_array_literal_ranges() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula(
            "Sheet1",
            "A1",
            r#"=SUMIFS({10,20,30,40},{"A","A","B","B"},"A",{1,2,3,4},">1")"#,
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "A2",
            r#"=AVERAGEIFS({10,20,30,40},{"A","A","B","B"},"A",{1,2,3,4},">1")"#,
        )
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A3", r#"=MAXIFS({10,20,30,40},{1,2,3,4},">2")"#)
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A4", r#"=MINIFS({10,20,30,40},{1,2,3,4},">2")"#)
        .unwrap();

    assert_eq!(
        engine.bytecode_program_count(),
        4,
        "expected criteria aggregates with array literal ranges to compile to bytecode"
    );

    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(20.0));
    assert_eq!(engine.get_cell_value("Sheet1", "A2"), Value::Number(20.0));
    assert_eq!(engine.get_cell_value("Sheet1", "A3"), Value::Number(40.0));
    assert_eq!(engine.get_cell_value("Sheet1", "A4"), Value::Number(30.0));

    assert_engine_matches_ast(
        &engine,
        r#"=SUMIFS({10,20,30,40},{"A","A","B","B"},"A",{1,2,3,4},">1")"#,
        "A1",
    );
    assert_engine_matches_ast(
        &engine,
        r#"=AVERAGEIFS({10,20,30,40},{"A","A","B","B"},"A",{1,2,3,4},">1")"#,
        "A2",
    );
    assert_engine_matches_ast(&engine, r#"=MAXIFS({10,20,30,40},{1,2,3,4},">2")"#, "A3");
    assert_engine_matches_ast(&engine, r#"=MINIFS({10,20,30,40},{1,2,3,4},">2")"#, "A4");
}

#[test]
fn bytecode_backend_supports_sumifs_averageifs_minifs_maxifs_array_literal_locals_via_let() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula(
            "Sheet1",
            "A1",
            r#"=LET(sum,{10,20,30,40},cats,{"A","A","B","B"},nums,{1,2,3,4},SUMIFS(sum,cats,"A",nums,">1"))"#,
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "A2",
            r#"=LET(avg,{10,20,30,40},cats,{"A","A","B","B"},nums,{1,2,3,4},AVERAGEIFS(avg,cats,"A",nums,">1"))"#,
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "A3",
            r#"=LET(vals,{10,20,30,40},nums,{1,2,3,4},MAXIFS(vals,nums,">2"))"#,
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "A4",
            r#"=LET(vals,{10,20,30,40},nums,{1,2,3,4},MINIFS(vals,nums,">2"))"#,
        )
        .unwrap();

    assert_eq!(
        engine.bytecode_program_count(),
        4,
        "expected criteria aggregates with LET-bound array literals to compile to bytecode"
    );

    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(20.0));
    assert_eq!(engine.get_cell_value("Sheet1", "A2"), Value::Number(20.0));
    assert_eq!(engine.get_cell_value("Sheet1", "A3"), Value::Number(40.0));
    assert_eq!(engine.get_cell_value("Sheet1", "A4"), Value::Number(30.0));

    assert_engine_matches_ast(
        &engine,
        r#"=LET(sum,{10,20,30,40},cats,{"A","A","B","B"},nums,{1,2,3,4},SUMIFS(sum,cats,"A",nums,">1"))"#,
        "A1",
    );
    assert_engine_matches_ast(
        &engine,
        r#"=LET(avg,{10,20,30,40},cats,{"A","A","B","B"},nums,{1,2,3,4},AVERAGEIFS(avg,cats,"A",nums,">1"))"#,
        "A2",
    );
    assert_engine_matches_ast(
        &engine,
        r#"=LET(vals,{10,20,30,40},nums,{1,2,3,4},MAXIFS(vals,nums,">2"))"#,
        "A3",
    );
    assert_engine_matches_ast(
        &engine,
        r#"=LET(vals,{10,20,30,40},nums,{1,2,3,4},MINIFS(vals,nums,">2"))"#,
        "A4",
    );
}

#[test]
fn array_literal_errors_propagate_in_sum() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=SUM({1,#DIV/0!})")
        .unwrap();

    // Ensure this stays on the bytecode backend (array literals can contain error literals).
    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();

    assert_engine_matches_ast(&engine, "=SUM({1,#DIV/0!})", "A1");
    assert_eq!(
        engine.get_cell_value("Sheet1", "A1"),
        Value::Error(ErrorKind::Div0)
    );
}

#[test]
fn array_literals_enable_bytecode_for_logical_functions() {
    // AND/OR have typed semantics over booleans within array literals, so ensure the bytecode
    // backend can compile and evaluate them without falling back to the AST path.
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=AND({TRUE,FALSE})")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A2", "=OR({TRUE,FALSE})")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A3", "=XOR({TRUE,FALSE})")
        .unwrap();

    assert_eq!(engine.bytecode_program_count(), 3);

    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Bool(false));
    assert_eq!(engine.get_cell_value("Sheet1", "A2"), Value::Bool(true));
    assert_eq!(engine.get_cell_value("Sheet1", "A3"), Value::Bool(true));
    assert_engine_matches_ast(&engine, "=AND({TRUE,FALSE})", "A1");
    assert_engine_matches_ast(&engine, "=OR({TRUE,FALSE})", "A2");
    assert_engine_matches_ast(&engine, "=XOR({TRUE,FALSE})", "A3");
}

#[test]
fn bytecode_backend_supports_countif_array_literal_range_arg() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=COUNTIF({1,,3}, \"<1\")")
        .unwrap();
    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(1.0));
    assert_engine_matches_ast(&engine, "=COUNTIF({1,,3}, \"<1\")", "A1");
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
fn bytecode_backend_inlines_defined_name_3d_sheet_span_ranges() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet2", "A1", 2.0).unwrap();
    engine
        .define_name(
            "My3D",
            NameScope::Workbook,
            NameDefinition::Reference("Sheet1:Sheet2!$A$1".to_string()),
        )
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B1", "=SUM(My3D)")
        .unwrap();

    // Ensure the named 3D reference was inlined and compiled to bytecode.
    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();
    let via_bytecode = engine.get_cell_value("Sheet1", "B1");
    assert_eq!(via_bytecode, Value::Number(3.0));

    engine.set_bytecode_enabled(false);
    engine.recalculate_single_threaded();
    let via_ast = engine.get_cell_value("Sheet1", "B1");

    assert_eq!(via_bytecode, via_ast);
}

#[test]
fn bytecode_backend_inlines_defined_name_3d_sheet_span_static_ranges() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 10.0).unwrap();
    engine.set_cell_value("Sheet2", "A1", 2.0).unwrap();
    engine.set_cell_value("Sheet2", "A2", 20.0).unwrap();
    engine
        .define_name(
            "My3DRange",
            NameScope::Workbook,
            NameDefinition::Reference("Sheet1:Sheet2!$A$1:$A$2".to_string()),
        )
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B1", "=SUM(My3DRange)")
        .unwrap();

    // Ensure the named 3D range was inlined and compiled to bytecode.
    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();
    let via_bytecode = engine.get_cell_value("Sheet1", "B1");
    assert_eq!(via_bytecode, Value::Number(33.0));

    engine.set_bytecode_enabled(false);
    engine.recalculate_single_threaded();
    let via_ast = engine.get_cell_value("Sheet1", "B1");

    assert_eq!(via_bytecode, via_ast);
}

#[test]
fn bytecode_backend_inlines_defined_name_reference_aliases_when_static_refs() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 2.0).unwrap();
    engine
        .define_name(
            "Base",
            NameScope::Workbook,
            NameDefinition::Reference("Sheet1!$A$1:$A$2".to_string()),
        )
        .unwrap();
    engine
        .define_name(
            "Alias",
            NameScope::Workbook,
            NameDefinition::Reference("Base".to_string()),
        )
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B1", "=SUM(Alias)")
        .unwrap();

    // Ensure the alias name was inlined (recursively) and compiled to bytecode.
    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();
    let via_bytecode = engine.get_cell_value("Sheet1", "B1");
    assert_eq!(via_bytecode, Value::Number(3.0));

    engine.set_bytecode_enabled(false);
    engine.recalculate_single_threaded();
    let via_ast = engine.get_cell_value("Sheet1", "B1");

    assert_eq!(via_bytecode, via_ast);
}

#[test]
fn bytecode_backend_inlines_defined_name_range_endpoints_through_aliases() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", 3.0).unwrap();

    engine
        .define_name(
            "Start",
            NameScope::Workbook,
            NameDefinition::Reference("Sheet1!$A$1".to_string()),
        )
        .unwrap();
    engine
        .define_name(
            "End",
            NameScope::Workbook,
            NameDefinition::Reference("Sheet1!$A$3".to_string()),
        )
        .unwrap();
    engine
        // Range definitions can be expressed via other defined names as long as they ultimately
        // resolve to static references.
        .define_name(
            "MyRange",
            NameScope::Workbook,
            NameDefinition::Reference("Start:End".to_string()),
        )
        .unwrap();

    engine
        .set_cell_formula("Sheet1", "B1", "=SUM(MyRange)")
        .unwrap();

    // Ensure the named range (built from aliased endpoints) was inlined and compiled to bytecode.
    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();
    let via_bytecode = engine.get_cell_value("Sheet1", "B1");
    assert_eq!(via_bytecode, Value::Number(6.0));

    engine.set_bytecode_enabled(false);
    engine.recalculate_single_threaded();
    let via_ast = engine.get_cell_value("Sheet1", "B1");

    assert_eq!(via_bytecode, via_ast);
}

#[test]
fn bytecode_backend_inlines_defined_name_union_static_refs() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 2.0).unwrap();
    engine
        .define_name(
            "MyUnion",
            NameScope::Workbook,
            NameDefinition::Reference("Sheet1!$A$1,Sheet1!$A$2".to_string()),
        )
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B1", "=SUM(MyUnion)")
        .unwrap();

    // Ensure the named union reference was inlined and compiled to bytecode.
    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();
    let via_bytecode = engine.get_cell_value("Sheet1", "B1");
    assert_eq!(via_bytecode, Value::Number(3.0));

    engine.set_bytecode_enabled(false);
    engine.recalculate_single_threaded();
    let via_ast = engine.get_cell_value("Sheet1", "B1");

    assert_eq!(via_bytecode, via_ast);
}

#[test]
fn bytecode_backend_inlines_defined_name_intersection_static_refs() {
    let mut engine = Engine::new();

    // Populate a small grid of numbers:
    // A1:C3 = 1..=9 (row-major).
    let mut n = 1.0;
    for row in 1..=3 {
        for col in ["A", "B", "C"] {
            engine
                .set_cell_value("Sheet1", &format!("{col}{row}"), n)
                .unwrap();
            n += 1.0;
        }
    }

    engine
        .define_name(
            "MyIntersect",
            NameScope::Workbook,
            NameDefinition::Reference("Sheet1!$A$1:$C$3 Sheet1!$B$2:$D$4".to_string()),
        )
        .unwrap();
    engine
        // Keep the formula outside the operand ranges to avoid spurious circular references in the
        // engine's conservative dependency analysis.
        .set_cell_formula("Sheet1", "E1", "=SUM(MyIntersect)")
        .unwrap();

    // Ensure the intersection name was inlined and compiled to bytecode.
    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();
    let via_bytecode = engine.get_cell_value("Sheet1", "E1");
    assert_eq!(via_bytecode, Value::Number(28.0));

    engine.set_bytecode_enabled(false);
    engine.recalculate_single_threaded();
    let via_ast = engine.get_cell_value("Sheet1", "E1");

    assert_eq!(via_bytecode, via_ast);
}

#[test]
fn bytecode_backend_inlines_defined_name_spill_range_refs() {
    let mut engine = Engine::new();

    // Create a spill on Sheet1 starting at A1 (spills into B1).
    engine
        .set_cell_formula("Sheet1", "A1", "={1,2}")
        .unwrap();
    engine
        .define_name(
            "MySpill",
            NameScope::Workbook,
            NameDefinition::Reference("Sheet1!A1#".to_string()),
        )
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "C1", "=SUM(MySpill)")
        .unwrap();

    // Ensure the dependent formula compiled to bytecode (even if the spill origin formula did too).
    let report = engine.bytecode_compile_report(usize::MAX);
    let c1 = parse_a1("C1").unwrap();
    assert!(
        report
            .iter()
            .find(|e| e.sheet == "Sheet1" && e.addr == c1)
            .is_none(),
        "expected C1 to compile to bytecode"
    );

    engine.recalculate_single_threaded();
    let via_bytecode = engine.get_cell_value("Sheet1", "C1");
    assert_eq!(via_bytecode, Value::Number(3.0));

    engine.set_bytecode_enabled(false);
    engine.recalculate_single_threaded();
    let via_ast = engine.get_cell_value("Sheet1", "C1");

    assert_eq!(via_bytecode, via_ast);
}

#[test]
fn bytecode_backend_sheet_qualified_defined_name_formula_uses_target_sheet_context() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet2", "A1", 2.0).unwrap();
    engine
        .define_name(
            "MyFormulaName",
            NameScope::Workbook,
            NameDefinition::Formula("=A1".to_string()),
        )
        .unwrap();

    // Evaluating `Sheet2!MyFormulaName` should use Sheet2 as the "current sheet" for `A1`.
    // The defined-name formula is inlined for bytecode compilation, and the engine normalizes
    // unprefixed references (like `A1`) to the target sheet so the resulting bytecode program can
    // evaluate correctly.
    engine
        .set_cell_formula("Sheet1", "B1", "=Sheet2!MyFormulaName")
        .unwrap();
    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(2.0));
}

#[test]
fn bytecode_backend_inlines_dynamic_defined_name_formulas() {
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

    // Defined-name formulas are inlined for bytecode compilation (bytecode does not resolve
    // workbook names at runtime). `INDIRECT` is now supported, so INDIRECT-based definitions can
    // still compile and evaluate.
    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();
    let via_bytecode = engine.get_cell_value("Sheet1", "B1");
    assert_eq!(via_bytecode, Value::Number(1.0));

    engine.set_bytecode_enabled(false);
    engine.recalculate_single_threaded();
    let via_ast = engine.get_cell_value("Sheet1", "B1");
    assert_eq!(via_bytecode, via_ast);
}

#[test]
fn bytecode_backend_compiles_and_evaluates_let_formulas() {
    // Simple LET (scalar-only) should compile to bytecode and match the AST evaluator.
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=LET(x, 1, x+2)")
        .unwrap();
    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(3.0));
    assert_engine_matches_ast(&engine, "=LET(x, 1, x+2)", "A1");
}

#[test]
fn bytecode_backend_let_supports_nested_scopes_and_shadowing() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=LET(x, 1, LET(x, 2, x+1))")
        .unwrap();
    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(3.0));
    assert_engine_matches_ast(&engine, "=LET(x, 1, LET(x, 2, x+1))", "A1");
}

#[test]
fn bytecode_backend_let_supports_reference_bindings() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 10.0).unwrap();
    engine
        .set_cell_formula("Sheet1", "B1", "=LET(x, A1, x+1)")
        .unwrap();
    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(11.0));
    assert_engine_matches_ast(&engine, "=LET(x, A1, x+1)", "B1");
}

#[test]
fn bytecode_backend_let_preserves_reference_semantics_for_sum() {
    // `SUM` treats scalar arguments differently from reference arguments:
    // - `SUM("5")` => 5
    // - `SUM(A1)` => 0 if A1 contains text
    //
    // LET should preserve reference semantics when binding a bare cell/range reference.
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", "5").unwrap();
    engine
        .set_cell_formula("Sheet1", "B1", "=LET(x, A1, SUM(x))")
        .unwrap();
    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(0.0));
    assert_engine_matches_ast(&engine, "=LET(x, A1, SUM(x))", "B1");
}

#[test]
fn bytecode_backend_choose_preserves_reference_semantics_for_sum() {
    // CHOOSE can return references; when the selected choice is a cell reference, outer aggregate
    // functions should observe reference semantics (e.g. SUM ignores text in referenced cells).
    //
    // Also verify the bytecode backend keeps CHOOSE lazy in this range context:
    // the unselected `1/0` branch must not be evaluated.
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", "5").unwrap();
    engine
        .set_cell_formula("Sheet1", "B1", "=SUM(CHOOSE(1, A1, 1/0))")
        .unwrap();
    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(0.0));
    assert_engine_matches_ast(&engine, "=SUM(CHOOSE(1, A1, 1/0))", "B1");
}

#[test]
fn bytecode_backend_choose_is_scalar_safe_for_concat() {
    // CONCAT evaluates its arguments in a scalar/value context in the bytecode compiler. Ensure
    // CHOOSE remains lazy in that context: the unselected `1/0` branch must not be evaluated.
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", "hello").unwrap();
    engine
        .set_cell_formula("Sheet1", "B1", "=CONCAT(CHOOSE(1, A1, 1/0))")
        .unwrap();
    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Sheet1", "B1"),
        Value::Text("hello".into())
    );
    assert_engine_matches_ast(&engine, "=CONCAT(CHOOSE(1, A1, 1/0))", "B1");
}

#[test]
fn bytecode_backend_choose_is_scalar_safe_for_abs() {
    // When CHOOSE is used in a scalar-only context (e.g. ABS), the selected value must behave like
    // a scalar value (not a reference/range). Otherwise ABS would treat it as a spill attempt.
    //
    // Also verify the bytecode backend keeps CHOOSE lazy in this scalar context: the unselected
    // `1/0` branch must not be evaluated.
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", -5.0).unwrap();
    engine
        .set_cell_formula("Sheet1", "B1", "=ABS(CHOOSE(1, A1, 1/0))")
        .unwrap();
    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(5.0));
    assert_engine_matches_ast(&engine, "=ABS(CHOOSE(1, A1, 1/0))", "B1");
}

#[test]
fn bytecode_backend_let_array_returning_abs_allows_concat_to_flatten_arrays() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", -1.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", -2.0).unwrap();

    // ABS supports array-lifting semantics in bytecode, and CONCAT flattens array/range arguments
    // into a single scalar text value. Ensure LET kind inference does not prevent compiling this to
    // bytecode.
    let formula = "=CONCAT(LET(x, ABS(A1:A2), x))";
    engine.set_cell_formula("Sheet1", "B1", formula).unwrap();

    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Text("12".into()));
    assert_engine_matches_ast(&engine, formula, "B1");
}

#[test]
fn bytecode_backend_let_xlookup_array_returning_xlookup_allows_concat_bytecode() {
    let mut engine = Engine::new();

    // XLOOKUP can spill when `return_array` is 2D (it returns a row/column slice). CONCAT flattens
    // array/range arguments into a single scalar text value, so it is safe to compile this to
    // bytecode.
    let formula = "=CONCAT(LET(x, XLOOKUP(2,{1;2;3},{10,11;20,21;30,31}), x))";
    engine.set_cell_formula("Sheet1", "B1", formula).unwrap();

    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Text("2021".into()));
    assert_engine_matches_ast(&engine, formula, "B1");
}

#[test]
fn bytecode_backend_let_xlookup_array_if_not_found_allows_concat_bytecode() {
    let mut engine = Engine::new();

    // XLOOKUP can return an array `if_not_found` value when no match is found. Ensure LET kind
    // inference still allows CONCAT to consume that array and flatten it into a scalar string.
    let formula = "=CONCAT(LET(x, XLOOKUP(99,{1;2;3},{10;20;30},{100;200}), x))";
    engine.set_cell_formula("Sheet1", "B1", formula).unwrap();
    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Text("100200".into()));
    assert_engine_matches_ast(&engine, formula, "B1");
}

#[test]
fn bytecode_backend_row_array_result_allows_concat_bytecode() {
    let mut engine = Engine::new();

    // ROW over a multi-cell range yields a dynamic array. CONCAT should flatten it into a single
    // scalar string value.
    let formula = "=CONCAT(ROW(A1:A2))";
    engine.set_cell_formula("Sheet1", "B1", formula).unwrap();
    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Text("12".into()));
    assert_engine_matches_ast(&engine, formula, "B1");
}

#[test]
fn bytecode_backend_if_array_result_allows_concat_bytecode() {
    let mut engine = Engine::new();

    // IF can return an array result when its condition is scalar. CONCAT should flatten that array
    // into a scalar string value.
    let formula = "=CONCAT(IF(TRUE, ROW(A1:A2), 0))";
    engine.set_cell_formula("Sheet1", "B1", formula).unwrap();
    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Text("12".into()));
    assert_engine_matches_ast(&engine, formula, "B1");
}

#[test]
fn bytecode_backend_xlookup_array_result_allows_concat_bytecode() {
    let mut engine = Engine::new();

    // XLOOKUP can return a 1D row/column slice (spilled array) when `return_array` is 2D. CONCAT
    // should flatten it into a scalar string value.
    let formula = "=CONCAT(XLOOKUP(2,{1;2;3},{10,11;20,21;30,31}))";
    engine.set_cell_formula("Sheet1", "B1", formula).unwrap();
    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Text("2021".into()));
    assert_engine_matches_ast(&engine, formula, "B1");
}

#[test]
fn bytecode_backend_let_single_cell_reference_local_is_scalar_safe_for_concat() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", "hello").unwrap();

    // LET binding values are evaluated in a reference context by the bytecode compiler, so `x`
    // is stored as a single-cell reference. CONCAT flattens range arguments, so consuming `x`
    // should still produce a scalar string value.
    let formula = "=LET(x, A1, CONCAT(x))";
    engine.set_cell_formula("Sheet1", "B1", formula).unwrap();
    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Text("hello".into()));
    assert_engine_matches_ast(&engine, formula, "B1");
}

#[test]
fn bytecode_backend_let_cell_ref_bindings_can_be_consumed_as_ranges_for_countif() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();

    // A LET binding to a bare cell reference is stored as a reference in bytecode (to preserve
    // reference semantics). Ensure it can still be consumed by range-taking functions like COUNTIF
    // without forcing an AST fallback.
    engine
        .set_cell_formula("Sheet1", "B1", r#"=LET(r, A1, COUNTIF(r, ">0"))"#)
        .unwrap();

    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(1.0));
    assert_engine_matches_ast(&engine, r#"=LET(r, A1, COUNTIF(r, ">0"))"#, "B1");
}

#[test]
fn bytecode_backend_let_cell_ref_bindings_can_be_consumed_as_ranges_for_match() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 10.0).unwrap();

    engine
        .set_cell_formula("Sheet1", "B1", "=LET(r, A1, MATCH(10, r, 0))")
        .unwrap();

    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(1.0));
    assert_engine_matches_ast(&engine, "=LET(r, A1, MATCH(10, r, 0))", "B1");
}

#[test]
fn bytecode_backend_let_supports_range_bindings_when_consumed_by_sum() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 2.0).unwrap();
    engine
        .set_cell_formula("Sheet1", "B1", "=LET(r, A1:A2, SUM(r))")
        .unwrap();
    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(3.0));
    assert_engine_matches_ast(&engine, "=LET(r, A1:A2, SUM(r))", "B1");
}

#[test]
fn bytecode_backend_let_supports_range_bindings_for_countif() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", 3.0).unwrap();
    engine
        .set_cell_formula("Sheet1", "B1", r#"=LET(r, A1:A3, COUNTIF(r, ">1"))"#)
        .unwrap();
    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(2.0));
    assert_engine_matches_ast(&engine, r#"=LET(r, A1:A3, COUNTIF(r, ">1"))"#, "B1");
}

#[test]
fn bytecode_backend_let_supports_array_literal_bindings_when_consumed_by_sum() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=LET(a, {1,2;3,4}, SUM(a))")
        .unwrap();
    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(10.0));
    assert_engine_matches_ast(&engine, "=LET(a, {1,2;3,4}, SUM(a))", "A1");
}

#[test]
fn bytecode_backend_sum_can_consume_array_literal_returned_from_let() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=SUM(LET(a, {1,2;3,4}, a))")
        .unwrap();
    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(10.0));
    assert_engine_matches_ast(&engine, "=SUM(LET(a, {1,2;3,4}, a))", "A1");
}

#[test]
fn bytecode_backend_supports_let_range_bindings_that_spill() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 2.0).unwrap();
    engine
        .set_cell_formula("Sheet1", "B1", "=LET(r, A1:A2, r)")
        .unwrap();
    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "B2"), Value::Number(2.0));
}

#[test]
fn bytecode_backend_rejects_invalid_let_arity() {
    // LET must have an odd number of args >= 3.
    // Invalid forms should fall back to the AST evaluator (and produce #VALUE!).
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=LET(x, 1)")
        .unwrap();
    assert_eq!(engine.bytecode_program_count(), 0);

    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Sheet1", "A1"),
        Value::Error(ErrorKind::Value)
    );
    assert_engine_matches_ast(&engine, "=LET(x, 1)", "A1");
}

#[test]
fn bytecode_backend_supports_3d_sheet_span_references() {
    let mut engine = Engine::new();

    for (sheet, values) in [
        ("Sheet1", [1.0, 2.0, 3.0]),
        ("Sheet2", [4.0, 5.0, 6.0]),
        ("Sheet3", [7.0, 8.0, 9.0]),
    ] {
        for (idx, v) in values.into_iter().enumerate() {
            engine
                .set_cell_value(sheet, &format!("A{}", idx + 1), v)
                .unwrap();
        }
    }

    engine
        .set_cell_formula("Sheet1", "B1", "=SUM(Sheet1:Sheet3!A1)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B2", "=COUNT(Sheet1:Sheet3!A1:A3)")
        .unwrap();

    // Both formulas should compile to bytecode (3D spans shouldn't force AST fallback).
    assert_eq!(engine.bytecode_program_count(), 2);

    engine.recalculate_single_threaded();

    let bc_sum = engine.get_cell_value("Sheet1", "B1");
    let bc_count = engine.get_cell_value("Sheet1", "B2");

    // Compare against the AST backend by disabling bytecode and re-evaluating.
    engine.set_bytecode_enabled(false);
    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Sheet1", "B1"), bc_sum);
    assert_eq!(engine.get_cell_value("Sheet1", "B2"), bc_count);

    assert_eq!(bc_sum, Value::Number(12.0)); // 1 + 4 + 7
    assert_eq!(bc_count, Value::Number(9.0));
}

#[test]
fn bytecode_backend_matches_ast_for_and_or_over_3d_sheet_spans() {
    fn setup(engine: &mut Engine) {
        // Ensure the sheet span exists so the bytecode lowerer can resolve the sheet names.
        for sheet in ["Sheet1", "Sheet2", "Sheet3"] {
            engine.ensure_sheet(sheet);
        }
        engine
            .set_cell_formula("Sheet1", "B1", "=AND(Sheet1:Sheet3!A1)")
            .unwrap();
        engine
            .set_cell_formula("Sheet1", "B2", "=OR(Sheet1:Sheet3!A1)")
            .unwrap();
    }

    let cases: Vec<([Option<Value>; 3], Value, Value)> = vec![
        // All referenced cells blank => no values => AND=TRUE, OR=FALSE.
        ([None, None, None], Value::Bool(true), Value::Bool(false)),
        // Text values in 3D references are treated like text in ranges (ignored).
        (
            [Some(Value::Text("hello".to_string())), None, None],
            Value::Bool(true),
            Value::Bool(false),
        ),
        // Numbers/bools across sheets participate.
        (
            [
                Some(Value::Number(1.0)),
                Some(Value::Number(0.0)),
                Some(Value::Number(1.0)),
            ],
            Value::Bool(false),
            Value::Bool(true),
        ),
        // Errors in the referenced cells propagate.
        (
            [
                Some(Value::Number(1.0)),
                Some(Value::Error(ErrorKind::Div0)),
                Some(Value::Number(1.0)),
            ],
            Value::Error(ErrorKind::Div0),
            Value::Error(ErrorKind::Div0),
        ),
    ];

    for (values, expected_and, expected_or) in cases {
        let mut bytecode_engine = Engine::new();
        setup(&mut bytecode_engine);
        assert_eq!(
            bytecode_engine.bytecode_program_count(),
            2,
            "expected AND/OR formulas to compile to bytecode"
        );

        for (sheet, value) in ["Sheet1", "Sheet2", "Sheet3"]
            .into_iter()
            .zip(values.iter())
        {
            match value {
                None => bytecode_engine.clear_cell(sheet, "A1").unwrap(),
                Some(v) => bytecode_engine
                    .set_cell_value(sheet, "A1", v.clone())
                    .unwrap(),
            }
        }

        bytecode_engine.recalculate_single_threaded();
        let bc_and = bytecode_engine.get_cell_value("Sheet1", "B1");
        let bc_or = bytecode_engine.get_cell_value("Sheet1", "B2");

        let mut ast_engine = Engine::new();
        ast_engine.set_bytecode_enabled(false);
        setup(&mut ast_engine);

        for (sheet, value) in ["Sheet1", "Sheet2", "Sheet3"]
            .into_iter()
            .zip(values.iter())
        {
            match value {
                None => ast_engine.clear_cell(sheet, "A1").unwrap(),
                Some(v) => ast_engine.set_cell_value(sheet, "A1", v.clone()).unwrap(),
            }
        }

        ast_engine.recalculate_single_threaded();
        let ast_and = ast_engine.get_cell_value("Sheet1", "B1");
        let ast_or = ast_engine.get_cell_value("Sheet1", "B2");

        assert_eq!(bc_and, ast_and, "AND mismatch");
        assert_eq!(bc_or, ast_or, "OR mismatch");
        assert_eq!(bc_and, expected_and, "unexpected AND result");
        assert_eq!(bc_or, expected_or, "unexpected OR result");
    }
}

#[test]
fn bytecode_backend_matches_ast_for_counta_and_countblank() {
    let mut engine = Engine::new();

    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    // A2 left blank
    engine.set_cell_value("Sheet1", "A3", "").unwrap(); // empty string
    engine.set_cell_value("Sheet1", "A4", "hello").unwrap();
    engine.set_cell_value("Sheet1", "A5", true).unwrap();
    engine
        .set_cell_value("Sheet1", "A6", Value::Error(ErrorKind::Div0))
        .unwrap();
    engine.set_cell_value("Sheet1", "A7", 0.0).unwrap();
    // A8 left blank
    engine.set_cell_value("Sheet1", "A9", "0").unwrap();
    engine.set_cell_value("Sheet1", "A10", 2.0).unwrap();

    engine
        .set_cell_formula("Sheet1", "B1", "=COUNTA(A1:A10)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B2", "=COUNTBLANK(A1:A10)")
        .unwrap();

    // Ensure the formulas compile to bytecode instead of falling back to AST evaluation.
    assert_eq!(engine.bytecode_program_count(), 2);

    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(8.0));
    assert_eq!(engine.get_cell_value("Sheet1", "B2"), Value::Number(3.0));
    assert_engine_matches_ast(&engine, "=COUNTA(A1:A10)", "B1");
    assert_engine_matches_ast(&engine, "=COUNTBLANK(A1:A10)", "B2");
}

#[test]
fn bytecode_backend_supports_counta_and_countblank_array_literals() {
    let mut engine = Engine::new();

    engine
        .set_cell_formula("Sheet1", "A1", "=COUNTA({\"a\",TRUE,\"\"})")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A2", "=COUNTBLANK({\"a\",TRUE,\"\"})")
        .unwrap();

    // Ensure these formulas compile to bytecode.
    assert_eq!(engine.bytecode_program_count(), 2);

    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(3.0));
    assert_eq!(engine.get_cell_value("Sheet1", "A2"), Value::Number(1.0));
    assert_engine_matches_ast(&engine, "=COUNTA({\"a\",TRUE,\"\"})", "A1");
    assert_engine_matches_ast(&engine, "=COUNTBLANK({\"a\",TRUE,\"\"})", "A2");
}

#[test]
fn bytecode_backend_counta_and_countblank_respect_non_numeric_reference_cells() {
    // If the bytecode backend incorrectly treats non-numeric reference values as NaN via column
    // slices, COUNTA/COUNTBLANK will miscount. This test ensures we fall back to per-cell scanning
    // when a referenced range contains text/bools.
    let mut engine = Engine::new();

    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", "x").unwrap();
    engine.set_cell_value("Sheet1", "A3", true).unwrap();
    engine.set_cell_value("Sheet1", "A4", "").unwrap(); // empty string: non-blank for COUNTA, blank for COUNTBLANK
                                                        // A5 left blank

    engine
        .set_cell_formula("Sheet1", "B1", "=COUNTA(A1:A5)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B2", "=COUNTBLANK(A1:A5)")
        .unwrap();

    assert_eq!(engine.bytecode_program_count(), 2);

    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(4.0));
    assert_eq!(engine.get_cell_value("Sheet1", "B2"), Value::Number(2.0));
    assert_engine_matches_ast(&engine, "=COUNTA(A1:A5)", "B1");
    assert_engine_matches_ast(&engine, "=COUNTBLANK(A1:A5)", "B2");
}

#[test]
fn bytecode_backend_supports_spill_range_operator() {
    let mut engine = Engine::new();

    // Create a spilled dynamic array (A1:A3).
    engine
        .set_cell_formula("Sheet1", "A1", "=SEQUENCE(3)")
        .unwrap();
    // Consume the spill range via `#` in a bytecode-eligible formula.
    engine
        .set_cell_formula("Sheet1", "B1", "=SUM(A1#)")
        .unwrap();
    // Also verify referencing a spill *child* (`A2#`) resolves to the same spill origin and still
    // compiles to bytecode.
    engine
        .set_cell_formula("Sheet1", "B2", "=SUM(A2#)")
        .unwrap();

    // Ensure both formulas compiled to bytecode (i.e. are absent from the fallback report).
    let report = engine.bytecode_compile_report(10);
    let b1 = parse_a1("B1").unwrap();
    let b2 = parse_a1("B2").unwrap();
    assert!(
        !report.iter().any(|e| e.addr == b1),
        "expected B1 to compile to bytecode (report={report:?})"
    );
    assert!(
        !report.iter().any(|e| e.addr == b2),
        "expected B2 to compile to bytecode (report={report:?})"
    );

    engine.recalculate_single_threaded();

    assert_eq!(engine.bytecode_program_count(), 1);
    assert_engine_matches_ast(&engine, "=SUM(A1#)", "B1");
    assert_engine_matches_ast(&engine, "=SUM(A2#)", "B2");
}

#[test]
fn bytecode_backend_let_spill_range_locals_allow_abs_to_spill_via_bytecode() {
    let mut engine = Engine::new();

    // Create a spilled dynamic array (A1:A2).
    engine
        .set_cell_formula("Sheet1", "A1", "=SEQUENCE(2)")
        .unwrap();

    // ABS supports array-lifting semantics in bytecode, so spill-range locals should be bytecode
    // eligible and spill correctly.
    engine
        .set_cell_formula("Sheet1", "B1", "=ABS(LET(r, A1#, r))")
        .unwrap();

    assert_eq!(engine.bytecode_program_count(), 1);
    engine.recalculate_single_threaded();

    assert_eq!(
        engine.spill_range("Sheet1", "B1"),
        Some((parse_a1("B1").unwrap(), parse_a1("B2").unwrap()))
    );
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "B2"), Value::Number(2.0));
}

#[test]
fn bytecode_backend_let_range_arithmetic_locals_allow_abs_to_spill_via_bytecode() {
    let mut engine = Engine::new();

    engine.set_cell_value("Sheet1", "A1", -1.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", -2.0).unwrap();

    // `A1:A2+0` produces an array result. ABS should lift over it and spill in bytecode mode.
    engine
        .set_cell_formula("Sheet1", "B1", "=ABS(LET(x, A1:A2+0, x))")
        .unwrap();

    assert_eq!(engine.bytecode_program_count(), 1);
    engine.recalculate_single_threaded();

    assert_eq!(
        engine.spill_range("Sheet1", "B1"),
        Some((parse_a1("B1").unwrap(), parse_a1("B2").unwrap()))
    );
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "B2"), Value::Number(2.0));
}

#[test]
fn bytecode_backend_let_array_literal_arithmetic_locals_allow_abs_to_spill_via_bytecode() {
    let mut engine = Engine::new();

    // `{-1;-2}+0` produces a dynamic array result (not a plain array literal). ABS should lift over
    // it and spill in bytecode mode.
    engine
        .set_cell_formula("Sheet1", "A1", "=ABS(LET(x, {-1;-2}+0, x))")
        .unwrap();

    assert_eq!(engine.bytecode_program_count(), 1);
    engine.recalculate_single_threaded();

    assert_eq!(
        engine.spill_range("Sheet1", "A1"),
        Some((parse_a1("A1").unwrap(), parse_a1("A2").unwrap()))
    );
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "A2"), Value::Number(2.0));
}

#[test]
fn bytecode_backend_let_nested_let_range_result_locals_allow_abs_to_spill_via_bytecode() {
    let mut engine = Engine::new();

    engine.set_cell_value("Sheet1", "A1", -1.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", -2.0).unwrap();

    // Nested LETs should propagate range/array-typed locals through the kind inference logic and
    // remain bytecode-eligible now that ABS lifts over arrays.
    engine
        .set_cell_formula("Sheet1", "B1", "=ABS(LET(x, LET(r, A1:A2, r)+0, x))")
        .unwrap();

    assert_eq!(engine.bytecode_program_count(), 1);
    engine.recalculate_single_threaded();

    assert_eq!(
        engine.spill_range("Sheet1", "B1"),
        Some((parse_a1("B1").unwrap(), parse_a1("B2").unwrap()))
    );
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "B2"), Value::Number(2.0));
}

#[test]
fn bytecode_backend_let_nested_let_cell_ref_result_is_scalar_safe() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", -1.0).unwrap();

    // Nested LETs that return a single-cell reference should still behave like a scalar in scalar
    // contexts (via implicit intersection), and remain bytecode-eligible.
    engine
        .set_cell_formula("Sheet1", "B1", "=ABS(LET(x, LET(r, A1, r), x))")
        .unwrap();

    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(1.0));
    assert_engine_matches_ast(&engine, "=ABS(LET(x, LET(r, A1, r), x))", "B1");
}

#[test]
fn bytecode_backend_let_choose_single_cell_reference_locals_are_scalar_safe() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", -1.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", -2.0).unwrap();

    // CHOOSE can return references. When those references are single-cell, LET locals should apply
    // implicit intersection in scalar contexts so scalar-only bytecode functions (like ABS) behave
    // correctly.
    engine
        .set_cell_formula("Sheet1", "B1", "=ABS(LET(x, CHOOSE(2, 1/0, A2), x))")
        .unwrap();

    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(2.0));
    assert_engine_matches_ast(&engine, "=ABS(LET(x, CHOOSE(2, 1/0, A2), x))", "B1");
}

#[test]
fn bytecode_backend_let_choose_range_locals_allow_abs_to_spill_via_bytecode() {
    let mut engine = Engine::new();

    engine.set_cell_value("Sheet1", "A1", -1.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", -2.0).unwrap();

    // CHOOSE can also return multi-cell ranges. ABS should be able to lift over those ranges and
    // spill in bytecode mode.
    engine
        .set_cell_formula("Sheet1", "C1", "=ABS(LET(x, CHOOSE(1, A1:A2, A1:A2), x))")
        .unwrap();

    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();

    assert_eq!(
        engine.spill_range("Sheet1", "C1"),
        Some((parse_a1("C1").unwrap(), parse_a1("C2").unwrap()))
    );
    assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "C2"), Value::Number(2.0));
}

#[test]
fn bytecode_backend_let_array_returning_function_locals_allow_abs_to_spill_via_bytecode() {
    let mut engine = Engine::new();

    engine.set_cell_value("Sheet1", "A1", -1.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", -2.0).unwrap();

    // N(A1:A2) returns an array result. ABS should lift over it and spill in bytecode mode.
    engine
        .set_cell_formula("Sheet1", "B1", "=ABS(LET(x, N(A1:A2), x))")
        .unwrap();

    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();

    assert_eq!(
        engine.spill_range("Sheet1", "B1"),
        Some((parse_a1("B1").unwrap(), parse_a1("B2").unwrap()))
    );
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "B2"), Value::Number(2.0));
}

#[test]
fn bytecode_backend_match_accepts_array_lookup_arrays_including_let_locals() {
    let mut engine = Engine::new();

    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", 3.0).unwrap();

    // MATCH accepts both reference-like lookup arrays and array values. Ensure it compiles to
    // bytecode when the lookup array is produced via range arithmetic (directly or via LET).
    engine
        .set_cell_formula("Sheet1", "B1", "=LET(x, A1:A3*10, MATCH(20, x, 0))")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B2", "=MATCH(20, A1:A3*10, 0)")
        .unwrap();

    assert_eq!(engine.bytecode_program_count(), 2);

    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(2.0));
    assert_engine_matches_ast(&engine, "=LET(x, A1:A3*10, MATCH(20, x, 0))", "B1");
    assert_eq!(engine.get_cell_value("Sheet1", "B2"), Value::Number(2.0));
    assert_engine_matches_ast(&engine, "=MATCH(20, A1:A3*10, 0)", "B2");
}

#[test]
fn bytecode_backend_match_accepts_array_literal_lookup_arrays() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=MATCH(2, {1,2,3}, 0)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A2", "=MATCH(2, {1;2;3}, 0)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A3", "=MATCH(2.5, {1,2,3}, 1)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A4", "=MATCH(2, {1,2;3,4}, 0)")
        .unwrap();

    assert_eq!(
        engine.bytecode_program_count(),
        4,
        "expected MATCH array literal formulas to compile to bytecode"
    );

    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "A2"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "A3"), Value::Number(2.0));
    assert_eq!(
        engine.get_cell_value("Sheet1", "A4"),
        Value::Error(ErrorKind::NA)
    );

    for (formula, cell) in [
        ("=MATCH(2, {1,2,3}, 0)", "A1"),
        ("=MATCH(2, {1;2;3}, 0)", "A2"),
        ("=MATCH(2.5, {1,2,3}, 1)", "A3"),
        ("=MATCH(2, {1,2;3,4}, 0)", "A4"),
    ] {
        assert_engine_matches_ast(&engine, formula, cell);
    }
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
fn bytecode_backend_style_only_cells_do_not_override_external_value_provider() {
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

    let style_id = engine.intern_style(Style {
        number_format: Some("0.00".to_string()),
        ..Default::default()
    });
    engine
        .set_cell_style_id("Sheet1", "A1", style_id)
        .expect("set style id");

    engine
        .set_cell_formula("Sheet1", "B1", "=SUM(A1:A3)")
        .unwrap();
    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(6.0));
    assert!(engine.bytecode_program_count() > 0);
}

#[test]
fn bytecode_backend_resolves_external_workbook_reference_via_provider() {
    struct Provider;

    impl ExternalValueProvider for Provider {
        fn get(&self, sheet: &str, addr: formula_engine::eval::CellAddr) -> Option<Value> {
            if !sheet.eq_ignore_ascii_case("[Book.xlsx]Sheet1") {
                return None;
            }
            match (addr.row, addr.col) {
                (0, 0) => Some(Value::Number(123.0)),
                _ => None,
            }
        }
    }

    let mut engine = Engine::new();
    engine.set_external_value_provider(Some(Arc::new(Provider)));
    let before = engine.bytecode_program_count();
    engine
        .set_cell_formula("Sheet1", "A1", "=[Book.xlsx]Sheet1!A1")
        .unwrap();

    // Ensure we're exercising the bytecode path.
    let stats = engine.bytecode_compile_stats();
    assert_eq!(stats.total_formula_cells, 1);
    assert_eq!(stats.compiled, 1);
    assert_eq!(stats.fallback, 0);
    assert!(
        engine.bytecode_program_count() > before,
        "expected bytecode program cache to grow for external workbook reference"
    );

    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(123.0));
    assert!(engine.bytecode_program_count() > 0);
}

#[test]
fn bytecode_backend_external_workbook_reference_without_provider_is_ref_error() {
    let mut engine = Engine::new();
    let before = engine.bytecode_program_count();
    engine
        .set_cell_formula("Sheet1", "A1", "=[Book.xlsx]Sheet1!A1")
        .unwrap();

    // Ensure the formula still compiles to bytecode even without an external provider.
    let stats = engine.bytecode_compile_stats();
    assert_eq!(stats.total_formula_cells, 1);
    assert_eq!(stats.compiled, 1);
    assert_eq!(stats.fallback, 0);
    assert!(
        engine.bytecode_program_count() > before,
        "expected bytecode program cache to grow for external workbook reference"
    );

    engine.recalculate_single_threaded();

    assert_eq!(
        engine.get_cell_value("Sheet1", "A1"),
        Value::Error(ErrorKind::Ref)
    );
}

#[test]
fn bytecode_backend_external_workbook_range_spills_via_provider() {
    struct Provider;

    impl ExternalValueProvider for Provider {
        fn get(&self, sheet: &str, addr: formula_engine::eval::CellAddr) -> Option<Value> {
            if sheet != "[Book.xlsx]Sheet1" {
                return None;
            }
            match (addr.row, addr.col) {
                (0, 0) => Some(Value::Number(1.0)),
                (0, 1) => Some(Value::Number(2.0)),
                (1, 0) => Some(Value::Number(3.0)),
                (1, 1) => Some(Value::Number(4.0)),
                _ => None,
            }
        }
    }

    let mut engine = Engine::new();
    engine.set_external_value_provider(Some(Arc::new(Provider)));
    engine
        .set_cell_formula("Sheet1", "C1", "=[Book.xlsx]Sheet1!A1:B2")
        .unwrap();

    // Ensure the formula compiles to bytecode.
    let stats = engine.bytecode_compile_stats();
    assert_eq!(stats.total_formula_cells, 1);
    assert_eq!(stats.compiled, 1);
    assert_eq!(stats.fallback, 0);

    engine.recalculate_single_threaded();

    assert_eq!(
        engine.spill_range("Sheet1", "C1"),
        Some((
            formula_engine::eval::CellAddr { row: 0, col: 2 },
            formula_engine::eval::CellAddr { row: 1, col: 3 }
        ))
    );
    assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "D1"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "C2"), Value::Number(3.0));
    assert_eq!(engine.get_cell_value("Sheet1", "D2"), Value::Number(4.0));
}

#[test]
fn sum_ignores_rich_values_in_references_via_bytecode_column_slices() {
    // Regression coverage: column-slice evaluation should tolerate non-numeric values in the
    // referenced range (ignored like text in Excel SUM semantics).
    //
    // This is also the intended behavior for future rich values (Entity/Record), which should be
    // classified like text/bool in the bytecode column cache.
    for rich in [
        Value::Entity(EntityValue::new("Entity display")),
        Value::Record(RecordValue::new("Record display")),
    ] {
        let mut engine = Engine::new();

        engine.set_cell_value("Sheet1", "A1", rich).unwrap();
        engine.set_cell_value("Sheet1", "A2", 3.0).unwrap();
        engine
            .set_cell_formula("Sheet1", "B1", "=SUM(A1:A2)")
            .unwrap();

        // Ensure we're exercising the bytecode path.
        assert_eq!(engine.bytecode_program_count(), 1);

        engine.recalculate_single_threaded();
        assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(3.0));
        assert_engine_matches_ast(&engine, "=SUM(A1:A2)", "B1");
    }
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
fn bytecode_backend_sumproduct_broadcasts_single_cell_ranges() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 3.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", 4.0).unwrap();
    engine
        .set_cell_formula("Sheet1", "B1", "=SUMPRODUCT(A1,A1:A3)")
        .unwrap();

    assert_eq!(
        engine.bytecode_program_count(),
        1,
        "expected SUMPRODUCT scalar-broadcasting to compile to bytecode"
    );

    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(18.0));
    assert_engine_matches_ast(&engine, "=SUMPRODUCT(A1,A1:A3)", "B1");
}

#[test]
fn bytecode_backend_sumproduct_preserves_error_precedence_for_broadcast_ranges() {
    let mut engine = Engine::new();

    engine
        .set_cell_value("Sheet1", "A1", Value::Error(ErrorKind::Value))
        .unwrap();
    engine
        .set_cell_value("Sheet1", "B1", Value::Error(ErrorKind::Div0))
        .unwrap();
    engine.set_cell_value("Sheet1", "B2", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "B3", 3.0).unwrap();

    engine
        .set_cell_formula("Sheet1", "C1", "=SUMPRODUCT(B1:B3,A1)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "C2", "=SUMPRODUCT(A1,B1:B3)")
        .unwrap();

    assert_eq!(
        engine.bytecode_program_count(),
        2,
        "expected SUMPRODUCT broadcast formulas to compile to bytecode"
    );

    engine.recalculate_single_threaded();

    // With broadcast, error precedence should follow per-index coercion order: coerce the first
    // argument before the second for each element.
    assert_eq!(
        engine.get_cell_value("Sheet1", "C1"),
        Value::Error(ErrorKind::Div0)
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "C2"),
        Value::Error(ErrorKind::Value)
    );
    assert_engine_matches_ast(&engine, "=SUMPRODUCT(B1:B3,A1)", "C1");
    assert_engine_matches_ast(&engine, "=SUMPRODUCT(A1,B1:B3)", "C2");
}

#[test]
fn bytecode_backend_sumproduct_flattens_mismatched_range_shapes_by_length() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "B1", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "C1", 3.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 4.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", 5.0).unwrap();
    engine
        .set_cell_formula("Sheet1", "D1", "=SUMPRODUCT(A1:C1,A1:A3)")
        .unwrap();

    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "D1"), Value::Number(24.0));
    assert_engine_matches_ast(&engine, "=SUMPRODUCT(A1:C1,A1:A3)", "D1");
}

#[test]
fn bytecode_backend_sumproduct_accepts_scalars_and_array_literals() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=SUMPRODUCT({1,2},{3,4})")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A2", "=SUMPRODUCT(2,{1,2,3})")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A3", "=LET(a,{1,2},SUMPRODUCT(a,{3,4}))")
        .unwrap();

    assert_eq!(
        engine.bytecode_program_count(),
        3,
        "expected array/scalar SUMPRODUCT formulas to compile to bytecode"
    );

    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(11.0));
    assert_eq!(engine.get_cell_value("Sheet1", "A2"), Value::Number(12.0));
    assert_eq!(engine.get_cell_value("Sheet1", "A3"), Value::Number(11.0));
    assert_engine_matches_ast(&engine, "=SUMPRODUCT({1,2},{3,4})", "A1");
    assert_engine_matches_ast(&engine, "=SUMPRODUCT(2,{1,2,3})", "A2");
    assert_engine_matches_ast(&engine, "=LET(a,{1,2},SUMPRODUCT(a,{3,4}))", "A3");
}

#[test]
fn bytecode_backend_sumproduct_accepts_array_expressions_and_range_args() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", -2.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", 3.0).unwrap();
    engine.set_cell_value("Sheet1", "B1", 10.0).unwrap();
    engine.set_cell_value("Sheet1", "B2", 20.0).unwrap();
    engine.set_cell_value("Sheet1", "B3", 30.0).unwrap();
    engine
        .set_cell_formula("Sheet1", "C1", "=SUMPRODUCT(A1:A3>0,B1:B3)")
        .unwrap();

    assert_eq!(
        engine.bytecode_program_count(),
        1,
        "expected SUMPRODUCT array+range formulas to compile to bytecode"
    );

    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::Number(40.0));
    assert_engine_matches_ast(&engine, "=SUMPRODUCT(A1:A3>0,B1:B3)", "C1");
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
fn bytecode_backend_matches_ast_for_scalar_information_functions() {
    let mut engine = Engine::new();

    // A1 left blank
    engine.set_cell_value("Sheet1", "A2", 3.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", "hello").unwrap();
    engine.set_cell_value("Sheet1", "A4", true).unwrap();
    engine
        .set_cell_value("Sheet1", "A5", Value::Error(ErrorKind::Div0))
        .unwrap();
    engine
        .set_cell_value("Sheet1", "A6", Value::Error(ErrorKind::NA))
        .unwrap();

    // IS* functions.
    engine
        .set_cell_formula("Sheet1", "B1", "=ISBLANK(A1)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B2", "=ISNUMBER(A2)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B3", "=ISTEXT(A3)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B4", "=ISLOGICAL(A4)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B5", "=ISERR(A5)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B6", "=ISERR(A6)")
        .unwrap();

    // TYPE.
    engine
        .set_cell_formula("Sheet1", "C1", "=TYPE(A1)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "C2", "=TYPE(A2)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "C3", "=TYPE(A3)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "C4", "=TYPE(A4)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "C5", "=TYPE(A5)")
        .unwrap();

    // ERROR.TYPE.
    engine
        .set_cell_formula("Sheet1", "D2", "=ERROR.TYPE(A2)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "D5", "=ERROR.TYPE(A5)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "D6", "=ERROR.TYPE(A6)")
        .unwrap();

    // N.
    engine.set_cell_formula("Sheet1", "E1", "=N(A1)").unwrap();
    engine.set_cell_formula("Sheet1", "E2", "=N(A2)").unwrap();
    engine.set_cell_formula("Sheet1", "E3", "=N(A3)").unwrap();
    engine.set_cell_formula("Sheet1", "E4", "=N(A4)").unwrap();
    engine.set_cell_formula("Sheet1", "E5", "=N(A5)").unwrap();

    // T.
    engine.set_cell_formula("Sheet1", "F1", "=T(A1)").unwrap();
    engine.set_cell_formula("Sheet1", "F2", "=T(A2)").unwrap();
    engine.set_cell_formula("Sheet1", "F3", "=T(A3)").unwrap();
    engine.set_cell_formula("Sheet1", "F5", "=T(A5)").unwrap();

    // Ensure all supported functions were compiled to bytecode (one program per distinct formula
    // pattern/function name).
    assert_eq!(engine.bytecode_program_count(), 9);

    engine.recalculate_single_threaded();

    for (formula, cell) in [
        ("=ISBLANK(A1)", "B1"),
        ("=ISNUMBER(A2)", "B2"),
        ("=ISTEXT(A3)", "B3"),
        ("=ISLOGICAL(A4)", "B4"),
        ("=ISERR(A5)", "B5"),
        ("=ISERR(A6)", "B6"),
        ("=TYPE(A1)", "C1"),
        ("=TYPE(A2)", "C2"),
        ("=TYPE(A3)", "C3"),
        ("=TYPE(A4)", "C4"),
        ("=TYPE(A5)", "C5"),
        ("=ERROR.TYPE(A2)", "D2"),
        ("=ERROR.TYPE(A5)", "D5"),
        ("=ERROR.TYPE(A6)", "D6"),
        ("=N(A1)", "E1"),
        ("=N(A2)", "E2"),
        ("=N(A3)", "E3"),
        ("=N(A4)", "E4"),
        ("=N(A5)", "E5"),
        ("=T(A1)", "F1"),
        ("=T(A2)", "F2"),
        ("=T(A3)", "F3"),
        ("=T(A5)", "F5"),
    ] {
        assert_engine_matches_ast(&engine, formula, cell);
    }
}

#[test]
fn bytecode_backend_matches_ast_for_depreciation_functions() {
    let mut engine = Engine::new();

    engine
        .set_cell_formula("Sheet1", "A1", "=DB(10000,1000,5,1)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A2", "=DB(10000,1000,5,6,7)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A3", "=VDB(2400,300,10,0,0.5)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A4", "=VDB(2400,0,10,6,10,2,FALSE)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A5", "=VDB(2400,0,10,6,10,2,TRUE)")
        .unwrap();

    // Ensure we're exercising the bytecode path for all formulas.
    assert_eq!(engine.bytecode_program_count(), 5);

    engine.recalculate_single_threaded();

    for (formula, cell) in [
        ("=DB(10000,1000,5,1)", "A1"),
        ("=DB(10000,1000,5,6,7)", "A2"),
        ("=VDB(2400,300,10,0,0.5)", "A3"),
        ("=VDB(2400,0,10,6,10,2,FALSE)", "A4"),
        ("=VDB(2400,0,10,6,10,2,TRUE)", "A5"),
    ] {
        assert_engine_matches_ast(&engine, formula, cell);
    }
}

#[test]
fn bytecode_backend_matches_ast_for_coupon_schedule_functions() {
    let mut engine = Engine::new();

    engine
        .set_cell_formula(
            "Sheet1",
            "A1",
            r#"=COUPDAYBS("2020-02-01","2025-01-15",2,)"#,
        )
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A2", r#"=COUPDAYS("2020-02-01","2025-01-15",2,)"#)
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "A3",
            r#"=COUPDAYSNC("2020-02-01","2025-01-15",2,)"#,
        )
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A4", r#"=COUPNCD("2020-02-01","2025-01-15",2,)"#)
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A5", r#"=COUPNUM("2020-02-01","2025-01-15",2,)"#)
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A6", r#"=COUPPCD("2020-02-01","2025-01-15",2,)"#)
        .unwrap();

    let stats = engine.bytecode_compile_stats();
    assert_eq!(stats.total_formula_cells, 6);
    assert_eq!(stats.compiled, 6);
    assert_eq!(stats.fallback, 0);
    assert_eq!(engine.bytecode_program_count(), 6);

    engine.recalculate_single_threaded();

    for (formula, cell) in [
        (r#"=COUPDAYBS("2020-02-01","2025-01-15",2,)"#, "A1"),
        (r#"=COUPDAYS("2020-02-01","2025-01-15",2,)"#, "A2"),
        (r#"=COUPDAYSNC("2020-02-01","2025-01-15",2,)"#, "A3"),
        (r#"=COUPNCD("2020-02-01","2025-01-15",2,)"#, "A4"),
        (r#"=COUPNUM("2020-02-01","2025-01-15",2,)"#, "A5"),
        (r#"=COUPPCD("2020-02-01","2025-01-15",2,)"#, "A6"),
    ] {
        assert_engine_matches_ast(&engine, formula, cell);
    }
}

#[test]
fn bytecode_backend_coupon_basis_4_uses_fixed_period_length_and_preserves_additivity() {
    let mut engine = Engine::new();

    // This schedule exercises the basis=4 (European 30E/360) quirk where:
    // - Day counts use DAYS360(..., TRUE)
    // - But the modeled coupon period length E used by COUPDAYS is fixed as 360/frequency
    //   (and COUPDAYSNC is computed as E - A to preserve the additivity invariant).
    //
    // settlement=2020-11-15, maturity=2021-02-28, frequency=2:
    // PCD=2020-08-31, NCD=2021-02-28, DAYS360_EU(PCD,NCD)=178 != 180 (=360/frequency)
    // A=DAYS360_EU(PCD,settlement)=75, so COUPDAYSNC should be 180-75=105 (not 103).
    engine
        .set_cell_formula("Sheet1", "A1", r#"=COUPPCD("2020-11-15","2021-02-28",2,4)"#)
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A2", r#"=COUPNCD("2020-11-15","2021-02-28",2,4)"#)
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A3", r#"=COUPDAYS("2020-11-15","2021-02-28",2,4)"#)
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A4", r#"=COUPDAYBS("2020-11-15","2021-02-28",2,4)"#)
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A5", r#"=COUPDAYSNC("2020-11-15","2021-02-28",2,4)"#)
        .unwrap();

    let stats = engine.bytecode_compile_stats();
    assert_eq!(stats.total_formula_cells, 5);
    assert_eq!(stats.compiled, 5);
    assert_eq!(stats.fallback, 0);
    assert_eq!(engine.bytecode_program_count(), 5);

    engine.recalculate_single_threaded();

    let system = engine.date_system();
    let expected_pcd = ymd_to_serial(ExcelDate::new(2020, 8, 31), system).unwrap() as f64;
    let expected_ncd = ymd_to_serial(ExcelDate::new(2021, 2, 28), system).unwrap() as f64;

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(expected_pcd));
    assert_eq!(engine.get_cell_value("Sheet1", "A2"), Value::Number(expected_ncd));

    // COUPDAYS uses the fixed modeled coupon period length for basis=4.
    assert_eq!(engine.get_cell_value("Sheet1", "A3"), Value::Number(180.0));
    // COUPDAYBS uses European DAYS360 for basis=4.
    assert_eq!(engine.get_cell_value("Sheet1", "A4"), Value::Number(75.0));
    // COUPDAYSNC is computed as E - A for basis=4 (preserving additivity).
    assert_eq!(engine.get_cell_value("Sheet1", "A5"), Value::Number(105.0));

    // Guard: ensure we really hit a schedule where DAYS360_EU(PCD,NCD) differs from E.
    let pcd = expected_pcd as i32;
    let ncd = expected_ncd as i32;
    let days360_eu = formula_engine::functions::date_time::days360(pcd, ncd, true, system).unwrap();
    assert_eq!(days360_eu, 178);
    assert_ne!(days360_eu as f64, 180.0);

    // Also guard that COUPDAYSNC differs from the direct European day-count (settlement->NCD).
    let settlement = ymd_to_serial(ExcelDate::new(2020, 11, 15), system).unwrap();
    let dsc_eu =
        formula_engine::functions::date_time::days360(settlement, ncd, true, system).unwrap();
    assert_eq!(dsc_eu, 103);

    // Bytecode-vs-AST parity (and a sanity check that additivity holds).
    assert_engine_matches_ast(&engine, r#"=COUPPCD("2020-11-15","2021-02-28",2,4)"#, "A1");
    assert_engine_matches_ast(&engine, r#"=COUPNCD("2020-11-15","2021-02-28",2,4)"#, "A2");
    assert_engine_matches_ast(&engine, r#"=COUPDAYS("2020-11-15","2021-02-28",2,4)"#, "A3");
    assert_engine_matches_ast(&engine, r#"=COUPDAYBS("2020-11-15","2021-02-28",2,4)"#, "A4");
    assert_engine_matches_ast(&engine, r#"=COUPDAYSNC("2020-11-15","2021-02-28",2,4)"#, "A5");
    let Value::Number(days) = engine.get_cell_value("Sheet1", "A3") else {
        panic!("expected COUPDAYS to return a number");
    };
    let Value::Number(daybs) = engine.get_cell_value("Sheet1", "A4") else {
        panic!("expected COUPDAYBS to return a number");
    };
    let Value::Number(daysnc) = engine.get_cell_value("Sheet1", "A5") else {
        panic!("expected COUPDAYSNC to return a number");
    };
    assert_eq!(days, daybs + daysnc);
}

#[test]
fn bytecode_backend_coupon_basis_0_preserves_additivity_even_when_days360_is_not_additive() {
    let mut engine = Engine::new();

    // Same schedule as the basis=4 test above, but exercising the US/NASD 30/360 method (basis=0).
    // US DAYS360 is not additive for some end-of-month schedules (the end-date adjustment depends
    // on the start-date day). Excel preserves the invariant COUPDAYBS + COUPDAYSNC == COUPDAYS by
    // computing COUPDAYSNC as E - A (where E is fixed at 360/frequency).
    engine
        .set_cell_formula("Sheet1", "A1", r#"=COUPDAYS("2020-11-15","2021-02-28",2,0)"#)
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A2", r#"=COUPDAYBS("2020-11-15","2021-02-28",2,0)"#)
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A3", r#"=COUPDAYSNC("2020-11-15","2021-02-28",2,0)"#)
        .unwrap();

    let stats = engine.bytecode_compile_stats();
    assert_eq!(stats.total_formula_cells, 3);
    assert_eq!(stats.compiled, 3);
    assert_eq!(stats.fallback, 0);
    assert_eq!(engine.bytecode_program_count(), 3);

    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(180.0));
    assert_eq!(engine.get_cell_value("Sheet1", "A2"), Value::Number(75.0));
    assert_eq!(engine.get_cell_value("Sheet1", "A3"), Value::Number(105.0));

    let Value::Number(days) = engine.get_cell_value("Sheet1", "A1") else {
        panic!("expected COUPDAYS to return a number");
    };
    let Value::Number(daybs) = engine.get_cell_value("Sheet1", "A2") else {
        panic!("expected COUPDAYBS to return a number");
    };
    let Value::Number(daysnc) = engine.get_cell_value("Sheet1", "A3") else {
        panic!("expected COUPDAYSNC to return a number");
    };
    assert_eq!(days, daybs + daysnc);

    // Guard: the direct US/NASD DAYS360(settlement, NCD) differs from E-A here.
    let system = engine.date_system();
    let settlement = ymd_to_serial(ExcelDate::new(2020, 11, 15), system).unwrap();
    let ncd = ymd_to_serial(ExcelDate::new(2021, 2, 28), system).unwrap();
    let dsc_us =
        formula_engine::functions::date_time::days360(settlement, ncd, false, system).unwrap();
    assert_eq!(dsc_us, 106);
    assert_ne!(dsc_us as f64, daysnc);

    assert_engine_matches_ast(&engine, r#"=COUPDAYS("2020-11-15","2021-02-28",2,0)"#, "A1");
    assert_engine_matches_ast(&engine, r#"=COUPDAYBS("2020-11-15","2021-02-28",2,0)"#, "A2");
    assert_engine_matches_ast(&engine, r#"=COUPDAYSNC("2020-11-15","2021-02-28",2,0)"#, "A3");
}

#[test]
fn bytecode_backend_matches_ast_for_standard_bond_functions() {
    let mut engine = Engine::new();

    engine
        .set_cell_formula(
            "Sheet1",
            "B1",
            r#"=PRICE("2020-02-01","2025-01-15",0.05,0.04,100,2,)"#,
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "B2",
            r#"=YIELD("2020-02-01","2025-01-15",0.05,95,100,2,)"#,
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "B3",
            r#"=DURATION("2020-02-01","2025-01-15",0.05,0.04,2,)"#,
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "B4",
            r#"=MDURATION("2020-02-01","2025-01-15",0.05,0.04,2,)"#,
        )
        .unwrap();

    let stats = engine.bytecode_compile_stats();
    assert_eq!(stats.total_formula_cells, 4);
    assert_eq!(stats.compiled, 4);
    assert_eq!(stats.fallback, 0);
    assert_eq!(engine.bytecode_program_count(), 4);

    engine.recalculate_single_threaded();

    for (formula, cell) in [
        (
            r#"=PRICE("2020-02-01","2025-01-15",0.05,0.04,100,2,)"#,
            "B1",
        ),
        (r#"=YIELD("2020-02-01","2025-01-15",0.05,95,100,2,)"#, "B2"),
        (r#"=DURATION("2020-02-01","2025-01-15",0.05,0.04,2,)"#, "B3"),
        (
            r#"=MDURATION("2020-02-01","2025-01-15",0.05,0.04,2,)"#,
            "B4",
        ),
    ] {
        assert_engine_matches_ast(&engine, formula, cell);
    }
}

#[test]
fn bytecode_backend_matches_ast_for_accrued_interest_functions() {
    let mut engine = Engine::new();

    engine
        .set_cell_formula(
            "Sheet1",
            "C1",
            r#"=ACCRINTM("2019-12-31","2020-03-31",0.05,1000,)"#,
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "C2",
            r#"=ACCRINT("2019-12-31","2020-06-30","2020-03-31",0.05,1000,2,,TRUE)"#,
        )
        .unwrap();

    let stats = engine.bytecode_compile_stats();
    assert_eq!(stats.total_formula_cells, 2);
    assert_eq!(stats.compiled, 2);
    assert_eq!(stats.fallback, 0);
    assert_eq!(engine.bytecode_program_count(), 2);

    engine.recalculate_single_threaded();

    for (formula, cell) in [
        (r#"=ACCRINTM("2019-12-31","2020-03-31",0.05,1000,)"#, "C1"),
        (
            r#"=ACCRINT("2019-12-31","2020-06-30","2020-03-31",0.05,1000,2,,TRUE)"#,
            "C2",
        ),
    ] {
        assert_engine_matches_ast(&engine, formula, cell);
    }
}

#[test]
fn bytecode_backend_matches_ast_for_discount_securities_and_tbill_functions() {
    let mut engine = Engine::new();

    engine
        .set_cell_formula(
            "Sheet1",
            "D1",
            r#"=DISC("2020-01-01","2020-12-31",97,100,)"#,
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "D2",
            r#"=PRICEDISC("2020-01-01","2020-12-31",0.05,100,)"#,
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "D3",
            r#"=YIELDDISC("2020-01-01","2020-12-31",97,100,)"#,
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "D4",
            r#"=INTRATE("2020-01-01","2020-12-31",97,100,)"#,
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "D5",
            r#"=RECEIVED("2020-01-01","2020-12-31",97,0.05,)"#,
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "D6",
            r#"=PRICEMAT("2020-01-01","2020-12-31","2019-12-31",0.05,0.04,)"#,
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "D7",
            r#"=YIELDMAT("2020-01-01","2020-12-31","2019-12-31",0.05,95,)"#,
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "D8",
            r#"=TBILLEQ("2020-01-01","2020-06-30",0.05)"#,
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "D9",
            r#"=TBILLPRICE("2020-01-01","2020-06-30",0.05)"#,
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "D10",
            r#"=TBILLYIELD("2020-01-01","2020-06-30",97)"#,
        )
        .unwrap();

    let stats = engine.bytecode_compile_stats();
    assert_eq!(stats.total_formula_cells, 10);
    assert_eq!(stats.compiled, 10);
    assert_eq!(stats.fallback, 0);
    assert_eq!(engine.bytecode_program_count(), 10);

    engine.recalculate_single_threaded();

    for (formula, cell) in [
        (r#"=DISC("2020-01-01","2020-12-31",97,100,)"#, "D1"),
        (r#"=PRICEDISC("2020-01-01","2020-12-31",0.05,100,)"#, "D2"),
        (r#"=YIELDDISC("2020-01-01","2020-12-31",97,100,)"#, "D3"),
        (r#"=INTRATE("2020-01-01","2020-12-31",97,100,)"#, "D4"),
        (r#"=RECEIVED("2020-01-01","2020-12-31",97,0.05,)"#, "D5"),
        (
            r#"=PRICEMAT("2020-01-01","2020-12-31","2019-12-31",0.05,0.04,)"#,
            "D6",
        ),
        (
            r#"=YIELDMAT("2020-01-01","2020-12-31","2019-12-31",0.05,95,)"#,
            "D7",
        ),
        (r#"=TBILLEQ("2020-01-01","2020-06-30",0.05)"#, "D8"),
        (r#"=TBILLPRICE("2020-01-01","2020-06-30",0.05)"#, "D9"),
        (r#"=TBILLYIELD("2020-01-01","2020-06-30",97)"#, "D10"),
    ] {
        assert_engine_matches_ast(&engine, formula, cell);
    }
}

#[test]
fn bytecode_backend_financial_functions_compile_with_omitted_optional_args() {
    let mut engine = Engine::new();

    // COUP* functions: omit `basis`.
    engine
        .set_cell_formula("Sheet1", "A1", r#"=COUPDAYBS("2020-02-01","2025-01-15",2)"#)
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A2", r#"=COUPDAYS("2020-02-01","2025-01-15",2)"#)
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A3", r#"=COUPDAYSNC("2020-02-01","2025-01-15",2)"#)
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A4", r#"=COUPNCD("2020-02-01","2025-01-15",2)"#)
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A5", r#"=COUPNUM("2020-02-01","2025-01-15",2)"#)
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A6", r#"=COUPPCD("2020-02-01","2025-01-15",2)"#)
        .unwrap();

    // Standard bond functions: omit `basis`.
    engine
        .set_cell_formula(
            "Sheet1",
            "B1",
            r#"=PRICE("2020-02-01","2025-01-15",0.05,0.04,100,2)"#,
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "B2",
            r#"=YIELD("2020-02-01","2025-01-15",0.05,95,100,2)"#,
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "B3",
            r#"=DURATION("2020-02-01","2025-01-15",0.05,0.04,2)"#,
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "B4",
            r#"=MDURATION("2020-02-01","2025-01-15",0.05,0.04,2)"#,
        )
        .unwrap();

    // Accrued interest functions: omit `basis` and `calc_method`.
    engine
        .set_cell_formula(
            "Sheet1",
            "C1",
            r#"=ACCRINTM("2019-12-31","2020-03-31",0.05,1000)"#,
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "C2",
            r#"=ACCRINT("2019-12-31","2020-06-30","2020-03-31",0.05,1000,2)"#,
        )
        .unwrap();

    // Discount securities: omit `basis`.
    engine
        .set_cell_formula("Sheet1", "D1", r#"=DISC("2020-01-01","2020-12-31",97,100)"#)
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "D2",
            r#"=PRICEDISC("2020-01-01","2020-12-31",0.05,100)"#,
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "D3",
            r#"=YIELDDISC("2020-01-01","2020-12-31",97,100)"#,
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "D4",
            r#"=INTRATE("2020-01-01","2020-12-31",97,100)"#,
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "D5",
            r#"=RECEIVED("2020-01-01","2020-12-31",97,0.05)"#,
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "D6",
            r#"=PRICEMAT("2020-01-01","2020-12-31","2019-12-31",0.05,0.04)"#,
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "D7",
            r#"=YIELDMAT("2020-01-01","2020-12-31","2019-12-31",0.05,95)"#,
        )
        .unwrap();

    let stats = engine.bytecode_compile_stats();
    assert_eq!(stats.total_formula_cells, 19);
    assert_eq!(stats.compiled, 19);
    assert_eq!(stats.fallback, 0);

    engine.recalculate_single_threaded();

    for (formula, cell) in [
        // COUP*
        (r#"=COUPDAYBS("2020-02-01","2025-01-15",2)"#, "A1"),
        (r#"=COUPDAYS("2020-02-01","2025-01-15",2)"#, "A2"),
        (r#"=COUPDAYSNC("2020-02-01","2025-01-15",2)"#, "A3"),
        (r#"=COUPNCD("2020-02-01","2025-01-15",2)"#, "A4"),
        (r#"=COUPNUM("2020-02-01","2025-01-15",2)"#, "A5"),
        (r#"=COUPPCD("2020-02-01","2025-01-15",2)"#, "A6"),
        // Standard bonds
        (
            r#"=PRICE("2020-02-01","2025-01-15",0.05,0.04,100,2)"#,
            "B1",
        ),
        (
            r#"=YIELD("2020-02-01","2025-01-15",0.05,95,100,2)"#,
            "B2",
        ),
        (
            r#"=DURATION("2020-02-01","2025-01-15",0.05,0.04,2)"#,
            "B3",
        ),
        (
            r#"=MDURATION("2020-02-01","2025-01-15",0.05,0.04,2)"#,
            "B4",
        ),
        // ACCRINT*
        (
            r#"=ACCRINTM("2019-12-31","2020-03-31",0.05,1000)"#,
            "C1",
        ),
        (
            r#"=ACCRINT("2019-12-31","2020-06-30","2020-03-31",0.05,1000,2)"#,
            "C2",
        ),
        // Discount securities
        (r#"=DISC("2020-01-01","2020-12-31",97,100)"#, "D1"),
        (
            r#"=PRICEDISC("2020-01-01","2020-12-31",0.05,100)"#,
            "D2",
        ),
        (r#"=YIELDDISC("2020-01-01","2020-12-31",97,100)"#, "D3"),
        (r#"=INTRATE("2020-01-01","2020-12-31",97,100)"#, "D4"),
        (r#"=RECEIVED("2020-01-01","2020-12-31",97,0.05)"#, "D5"),
        (
            r#"=PRICEMAT("2020-01-01","2020-12-31","2019-12-31",0.05,0.04)"#,
            "D6",
        ),
        (
            r#"=YIELDMAT("2020-01-01","2020-12-31","2019-12-31",0.05,95)"#,
            "D7",
        ),
    ] {
        assert_engine_matches_ast(&engine, formula, cell);
    }
}

#[test]
fn bytecode_backend_financial_date_text_uses_datevalue_semantics() {
    let mut engine = Engine::new();
    // For these financial functions, text date arguments must use DATEVALUE-like parsing (not VALUE).
    // A numeric-looking string like "1" should be rejected as a date string (#VALUE!).
    engine
        .set_cell_formula("Sheet1", "A1", r#"=COUPDAYBS("1","2025-01-15",2)"#)
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "A2",
            r#"=PRICE("1","2025-01-15",0.05,0.04,100,2)"#,
        )
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A3", r#"=DISC("1","2020-12-31",97,100)"#)
        .unwrap();

    let stats = engine.bytecode_compile_stats();
    assert_eq!(stats.total_formula_cells, 3);
    assert_eq!(stats.compiled, 3);
    assert_eq!(stats.fallback, 0);
    assert_eq!(engine.bytecode_program_count(), 3);

    engine.recalculate_single_threaded();
    for cell in ["A1", "A2", "A3"] {
        assert_eq!(
            engine.get_cell_value("Sheet1", cell),
            Value::Error(ErrorKind::Value),
            "expected {cell} to error for invalid text date input"
        );
    }

    for (formula, cell) in [
        (r#"=COUPDAYBS("1","2025-01-15",2)"#, "A1"),
        (r#"=PRICE("1","2025-01-15",0.05,0.04,100,2)"#, "A2"),
        (r#"=DISC("1","2020-12-31",97,100)"#, "A3"),
    ] {
        assert_engine_matches_ast(&engine, formula, cell);
    }
}

#[test]
fn bytecode_backend_financial_date_text_uses_datevalue_semantics_for_accrint() {
    let mut engine = Engine::new();

    engine
        .set_cell_formula("Sheet1", "A1", r#"=ACCRINTM("1","2020-03-31",0.05,1000)"#)
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "A2",
            r#"=ACCRINT("1","2020-06-30","2020-03-31",0.05,1000,2)"#,
        )
        .unwrap();

    let stats = engine.bytecode_compile_stats();
    assert_eq!(stats.total_formula_cells, 2);
    assert_eq!(stats.compiled, 2);
    assert_eq!(stats.fallback, 0);
    assert_eq!(engine.bytecode_program_count(), 2);

    engine.recalculate_single_threaded();

    for cell in ["A1", "A2"] {
        assert_eq!(
            engine.get_cell_value("Sheet1", cell),
            Value::Error(ErrorKind::Value),
            "expected {cell} to error for invalid text date input"
        );
    }

    for (formula, cell) in [
        (r#"=ACCRINTM("1","2020-03-31",0.05,1000)"#, "A1"),
        (
            r#"=ACCRINT("1","2020-06-30","2020-03-31",0.05,1000,2)"#,
            "A2",
        ),
    ] {
        assert_engine_matches_ast(&engine, formula, cell);
    }
}

#[test]
fn bytecode_backend_discount_security_text_dates_reject_numeric_strings() {
    let mut engine = Engine::new();

    // Discount security + T-Bill functions should use DATEVALUE-style parsing for *text* date
    // arguments. Numeric-looking strings like "1" must be rejected as dates (#VALUE!), not coerced
    // as date serials.
    engine
        .set_cell_formula("Sheet1", "A1", r#"=TBILLPRICE("1","2020-06-30",0.05)"#)
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A2", r#"=TBILLYIELD("1","2020-06-30",97)"#)
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A3", r#"=TBILLEQ("1","2020-06-30",0.05)"#)
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A4", r#"=PRICEDISC("1","2020-12-31",0.05,100)"#)
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A5", r#"=YIELDDISC("1","2020-12-31",97,100)"#)
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A6", r#"=INTRATE("1","2020-12-31",97,100)"#)
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A7", r#"=RECEIVED("1","2020-12-31",97,0.05)"#)
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "A8",
            r#"=PRICEMAT("1","2020-12-31","2019-12-31",0.05,0.04)"#,
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "A9",
            r#"=YIELDMAT("1","2020-12-31","2019-12-31",0.05,95)"#,
        )
        .unwrap();

    let stats = engine.bytecode_compile_stats();
    assert_eq!(stats.total_formula_cells, 9);
    assert_eq!(stats.compiled, 9);
    assert_eq!(stats.fallback, 0);
    assert_eq!(engine.bytecode_program_count(), 9);

    engine.recalculate_single_threaded();

    for cell in ["A1", "A2", "A3", "A4", "A5", "A6", "A7", "A8", "A9"] {
        assert_eq!(
            engine.get_cell_value("Sheet1", cell),
            Value::Error(ErrorKind::Value),
            "expected {cell} to error for invalid text date input"
        );
    }

    for (formula, cell) in [
        (r#"=TBILLPRICE("1","2020-06-30",0.05)"#, "A1"),
        (r#"=TBILLYIELD("1","2020-06-30",97)"#, "A2"),
        (r#"=TBILLEQ("1","2020-06-30",0.05)"#, "A3"),
        (r#"=PRICEDISC("1","2020-12-31",0.05,100)"#, "A4"),
        (r#"=YIELDDISC("1","2020-12-31",97,100)"#, "A5"),
        (r#"=INTRATE("1","2020-12-31",97,100)"#, "A6"),
        (r#"=RECEIVED("1","2020-12-31",97,0.05)"#, "A7"),
        (
            r#"=PRICEMAT("1","2020-12-31","2019-12-31",0.05,0.04)"#,
            "A8",
        ),
        (
            r#"=YIELDMAT("1","2020-12-31","2019-12-31",0.05,95)"#,
            "A9",
        ),
    ] {
        assert_engine_matches_ast(&engine, formula, cell);
    }
}

#[test]
fn bytecode_backend_bond_text_dates_reject_numeric_strings() {
    let mut engine = Engine::new();

    // Standard and odd-coupon bond functions should use DATEVALUE-style parsing for *text* date
    // arguments. Numeric-looking strings like "1" must be rejected as dates (#VALUE!), not coerced
    // as date serials.
    engine
        .set_cell_formula("Sheet1", "A1", r#"=YIELD("1","2025-01-15",0.05,95,100,2)"#)
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A2", r#"=DURATION("1","2025-01-15",0.05,0.04,2)"#)
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A3", r#"=MDURATION("1","2025-01-15",0.05,0.04,2)"#)
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "A4",
            r#"=ODDFPRICE("1","2025-01-01","2020-01-01","2020-07-01",0.05,0.04,100,2)"#,
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "A5",
            r#"=ODDFYIELD("1","2025-01-01","2020-01-01","2020-07-01",0.05,95,100,2)"#,
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "A6",
            r#"=ODDLPRICE("1","2021-03-01","2020-10-15",0.05,0.06,100,2)"#,
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "A7",
            r#"=ODDLYIELD("1","2021-03-01","2020-10-15",0.05,95,100,2)"#,
        )
        .unwrap();

    let stats = engine.bytecode_compile_stats();
    assert_eq!(stats.total_formula_cells, 7);
    assert_eq!(stats.compiled, 7);
    assert_eq!(stats.fallback, 0);
    assert_eq!(engine.bytecode_program_count(), 7);

    engine.recalculate_single_threaded();

    for cell in ["A1", "A2", "A3", "A4", "A5", "A6", "A7"] {
        assert_eq!(
            engine.get_cell_value("Sheet1", cell),
            Value::Error(ErrorKind::Value),
            "expected {cell} to error for invalid text date input"
        );
    }

    for (formula, cell) in [
        (r#"=YIELD("1","2025-01-15",0.05,95,100,2)"#, "A1"),
        (r#"=DURATION("1","2025-01-15",0.05,0.04,2)"#, "A2"),
        (r#"=MDURATION("1","2025-01-15",0.05,0.04,2)"#, "A3"),
        (
            r#"=ODDFPRICE("1","2025-01-01","2020-01-01","2020-07-01",0.05,0.04,100,2)"#,
            "A4",
        ),
        (
            r#"=ODDFYIELD("1","2025-01-01","2020-01-01","2020-07-01",0.05,95,100,2)"#,
            "A5",
        ),
        (
            r#"=ODDLPRICE("1","2021-03-01","2020-10-15",0.05,0.06,100,2)"#,
            "A6",
        ),
        (
            r#"=ODDLYIELD("1","2021-03-01","2020-10-15",0.05,95,100,2)"#,
            "A7",
        ),
    ] {
        assert_engine_matches_ast(&engine, formula, cell);
    }
}

#[test]
fn bytecode_backend_bond_numeric_text_respects_value_locale() {
    let mut engine = Engine::new();
    engine.set_value_locale(ValueLocaleConfig::de_de());

    // Use ISO dates to avoid locale-dependent date order ambiguity; focus this test on numeric text
    // coercion ("0,05" -> 0.05 under de-DE).
    engine
        .set_cell_formula(
            "Sheet1",
            "A1",
            r#"=YIELD("2020-02-01","2025-01-15","0,05","95","100","2")"#,
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "A2",
            r#"=YIELD("2020-02-01","2025-01-15",0.05,95,100,2)"#,
        )
        .unwrap();

    let stats = engine.bytecode_compile_stats();
    assert_eq!(stats.total_formula_cells, 2);
    assert_eq!(stats.compiled, 2);
    assert_eq!(stats.fallback, 0);

    engine.recalculate_single_threaded();

    let Value::Number(text_coerced) = engine.get_cell_value("Sheet1", "A1") else {
        panic!("expected YIELD with numeric text args to return a number");
    };
    let Value::Number(numeric_literal) = engine.get_cell_value("Sheet1", "A2") else {
        panic!("expected YIELD with numeric literals to return a number");
    };
    assert!(
        (text_coerced - numeric_literal).abs() <= 1e-12,
        "expected numeric text coercion to match numeric literals under de-DE; got {text_coerced} vs {numeric_literal}"
    );

    assert_engine_matches_ast(
        &engine,
        r#"=YIELD("2020-02-01","2025-01-15","0,05","95","100","2")"#,
        "A1",
    );
    assert_engine_matches_ast(
        &engine,
        r#"=YIELD("2020-02-01","2025-01-15",0.05,95,100,2)"#,
        "A2",
    );
}

#[test]
fn bytecode_backend_tbill_numeric_text_respects_value_locale() {
    let mut engine = Engine::new();
    engine.set_value_locale(ValueLocaleConfig::de_de());

    engine
        .set_cell_formula(
            "Sheet1",
            "A1",
            r#"=TBILLPRICE("2020-01-01","2020-07-01","0,05")"#,
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "A2",
            r#"=TBILLPRICE("2020-01-01","2020-07-01",0.05)"#,
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "A3",
            r#"=TBILLYIELD("2020-01-01","2020-07-01","97,5")"#,
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "A4",
            r#"=TBILLYIELD("2020-01-01","2020-07-01",97.5)"#,
        )
        .unwrap();

    let stats = engine.bytecode_compile_stats();
    assert_eq!(stats.total_formula_cells, 4);
    assert_eq!(stats.compiled, 4);
    assert_eq!(stats.fallback, 0);

    engine.recalculate_single_threaded();

    let Value::Number(p_text) = engine.get_cell_value("Sheet1", "A1") else {
        panic!("expected TBILLPRICE with numeric text to return a number");
    };
    let Value::Number(p_num) = engine.get_cell_value("Sheet1", "A2") else {
        panic!("expected TBILLPRICE with numeric literals to return a number");
    };
    assert!(
        (p_text - p_num).abs() <= 1e-12,
        "expected TBILLPRICE numeric text coercion to match numeric literals; got {p_text} vs {p_num}"
    );

    let Value::Number(y_text) = engine.get_cell_value("Sheet1", "A3") else {
        panic!("expected TBILLYIELD with numeric text to return a number");
    };
    let Value::Number(y_num) = engine.get_cell_value("Sheet1", "A4") else {
        panic!("expected TBILLYIELD with numeric literals to return a number");
    };
    assert!(
        (y_text - y_num).abs() <= 1e-12,
        "expected TBILLYIELD numeric text coercion to match numeric literals; got {y_text} vs {y_num}"
    );

    assert_engine_matches_ast(
        &engine,
        r#"=TBILLPRICE("2020-01-01","2020-07-01","0,05")"#,
        "A1",
    );
    assert_engine_matches_ast(
        &engine,
        r#"=TBILLPRICE("2020-01-01","2020-07-01",0.05)"#,
        "A2",
    );
    assert_engine_matches_ast(
        &engine,
        r#"=TBILLYIELD("2020-01-01","2020-07-01","97,5")"#,
        "A3",
    );
    assert_engine_matches_ast(
        &engine,
        r#"=TBILLYIELD("2020-01-01","2020-07-01",97.5)"#,
        "A4",
    );
}

#[test]
fn bytecode_backend_discount_securities_numeric_text_respects_value_locale() {
    let mut engine = Engine::new();
    engine.set_value_locale(ValueLocaleConfig::de_de());

    // Use ISO date text to avoid locale-dependent date order ambiguity; focus this test on numeric
    // text coercion ("," decimal separator under de-DE).
    engine
        .set_cell_formula(
            "Sheet1",
            "A1",
            r#"=PRICEDISC("2020-01-01","2020-12-31","0,05",100)"#,
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "A2",
            r#"=PRICEDISC("2020-01-01","2020-12-31",0.05,100)"#,
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "A3",
            r#"=DISC("2020-01-01","2020-12-31","97,5",100)"#,
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "A4",
            r#"=DISC("2020-01-01","2020-12-31",97.5,100)"#,
        )
        .unwrap();

    let stats = engine.bytecode_compile_stats();
    assert_eq!(stats.total_formula_cells, 4);
    assert_eq!(stats.compiled, 4);
    assert_eq!(stats.fallback, 0);

    engine.recalculate_single_threaded();

    let Value::Number(pd_text) = engine.get_cell_value("Sheet1", "A1") else {
        panic!("expected PRICEDISC with numeric text to return a number");
    };
    let Value::Number(pd_num) = engine.get_cell_value("Sheet1", "A2") else {
        panic!("expected PRICEDISC with numeric literals to return a number");
    };
    assert!(
        (pd_text - pd_num).abs() <= 1e-12,
        "expected PRICEDISC numeric text coercion to match numeric literals; got {pd_text} vs {pd_num}"
    );

    let Value::Number(d_text) = engine.get_cell_value("Sheet1", "A3") else {
        panic!("expected DISC with numeric text to return a number");
    };
    let Value::Number(d_num) = engine.get_cell_value("Sheet1", "A4") else {
        panic!("expected DISC with numeric literals to return a number");
    };
    assert!(
        (d_text - d_num).abs() <= 1e-12,
        "expected DISC numeric text coercion to match numeric literals; got {d_text} vs {d_num}"
    );

    for (formula, cell) in [
        (r#"=PRICEDISC("2020-01-01","2020-12-31","0,05",100)"#, "A1"),
        (r#"=PRICEDISC("2020-01-01","2020-12-31",0.05,100)"#, "A2"),
        (r#"=DISC("2020-01-01","2020-12-31","97,5",100)"#, "A3"),
        (r#"=DISC("2020-01-01","2020-12-31",97.5,100)"#, "A4"),
    ] {
        assert_engine_matches_ast(&engine, formula, cell);
    }
}

#[test]
fn bytecode_backend_standard_bond_basis_0_and_4_use_different_day_counts() {
    let mut engine = Engine::new();

    // Schedule designed so US vs EU DAYS360 diverge:
    // maturity is month-end Aug 31, so the maturity-anchored schedule is EOM-pinned and the PCD is
    // Feb 28. From Feb 28 to May 1, US DAYS360 counts 61 days, while EU counts 63.
    let settlement = "2021-05-01";
    let maturity = "2021-08-31";
    let rate = 0.05;
    let yld = 0.06;
    let redemption = 100.0;
    let frequency = 2;

    engine
        .set_cell_formula(
            "Sheet1",
            "A1",
            &format!(r#"=PRICE("{settlement}","{maturity}",{rate},{yld},{redemption},{frequency},0)"#),
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "A2",
            &format!(r#"=PRICE("{settlement}","{maturity}",{rate},{yld},{redemption},{frequency},4)"#),
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "B1",
            &format!(
                r#"=YIELD("{settlement}","{maturity}",{rate},A1,{redemption},{frequency},0)"#
            ),
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "B2",
            &format!(
                r#"=YIELD("{settlement}","{maturity}",{rate},A2,{redemption},{frequency},4)"#
            ),
        )
        .unwrap();

    let stats = engine.bytecode_compile_stats();
    assert_eq!(stats.total_formula_cells, 4);
    assert_eq!(stats.compiled, 4);
    assert_eq!(stats.fallback, 0);
    assert_eq!(engine.bytecode_program_count(), 4);

    engine.recalculate_single_threaded();

    let Value::Number(price0) = engine.get_cell_value("Sheet1", "A1") else {
        panic!("expected PRICE basis=0 to return a number");
    };
    let Value::Number(price4) = engine.get_cell_value("Sheet1", "A2") else {
        panic!("expected PRICE basis=4 to return a number");
    };
    assert!(
        price0.is_finite() && price4.is_finite(),
        "expected finite prices"
    );
    assert!(
        (price0 - price4).abs() > 0.0,
        "expected PRICE basis=0 vs basis=4 to differ for this Feb/EOM schedule"
    );

    let Value::Number(y0) = engine.get_cell_value("Sheet1", "B1") else {
        panic!("expected YIELD basis=0 to return a number");
    };
    let Value::Number(y4) = engine.get_cell_value("Sheet1", "B2") else {
        panic!("expected YIELD basis=4 to return a number");
    };
    assert!(
        (y0 - yld).abs() <= 1e-10,
        "expected YIELD basis=0 to recover yld={yld}, got {y0}"
    );
    assert!(
        (y4 - yld).abs() <= 1e-10,
        "expected YIELD basis=4 to recover yld={yld}, got {y4}"
    );

    // Guard: ensure US vs EU DAYS360 actually diverge for the PCD->settlement day count.
    let system = engine.date_system();
    let pcd = ymd_to_serial(ExcelDate::new(2021, 2, 28), system).unwrap();
    let settlement_serial = ymd_to_serial(ExcelDate::new(2021, 5, 1), system).unwrap();
    let a_us = formula_engine::functions::date_time::days360(pcd, settlement_serial, false, system)
        .unwrap();
    let a_eu =
        formula_engine::functions::date_time::days360(pcd, settlement_serial, true, system).unwrap();
    assert_eq!(a_us, 61);
    assert_eq!(a_eu, 63);

    // Bytecode-vs-AST parity.
    assert_engine_matches_ast(
        &engine,
        &format!(r#"=PRICE("{settlement}","{maturity}",{rate},{yld},{redemption},{frequency},0)"#),
        "A1",
    );
    assert_engine_matches_ast(
        &engine,
        &format!(r#"=PRICE("{settlement}","{maturity}",{rate},{yld},{redemption},{frequency},4)"#),
        "A2",
    );
    assert_engine_matches_ast(
        &engine,
        &format!(
            r#"=YIELD("{settlement}","{maturity}",{rate},A1,{redemption},{frequency},0)"#
        ),
        "B1",
    );
    assert_engine_matches_ast(
        &engine,
        &format!(
            r#"=YIELD("{settlement}","{maturity}",{rate},A2,{redemption},{frequency},4)"#
        ),
        "B2",
    );
}

#[test]
fn bytecode_backend_financial_date_text_respects_value_locale() {
    let mut engine = Engine::new();
    let formula = r#"=DISC("01/02/2020","01/03/2020",97,100)"#;
    engine.set_cell_formula("Sheet1", "A1", formula).unwrap();

    // Ensure we're exercising the bytecode path.
    let stats = engine.bytecode_compile_stats();
    assert_eq!(stats.total_formula_cells, 1);
    assert_eq!(stats.compiled, 1);
    assert_eq!(stats.fallback, 0);
    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();

    // en-US (default) is MDY: "01/02/2020" == Jan 2, 2020.
    let system = engine.date_system();
    let settlement_us = ymd_to_serial(ExcelDate::new(2020, 1, 2), system).unwrap();
    let maturity_us = ymd_to_serial(ExcelDate::new(2020, 1, 3), system).unwrap();
    let expected_us =
        formula_engine::functions::financial::disc(settlement_us, maturity_us, 97.0, 100.0, 0, system)
            .unwrap();
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(expected_us));
    assert_engine_matches_ast(&engine, formula, "A1");

    // Switch to a DMY locale and ensure the same text dates are parsed differently.
    assert!(engine.set_value_locale_id("de-DE"));
    engine.recalculate_single_threaded();
    assert_eq!(engine.bytecode_program_count(), 1);

    // de-DE is DMY: "01/02/2020" == Feb 1, 2020.
    let settlement_de = ymd_to_serial(ExcelDate::new(2020, 2, 1), system).unwrap();
    let maturity_de = ymd_to_serial(ExcelDate::new(2020, 3, 1), system).unwrap();
    let expected_de =
        formula_engine::functions::financial::disc(settlement_de, maturity_de, 97.0, 100.0, 0, system)
            .unwrap();
    assert_ne!(expected_de, expected_us);
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(expected_de));
    assert_engine_matches_ast(&engine, formula, "A1");
}

#[test]
fn bytecode_backend_coupon_date_text_respects_value_locale() {
    let mut engine = Engine::new();
    let formula = r#"=COUPDAYBS("01/02/2020","01/03/2020",2)"#;
    engine.set_cell_formula("Sheet1", "A1", formula).unwrap();

    // Ensure we're exercising the bytecode path.
    let stats = engine.bytecode_compile_stats();
    assert_eq!(stats.total_formula_cells, 1);
    assert_eq!(stats.compiled, 1);
    assert_eq!(stats.fallback, 0);
    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();

    // en-US (default) is MDY: "01/02/2020" == Jan 2, 2020.
    let system = engine.date_system();
    let settlement_us = ymd_to_serial(ExcelDate::new(2020, 1, 2), system).unwrap();
    let maturity_us = ymd_to_serial(ExcelDate::new(2020, 1, 3), system).unwrap();
    let expected_us =
        formula_engine::functions::financial::coupdaybs(settlement_us, maturity_us, 2, 0, system)
            .unwrap();
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(expected_us));
    assert_engine_matches_ast(&engine, formula, "A1");

    // Switch to a DMY locale and ensure the same text dates are parsed differently.
    assert!(engine.set_value_locale_id("de-DE"));
    engine.recalculate_single_threaded();
    assert_eq!(engine.bytecode_program_count(), 1);

    // de-DE is DMY: "01/02/2020" == Feb 1, 2020.
    let settlement_de = ymd_to_serial(ExcelDate::new(2020, 2, 1), system).unwrap();
    let maturity_de = ymd_to_serial(ExcelDate::new(2020, 3, 1), system).unwrap();
    let expected_de =
        formula_engine::functions::financial::coupdaybs(settlement_de, maturity_de, 2, 0, system)
            .unwrap();
    assert_ne!(expected_de, expected_us);
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(expected_de));
    assert_engine_matches_ast(&engine, formula, "A1");
}

#[test]
fn bytecode_backend_financial_date_text_respects_engine_date_system() {
    let mut engine = Engine::new();
    // Excel's 1900 date system can (optionally) emulate the Lotus 1-2-3 leap-year bug and accept
    // the fictitious date 1900-02-29. Ensure financial functions' text date coercion respects the
    // workbook date system.
    let formula = r#"=DISC("1900-02-29","1900-03-01",97,100)"#;
    engine.set_cell_formula("Sheet1", "A1", formula).unwrap();

    // Ensure we're exercising the bytecode path.
    let stats = engine.bytecode_compile_stats();
    assert_eq!(stats.total_formula_cells, 1);
    assert_eq!(stats.compiled, 1);
    assert_eq!(stats.fallback, 0);
    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();

    let system = engine.date_system();
    let settlement = ymd_to_serial(ExcelDate::new(1900, 2, 29), system).unwrap();
    let maturity = ymd_to_serial(ExcelDate::new(1900, 3, 1), system).unwrap();
    let expected =
        formula_engine::functions::financial::disc(settlement, maturity, 97.0, 100.0, 0, system)
            .unwrap();
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(expected));
    assert_engine_matches_ast(&engine, formula, "A1");

    // Flip date system after the formula has been compiled to ensure the runtime context is used.
    engine.set_date_system(ExcelDateSystem::Excel1904);
    engine.recalculate_single_threaded();
    assert_eq!(engine.bytecode_program_count(), 1);

    // 1900-02-29 is not valid under the 1904 date system.
    assert_eq!(
        engine.get_cell_value("Sheet1", "A1"),
        Value::Error(ErrorKind::Value)
    );
    assert_engine_matches_ast(&engine, formula, "A1");
}

#[test]
fn bytecode_backend_financial_numeric_dates_respect_engine_date_system() {
    let mut engine = Engine::new();

    // Numeric date serial arguments must be interpreted under the workbook's configured date
    // system. For serial 60, this means:
    // - Excel1900 (Lotus-compatible) => 1900-02-29
    // - Excel1904 => 1904-03-01
    //
    // Use a maturity date whose coupon schedule is easy to reason about (semiannual Jan/Jul).
    let settlement_serial = 60;
    let maturity = "1905-01-01";

    let pcd_formula = &format!(r#"=COUPPCD({settlement_serial},"{maturity}",2)"#);
    let ncd_formula = &format!(r#"=COUPNCD({settlement_serial},"{maturity}",2)"#);

    engine.set_cell_formula("Sheet1", "A1", pcd_formula).unwrap();
    engine.set_cell_formula("Sheet1", "A2", ncd_formula).unwrap();

    // Ensure we're exercising the bytecode path.
    let stats = engine.bytecode_compile_stats();
    assert_eq!(stats.total_formula_cells, 2);
    assert_eq!(stats.compiled, 2);
    assert_eq!(stats.fallback, 0);
    assert_eq!(engine.bytecode_program_count(), 2);

    // Excel1900: settlement serial 60 is the fictitious 1900-02-29, which falls in the Jan->Jul
    // coupon period for maturity 1905-01-01.
    engine.recalculate_single_threaded();

    let system = engine.date_system();
    let settlement_date = serial_to_ymd(settlement_serial, system).unwrap();
    let expected_pcd =
        ymd_to_serial(ExcelDate::new(settlement_date.year, 1, 1), system).unwrap() as f64;
    let expected_ncd =
        ymd_to_serial(ExcelDate::new(settlement_date.year, 7, 1), system).unwrap() as f64;
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(expected_pcd));
    assert_eq!(engine.get_cell_value("Sheet1", "A2"), Value::Number(expected_ncd));

    assert_engine_matches_ast(&engine, pcd_formula, "A1");
    assert_engine_matches_ast(&engine, ncd_formula, "A2");

    // Flip date system after the formulas have been compiled to ensure the runtime context is used.
    engine.set_date_system(ExcelDateSystem::Excel1904);
    engine.recalculate_single_threaded();
    assert_eq!(engine.bytecode_program_count(), 2);

    let system = engine.date_system();
    let settlement_date = serial_to_ymd(settlement_serial, system).unwrap();
    let expected_pcd =
        ymd_to_serial(ExcelDate::new(settlement_date.year, 1, 1), system).unwrap() as f64;
    let expected_ncd =
        ymd_to_serial(ExcelDate::new(settlement_date.year, 7, 1), system).unwrap() as f64;
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(expected_pcd));
    assert_eq!(engine.get_cell_value("Sheet1", "A2"), Value::Number(expected_ncd));

    assert_engine_matches_ast(&engine, pcd_formula, "A1");
    assert_engine_matches_ast(&engine, ncd_formula, "A2");
}

#[test]
fn bytecode_backend_matches_ast_for_odd_coupon_bond_functions() {
    let mut engine = Engine::new();

    engine
        .set_cell_formula(
            "Sheet1",
            "E1",
            r#"=ODDFPRICE("2020-01-01","2025-01-01","2020-01-01","2020-07-01",0.05,0.04,100,2,)"#,
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "E2",
            r#"=ODDFYIELD("2020-03-01","2020-07-01","2020-01-01","2020-07-01",0.05,ODDFPRICE("2020-03-01","2020-07-01","2020-01-01","2020-07-01",0.05,0.04,100,2,),100,2,)"#,
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "E3",
            r#"=ODDLPRICE("2020-10-15","2021-03-01","2020-10-15",0.05,0.06,100,2,)"#,
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "E4",
            r#"=ODDLYIELD("2020-10-15","2021-03-01","2020-10-15",0.05,ODDLPRICE("2020-10-15","2021-03-01","2020-10-15",0.05,0.06,100,2,),100,2,)"#,
        )
        .unwrap();

    let stats = engine.bytecode_compile_stats();
    assert_eq!(stats.total_formula_cells, 4);
    assert_eq!(stats.compiled, 4);
    assert_eq!(stats.fallback, 0);
    assert_eq!(engine.bytecode_program_count(), 4);

    engine.recalculate_single_threaded();

    for (formula, cell) in [
        (
            r#"=ODDFPRICE("2020-01-01","2025-01-01","2020-01-01","2020-07-01",0.05,0.04,100,2,)"#,
            "E1",
        ),
        (
            r#"=ODDFYIELD("2020-03-01","2020-07-01","2020-01-01","2020-07-01",0.05,ODDFPRICE("2020-03-01","2020-07-01","2020-01-01","2020-07-01",0.05,0.04,100,2,),100,2,)"#,
            "E2",
        ),
        (
            r#"=ODDLPRICE("2020-10-15","2021-03-01","2020-10-15",0.05,0.06,100,2,)"#,
            "E3",
        ),
        (
            r#"=ODDLYIELD("2020-10-15","2021-03-01","2020-10-15",0.05,ODDLPRICE("2020-10-15","2021-03-01","2020-10-15",0.05,0.06,100,2,),100,2,)"#,
            "E4",
        ),
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
    engine.set_cell_formula("Sheet1", "C2", "=A2&B2").unwrap();

    // Ensure we're exercising the bytecode path for both formulas.
    assert_eq!(engine.bytecode_program_count(), 2);

    engine.recalculate_single_threaded();
    assert_engine_matches_ast(&engine, "=\"a\"&\"b\"", "A1");
    assert_engine_matches_ast(&engine, "=A2&B2", "C2");
}

#[test]
fn bytecode_backend_compiles_rand_functions_and_matches_ast() {
    let mut engine = Engine::new();

    // Multiple draws within a single formula should be distinct (per-eval RNG counter).
    engine
        .set_cell_formula("Sheet1", "A1", "=RAND()+RAND()")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A2", "=RANDBETWEEN(1,100)+RANDBETWEEN(1,100)")
        .unwrap();

    // RAND/RANDBETWEEN should both compile to bytecode.
    assert_eq!(engine.bytecode_program_count(), 2);

    engine.recalculate_single_threaded();

    assert_engine_matches_ast(&engine, "=RAND()+RAND()", "A1");
    assert_engine_matches_ast(&engine, "=RANDBETWEEN(1,100)+RANDBETWEEN(1,100)", "A2");

    // Sanity check basic invariants.
    match engine.get_cell_value("Sheet1", "A1") {
        Value::Number(n) => assert!((0.0..2.0).contains(&n), "got {n}"),
        other => panic!("expected RAND()+RAND() to return a number, got {other:?}"),
    }
    match engine.get_cell_value("Sheet1", "A2") {
        Value::Number(n) => {
            assert!((2.0..=200.0).contains(&n), "got {n}");
            assert_eq!(n.fract(), 0.0, "expected integer result, got {n}");
        }
        other => panic!("expected RANDBETWEEN()+RANDBETWEEN() to return a number, got {other:?}"),
    }
}

#[test]
fn bytecode_backend_matches_ast_for_concatenate_function() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=CONCATENATE(\"a\", \"b\")")
        .unwrap();

    // Ensure we're exercising the bytecode path.
    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();
    assert_engine_matches_ast(&engine, "=CONCATENATE(\"a\", \"b\")", "A1");
}

#[test]
fn bytecode_backend_concat_function_uses_engine_value_locale_for_number_to_text() {
    let mut engine = Engine::new();
    engine.set_value_locale(ValueLocaleConfig::de_de());
    engine.set_cell_value("Sheet1", "A1", 1.5).unwrap();
    engine
        .set_cell_formula("Sheet1", "B1", r#"=CONCAT(A1,"x")"#)
        .unwrap();
    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Sheet1", "B1"),
        Value::Text("1,5x".to_string())
    );
    assert_engine_matches_ast(&engine, r#"=CONCAT(A1,"x")"#, "B1");
}

#[test]
fn bytecode_backend_matches_ast_for_concat_operator_with_numeric_literals() {
    let mut engine = Engine::new();
    engine.set_cell_formula("Sheet1", "A1", "=1&2").unwrap();

    // Ensure we're exercising the bytecode path.
    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();
    assert_engine_matches_ast(&engine, "=1&2", "A1");
}

#[test]
fn bytecode_backend_matches_ast_for_concat_operator_with_cell_and_string_literal() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 123.0).unwrap();
    engine
        .set_cell_formula("Sheet1", "B1", r#"=A1&"x""#)
        .unwrap();

    // Ensure we're exercising the bytecode path.
    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();
    assert_engine_matches_ast(&engine, r#"=A1&"x""#, "B1");
}

#[test]
fn bytecode_backend_concat_operator_uses_engine_value_locale_for_number_to_text() {
    let mut engine = Engine::new();
    engine.set_value_locale(ValueLocaleConfig::de_de());
    engine.set_cell_value("Sheet1", "A1", 1.5).unwrap();
    engine
        .set_cell_formula("Sheet1", "B1", r#"=A1&"x""#)
        .unwrap();
    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Sheet1", "B1"),
        Value::Text("1,5x".to_string())
    );
    assert_engine_matches_ast(&engine, r#"=A1&"x""#, "B1");
}

#[test]
fn bytecode_backend_matches_ast_for_concat_operator_blank_and_error_operands() {
    let mut engine = Engine::new();

    // A1 left blank: concat should treat blanks as empty text.
    engine
        .set_cell_formula("Sheet1", "B1", r#"=A1&"x""#)
        .unwrap();
    assert_eq!(engine.bytecode_program_count(), 1);

    // Errors should propagate through concatenation.
    engine
        .set_cell_formula("Sheet1", "B2", r#"=#DIV/0!&"x""#)
        .unwrap();
    assert_eq!(engine.bytecode_program_count(), 2);

    engine.recalculate_single_threaded();
    assert_engine_matches_ast(&engine, r#"=A1&"x""#, "B1");
    assert_engine_matches_ast(&engine, r#"=#DIV/0!&"x""#, "B2");
}

#[test]
fn bytecode_lower_flattens_concat_operator_chains() {
    use formula_engine::bytecode;
    use formula_engine::{LocaleConfig, ParseOptions, ReferenceStyle};

    let origin = parse_a1("A1").expect("origin");
    let origin = formula_engine::CellAddr::new(origin.row, origin.col);
    let ast = formula_engine::parse_formula(
        "=\"a\"&\"b\"&\"c\"",
        ParseOptions {
            locale: LocaleConfig::en_us(),
            reference_style: ReferenceStyle::A1,
            normalize_relative_to: Some(origin),
        },
    )
    .expect("parse canonical formula");

    let mut resolve_sheet = |_name: &str| Some(0usize);
    let mut sheet_dimensions =
        |_sheet_id: usize| Some((formula_model::EXCEL_MAX_ROWS, formula_model::EXCEL_MAX_COLS));
    let expr = bytecode::lower_canonical_expr(
        &ast.expr,
        origin,
        0,
        &mut resolve_sheet,
        &mut sheet_dimensions,
    )
        .expect("lower to bytecode expr");

    let bytecode::Expr::FuncCall { func, args } = expr else {
        panic!("expected FuncCall");
    };
    assert_eq!(func, bytecode::ast::Function::ConcatOp);
    assert_eq!(args.len(), 3, "expected concat chain to flatten");
}

#[test]
fn bytecode_backend_supports_concat_and_concatenate_with_ranges() {
    let mut engine = Engine::new();

    engine
        .set_cell_value("Sheet1", "A1", Value::Text("a".to_string()))
        .unwrap();
    engine
        .set_cell_value("Sheet1", "A2", Value::Text("b".to_string()))
        .unwrap();

    engine
        .set_cell_formula("Sheet1", "B1", r#"=CONCAT(A1:A2,"c")"#)
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "C1", r#"=CONCATENATE(A1:A2,"c")"#)
        .unwrap();

    // Ensure we're exercising the bytecode path.
    assert_eq!(engine.bytecode_program_count(), 2);

    engine.recalculate_single_threaded();

    assert_eq!(
        engine.get_cell_value("Sheet1", "B1"),
        Value::Text("abc".to_string())
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "C1"),
        Value::Text("ac".to_string())
    );

    assert_engine_matches_ast(&engine, r#"=CONCAT(A1:A2,"c")"#, "B1");
    assert_engine_matches_ast(&engine, r#"=CONCATENATE(A1:A2,"c")"#, "C1");
}

#[test]
fn bytecode_backend_supports_concat_operator_spilling_over_ranges() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", "a").unwrap();
    engine.set_cell_value("Sheet1", "A2", "b").unwrap();

    engine
        .set_cell_formula("Sheet1", "B1", r#"=A1:A2&"c""#)
        .unwrap();

    // Ensure we're exercising the bytecode path.
    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();
    assert_engine_spill_matches_ast(&engine, r#"=A1:A2&"c""#, "B1");
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
fn bytecode_backend_matches_ast_for_postfix_percent_blank_and_error_literals() {
    let mut engine = Engine::new();

    // A1 left blank.
    engine.set_cell_formula("Sheet1", "B1", "=A1%").unwrap();
    engine
        .set_cell_formula("Sheet1", "B2", "=#DIV/0!%")
        .unwrap();
    engine.set_cell_formula("Sheet1", "B3", "=200*50%").unwrap();

    // Ensure we're exercising the bytecode path.
    assert_eq!(engine.bytecode_program_count(), 3);

    engine.recalculate_single_threaded();

    assert_engine_matches_ast(&engine, "=A1%", "B1");
    assert_engine_matches_ast(&engine, "=#DIV/0!%", "B2");
    assert_engine_matches_ast(&engine, "=200*50%", "B3");
}

#[test]
fn bytecode_backend_matches_ast_for_double_percent_postfix() {
    let mut engine = Engine::new();
    engine.set_cell_formula("Sheet1", "A1", "=10%%").unwrap();

    // Ensure we're exercising the bytecode path.
    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();
    assert_engine_matches_ast(&engine, "=10%%", "A1");

    // Also check a locale-sensitive numeric text with a trailing percent.
    engine
        .set_cell_formula("Sheet1", "A2", r#"="10%"%"#)
        .unwrap();
    assert_eq!(engine.bytecode_program_count(), 2);
    engine.recalculate_single_threaded();
    assert_engine_matches_ast(&engine, r#"="10%"%"#, "A2");
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
fn bytecode_backend_matches_ast_for_implicit_intersection_over_sheet_span_when_only_one_area_hits()
{
    let mut engine = Engine::new();

    // Use a 3D span and custom sheet dimensions to create a case where only one sheet contributes
    // a non-#VALUE result under implicit intersection.
    engine.ensure_sheet("Sheet3");
    engine
        .set_sheet_dimensions("Sheet2", 1, formula_model::EXCEL_MAX_COLS)
        .unwrap();

    engine
        .set_cell_formula("Sheet1", "A5", "=@Sheet1:Sheet3!A1:A3")
        .unwrap();

    // Ensure this formula is compiled to bytecode (and isn't forcing AST fallback).
    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();
    let via_bytecode = engine.get_cell_value("Sheet1", "A5");
    assert_eq!(via_bytecode, Value::Error(ErrorKind::Ref));

    // Compare to AST-only evaluation.
    engine.set_bytecode_enabled(false);
    engine.recalculate_single_threaded();
    let via_ast = engine.get_cell_value("Sheet1", "A5");
    assert_eq!(via_bytecode, via_ast);
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
    grid.set_value(
        CellCoord::new(0, 0),
        formula_engine::bytecode::Value::Number(1.0),
    );
    grid.set_value(
        CellCoord::new(1, 0),
        formula_engine::bytecode::Value::Number(2.0),
    );
    grid.set_value(
        CellCoord::new(0, 1),
        formula_engine::bytecode::Value::Number(10.0),
    );
    grid.set_value(
        CellCoord::new(1, 1),
        formula_engine::bytecode::Value::Number(42.0),
    );

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
    let mut sheet_dimensions = |_sheet_id: usize| Some((10u32, 10u32));
    let bc_expr =
        formula_engine::bytecode::lower_canonical_expr(
            &ast.expr,
            origin,
            0,
            &mut resolve_sheet,
            &mut sheet_dimensions,
        )
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
    let bc_value = vm.eval(
        &program,
        &grid,
        0,
        base,
        &formula_engine::LocaleConfig::en_us(),
    );

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

#[test]
fn bytecode_backend_vlookup_hlookup_accept_array_literal_tables() {
    let mut engine = Engine::new();

    engine
        .set_cell_formula("Sheet1", "A1", "=VLOOKUP(3,{1,10;3,30;5,50},2,FALSE)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A2", "=VLOOKUP(4,{1,10;3,30;5,50},2)")
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "A3",
            "=LET(t,{1,10;3,30;5,50},VLOOKUP(5,t,2,FALSE))",
        )
        .unwrap();

    engine
        .set_cell_formula("Sheet1", "B1", "=HLOOKUP(3,{1,3,5;10,30,50},2,FALSE)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B2", "=HLOOKUP(4,{1,3,5;10,30,50},2)")
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "B3",
            "=LET(t,{1,3,5;10,30,50},HLOOKUP(5,t,2,FALSE))",
        )
        .unwrap();

    let stats = engine.bytecode_compile_stats();
    assert_eq!(stats.total_formula_cells, 6);
    assert_eq!(
        stats.compiled, 6,
        "expected VLOOKUP/HLOOKUP to compile for array table literals and LET array locals"
    );
    assert_eq!(stats.fallback, 0);

    engine.recalculate_single_threaded();

    for (formula, cell) in [
        ("=VLOOKUP(3,{1,10;3,30;5,50},2,FALSE)", "A1"),
        ("=VLOOKUP(4,{1,10;3,30;5,50},2)", "A2"),
        ("=LET(t,{1,10;3,30;5,50},VLOOKUP(5,t,2,FALSE))", "A3"),
        ("=HLOOKUP(3,{1,3,5;10,30,50},2,FALSE)", "B1"),
        ("=HLOOKUP(4,{1,3,5;10,30,50},2)", "B2"),
        ("=LET(t,{1,3,5;10,30,50},HLOOKUP(5,t,2,FALSE))", "B3"),
    ] {
        assert_engine_matches_ast(&engine, formula, cell);
    }
}

#[test]
fn bytecode_backend_applies_implicit_intersection_for_lookup_value_ranges() {
    let mut engine = Engine::new();

    // Lookup values in a vertical range (A1:A3) and a horizontal range (A20:C20).
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 3.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", 5.0).unwrap();

    engine.set_cell_value("Sheet1", "A20", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "B20", 3.0).unwrap();
    engine.set_cell_value("Sheet1", "C20", 5.0).unwrap();

    // VLOOKUP table in D:E: key/value pairs.
    engine.set_cell_value("Sheet1", "D1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "E1", 10.0).unwrap();
    engine.set_cell_value("Sheet1", "D2", 3.0).unwrap();
    engine.set_cell_value("Sheet1", "E2", 30.0).unwrap();
    engine.set_cell_value("Sheet1", "D3", 5.0).unwrap();
    engine.set_cell_value("Sheet1", "E3", 50.0).unwrap();

    // HLOOKUP table in A10:C11: keys on the first row, values on the second row.
    engine.set_cell_value("Sheet1", "A10", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "B10", 3.0).unwrap();
    engine.set_cell_value("Sheet1", "C10", 5.0).unwrap();
    engine.set_cell_value("Sheet1", "A11", 10.0).unwrap();
    engine.set_cell_value("Sheet1", "B11", 30.0).unwrap();
    engine.set_cell_value("Sheet1", "C11", 50.0).unwrap();

    // When lookup_value is passed as a range, Excel implicitly intersects it with the formula
    // cell's row/column. The AST evaluator implements this via `eval_scalar_arg`; the bytecode
    // backend should match.
    engine
        .set_cell_formula("Sheet1", "B2", "=VLOOKUP(A1:A3, D1:E3, 2, FALSE)")
        .unwrap();
    // Row 5 does not intersect A1:A3 -> #VALUE!.
    engine
        .set_cell_formula("Sheet1", "B5", "=VLOOKUP(A1:A3, D1:E3, 2, FALSE)")
        .unwrap();

    engine
        .set_cell_formula("Sheet1", "B21", "=HLOOKUP(A20:C20, A10:C11, 2, FALSE)")
        .unwrap();
    // Column D does not intersect A20:C20 -> #VALUE!.
    engine
        .set_cell_formula("Sheet1", "D21", "=HLOOKUP(A20:C20, A10:C11, 2, FALSE)")
        .unwrap();

    engine
        .set_cell_formula("Sheet1", "C2", "=MATCH(A1:A3, A1:A3, 0)")
        .unwrap();

    let stats = engine.bytecode_compile_stats();
    assert_eq!(stats.total_formula_cells, 5);
    assert_eq!(
        stats.compiled, 5,
        "expected VLOOKUP/HLOOKUP/MATCH to compile with range lookup_value args"
    );
    assert_eq!(stats.fallback, 0);

    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Sheet1", "B2"), Value::Number(30.0));
    assert_eq!(
        engine.get_cell_value("Sheet1", "B5"),
        Value::Error(ErrorKind::Value)
    );
    assert_eq!(engine.get_cell_value("Sheet1", "B21"), Value::Number(30.0));
    assert_eq!(
        engine.get_cell_value("Sheet1", "D21"),
        Value::Error(ErrorKind::Value)
    );
    assert_eq!(engine.get_cell_value("Sheet1", "C2"), Value::Number(2.0));

    for (formula, cell) in [
        ("=VLOOKUP(A1:A3, D1:E3, 2, FALSE)", "B2"),
        ("=VLOOKUP(A1:A3, D1:E3, 2, FALSE)", "B5"),
        ("=HLOOKUP(A20:C20, A10:C11, 2, FALSE)", "B21"),
        ("=HLOOKUP(A20:C20, A10:C11, 2, FALSE)", "D21"),
        ("=MATCH(A1:A3, A1:A3, 0)", "C2"),
    ] {
        assert_engine_matches_ast(&engine, formula, cell);
    }
}

#[test]
fn bytecode_backend_matches_ast_for_match() {
    let mut engine = Engine::new();

    // MATCH: ascending numeric values (match_type=1 default).
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 3.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", 5.0).unwrap();

    // MATCH: descending numeric values (match_type=-1).
    engine.set_cell_value("Sheet1", "B1", 5.0).unwrap();
    engine.set_cell_value("Sheet1", "B2", 3.0).unwrap();
    engine.set_cell_value("Sheet1", "B3", 1.0).unwrap();

    // MATCH: mixed numeric/text ascending order (numbers < text).
    engine.set_cell_value("Sheet1", "C1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "C2", 3.0).unwrap();
    engine.set_cell_value("Sheet1", "C3", "A").unwrap();

    // MATCH: wildcard text matching.
    engine.set_cell_value("Sheet1", "F1", "apple").unwrap();
    engine.set_cell_value("Sheet1", "F2", "apricot").unwrap();
    engine.set_cell_value("Sheet1", "F3", "banana").unwrap();

    // Error propagation from lookup_value.
    engine
        .set_cell_value("Sheet1", "D1", Value::Error(ErrorKind::Div0))
        .unwrap();

    // MATCH: exact match.
    engine
        .set_cell_formula("Sheet1", "E1", "=MATCH(3, A1:A3, 0)")
        .unwrap();
    // MATCH: approximate match (match_type omitted => 1).
    engine
        .set_cell_formula("Sheet1", "E2", "=MATCH(4, A1:A3)")
        .unwrap();
    // MATCH: missing => #N/A.
    engine
        .set_cell_formula("Sheet1", "E3", "=MATCH(0, A1:A3, 0)")
        .unwrap();
    // MATCH: descending approximate match.
    engine
        .set_cell_formula("Sheet1", "E4", "=MATCH(2, B1:B3, -1)")
        .unwrap();
    // MATCH: approximate match with mixed-type array.
    engine
        .set_cell_formula("Sheet1", "E5", "=MATCH(4, C1:C3)")
        .unwrap();
    // MATCH: propagate lookup_value error.
    engine
        .set_cell_formula("Sheet1", "E6", "=MATCH(D1, A1:A3, 0)")
        .unwrap();
    // MATCH: 2D lookup array => #N/A.
    engine
        .set_cell_formula("Sheet1", "E7", "=MATCH(3, A1:B2, 0)")
        .unwrap();
    // MATCH: invalid match_type => #N/A.
    engine
        .set_cell_formula("Sheet1", "E8", "=MATCH(3, A1:A3, 2)")
        .unwrap();
    // MATCH: wildcard matching in exact mode.
    engine
        .set_cell_formula("Sheet1", "E9", r#"=MATCH("ap*", F1:F3, 0)"#)
        .unwrap();

    // Horizontal match range.
    engine.set_cell_value("Sheet1", "A10", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "B10", 3.0).unwrap();
    engine.set_cell_value("Sheet1", "C10", 5.0).unwrap();
    engine
        .set_cell_formula("Sheet1", "E10", "=MATCH(3, A10:C10, 0)")
        .unwrap();

    // Single-cell lookup array.
    engine
        .set_cell_formula("Sheet1", "E11", "=MATCH(3, A2, 0)")
        .unwrap();

    let stats = engine.bytecode_compile_stats();
    assert_eq!(stats.total_formula_cells, 11);
    assert_eq!(
        stats.compiled, 11,
        "expected MATCH formulas to compile to bytecode"
    );
    assert_eq!(stats.fallback, 0);

    engine.recalculate_single_threaded();

    // Explicit expectations for key MATCH behaviors.
    assert_eq!(engine.get_cell_value("Sheet1", "E1"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "E2"), Value::Number(2.0));
    assert_eq!(
        engine.get_cell_value("Sheet1", "E3"),
        Value::Error(ErrorKind::NA)
    );
    assert_eq!(engine.get_cell_value("Sheet1", "E4"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "E5"), Value::Number(2.0));
    assert_eq!(
        engine.get_cell_value("Sheet1", "E6"),
        Value::Error(ErrorKind::Div0)
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "E7"),
        Value::Error(ErrorKind::NA)
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "E8"),
        Value::Error(ErrorKind::NA)
    );
    assert_eq!(engine.get_cell_value("Sheet1", "E9"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "E10"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "E11"), Value::Number(1.0));

    for (formula, cell) in [
        ("=MATCH(3, A1:A3, 0)", "E1"),
        ("=MATCH(4, A1:A3)", "E2"),
        ("=MATCH(0, A1:A3, 0)", "E3"),
        ("=MATCH(2, B1:B3, -1)", "E4"),
        ("=MATCH(4, C1:C3)", "E5"),
        ("=MATCH(D1, A1:A3, 0)", "E6"),
        ("=MATCH(3, A1:B2, 0)", "E7"),
        ("=MATCH(3, A1:A3, 2)", "E8"),
        (r#"=MATCH("ap*", F1:F3, 0)"#, "E9"),
        ("=MATCH(3, A10:C10, 0)", "E10"),
        ("=MATCH(3, A2, 0)", "E11"),
    ] {
        assert_engine_matches_ast(&engine, formula, cell);
    }
}

#[test]
fn bytecode_backend_matches_ast_for_xlookup_and_xmatch() {
    let mut engine = Engine::new();

    engine.set_cell_value("Sheet1", "B1", "a").unwrap();
    engine.set_cell_value("Sheet1", "B2", "b").unwrap();
    engine.set_cell_value("Sheet1", "B3", "c").unwrap();
    // Duplicate "b" to exercise last-to-first search mode.
    engine.set_cell_value("Sheet1", "B4", "b").unwrap();

    engine.set_cell_value("Sheet1", "C1", 10.0).unwrap();
    engine.set_cell_value("Sheet1", "C2", 20.0).unwrap();
    engine.set_cell_value("Sheet1", "C3", 30.0).unwrap();
    engine.set_cell_value("Sheet1", "C4", 40.0).unwrap();

    engine
        .set_cell_formula("Sheet1", "D1", r#"=XLOOKUP("b",B1:B3,C1:C3)"#)
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "D2", r#"=XLOOKUP("z",B1:B3,C1:C3)"#)
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "D3", r#"=XLOOKUP("z",B1:B3,C1:C3,"missing")"#)
        .unwrap();

    engine
        .set_cell_formula("Sheet1", "D4", r#"=XMATCH("b",B1:B3)"#)
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "D5", r#"=XMATCH("z",B1:B3)"#)
        .unwrap();

    // Omit if_not_found (empty arg) while specifying match/search modes.
    engine
        .set_cell_formula("Sheet1", "D6", r#"=XLOOKUP("b",B1:B4,C1:C4,,0,-1)"#)
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "D7", r#"=XMATCH("b",B1:B4,0,-1)"#)
        .unwrap();

    // Wildcard mode.
    engine
        .set_cell_formula("Sheet1", "D8", r#"=XMATCH("b*",B1:B3,2)"#)
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "D9", r#"=XLOOKUP("b*",B1:B3,C1:C3,"missing",2)"#)
        .unwrap();
    // Wildcard + last-to-first should find the last matching value.
    engine
        .set_cell_formula(
            "Sheet1",
            "D10",
            r#"=XLOOKUP("b*",B1:B4,C1:C4,"missing",2,-1)"#,
        )
        .unwrap();

    // 2D lookup arrays should error with #VALUE!
    engine
        .set_cell_formula("Sheet1", "D11", r#"=XMATCH("b",B1:C2)"#)
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "D12", r#"=XLOOKUP("b",B1:C2,C1:C2)"#)
        .unwrap();
    // Mismatched lookup/return lengths should also error with #VALUE!
    engine
        .set_cell_formula("Sheet1", "D13", r#"=XLOOKUP("b",B1:B3,C1:C2)"#)
        .unwrap();

    // Ensure we're exercising the bytecode path for all of the above formulas.
    assert_eq!(engine.bytecode_program_count(), 13);

    engine.recalculate_single_threaded();

    for (formula, cell) in [
        (r#"=XLOOKUP("b",B1:B3,C1:C3)"#, "D1"),
        (r#"=XLOOKUP("z",B1:B3,C1:C3)"#, "D2"),
        (r#"=XLOOKUP("z",B1:B3,C1:C3,"missing")"#, "D3"),
        (r#"=XMATCH("b",B1:B3)"#, "D4"),
        (r#"=XMATCH("z",B1:B3)"#, "D5"),
        (r#"=XLOOKUP("b",B1:B4,C1:C4,,0,-1)"#, "D6"),
        (r#"=XMATCH("b",B1:B4,0,-1)"#, "D7"),
        (r#"=XMATCH("b*",B1:B3,2)"#, "D8"),
        (r#"=XLOOKUP("b*",B1:B3,C1:C3,"missing",2)"#, "D9"),
        (r#"=XLOOKUP("b*",B1:B4,C1:C4,"missing",2,-1)"#, "D10"),
        (r#"=XMATCH("b",B1:C2)"#, "D11"),
        (r#"=XLOOKUP("b",B1:C2,C1:C2)"#, "D12"),
        (r#"=XLOOKUP("b",B1:B3,C1:C2)"#, "D13"),
    ] {
        assert_engine_matches_ast(&engine, formula, cell);
    }
}

#[test]
fn bytecode_backend_applies_implicit_intersection_for_xlookup_xmatch_lookup_value_ranges() {
    let mut engine = Engine::new();

    // Lookup values in a vertical range (A1:A3) and a horizontal range (A20:C20).
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 3.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", 5.0).unwrap();

    engine.set_cell_value("Sheet1", "A20", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "B20", 3.0).unwrap();
    engine.set_cell_value("Sheet1", "C20", 5.0).unwrap();

    // Vertical lookup vectors in D:E (lookup_array/return_array).
    engine.set_cell_value("Sheet1", "D1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "E1", 10.0).unwrap();
    engine.set_cell_value("Sheet1", "D2", 3.0).unwrap();
    engine.set_cell_value("Sheet1", "E2", 30.0).unwrap();
    engine.set_cell_value("Sheet1", "D3", 5.0).unwrap();
    engine.set_cell_value("Sheet1", "E3", 50.0).unwrap();

    // Horizontal lookup vectors in rows 10-11.
    engine.set_cell_value("Sheet1", "A10", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "B10", 3.0).unwrap();
    engine.set_cell_value("Sheet1", "C10", 5.0).unwrap();
    engine.set_cell_value("Sheet1", "A11", 10.0).unwrap();
    engine.set_cell_value("Sheet1", "B11", 30.0).unwrap();
    engine.set_cell_value("Sheet1", "C11", 50.0).unwrap();

    // When lookup_value is passed as a range, Excel implicitly intersects it with the formula
    // cell's row/column. The AST evaluator implements this via `eval_scalar_arg`; the bytecode
    // backend should match.
    engine
        .set_cell_formula("Sheet1", "B2", "=XMATCH(A1:A3, D1:D3)")
        .unwrap();
    // Row 5 does not intersect A1:A3 -> #VALUE!.
    engine
        .set_cell_formula("Sheet1", "B5", "=XMATCH(A1:A3, D1:D3)")
        .unwrap();

    engine
        .set_cell_formula("Sheet1", "C2", "=XLOOKUP(A1:A3, D1:D3, E1:E3)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "C5", "=XLOOKUP(A1:A3, D1:D3, E1:E3)")
        .unwrap();

    engine
        .set_cell_formula("Sheet1", "B21", "=XMATCH(A20:C20, A10:C10)")
        .unwrap();
    // Column D does not intersect A20:C20 -> #VALUE!.
    engine
        .set_cell_formula("Sheet1", "D21", "=XMATCH(A20:C20, A10:C10)")
        .unwrap();

    engine
        .set_cell_formula("Sheet1", "C21", "=XLOOKUP(A20:C20, A10:C10, A11:C11)")
        .unwrap();
    // Column E does not intersect A20:C20 -> #VALUE!.
    engine
        .set_cell_formula("Sheet1", "E21", "=XLOOKUP(A20:C20, A10:C10, A11:C11)")
        .unwrap();

    let stats = engine.bytecode_compile_stats();
    assert_eq!(stats.total_formula_cells, 8);
    assert_eq!(
        stats.compiled, 8,
        "expected XLOOKUP/XMATCH to compile with range lookup_value args"
    );
    assert_eq!(stats.fallback, 0);

    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Sheet1", "B2"), Value::Number(2.0));
    assert_eq!(
        engine.get_cell_value("Sheet1", "B5"),
        Value::Error(ErrorKind::Value)
    );
    assert_eq!(engine.get_cell_value("Sheet1", "C2"), Value::Number(30.0));
    assert_eq!(
        engine.get_cell_value("Sheet1", "C5"),
        Value::Error(ErrorKind::Value)
    );

    assert_eq!(engine.get_cell_value("Sheet1", "B21"), Value::Number(2.0));
    assert_eq!(
        engine.get_cell_value("Sheet1", "D21"),
        Value::Error(ErrorKind::Value)
    );
    // In C21, the horizontal lookup_value range intersects at C20 (=5).
    assert_eq!(engine.get_cell_value("Sheet1", "C21"), Value::Number(50.0));
    assert_eq!(
        engine.get_cell_value("Sheet1", "E21"),
        Value::Error(ErrorKind::Value)
    );

    for (formula, cell) in [
        ("=XMATCH(A1:A3, D1:D3)", "B2"),
        ("=XMATCH(A1:A3, D1:D3)", "B5"),
        ("=XLOOKUP(A1:A3, D1:D3, E1:E3)", "C2"),
        ("=XLOOKUP(A1:A3, D1:D3, E1:E3)", "C5"),
        ("=XMATCH(A20:C20, A10:C10)", "B21"),
        ("=XMATCH(A20:C20, A10:C10)", "D21"),
        ("=XLOOKUP(A20:C20, A10:C10, A11:C11)", "C21"),
        ("=XLOOKUP(A20:C20, A10:C10, A11:C11)", "E21"),
    ] {
        assert_engine_matches_ast(&engine, formula, cell);
    }
}

#[test]
fn bytecode_backend_vlookup_hlookup_match_reject_3d_sheet_span_table_ranges() {
    let mut engine = Engine::new();

    // Ensure the sheet span exists so the bytecode lowerer can resolve the sheet names.
    for sheet in ["Sheet1", "Sheet2", "Sheet3"] {
        engine.ensure_sheet(sheet);
    }

    engine
        .set_cell_formula(
            "Sheet1",
            "B1",
            "=VLOOKUP(1, Sheet1:Sheet3!A1:B1, 2, FALSE)",
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "B2",
            "=HLOOKUP(1, Sheet1:Sheet3!A1:B1, 2, FALSE)",
        )
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B3", "=MATCH(1, Sheet1:Sheet3!A1, 0)")
        .unwrap();

    // These formulas should still compile to bytecode even though 3D table ranges are not valid
    // lookup tables for VLOOKUP/HLOOKUP/MATCH.
    let stats = engine.bytecode_compile_stats();
    assert_eq!(stats.total_formula_cells, 3);
    assert_eq!(
        stats.compiled, 3,
        "expected VLOOKUP/HLOOKUP/MATCH to compile for 3D table ranges"
    );
    assert_eq!(stats.fallback, 0);

    engine.recalculate_single_threaded();

    let bc_values = [
        ("B1", engine.get_cell_value("Sheet1", "B1")),
        ("B2", engine.get_cell_value("Sheet1", "B2")),
        ("B3", engine.get_cell_value("Sheet1", "B3")),
    ];

    // Compare against the AST backend by disabling bytecode and re-evaluating.
    engine.set_bytecode_enabled(false);
    engine.recalculate_single_threaded();

    for (cell, expected) in bc_values {
        assert_eq!(engine.get_cell_value("Sheet1", cell), expected);
    }
}

#[test]
fn bytecode_backend_xlookup_xmatch_accept_spill_ranges_and_let_range_locals() {
    let mut engine = Engine::new();

    // Create spilled lookup/return arrays from literals so `A1#` / `B1#` are valid range arguments.
    engine
        .set_cell_formula("Sheet1", "A1", r#"={"a";"b";"c";"b"}"#)
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B1", "={10;20;30;40}")
        .unwrap();

    engine
        .set_cell_formula("Sheet1", "C1", r#"=XMATCH("b",A1#)"#)
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "C2", r#"=XMATCH("b",A1#,0,-1)"#)
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "D1", r#"=XLOOKUP("b",A1#,B1#)"#)
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "D2", r#"=XLOOKUP("b",A1#,B1#,,0,-1)"#)
        .unwrap();

    // LET-bound range locals should be accepted for lookup_array / return_array arguments.
    engine
        .set_cell_formula("Sheet1", "E1", r#"=LET(arr,A1:A4,XMATCH("b",arr,0,-1))"#)
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "E2",
            r#"=LET(ret,B1:B4,XLOOKUP("b",A1:A4,ret,,0,-1))"#,
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "E3",
            r#"=LET(arr,A1:A4,ret,B1:B4,XLOOKUP("b",arr,ret))"#,
        )
        .unwrap();

    // Ensure we're exercising the bytecode path for all of the above formulas.
    assert_eq!(engine.bytecode_program_count(), 9);

    engine.recalculate_single_threaded();

    for (formula, cell) in [
        (r#"=XMATCH("b",A1#)"#, "C1"),
        (r#"=XMATCH("b",A1#,0,-1)"#, "C2"),
        (r#"=XLOOKUP("b",A1#,B1#)"#, "D1"),
        (r#"=XLOOKUP("b",A1#,B1#,,0,-1)"#, "D2"),
        (r#"=LET(arr,A1:A4,XMATCH("b",arr,0,-1))"#, "E1"),
        (r#"=LET(ret,B1:B4,XLOOKUP("b",A1:A4,ret,,0,-1))"#, "E2"),
        (r#"=LET(arr,A1:A4,ret,B1:B4,XLOOKUP("b",arr,ret))"#, "E3"),
    ] {
        assert_engine_matches_ast(&engine, formula, cell);
    }
}

#[test]
fn bytecode_backend_xlookup_xmatch_accept_array_literals() {
    let mut engine = Engine::new();

    engine
        .set_cell_formula("Sheet1", "A1", "=XMATCH(2,{1;2;3})")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A2", "=XLOOKUP(2,{1;2;3},{10;20;30})")
        .unwrap();

    // 2D return array should spill horizontally (row slice).
    engine
        .set_cell_formula("Sheet1", "A3", "=XLOOKUP(2,{1;2;3},{10,11;20,21;30,31})")
        .unwrap();

    // Horizontal lookup vector + 2D return array should spill vertically (column slice).
    engine
        .set_cell_formula("Sheet1", "A5", "=XLOOKUP(2,{1,2,3},{10,20,30;40,50,60})")
        .unwrap();

    assert_eq!(
        engine.bytecode_program_count(),
        4,
        "expected all formulas to compile to bytecode"
    );

    engine.recalculate_single_threaded();

    assert_engine_matches_ast(&engine, "=XMATCH(2,{1;2;3})", "A1");
    assert_engine_matches_ast(&engine, "=XLOOKUP(2,{1;2;3},{10;20;30})", "A2");
    assert_engine_spill_matches_ast(&engine, "=XLOOKUP(2,{1;2;3},{10,11;20,21;30,31})", "A3");
    assert_engine_spill_matches_ast(&engine, "=XLOOKUP(2,{1,2,3},{10,20,30;40,50,60})", "A5");
}

#[test]
fn bytecode_backend_xlookup_accepts_array_if_not_found() {
    let mut engine = Engine::new();

    // When no match is found, XLOOKUP can return an array literal which should spill.
    engine
        .set_cell_formula("Sheet1", "A1", "=XLOOKUP(99,{1;2;3},{10;20;30},{100;200})")
        .unwrap();

    assert_eq!(
        engine.bytecode_program_count(),
        1,
        "expected formula to compile to bytecode"
    );

    engine.recalculate_single_threaded();
    assert_engine_spill_matches_ast(&engine, "=XLOOKUP(99,{1;2;3},{10;20;30},{100;200})", "A1");
}

#[test]
fn bytecode_backend_xlookup_applies_implicit_intersection_for_if_not_found_ranges() {
    let mut engine = Engine::new();

    engine.set_cell_value("Sheet1", "B1", 10.0).unwrap();
    engine.set_cell_value("Sheet1", "B2", 20.0).unwrap();
    engine.set_cell_value("Sheet1", "B3", 30.0).unwrap();

    engine
        .set_cell_formula("Sheet1", "C2", "=XLOOKUP(99,{1;2;3},{10;20;30},B1:B3)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "C5", "=XLOOKUP(99,{1;2;3},{10;20;30},B1:B3)")
        .unwrap();

    let stats = engine.bytecode_compile_stats();
    assert_eq!(stats.total_formula_cells, 2);
    assert_eq!(
        stats.compiled, 2,
        "expected XLOOKUP to compile with range if_not_found args"
    );
    assert_eq!(stats.fallback, 0);

    engine.recalculate_single_threaded();

    // In C2, the vertical range intersects on row 2 => B2.
    assert_eq!(engine.get_cell_value("Sheet1", "C2"), Value::Number(20.0));
    // Row 5 does not intersect B1:B3 => #VALUE!.
    assert_eq!(
        engine.get_cell_value("Sheet1", "C5"),
        Value::Error(ErrorKind::Value)
    );

    for cell in ["C2", "C5"] {
        assert_engine_matches_ast(&engine, "=XLOOKUP(99,{1;2;3},{10;20;30},B1:B3)", cell);
    }
}

#[test]
fn bytecode_backend_applies_implicit_intersection_for_xlookup_xmatch_mode_ranges() {
    let mut engine = Engine::new();

    // match_mode varies by row via implicit intersection.
    engine.set_cell_value("Sheet1", "D1", 0.0).unwrap(); // exact
    engine.set_cell_value("Sheet1", "D2", 1.0).unwrap(); // exact or next larger
    engine.set_cell_value("Sheet1", "D3", 0.0).unwrap(); // exact

    // search_mode varies by row via implicit intersection.
    engine.set_cell_value("Sheet1", "E1", 1.0).unwrap(); // first-to-last
    engine.set_cell_value("Sheet1", "E2", -1.0).unwrap(); // last-to-first
    engine.set_cell_value("Sheet1", "E3", 1.0).unwrap(); // first-to-last

    // match_mode implicit intersection.
    for cell in ["B1", "B2", "B5"] {
        engine
            .set_cell_formula("Sheet1", cell, "=XMATCH(2.5,{1;2;3},D1:D3)")
            .unwrap();
    }
    for cell in ["C1", "C2", "C5"] {
        engine
            .set_cell_formula(
                "Sheet1",
                cell,
                r#"=XLOOKUP(2.5,{1;2;3},{10;20;30},"missing",D1:D3)"#,
            )
            .unwrap();
    }

    // search_mode implicit intersection.
    for cell in ["F1", "F2", "F5"] {
        engine
            .set_cell_formula("Sheet1", cell, "=XMATCH(2,{1;2;2;3},,E1:E3)")
            .unwrap();
    }
    for cell in ["G1", "G2", "G5"] {
        engine
            .set_cell_formula(
                "Sheet1",
                cell,
                "=XLOOKUP(2,{1;2;2;3},{10;20;21;30},,0,E1:E3)",
            )
            .unwrap();
    }

    let stats = engine.bytecode_compile_stats();
    assert_eq!(stats.total_formula_cells, 12);
    assert_eq!(
        stats.compiled, 12,
        "expected XLOOKUP/XMATCH to compile with match/search mode range args"
    );
    assert_eq!(stats.fallback, 0);

    engine.recalculate_single_threaded();

    // XMATCH match_mode varies per row.
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Error(ErrorKind::NA));
    assert_eq!(engine.get_cell_value("Sheet1", "B2"), Value::Number(3.0));
    assert_eq!(
        engine.get_cell_value("Sheet1", "B5"),
        Value::Error(ErrorKind::Value)
    );

    // XLOOKUP match_mode varies per row.
    assert_eq!(
        engine.get_cell_value("Sheet1", "C1"),
        Value::Text("missing".into())
    );
    assert_eq!(engine.get_cell_value("Sheet1", "C2"), Value::Number(30.0));
    assert_eq!(
        engine.get_cell_value("Sheet1", "C5"),
        Value::Error(ErrorKind::Value)
    );

    // XMATCH search_mode varies per row (first vs last match).
    assert_eq!(engine.get_cell_value("Sheet1", "F1"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "F2"), Value::Number(3.0));
    assert_eq!(
        engine.get_cell_value("Sheet1", "F5"),
        Value::Error(ErrorKind::Value)
    );

    // XLOOKUP search_mode varies per row (first vs last return value).
    assert_eq!(engine.get_cell_value("Sheet1", "G1"), Value::Number(20.0));
    assert_eq!(engine.get_cell_value("Sheet1", "G2"), Value::Number(21.0));
    assert_eq!(
        engine.get_cell_value("Sheet1", "G5"),
        Value::Error(ErrorKind::Value)
    );

    for (formula, cell) in [
        ("=XMATCH(2.5,{1;2;3},D1:D3)", "B1"),
        ("=XMATCH(2.5,{1;2;3},D1:D3)", "B2"),
        ("=XMATCH(2.5,{1;2;3},D1:D3)", "B5"),
        (r#"=XLOOKUP(2.5,{1;2;3},{10;20;30},"missing",D1:D3)"#, "C1"),
        (r#"=XLOOKUP(2.5,{1;2;3},{10;20;30},"missing",D1:D3)"#, "C2"),
        (r#"=XLOOKUP(2.5,{1;2;3},{10;20;30},"missing",D1:D3)"#, "C5"),
        ("=XMATCH(2,{1;2;2;3},,E1:E3)", "F1"),
        ("=XMATCH(2,{1;2;2;3},,E1:E3)", "F2"),
        ("=XMATCH(2,{1;2;2;3},,E1:E3)", "F5"),
        ("=XLOOKUP(2,{1;2;2;3},{10;20;21;30},,0,E1:E3)", "G1"),
        ("=XLOOKUP(2,{1;2;2;3},{10;20;21;30},,0,E1:E3)", "G2"),
        ("=XLOOKUP(2,{1;2;2;3},{10;20;21;30},,0,E1:E3)", "G5"),
    ] {
        assert_engine_matches_ast(&engine, formula, cell);
    }
}

#[test]
fn bytecode_backend_xlookup_xmatch_mode_args_error_on_array_expressions_like_ast() {
    let mut engine = Engine::new();

    // Produce array results via range arithmetic.
    engine.set_cell_value("Sheet1", "D1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "D2", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "D3", 3.0).unwrap();

    // match_mode as array expression.
    engine
        .set_cell_formula("Sheet1", "A1", "=XMATCH(2,{1;2;3},D1:D3*0)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A2", "=XLOOKUP(2,{1;2;3},{10;20;30},,D1:D3*0)")
        .unwrap();

    // search_mode as array expression.
    engine
        .set_cell_formula("Sheet1", "A3", "=XMATCH(2,{1;2;3},0,D1:D3*0)")
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "A4",
            "=XLOOKUP(2,{1;2;3},{10;20;30},,0,D1:D3*0)",
        )
        .unwrap();

    let stats = engine.bytecode_compile_stats();
    assert_eq!(stats.total_formula_cells, 4);
    assert_eq!(
        stats.compiled, 4,
        "expected XLOOKUP/XMATCH to compile with array-expr mode args"
    );
    assert_eq!(stats.fallback, 0);

    engine.recalculate_single_threaded();

    for cell in ["A1", "A2", "A3", "A4"] {
        assert_eq!(
            engine.get_cell_value("Sheet1", cell),
            Value::Error(ErrorKind::Value)
        );
    }

    for (formula, cell) in [
        ("=XMATCH(2,{1;2;3},D1:D3*0)", "A1"),
        ("=XLOOKUP(2,{1;2;3},{10;20;30},,D1:D3*0)", "A2"),
        ("=XMATCH(2,{1;2;3},0,D1:D3*0)", "A3"),
        ("=XLOOKUP(2,{1;2;3},{10;20;30},,0,D1:D3*0)", "A4"),
    ] {
        assert_engine_matches_ast(&engine, formula, cell);
    }
}

#[test]
fn bytecode_backend_xlookup_xmatch_accept_let_single_cell_reference_locals() {
    let mut engine = Engine::new();

    engine.set_cell_value("Sheet1", "A1", "a").unwrap();
    engine.set_cell_value("Sheet1", "A2", "b").unwrap();
    engine.set_cell_value("Sheet1", "B1", 10.0).unwrap();
    engine.set_cell_value("Sheet1", "B2", 20.0).unwrap();

    // LET-bound single-cell references should be accepted for lookup_array / return_array.
    engine
        .set_cell_formula("Sheet1", "C1", r#"=LET(arr,A2,XMATCH("b",arr))"#)
        .unwrap();

    // Ensure CHOOSE can still produce reference locals that are consumed by XMATCH/XLOOKUP, while
    // preserving lazy evaluation and error propagation semantics.
    engine
        .set_cell_formula(
            "Sheet1",
            "C2",
            r#"=LET(arr,CHOOSE(1,A2,1/0),XMATCH("b",arr))"#,
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "C3",
            r#"=LET(arr,CHOOSE(2,A2,1/0),XMATCH("b",arr))"#,
        )
        .unwrap();

    engine
        .set_cell_formula(
            "Sheet1",
            "D1",
            r#"=LET(arr,A2,ret,B2,XLOOKUP("b",arr,ret))"#,
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "D2",
            r#"=LET(arr,CHOOSE(1,A2,1/0),ret,CHOOSE(1,B2,1/0),XLOOKUP("b",arr,ret))"#,
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "D3",
            r#"=LET(arr,CHOOSE(2,A2,1/0),ret,B2,XLOOKUP("b",arr,ret))"#,
        )
        .unwrap();

    // Ensure we're exercising the bytecode path for all of the above formulas.
    assert_eq!(engine.bytecode_program_count(), 6);

    engine.recalculate_single_threaded();

    for (formula, cell) in [
        (r#"=LET(arr,A2,XMATCH("b",arr))"#, "C1"),
        (r#"=LET(arr,CHOOSE(1,A2,1/0),XMATCH("b",arr))"#, "C2"),
        (r#"=LET(arr,CHOOSE(2,A2,1/0),XMATCH("b",arr))"#, "C3"),
        (r#"=LET(arr,A2,ret,B2,XLOOKUP("b",arr,ret))"#, "D1"),
        (
            r#"=LET(arr,CHOOSE(1,A2,1/0),ret,CHOOSE(1,B2,1/0),XLOOKUP("b",arr,ret))"#,
            "D2",
        ),
        (
            r#"=LET(arr,CHOOSE(2,A2,1/0),ret,B2,XLOOKUP("b",arr,ret))"#,
            "D3",
        ),
    ] {
        assert_engine_matches_ast(&engine, formula, cell);
    }
}

#[test]
fn bytecode_backend_row_column_accept_let_reference_locals() {
    let mut engine = Engine::new();

    // Ensure the referenced grid locations exist.
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "B2", 2.0).unwrap();

    // LET-bound range locals should be accepted by ROW/COLUMN.
    engine
        .set_cell_formula("Sheet1", "C1", "=LET(r, A1:B2, ROW(r))")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "E1", "=LET(r, A1:B2, COLUMN(r))")
        .unwrap();

    // LET-bound single-cell reference locals should also be accepted.
    engine
        .set_cell_formula("Sheet1", "G1", "=LET(r, A1, ROW(r))")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "G2", "=LET(r, B2, COLUMN(r))")
        .unwrap();

    let stats = engine.bytecode_compile_stats();
    assert_eq!(stats.total_formula_cells, 4);
    assert_eq!(
        stats.compiled,
        4,
        "expected all ROW/COLUMN LET formulas to compile to bytecode (report={:?})",
        engine.bytecode_compile_report(32)
    );
    assert_eq!(stats.fallback, 0);

    engine.recalculate_single_threaded();

    assert_engine_spill_matches_ast(&engine, "=LET(r, A1:B2, ROW(r))", "C1");
    assert_engine_spill_matches_ast(&engine, "=LET(r, A1:B2, COLUMN(r))", "E1");
    assert_engine_matches_ast(&engine, "=LET(r, A1, ROW(r))", "G1");
    assert_engine_matches_ast(&engine, "=LET(r, B2, COLUMN(r))", "G2");
}

#[test]
fn bytecode_backend_matches_ast_for_common_logical_error_functions() {
    let mut engine = Engine::new();

    // IF lazy error propagation.
    engine
        .set_cell_formula("Sheet1", "B1", "=IF(FALSE, 1/0, 7)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B2", "=IF(TRUE, \"x\", 1/0)")
        .unwrap();

    // IFERROR / IFNA.
    engine
        .set_cell_formula("Sheet1", "B3", "=IFERROR(1, 1/0)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B4", "=IFERROR(1/0, 7)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B5", "=IFNA(NA(), 7)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B9", "=IFNA(1/0, 7)")
        .unwrap();

    // Error helpers.
    engine.set_cell_formula("Sheet1", "B6", "=NA()").unwrap();
    engine
        .set_cell_formula("Sheet1", "B7", "=ISERROR(1/0)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B8", "=ISNA(NA())")
        .unwrap();
    engine
        // ISNA must not propagate non-#N/A errors.
        .set_cell_formula("Sheet1", "B10", "=ISNA(1/0)")
        .unwrap();

    assert_eq!(
        engine.bytecode_program_count(),
        10,
        "expected all formulas to compile to bytecode"
    );

    engine.recalculate_single_threaded();

    for (formula, cell) in [
        ("=IF(FALSE, 1/0, 7)", "B1"),
        ("=IF(TRUE, \"x\", 1/0)", "B2"),
        ("=IFERROR(1, 1/0)", "B3"),
        ("=IFERROR(1/0, 7)", "B4"),
        ("=IFNA(NA(), 7)", "B5"),
        ("=IFNA(1/0, 7)", "B9"),
        ("=NA()", "B6"),
        ("=ISERROR(1/0)", "B7"),
        ("=ISNA(NA())", "B8"),
        ("=ISNA(1/0)", "B10"),
    ] {
        assert_engine_matches_ast(&engine, formula, cell);
    }
}

#[test]
fn bytecode_backend_compiles_choose_and_is_lazy() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=CHOOSE(2, 1/0, 5)")
        .unwrap();

    assert_eq!(
        engine.bytecode_program_count(),
        1,
        "expected CHOOSE formula to compile to bytecode"
    );

    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(5.0));
    assert_engine_matches_ast(&engine, "=CHOOSE(2, 1/0, 5)", "A1");
}

#[test]
fn bytecode_backend_choose_out_of_range_returns_value_error() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=CHOOSE(3, 1, 2)")
        .unwrap();

    assert_eq!(
        engine.bytecode_program_count(),
        1,
        "expected CHOOSE formula to compile to bytecode"
    );

    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Sheet1", "A1"),
        Value::Error(ErrorKind::Value)
    );
    assert_engine_matches_ast(&engine, "=CHOOSE(3, 1, 2)", "A1");
}

#[test]
fn bytecode_backend_matches_ast_for_choose_scalar_index_matrix() {
    let mut engine = Engine::new();

    engine
        .set_cell_formula("Sheet1", "A1", "=CHOOSE(1, 10, 20)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A2", "=CHOOSE(2, 10, 20)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A3", "=CHOOSE(TRUE, 10, 20)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A4", r#"=CHOOSE("2", 10, 20)"#)
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A5", "=CHOOSE(2.9, 10, 20)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A6", "=CHOOSE(0, 10, 20)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A7", r#"=CHOOSE("x", 10, 20)"#)
        .unwrap();

    let stats = engine.bytecode_compile_stats();
    assert_eq!(
        stats.fallback,
        0,
        "expected all CHOOSE formulas to compile to bytecode (report={:?})",
        engine.bytecode_compile_report(100)
    );
    assert_eq!(stats.total_formula_cells, 7);
    assert_eq!(stats.compiled, 7);

    engine.recalculate_single_threaded();

    for (formula, cell) in [
        ("=CHOOSE(1, 10, 20)", "A1"),
        ("=CHOOSE(2, 10, 20)", "A2"),
        ("=CHOOSE(TRUE, 10, 20)", "A3"),
        (r#"=CHOOSE("2", 10, 20)"#, "A4"),
        ("=CHOOSE(2.9, 10, 20)", "A5"),
        ("=CHOOSE(0, 10, 20)", "A6"),
        (r#"=CHOOSE("x", 10, 20)"#, "A7"),
    ] {
        assert_engine_matches_ast(&engine, formula, cell);
    }
}

#[test]
fn bytecode_backend_if_two_arg_default_false_branch_matches_ast() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=IF(FALSE,1/0)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A2", "=IF(TRUE,7)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A3", "=IF(TRUE,1/0)")
        .unwrap();

    let stats = engine.bytecode_compile_stats();
    assert_eq!(
        stats.fallback,
        0,
        "expected all formulas to compile to bytecode (report={:?})",
        engine.bytecode_compile_report(100)
    );
    assert_eq!(stats.total_formula_cells, 3);
    assert_eq!(stats.compiled, 3);

    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Bool(false));
    assert_eq!(engine.get_cell_value("Sheet1", "A2"), Value::Number(7.0));
    assert_eq!(
        engine.get_cell_value("Sheet1", "A3"),
        Value::Error(ErrorKind::Div0)
    );

    for (formula, cell) in [
        ("=IF(FALSE,1/0)", "A1"),
        ("=IF(TRUE,7)", "A2"),
        ("=IF(TRUE,1/0)", "A3"),
    ] {
        assert_engine_matches_ast(&engine, formula, cell);
    }
}

#[test]
fn bytecode_backend_if_family_trailing_blank_args_match_ast() {
    let mut engine = Engine::new();
    engine
        // Trailing blank else-branch should behave like an explicit blank, not the 2-arg default.
        .set_cell_formula("Sheet1", "A1", "=IF(FALSE,1,)")
        .unwrap();
    engine
        // IF should still short-circuit the unused branch.
        .set_cell_formula("Sheet1", "A2", "=IF(FALSE,1/0,)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A3", "=IFERROR(1/0,)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A4", "=IFERROR(1,)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A5", "=IFNA(NA(),)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A6", "=IFNA(1/0,)")
        .unwrap();

    let stats = engine.bytecode_compile_stats();
    assert_eq!(
        stats.fallback,
        0,
        "expected all formulas to compile to bytecode (report={:?})",
        engine.bytecode_compile_report(100)
    );
    assert_eq!(stats.total_formula_cells, 6);
    assert_eq!(stats.compiled, 6);

    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Blank);
    assert_eq!(engine.get_cell_value("Sheet1", "A2"), Value::Blank);
    assert_eq!(engine.get_cell_value("Sheet1", "A3"), Value::Blank);
    assert_eq!(engine.get_cell_value("Sheet1", "A4"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "A5"), Value::Blank);
    assert_eq!(
        engine.get_cell_value("Sheet1", "A6"),
        Value::Error(ErrorKind::Div0)
    );

    for (formula, cell) in [
        ("=IF(FALSE,1,)", "A1"),
        ("=IF(FALSE,1/0,)", "A2"),
        ("=IFERROR(1/0,)", "A3"),
        ("=IFERROR(1,)", "A4"),
        ("=IFNA(NA(),)", "A5"),
        ("=IFNA(1/0,)", "A6"),
    ] {
        assert_engine_matches_ast(&engine, formula, cell);
    }
}

#[test]
fn bytecode_backend_if_family_missing_args_match_ast() {
    let mut engine = Engine::new();
    engine
        // Missing value_if_true returns blank when the condition is TRUE.
        .set_cell_formula("Sheet1", "A1", "=IF(TRUE,)")
        .unwrap();
    engine
        // Missing value_if_true does not affect the default FALSE value_if_false behavior.
        .set_cell_formula("Sheet1", "A2", "=IF(FALSE,)")
        .unwrap();
    engine
        // Missing value_if_true with an explicit else branch.
        .set_cell_formula("Sheet1", "A3", "=IF(FALSE,,7)")
        .unwrap();
    engine
        // IF should remain lazy when the selected branch is missing/blank.
        .set_cell_formula("Sheet1", "A4", "=IF(TRUE,,1/0)")
        .unwrap();
    engine
        // IFERROR should treat missing arg0 as blank (not error) and still short-circuit.
        .set_cell_formula("Sheet1", "A5", "=IFERROR(,1/0)")
        .unwrap();
    engine
        // IFNA should treat missing arg0 as blank (not #N/A) and still short-circuit.
        .set_cell_formula("Sheet1", "A6", "=IFNA(,1/0)")
        .unwrap();

    let stats = engine.bytecode_compile_stats();
    assert_eq!(
        stats.fallback,
        0,
        "expected all formulas to compile to bytecode (report={:?})",
        engine.bytecode_compile_report(100)
    );
    assert_eq!(stats.total_formula_cells, 6);
    assert_eq!(stats.compiled, 6);

    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Blank);
    assert_eq!(engine.get_cell_value("Sheet1", "A2"), Value::Bool(false));
    assert_eq!(engine.get_cell_value("Sheet1", "A3"), Value::Number(7.0));
    assert_eq!(engine.get_cell_value("Sheet1", "A4"), Value::Blank);
    assert_eq!(engine.get_cell_value("Sheet1", "A5"), Value::Blank);
    assert_eq!(engine.get_cell_value("Sheet1", "A6"), Value::Blank);

    for (formula, cell) in [
        ("=IF(TRUE,)", "A1"),
        ("=IF(FALSE,)", "A2"),
        ("=IF(FALSE,,7)", "A3"),
        ("=IF(TRUE,,1/0)", "A4"),
        ("=IFERROR(,1/0)", "A5"),
        ("=IFNA(,1/0)", "A6"),
    ] {
        assert_engine_matches_ast(&engine, formula, cell);
    }
}

#[test]
fn bytecode_backend_matches_ast_for_logical_error_functions_with_error_literals() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=IF(FALSE, #N/A, 1)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A2", "=IF(TRUE, 1, #DIV/0!)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A3", "=IFERROR(#DIV/0!, 7)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A4", "=IFERROR(#N/A, 7)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A5", "=IFNA(#N/A, 7)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A6", "=IFNA(#DIV/0!, 7)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A7", "=IF(#N/A, 1, 2)")
        .unwrap();
    engine
        // ISNA must not propagate non-#N/A error literals.
        .set_cell_formula("Sheet1", "A8", "=ISNA(#DIV/0!)")
        .unwrap();

    assert_eq!(
        engine.bytecode_program_count(),
        8,
        "expected all formulas to compile to bytecode"
    );

    engine.recalculate_single_threaded();

    for (formula, cell) in [
        ("=IF(FALSE, #N/A, 1)", "A1"),
        ("=IF(TRUE, 1, #DIV/0!)", "A2"),
        ("=IFERROR(#DIV/0!, 7)", "A3"),
        ("=IFERROR(#N/A, 7)", "A4"),
        ("=IFNA(#N/A, 7)", "A5"),
        ("=IFNA(#DIV/0!, 7)", "A6"),
        ("=IF(#N/A, 1, 2)", "A7"),
        ("=ISNA(#DIV/0!)", "A8"),
    ] {
        assert_engine_matches_ast(&engine, formula, cell);
    }
}

#[test]
fn bytecode_backend_matches_ast_for_information_functions_with_error_literals() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=ISERROR(#DIV/0!)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A2", "=ISERROR(#N/A)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A3", "=ISNA(#DIV/0!)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A4", "=ISNA(#N/A)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A5", "=ISERR(#DIV/0!)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A6", "=ISERR(#N/A)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A7", "=ERROR.TYPE(#DIV/0!)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A8", "=ERROR.TYPE(#N/A)")
        .unwrap();

    let stats = engine.bytecode_compile_stats();
    assert_eq!(
        stats.fallback,
        0,
        "expected all formulas to compile to bytecode (report={:?})",
        engine.bytecode_compile_report(100)
    );
    assert_eq!(stats.total_formula_cells, 8);
    assert_eq!(stats.compiled, 8);

    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Bool(true));
    assert_eq!(engine.get_cell_value("Sheet1", "A2"), Value::Bool(true));
    assert_eq!(engine.get_cell_value("Sheet1", "A3"), Value::Bool(false));
    assert_eq!(engine.get_cell_value("Sheet1", "A4"), Value::Bool(true));
    assert_eq!(engine.get_cell_value("Sheet1", "A5"), Value::Bool(true));
    assert_eq!(engine.get_cell_value("Sheet1", "A6"), Value::Bool(false));
    assert_eq!(engine.get_cell_value("Sheet1", "A7"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "A8"), Value::Number(7.0));

    for (formula, cell) in [
        ("=ISERROR(#DIV/0!)", "A1"),
        ("=ISERROR(#N/A)", "A2"),
        ("=ISNA(#DIV/0!)", "A3"),
        ("=ISNA(#N/A)", "A4"),
        ("=ISERR(#DIV/0!)", "A5"),
        ("=ISERR(#N/A)", "A6"),
        ("=ERROR.TYPE(#DIV/0!)", "A7"),
        ("=ERROR.TYPE(#N/A)", "A8"),
    ] {
        assert_engine_matches_ast(&engine, formula, cell);
    }
}

#[test]
fn bytecode_backend_matches_ast_for_choose_ifs_and_switch() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "B1", "=IFS(TRUE, 1, 1/0, 2)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B2", "=IFS(FALSE, 1/0, TRUE, 2)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B3", "=IFS(FALSE, 1, FALSE, 2)")
        .unwrap();

    engine
        .set_cell_formula("Sheet1", "B4", "=SWITCH(1, 1, 10, 1/0, 20)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B5", "=SWITCH(2, 1, 1/0, 2, 20)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B6", "=SWITCH(3, 1, 10, 2, 20, 99)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B7", "=SWITCH(3, 1, 10, 2, 20)")
        .unwrap();

    engine
        .set_cell_formula("Sheet1", "B8", "=CHOOSE(2, 1/0, 7)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B9", "=CHOOSE(1, 1/0, 7)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B10", "=CHOOSE(3, 1, 2)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B11", "=CHOOSE(1/0, 1, 2)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B12", "=CHOOSE(\"2\", 1/0, 7)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B13", "=CHOOSE(TRUE, 10, 20)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B14", "=CHOOSE(FALSE, 10, 20)")
        .unwrap();

    let stats = engine.bytecode_compile_stats();
    assert_eq!(
        stats.fallback,
        0,
        "expected all formulas to compile to bytecode (report={:?})",
        engine.bytecode_compile_report(100)
    );
    assert_eq!(stats.total_formula_cells, 14);
    assert_eq!(stats.compiled, 14);

    engine.recalculate_single_threaded();

    for (formula, cell) in [
        ("=IFS(TRUE, 1, 1/0, 2)", "B1"),
        ("=IFS(FALSE, 1/0, TRUE, 2)", "B2"),
        ("=IFS(FALSE, 1, FALSE, 2)", "B3"),
        ("=SWITCH(1, 1, 10, 1/0, 20)", "B4"),
        ("=SWITCH(2, 1, 1/0, 2, 20)", "B5"),
        ("=SWITCH(3, 1, 10, 2, 20, 99)", "B6"),
        ("=SWITCH(3, 1, 10, 2, 20)", "B7"),
        ("=CHOOSE(2, 1/0, 7)", "B8"),
        ("=CHOOSE(1, 1/0, 7)", "B9"),
        ("=CHOOSE(3, 1, 2)", "B10"),
        ("=CHOOSE(1/0, 1, 2)", "B11"),
        ("=CHOOSE(\"2\", 1/0, 7)", "B12"),
        ("=CHOOSE(TRUE, 10, 20)", "B13"),
        ("=CHOOSE(FALSE, 10, 20)", "B14"),
    ] {
        assert_engine_matches_ast(&engine, formula, cell);
    }
}

#[test]
fn bytecode_backend_choose_can_return_ranges() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", 3.0).unwrap();
    engine.set_cell_value("Sheet1", "B1", 10.0).unwrap();
    engine.set_cell_value("Sheet1", "B2", 20.0).unwrap();
    engine.set_cell_value("Sheet1", "B3", 30.0).unwrap();

    engine
        .set_cell_formula("Sheet1", "C1", "=SUM(CHOOSE(1, A1:A3, B1:B3))")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "C2", "=SUM(CHOOSE(2, A1:A3, B1:B3))")
        .unwrap();

    // Ensure we're exercising the bytecode compiler path for CHOOSE returning a reference.
    assert_eq!(engine.bytecode_program_count(), 2);

    engine.recalculate_single_threaded();
    assert_engine_matches_ast(&engine, "=SUM(CHOOSE(1, A1:A3, B1:B3))", "C1");
    assert_engine_matches_ast(&engine, "=SUM(CHOOSE(2, A1:A3, B1:B3))", "C2");
}

#[test]
fn bytecode_backend_and_or_reference_semantics_match_ast() {
    let mut engine = Engine::new();
    engine.set_cell_formula("Sheet1", "B1", "=AND(A1)").unwrap();
    engine.set_cell_formula("Sheet1", "B2", "=OR(A1)").unwrap();

    // Ensure we're exercising the bytecode path.
    assert_eq!(engine.bytecode_program_count(), 2);

    let cases: &[(Option<Value>, Value, Value)] = &[
        // Blank cell refs are ignored.
        (None, Value::Bool(true), Value::Bool(false)),
        // Text cell refs behave like scalar text arguments (#VALUE!).
        (
            Some(Value::Text("hello".to_string())),
            Value::Error(ErrorKind::Value),
            Value::Error(ErrorKind::Value),
        ),
        // Numbers/bools are included.
        (
            Some(Value::Number(0.0)),
            Value::Bool(false),
            Value::Bool(false),
        ),
        (
            Some(Value::Number(2.0)),
            Value::Bool(true),
            Value::Bool(true),
        ),
        (
            Some(Value::Bool(false)),
            Value::Bool(false),
            Value::Bool(false),
        ),
        (
            Some(Value::Bool(true)),
            Value::Bool(true),
            Value::Bool(true),
        ),
    ];

    for (a1, expected_and, expected_or) in cases {
        match a1 {
            None => engine.clear_cell("Sheet1", "A1").unwrap(),
            Some(v) => engine.set_cell_value("Sheet1", "A1", v.clone()).unwrap(),
        };

        engine.recalculate_single_threaded();

        assert_eq!(engine.get_cell_value("Sheet1", "B1"), *expected_and);
        assert_eq!(engine.get_cell_value("Sheet1", "B2"), *expected_or);
        assert_engine_matches_ast(&engine, "=AND(A1)", "B1");
        assert_engine_matches_ast(&engine, "=OR(A1)", "B2");
    }
}

#[test]
fn bytecode_backend_and_or_single_cell_range_semantics_match_ast() {
    let mut engine = Engine::new();
    engine.set_cell_formula("Sheet1", "B1", "=AND(A1)").unwrap();
    engine.set_cell_formula("Sheet1", "B2", "=OR(A1)").unwrap();
    engine
        .set_cell_formula("Sheet1", "C1", "=AND(A1:A1)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "C2", "=OR(A1:A1)")
        .unwrap();

    let stats = engine.bytecode_compile_stats();
    assert_eq!(
        stats.fallback,
        0,
        "expected all AND/OR formulas to compile to bytecode (report={:?})",
        engine.bytecode_compile_report(100)
    );
    assert_eq!(stats.total_formula_cells, 4);
    assert_eq!(stats.compiled, 4);

    let cases: &[(Option<Value>, Value, Value, Value, Value)] = &[
        // Blank cell refs are ignored for both scalar and range semantics.
        (
            None,
            Value::Bool(true),
            Value::Bool(false),
            Value::Bool(true),
            Value::Bool(false),
        ),
        // Text cell refs behave like scalar text arguments (#VALUE!) but are ignored in ranges.
        (
            Some(Value::Text("hello".to_string())),
            Value::Error(ErrorKind::Value),
            Value::Error(ErrorKind::Value),
            Value::Bool(true),
            Value::Bool(false),
        ),
        // Entity/record values behave like text: scalar refs error, but ranges ignore them.
        (
            Some(Value::Entity(EntityValue::new("Entity"))),
            Value::Error(ErrorKind::Value),
            Value::Error(ErrorKind::Value),
            Value::Bool(true),
            Value::Bool(false),
        ),
        (
            Some(Value::Record(RecordValue::new("Record"))),
            Value::Error(ErrorKind::Value),
            Value::Error(ErrorKind::Value),
            Value::Bool(true),
            Value::Bool(false),
        ),
        // Numbers/bools are included in both scalar and range semantics.
        (
            Some(Value::Number(0.0)),
            Value::Bool(false),
            Value::Bool(false),
            Value::Bool(false),
            Value::Bool(false),
        ),
        (
            Some(Value::Number(2.0)),
            Value::Bool(true),
            Value::Bool(true),
            Value::Bool(true),
            Value::Bool(true),
        ),
        (
            Some(Value::Bool(false)),
            Value::Bool(false),
            Value::Bool(false),
            Value::Bool(false),
            Value::Bool(false),
        ),
        (
            Some(Value::Bool(true)),
            Value::Bool(true),
            Value::Bool(true),
            Value::Bool(true),
            Value::Bool(true),
        ),
        // Errors always propagate (even if the result is otherwise known).
        (
            Some(Value::Error(ErrorKind::Div0)),
            Value::Error(ErrorKind::Div0),
            Value::Error(ErrorKind::Div0),
            Value::Error(ErrorKind::Div0),
            Value::Error(ErrorKind::Div0),
        ),
    ];

    for (a1, expected_b1, expected_b2, expected_c1, expected_c2) in cases {
        match a1 {
            None => engine.clear_cell("Sheet1", "A1").unwrap(),
            Some(v) => engine.set_cell_value("Sheet1", "A1", v.clone()).unwrap(),
        };

        engine.recalculate_single_threaded();

        assert_eq!(engine.get_cell_value("Sheet1", "B1"), *expected_b1);
        assert_eq!(engine.get_cell_value("Sheet1", "B2"), *expected_b2);
        assert_eq!(engine.get_cell_value("Sheet1", "C1"), *expected_c1);
        assert_eq!(engine.get_cell_value("Sheet1", "C2"), *expected_c2);

        for (formula, cell) in [
            ("=AND(A1)", "B1"),
            ("=OR(A1)", "B2"),
            ("=AND(A1:A1)", "C1"),
            ("=OR(A1:A1)", "C2"),
        ] {
            assert_engine_matches_ast(&engine, formula, cell);
        }
    }
}

#[test]
fn bytecode_backend_propagates_error_literals_through_and_or() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=AND(TRUE, #DIV/0!)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A2", "=OR(FALSE, #DIV/0!)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A3", "=AND(FALSE, #DIV/0!)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A4", "=OR(TRUE, #DIV/0!)")
        .unwrap();

    let stats = engine.bytecode_compile_stats();
    assert_eq!(
        stats.fallback,
        0,
        "expected all AND/OR formulas to compile to bytecode (report={:?})",
        engine.bytecode_compile_report(100)
    );
    assert_eq!(stats.total_formula_cells, 4);
    assert_eq!(stats.compiled, 4);

    engine.recalculate_single_threaded();

    for (formula, cell) in [
        ("=AND(TRUE, #DIV/0!)", "A1"),
        ("=OR(FALSE, #DIV/0!)", "A2"),
        ("=AND(FALSE, #DIV/0!)", "A3"),
        ("=OR(TRUE, #DIV/0!)", "A4"),
    ] {
        assert_eq!(
            engine.get_cell_value("Sheet1", cell),
            Value::Error(ErrorKind::Div0)
        );
        assert_engine_matches_ast(&engine, formula, cell);
    }
}

#[test]
fn bytecode_backend_propagates_error_literals_through_xor() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=XOR(TRUE, #DIV/0!)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A2", "=XOR(FALSE, #DIV/0!)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A3", "=XOR(#DIV/0!)")
        .unwrap();

    let stats = engine.bytecode_compile_stats();
    assert_eq!(
        stats.fallback,
        0,
        "expected all XOR formulas to compile to bytecode (report={:?})",
        engine.bytecode_compile_report(100)
    );
    assert_eq!(stats.total_formula_cells, 3);
    assert_eq!(stats.compiled, 3);

    engine.recalculate_single_threaded();

    for (formula, cell) in [
        ("=XOR(TRUE, #DIV/0!)", "A1"),
        ("=XOR(FALSE, #DIV/0!)", "A2"),
        ("=XOR(#DIV/0!)", "A3"),
    ] {
        assert_eq!(
            engine.get_cell_value("Sheet1", cell),
            Value::Error(ErrorKind::Div0)
        );
        assert_engine_matches_ast(&engine, formula, cell);
    }
}

#[test]
fn bytecode_backend_matches_ast_for_reference_union_and_intersection() {
    let mut engine = Engine::new();

    // Populate a small grid of numbers:
    // A1:C3 = 1..=9 (row-major).
    let mut n = 1.0;
    for row in 1..=3 {
        for col in ["A", "B", "C"] {
            engine
                .set_cell_value("Sheet1", &format!("{col}{row}"), n)
                .unwrap();
            n += 1.0;
        }
    }

    engine
        .set_cell_formula("Sheet1", "D1", "=SUM((A1:A3,B1:B3))")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "D2", "=COUNT((A1:A3,B1:B3))")
        .unwrap();
    engine
        // Place this formula outside the operand ranges so the engine's conservative dependency
        // analysis (which treats the full operand ranges as precedents) does not create a spurious
        // circular reference.
        .set_cell_formula("Sheet1", "E3", "=SUM((A1:C3 B2:D4))")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "D4", "=SUM((A1:A3 B1:B3))")
        .unwrap();

    assert_eq!(
        engine.bytecode_program_count(),
        4,
        "expected all union/intersection formulas to compile to bytecode"
    );

    engine.recalculate_single_threaded();

    for (formula, cell) in [
        ("=SUM((A1:A3,B1:B3))", "D1"),
        ("=COUNT((A1:A3,B1:B3))", "D2"),
        ("=SUM((A1:C3 B2:D4))", "E3"),
        ("=SUM((A1:A3 B1:B3))", "D4"),
    ] {
        assert_engine_matches_ast(&engine, formula, cell);
    }

    // Sanity check expected results:
    // - SUM((A1:A3,B1:B3)) = (1+4+7) + (2+5+8) = 27
    // - COUNT(...) = 6
    // - SUM((A1:C3 B2:D4)) = SUM(B2:C3) = 5+6+8+9 = 28
    // - SUM((A1:A3 B1:B3)) => #NULL! (empty intersection)
    assert_eq!(engine.get_cell_value("Sheet1", "D1"), Value::Number(27.0));
    assert_eq!(engine.get_cell_value("Sheet1", "D2"), Value::Number(6.0));
    assert_eq!(engine.get_cell_value("Sheet1", "E3"), Value::Number(28.0));
    assert_eq!(
        engine.get_cell_value("Sheet1", "D4"),
        Value::Error(ErrorKind::Null)
    );
}

#[test]
fn bytecode_backend_reference_union_error_precedence_matches_ast() {
    let mut engine = Engine::new();

    // First union area: C1:C10 (vertical strip).
    for row in 1..=10 {
        engine
            .set_cell_value("Sheet1", &format!("C{row}"), row as f64)
            .unwrap();
    }

    // Second union area: A2:D3 overlaps C2:C3 with the first area. Place two distinct errors in
    // the *unique* portion so error precedence depends on union iteration order:
    // - D2 should be visited before A3 under the AST evaluator's row-major scan.
    engine.set_cell_value("Sheet1", "A2", 0.0).unwrap();
    engine.set_cell_value("Sheet1", "B2", 0.0).unwrap();
    engine.set_cell_value("Sheet1", "C2", 0.0).unwrap();
    engine
        .set_cell_value("Sheet1", "D2", Value::Error(ErrorKind::Div0))
        .unwrap();
    engine
        .set_cell_value("Sheet1", "A3", Value::Error(ErrorKind::Value))
        .unwrap();
    engine.set_cell_value("Sheet1", "B3", 0.0).unwrap();
    engine.set_cell_value("Sheet1", "C3", 0.0).unwrap();
    engine.set_cell_value("Sheet1", "D3", 0.0).unwrap();

    engine
        .set_cell_formula("Sheet1", "E11", "=XOR((C1:C10,A2:D3))")
        .unwrap();

    let stats = engine.bytecode_compile_stats();
    assert_eq!(
        stats.fallback,
        0,
        "expected XOR + union formula to compile to bytecode (report={:?})",
        engine.bytecode_compile_report(100)
    );
    assert_eq!(stats.total_formula_cells, 1);
    assert_eq!(stats.compiled, 1);

    engine.recalculate_single_threaded();

    assert_engine_matches_ast(&engine, "=XOR((C1:C10,A2:D3))", "E11");
    assert_eq!(
        engine.get_cell_value("Sheet1", "E11"),
        Value::Error(ErrorKind::Div0)
    );
}

#[test]
fn bytecode_backend_reference_union_sum_error_precedence_matches_ast() {
    let mut engine = Engine::new();

    // First union area: C1:C10 (vertical strip).
    for row in 1..=10 {
        engine
            .set_cell_value("Sheet1", &format!("C{row}"), row as f64)
            .unwrap();
    }

    // Second union area: A2:D3 overlaps C2:C3 with the first area. Place two distinct errors in
    // the *unique* portion so error precedence depends on union iteration order:
    // - D2 should be visited before A3 under the AST evaluator's row-major scan.
    engine.set_cell_value("Sheet1", "A2", 0.0).unwrap();
    engine.set_cell_value("Sheet1", "B2", 0.0).unwrap();
    engine.set_cell_value("Sheet1", "C2", 0.0).unwrap();
    engine
        .set_cell_value("Sheet1", "D2", Value::Error(ErrorKind::Div0))
        .unwrap();
    engine
        .set_cell_value("Sheet1", "A3", Value::Error(ErrorKind::Value))
        .unwrap();
    engine.set_cell_value("Sheet1", "B3", 0.0).unwrap();
    engine.set_cell_value("Sheet1", "C3", 0.0).unwrap();
    engine.set_cell_value("Sheet1", "D3", 0.0).unwrap();

    engine
        .set_cell_formula("Sheet1", "E11", "=SUM((C1:C10,A2:D3))")
        .unwrap();

    let stats = engine.bytecode_compile_stats();
    assert_eq!(
        stats.fallback,
        0,
        "expected SUM + union formula to compile to bytecode (report={:?})",
        engine.bytecode_compile_report(100)
    );
    assert_eq!(stats.total_formula_cells, 1);
    assert_eq!(stats.compiled, 1);

    engine.recalculate_single_threaded();

    assert_engine_matches_ast(&engine, "=SUM((C1:C10,A2:D3))", "E11");
    assert_eq!(
        engine.get_cell_value("Sheet1", "E11"),
        Value::Error(ErrorKind::Div0)
    );
}

#[test]
fn bytecode_backend_reference_union_concat_dedups_overlaps_and_preserves_row_major_order() {
    let mut engine = Engine::new();

    // First union area: C1:C10 (vertical strip).
    for row in 1..=10 {
        engine
            .set_cell_value("Sheet1", &format!("C{row}"), row as f64)
            .unwrap();
    }

    // Second union area: A2:D3 overlaps C2:C3 with the first area. Populate the unique cells with
    // distinct text so both overlap-dedup and row-major visitation order are observable.
    engine.set_cell_value("Sheet1", "A2", "A2").unwrap();
    engine.set_cell_value("Sheet1", "B2", "B2").unwrap();
    engine.set_cell_value("Sheet1", "D2", "D2").unwrap();
    engine.set_cell_value("Sheet1", "A3", "A3").unwrap();
    engine.set_cell_value("Sheet1", "B3", "B3").unwrap();
    engine.set_cell_value("Sheet1", "D3", "D3").unwrap();

    engine
        .set_cell_formula("Sheet1", "E11", "=CONCAT((C1:C10,A2:D3))")
        .unwrap();

    let stats = engine.bytecode_compile_stats();
    assert_eq!(
        stats.fallback,
        0,
        "expected CONCAT + union formula to compile to bytecode (report={:?})",
        engine.bytecode_compile_report(100)
    );
    assert_eq!(stats.total_formula_cells, 1);
    assert_eq!(stats.compiled, 1);

    engine.recalculate_single_threaded();

    assert_engine_matches_ast(&engine, "=CONCAT((C1:C10,A2:D3))", "E11");
    assert_eq!(
        engine.get_cell_value("Sheet1", "E11"),
        Value::Text("12345678910A2B2D2A3B3D3".into())
    );
}

#[test]
fn bytecode_backend_reference_union_aggregates_error_precedence_matches_ast() {
    let mut engine = Engine::new();

    // First union area: C1:C10 (vertical strip).
    for row in 1..=10 {
        engine
            .set_cell_value("Sheet1", &format!("C{row}"), row as f64)
            .unwrap();
    }

    // Second union area: A2:D3 overlaps C2:C3 with the first area. Place two distinct errors in
    // the *unique* portion so error precedence depends on correct row-major union iteration:
    // - D2 should be visited before A3 under the AST evaluator's row-major scan.
    engine.set_cell_value("Sheet1", "A2", 0.0).unwrap();
    engine.set_cell_value("Sheet1", "B2", 0.0).unwrap();
    engine.set_cell_value("Sheet1", "C2", 0.0).unwrap();
    engine
        .set_cell_value("Sheet1", "D2", Value::Error(ErrorKind::Div0))
        .unwrap();
    engine
        .set_cell_value("Sheet1", "A3", Value::Error(ErrorKind::Value))
        .unwrap();
    engine.set_cell_value("Sheet1", "B3", 0.0).unwrap();
    engine.set_cell_value("Sheet1", "C3", 0.0).unwrap();
    engine.set_cell_value("Sheet1", "D3", 0.0).unwrap();

    engine
        .set_cell_formula("Sheet1", "E11", "=AVERAGE((C1:C10,A2:D3))")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "E12", "=MIN((C1:C10,A2:D3))")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "E13", "=MAX((C1:C10,A2:D3))")
        .unwrap();

    let stats = engine.bytecode_compile_stats();
    assert_eq!(
        stats.fallback,
        0,
        "expected aggregate + union formulas to compile to bytecode (report={:?})",
        engine.bytecode_compile_report(100)
    );
    assert_eq!(stats.total_formula_cells, 3);
    assert_eq!(stats.compiled, 3);

    engine.recalculate_single_threaded();

    for (formula, cell) in [
        ("=AVERAGE((C1:C10,A2:D3))", "E11"),
        ("=MIN((C1:C10,A2:D3))", "E12"),
        ("=MAX((C1:C10,A2:D3))", "E13"),
    ] {
        assert_engine_matches_ast(&engine, formula, cell);
        assert_eq!(
            engine.get_cell_value("Sheet1", cell),
            Value::Error(ErrorKind::Div0)
        );
    }
}

#[test]
fn bytecode_backend_reference_algebra_accepts_let_single_cell_reference_locals() {
    let mut engine = Engine::new();

    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "B1", 2.0).unwrap();

    engine
        .set_cell_formula("Sheet1", "C1", "=LET(x,A1,SUM((x,B1)))")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "C2", "=LET(x,A1,SUM((x A1:B1)))")
        .unwrap();

    let stats = engine.bytecode_compile_stats();
    assert_eq!(
        stats.fallback,
        0,
        "expected LET + union/intersection formulas to compile to bytecode (report={:?})",
        engine.bytecode_compile_report(100)
    );
    assert_eq!(stats.total_formula_cells, 2);
    assert_eq!(stats.compiled, 2);

    engine.recalculate_single_threaded();

    for (formula, cell) in [
        ("=LET(x,A1,SUM((x,B1)))", "C1"),
        ("=LET(x,A1,SUM((x A1:B1)))", "C2"),
    ] {
        assert_engine_matches_ast(&engine, formula, cell);
    }

    assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::Number(3.0));
    assert_eq!(engine.get_cell_value("Sheet1", "C2"), Value::Number(1.0));
}

#[test]
fn bytecode_backend_reference_algebra_as_formula_result_matches_ast() {
    let mut engine = Engine::new();

    // Populate a small grid of numbers:
    // A1:C3 = 1..=9 (row-major).
    let mut n = 1.0;
    for row in 1..=3 {
        for col in ["A", "B", "C"] {
            engine
                .set_cell_value("Sheet1", &format!("{col}{row}"), n)
                .unwrap();
            n += 1.0;
        }
    }

    engine
        .set_cell_formula("Sheet1", "E1", "=(A1:A3,B1:B3)")
        .unwrap();
    // Spill result: intersection of A1:C3 and B2:D4 is B2:C3 (2x2).
    engine
        .set_cell_formula("Sheet1", "E2", "=(A1:C3 B2:D4)")
        .unwrap();

    let stats = engine.bytecode_compile_stats();
    assert_eq!(
        stats.fallback,
        0,
        "expected reference algebra formula results to compile to bytecode (report={:?})",
        engine.bytecode_compile_report(100)
    );
    assert_eq!(stats.total_formula_cells, 2);
    assert_eq!(stats.compiled, 2);

    engine.recalculate_single_threaded();

    assert_engine_matches_ast(&engine, "=(A1:A3,B1:B3)", "E1");
    assert_engine_spill_matches_ast(&engine, "=(A1:C3 B2:D4)", "E2");

    // Discontiguous unions cannot be spilled as a single rectangle.
    assert_eq!(engine.get_cell_value("Sheet1", "E1"), Value::Error(ErrorKind::Value));
}

#[test]
fn bytecode_backend_matches_ast_for_information_functions_scalar() {
    let mut engine = Engine::new();

    // Inputs for the information functions.
    // A1 left blank.
    engine.set_cell_value("Sheet1", "A2", "").unwrap(); // empty string is not blank

    engine.set_cell_value("Sheet1", "A3", 123.0).unwrap();
    engine.set_cell_value("Sheet1", "A4", "123").unwrap();

    engine.set_cell_value("Sheet1", "A5", "foo").unwrap();
    engine.set_cell_value("Sheet1", "A6", 1.0).unwrap();

    engine.set_cell_value("Sheet1", "A7", true).unwrap();
    engine.set_cell_value("Sheet1", "A8", 0.0).unwrap();

    engine
        .set_cell_value("Sheet1", "A9", Value::Error(ErrorKind::NA))
        .unwrap();
    engine
        .set_cell_value("Sheet1", "A10", Value::Error(ErrorKind::Div0))
        .unwrap();

    engine.set_cell_value("Sheet1", "A12", 1.0).unwrap();
    engine
        .set_cell_value("Sheet1", "A13", Value::Error(ErrorKind::Ref))
        .unwrap();

    engine.set_cell_value("Sheet1", "A14", true).unwrap();
    engine.set_cell_value("Sheet1", "A15", "5").unwrap();
    engine
        .set_cell_value("Sheet1", "A16", Value::Error(ErrorKind::Div0))
        .unwrap();

    engine.set_cell_value("Sheet1", "A18", "hi").unwrap();
    engine.set_cell_value("Sheet1", "A19", 1.0).unwrap();
    engine
        .set_cell_value("Sheet1", "A20", Value::Error(ErrorKind::Div0))
        .unwrap();

    // TYPE inputs.
    // A21 left blank.
    engine.set_cell_value("Sheet1", "A22", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A23", "x").unwrap();
    engine.set_cell_value("Sheet1", "A24", true).unwrap();
    engine
        .set_cell_value("Sheet1", "A25", Value::Error(ErrorKind::Value))
        .unwrap();
    engine.set_cell_value("Sheet1", "A26", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A27", 2.0).unwrap();

    // Formulas are all placed in column B so they share the same relative reference pattern; the
    // bytecode cache should compile one program per distinct function pattern.
    engine
        .set_cell_formula("Sheet1", "B1", "=ISBLANK(A1)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B2", "=ISBLANK(A2)")
        .unwrap();

    engine
        .set_cell_formula("Sheet1", "B3", "=ISNUMBER(A3)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B4", "=ISNUMBER(A4)")
        .unwrap();

    engine
        .set_cell_formula("Sheet1", "B5", "=ISTEXT(A5)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B6", "=ISTEXT(A6)")
        .unwrap();

    engine
        .set_cell_formula("Sheet1", "B7", "=ISLOGICAL(A7)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B8", "=ISLOGICAL(A8)")
        .unwrap();

    engine
        .set_cell_formula("Sheet1", "B9", "=ISERR(A9)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B10", "=ISERR(A10)")
        .unwrap();

    engine
        .set_cell_formula("Sheet1", "B12", "=ERROR.TYPE(A12)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B13", "=ERROR.TYPE(A13)")
        .unwrap();

    engine.set_cell_formula("Sheet1", "B14", "=N(A14)").unwrap();
    engine.set_cell_formula("Sheet1", "B15", "=N(A15)").unwrap();
    engine.set_cell_formula("Sheet1", "B16", "=N(A16)").unwrap();

    engine.set_cell_formula("Sheet1", "B18", "=T(A18)").unwrap();
    engine.set_cell_formula("Sheet1", "B19", "=T(A19)").unwrap();
    engine.set_cell_formula("Sheet1", "B20", "=T(A20)").unwrap();

    engine
        .set_cell_formula("Sheet1", "B21", "=TYPE(A21)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B22", "=TYPE(A22)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B23", "=TYPE(A23)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B24", "=TYPE(A24)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B25", "=TYPE(A25)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B26", "=TYPE(A26:A27)")
        .unwrap();

    // 9 information functions + TYPE has 2 distinct shapes (scalar vs multi-cell range).
    assert_eq!(engine.bytecode_program_count(), 10);

    engine.recalculate_single_threaded();

    // ISBLANK: blank vs empty-string.
    assert_engine_matches_ast(&engine, "=ISBLANK(A1)", "B1");
    assert_engine_matches_ast(&engine, "=ISBLANK(A2)", "B2");

    // ISNUMBER/ISTEXT/ISLOGICAL scalar behavior.
    assert_engine_matches_ast(&engine, "=ISNUMBER(A3)", "B3");
    assert_engine_matches_ast(&engine, "=ISNUMBER(A4)", "B4");
    assert_engine_matches_ast(&engine, "=ISTEXT(A5)", "B5");
    assert_engine_matches_ast(&engine, "=ISTEXT(A6)", "B6");
    assert_engine_matches_ast(&engine, "=ISLOGICAL(A7)", "B7");
    assert_engine_matches_ast(&engine, "=ISLOGICAL(A8)", "B8");

    // ISERR distinguishes #N/A.
    assert_engine_matches_ast(&engine, "=ISERR(A9)", "B9");
    assert_engine_matches_ast(&engine, "=ISERR(A10)", "B10");

    // ERROR.TYPE returns #N/A for non-errors; returns a numeric code for errors.
    assert_engine_matches_ast(&engine, "=ERROR.TYPE(A12)", "B12");
    assert_engine_matches_ast(&engine, "=ERROR.TYPE(A13)", "B13");

    // N/T propagate errors.
    assert_engine_matches_ast(&engine, "=N(A14)", "B14");
    assert_engine_matches_ast(&engine, "=N(A15)", "B15");
    assert_engine_matches_ast(&engine, "=N(A16)", "B16");
    assert_engine_matches_ast(&engine, "=T(A18)", "B18");
    assert_engine_matches_ast(&engine, "=T(A19)", "B19");
    assert_engine_matches_ast(&engine, "=T(A20)", "B20");

    // TYPE on scalar values + multi-cell range (64).
    assert_engine_matches_ast(&engine, "=TYPE(A21)", "B21");
    assert_engine_matches_ast(&engine, "=TYPE(A22)", "B22");
    assert_engine_matches_ast(&engine, "=TYPE(A23)", "B23");
    assert_engine_matches_ast(&engine, "=TYPE(A24)", "B24");
    assert_engine_matches_ast(&engine, "=TYPE(A25)", "B25");
    assert_engine_matches_ast(&engine, "=TYPE(A26:A27)", "B26");
}

#[test]
fn bytecode_backend_matches_ast_for_information_functions_with_range_args() {
    let mut engine = Engine::new();
    // A1 left blank.
    engine.set_cell_value("Sheet1", "A2", "").unwrap();
    engine.set_cell_value("Sheet1", "A3", 3.0).unwrap();
    engine.set_cell_value("Sheet1", "A4", "hello").unwrap();
    engine.set_cell_value("Sheet1", "A5", true).unwrap();
    engine
        .set_cell_value("Sheet1", "A6", Value::Error(ErrorKind::Div0))
        .unwrap();
    engine
        .set_cell_value("Sheet1", "A7", Value::Error(ErrorKind::NA))
        .unwrap();

    // Each formula spills down 7 rows, so place them in separate columns to avoid overlap.
    engine
        .set_cell_formula("Sheet1", "B1", "=ISBLANK(A1:A7)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "C1", "=ISNUMBER(A1:A7)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "D1", "=ISTEXT(A1:A7)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "E1", "=ISLOGICAL(A1:A7)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "F1", "=ISERROR(A1:A7)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "G1", "=ISNA(A1:A7)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "H1", "=ISERR(A1:A7)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "I1", "=ERROR.TYPE(A1:A7)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "J1", "=N(A1:A7)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "K1", "=T(A1:A7)")
        .unwrap();

    // Array literals should also be eligible for these functions.
    engine
        .set_cell_formula("Sheet1", "B10", "=ISNUMBER({1,\"a\"})")
        .unwrap();

    // All formulas in this fixture should be bytecode-eligible.
    assert!(
        engine.bytecode_compile_report(32).is_empty(),
        "expected all information function formulas to compile to bytecode"
    );

    engine.recalculate_single_threaded();

    for (formula, cell) in [
        ("=ISBLANK(A1:A7)", "B1"),
        ("=ISNUMBER(A1:A7)", "C1"),
        ("=ISTEXT(A1:A7)", "D1"),
        ("=ISLOGICAL(A1:A7)", "E1"),
        ("=ISERROR(A1:A7)", "F1"),
        ("=ISNA(A1:A7)", "G1"),
        ("=ISERR(A1:A7)", "H1"),
        ("=ERROR.TYPE(A1:A7)", "I1"),
        ("=N(A1:A7)", "J1"),
        ("=T(A1:A7)", "K1"),
        ("=ISNUMBER({1,\"a\"})", "B10"),
    ] {
        assert_engine_spill_matches_ast(&engine, formula, cell);
    }
}

#[test]
fn bytecode_backend_matches_ast_for_reference_functions() {
    let mut engine = Engine::new();

    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();

    engine.set_cell_formula("Sheet1", "C1", "=ROW()").unwrap();
    engine
        .set_cell_formula("Sheet1", "C2", "=COLUMN()")
        .unwrap();
    engine.set_cell_formula("Sheet1", "C3", "=ROW(A1)").unwrap();
    engine
        .set_cell_formula("Sheet1", "C4", "=COLUMN(A1)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "C5", "=ROWS(A1:B3)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "C6", "=COLUMNS(A1:B3)")
        .unwrap();

    engine
        .set_cell_formula("Sheet1", "C7", "=ADDRESS(1,1)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "C8", "=ADDRESS(1,1,4)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "C9", "=ADDRESS(1,1,1,FALSE)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "C10", "=ADDRESS(1,1,1,TRUE,\"1\")")
        .unwrap();

    // Ensure all formulas were eligible for bytecode compilation.
    assert_eq!(engine.bytecode_program_count(), 10);

    engine.recalculate_single_threaded();

    for (formula, cell) in [
        ("=ROW()", "C1"),
        ("=COLUMN()", "C2"),
        ("=ROW(A1)", "C3"),
        ("=COLUMN(A1)", "C4"),
        ("=ROWS(A1:B3)", "C5"),
        ("=COLUMNS(A1:B3)", "C6"),
        ("=ADDRESS(1,1)", "C7"),
        ("=ADDRESS(1,1,4)", "C8"),
        ("=ADDRESS(1,1,1,FALSE)", "C9"),
        ("=ADDRESS(1,1,1,TRUE,\"1\")", "C10"),
    ] {
        assert_engine_matches_ast(&engine, formula, cell);
    }
}

#[test]
fn bytecode_backend_allows_row_column_on_whole_row_and_column_ranges() {
    let mut engine = Engine::new();
    engine
        // Place the formulas outside of the referenced whole-row/whole-column ranges so we don't
        // create an accidental circular reference (range nodes participate in calc ordering).
        .set_cell_formula("Sheet1", "AA6001", "=SUM(ROW(1:5000))")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "AA6002", "=SUM(COLUMN(A:Z))")
        .unwrap();

    // Regression: whole-row / whole-column references expand to the full sheet width/height in the
    // lowered bytecode AST, but ROW/COLUMN treat them as 1-D arrays and should remain bytecode-eligible.
    assert_eq!(engine.bytecode_program_count(), 2);

    engine.recalculate_single_threaded();

    assert_engine_matches_ast(&engine, "=SUM(ROW(1:5000))", "AA6001");
    assert_engine_matches_ast(&engine, "=SUM(COLUMN(A:Z))", "AA6002");
}

#[test]
fn bytecode_backend_distinguishes_missing_args_from_blank_cells_for_address() {
    let mut engine = Engine::new();
    // A1 is left blank (unset).

    engine
        .set_cell_formula("Sheet1", "B1", "=ADDRESS(1,1,,FALSE)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B2", "=ADDRESS(1,1,A1,FALSE)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B3", "=ADDRESS(1,1,IF(FALSE,1,),FALSE)")
        .unwrap();

    // Ensure both formulas compiled to bytecode.
    assert_eq!(engine.bytecode_program_count(), 3);

    engine.recalculate_single_threaded();

    assert_engine_matches_ast(&engine, "=ADDRESS(1,1,,FALSE)", "B1");
    assert_engine_matches_ast(&engine, "=ADDRESS(1,1,A1,FALSE)", "B2");
    assert_engine_matches_ast(&engine, "=ADDRESS(1,1,IF(FALSE,1,),FALSE)", "B3");

    // Regression: omitted abs_num defaults to 1, but a blank cell passed for abs_num should error.
    assert_ne!(
        engine.get_cell_value("Sheet1", "B1"),
        engine.get_cell_value("Sheet1", "B2")
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "B2"),
        Value::Error(ErrorKind::Value)
    );

    // A blank *value* produced by an expression is not the same as an omitted argument; it should
    // not trigger ADDRESS's defaulting behavior.
    assert_eq!(
        engine.get_cell_value("Sheet1", "B3"),
        Value::Error(ErrorKind::Value)
    );
}

#[test]
fn bytecode_backend_rows_and_columns_accept_array_literals() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=ROWS({1,2;3,4})")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A2", "=COLUMNS({1,2;3,4})")
        .unwrap();

    assert_eq!(engine.bytecode_program_count(), 2);

    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "A2"), Value::Number(2.0));
    assert_engine_matches_ast(&engine, "=ROWS({1,2;3,4})", "A1");
    assert_engine_matches_ast(&engine, "=COLUMNS({1,2;3,4})", "A2");
}

#[test]
fn bytecode_backend_xor_reference_semantics_match_ast() {
    let mut engine = Engine::new();
    engine.set_cell_formula("Sheet1", "B1", "=XOR(A1)").unwrap();
    // Scalar text values coerce like NOT(); reference text values are ignored.
    engine
        .set_cell_formula("Sheet1", "B2", "=XOR(\"TRUE\")")
        .unwrap();

    assert_eq!(engine.bytecode_program_count(), 2);

    engine.recalculate_single_threaded();
    assert_engine_matches_ast(&engine, "=XOR(\"TRUE\")", "B2");
    assert_eq!(engine.get_cell_value("Sheet1", "B2"), Value::Bool(true));

    let cases: &[(Option<Value>, Value)] = &[
        (None, Value::Bool(false)),
        (Some(Value::Text("TRUE".to_string())), Value::Bool(false)),
        (Some(Value::Text("hello".to_string())), Value::Bool(false)),
        (Some(Value::Number(0.0)), Value::Bool(false)),
        (Some(Value::Number(2.0)), Value::Bool(true)),
        (Some(Value::Bool(false)), Value::Bool(false)),
        (Some(Value::Bool(true)), Value::Bool(true)),
        (
            Some(Value::Error(ErrorKind::Div0)),
            Value::Error(ErrorKind::Div0),
        ),
    ];

    for (a1, expected) in cases {
        match a1 {
            None => engine.clear_cell("Sheet1", "A1").unwrap(),
            Some(v) => engine.set_cell_value("Sheet1", "A1", v.clone()).unwrap(),
        };

        engine.recalculate_single_threaded();
        assert_eq!(engine.get_cell_value("Sheet1", "B1"), *expected);
        assert_engine_matches_ast(&engine, "=XOR(A1)", "B1");
    }
}

#[test]
fn bytecode_backend_xor_array_semantics_match_ast() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=XOR({TRUE,FALSE,TRUE})")
        .unwrap();
    // Text values inside arrays are ignored (unlike scalar text args, which coerce).
    engine
        .set_cell_formula("Sheet1", "A2", "=XOR({\"TRUE\"})")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A3", "=XOR({\"TRUE\",0,1,\"x\"})")
        .unwrap();
    // Errors in arrays should propagate.
    engine
        .set_cell_formula("Sheet1", "A4", "=XOR({#DIV/0!,TRUE})")
        .unwrap();

    assert_eq!(engine.bytecode_program_count(), 4);

    engine.recalculate_single_threaded();

    assert_engine_matches_ast(&engine, "=XOR({TRUE,FALSE,TRUE})", "A1");
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Bool(false));

    assert_engine_matches_ast(&engine, "=XOR({\"TRUE\"})", "A2");
    assert_eq!(engine.get_cell_value("Sheet1", "A2"), Value::Bool(false));

    assert_engine_matches_ast(&engine, "=XOR({\"TRUE\",0,1,\"x\"})", "A3");
    assert_eq!(engine.get_cell_value("Sheet1", "A3"), Value::Bool(true));

    assert_engine_matches_ast(&engine, "=XOR({#DIV/0!,TRUE})", "A4");
    assert_eq!(
        engine.get_cell_value("Sheet1", "A4"),
        Value::Error(ErrorKind::Div0)
    );
}

#[test]
fn bytecode_backend_and_array_semantics_match_ast() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=AND({TRUE,FALSE,TRUE})")
        .unwrap();
    // Text values inside arrays are ignored (unlike scalar text args, which error).
    engine
        .set_cell_formula("Sheet1", "A2", "=AND({\"TRUE\"})")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A3", "=AND({\"TRUE\",0,1,\"x\"})")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A4", "=AND({#DIV/0!,TRUE})")
        .unwrap();

    assert_eq!(engine.bytecode_program_count(), 4);

    engine.recalculate_single_threaded();

    assert_engine_matches_ast(&engine, "=AND({TRUE,FALSE,TRUE})", "A1");
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Bool(false));

    assert_engine_matches_ast(&engine, "=AND({\"TRUE\"})", "A2");
    assert_eq!(engine.get_cell_value("Sheet1", "A2"), Value::Bool(true));

    assert_engine_matches_ast(&engine, "=AND({\"TRUE\",0,1,\"x\"})", "A3");
    assert_eq!(engine.get_cell_value("Sheet1", "A3"), Value::Bool(false));

    assert_engine_matches_ast(&engine, "=AND({#DIV/0!,TRUE})", "A4");
    assert_eq!(
        engine.get_cell_value("Sheet1", "A4"),
        Value::Error(ErrorKind::Div0)
    );
}

#[test]
fn bytecode_backend_or_array_semantics_match_ast() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=OR({TRUE,FALSE,TRUE})")
        .unwrap();
    // Text values inside arrays are ignored (unlike scalar text args, which error).
    engine
        .set_cell_formula("Sheet1", "A2", "=OR({\"TRUE\"})")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A3", "=OR({\"TRUE\",0,1,\"x\"})")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A4", "=OR({#DIV/0!,TRUE})")
        .unwrap();

    assert_eq!(engine.bytecode_program_count(), 4);

    engine.recalculate_single_threaded();

    assert_engine_matches_ast(&engine, "=OR({TRUE,FALSE,TRUE})", "A1");
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Bool(true));

    assert_engine_matches_ast(&engine, "=OR({\"TRUE\"})", "A2");
    assert_eq!(engine.get_cell_value("Sheet1", "A2"), Value::Bool(false));

    assert_engine_matches_ast(&engine, "=OR({\"TRUE\",0,1,\"x\"})", "A3");
    assert_eq!(engine.get_cell_value("Sheet1", "A3"), Value::Bool(true));

    assert_engine_matches_ast(&engine, "=OR({#DIV/0!,TRUE})", "A4");
    assert_eq!(
        engine.get_cell_value("Sheet1", "A4"),
        Value::Error(ErrorKind::Div0)
    );
}

#[test]
fn bytecode_backend_rng_is_stable_within_one_recalc_tick_and_changes_across_ticks() {
    let mut engine = Engine::new();
    engine.set_cell_formula("Sheet1", "A1", "=RAND()").unwrap();
    // Ensure we're exercising the bytecode backend for the RAND() formula.
    assert_eq!(engine.bytecode_program_count(), 1);

    engine.set_cell_formula("Sheet1", "B1", "=A1").unwrap();
    engine.set_cell_formula("Sheet1", "C1", "=A1").unwrap();

    engine.recalculate_single_threaded();

    let first = engine.get_cell_value("Sheet1", "A1");
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), first);
    assert_eq!(engine.get_cell_value("Sheet1", "C1"), first);

    // RNG should change on each recalc tick; allow a few attempts to avoid pathological
    // collisions in the float representation.
    let mut changed = false;
    for _ in 0..5 {
        engine.recalculate_single_threaded();
        let next = engine.get_cell_value("Sheet1", "A1");
        assert_eq!(engine.get_cell_value("Sheet1", "B1"), next);
        assert_eq!(engine.get_cell_value("Sheet1", "C1"), next);
        if next != first {
            changed = true;
            break;
        }
    }
    assert!(changed, "expected RAND() to change across recalculations");
}

proptest! {
    #![proptest_config(ProptestConfig { cases: 32, .. ProptestConfig::default() })]
    #[test]
    fn bytecode_backend_matches_ast_for_random_supported_formulas(
        a in -1000f64..1000f64,
        b in -1000f64..1000f64,
        digits in -6i32..6i32,
        choice in 0u8..20u8,
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
            14 => "=IF(A1>0, A1/B1, 1)".to_string(),
            15 => "=IFERROR(A1/B1, 7)".to_string(),
            16 => "=IFNA(NA(), A1)".to_string(),
            17 => "=AND(A1, B1)".to_string(),
            18 => "=OR(A1, B1)".to_string(),
            19 => "=ISERROR(A1/B1)".to_string(),
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
    // Volatile (thread-safe; should compile to bytecode).
    engine.set_cell_formula("Sheet1", "A2", "=RAND()").unwrap();
    // Cross-sheet reference.
    engine.set_cell_value("Sheet2", "A1", 42.0).unwrap();
    engine.set_cell_formula("Sheet1", "A3", "=Sheet2!A1").unwrap();
    // Lowering error (unsupported expression): sheet-qualified defined name.
    engine.set_cell_formula("Sheet1", "A4", "=Sheet1!Foo").unwrap();

    let stats = engine.bytecode_compile_stats();
    assert_eq!(stats.total_formula_cells, 4);
    assert_eq!(stats.compiled, 3);
    assert_eq!(stats.fallback, 1);

    // LowerError::Unsupported is mapped to `IneligibleExpr` so compile stats can distinguish
    // structural lowering errors (e.g. unknown sheets/external refs) from missing bytecode
    // implementation.
    let ineligible = stats
        .fallback_reasons
        .get(&BytecodeCompileReason::IneligibleExpr)
        .copied()
        .unwrap_or(0);
    assert_eq!(
        ineligible, 1,
        "expected Sheet1!Foo to be ineligible for bytecode"
    );

    let report = engine.bytecode_compile_report(usize::MAX);
    assert_eq!(report.len(), 1);

    let a2 = parse_a1("A2").unwrap();
    let a3 = parse_a1("A3").unwrap();
    let a4 = parse_a1("A4").unwrap();

    let reason_for = |addr| {
        report
            .iter()
            .find(|e| e.sheet == "Sheet1" && e.addr == addr)
            .map(|e| e.reason.clone())
    };

    assert_eq!(reason_for(a2), None);
    assert_eq!(reason_for(a3), None);
    let a4_reason = reason_for(a4).expect("A4 should appear in fallback report");
    assert!(
        matches!(
            a4_reason,
            BytecodeCompileReason::IneligibleExpr
                | BytecodeCompileReason::LowerError(
                    formula_engine::bytecode::LowerError::Unsupported
                )
        ),
        "unexpected A4 bytecode compile reason: {a4_reason:?}"
    );
}

#[test]
fn bytecode_compile_diagnostics_compiles_indirect() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A2", 1.0).unwrap();
    engine
        .set_cell_formula("Sheet1", "A1", "=SUM(INDIRECT(\"A2\"))")
        .unwrap();

    let stats = engine.bytecode_compile_stats();
    assert_eq!(stats.total_formula_cells, 1);
    assert_eq!(stats.compiled, 1);
    assert_eq!(stats.fallback, 0);
    assert!(
        stats.fallback_reasons.is_empty(),
        "unexpected bytecode fallback reasons: {:?}",
        stats.fallback_reasons
    );

    let report = engine.bytecode_compile_report(10);
    assert!(report.is_empty(), "unexpected bytecode fallback report: {report:?}");

    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(1.0));
}

#[test]
fn bytecode_compile_diagnostics_compiles_offset() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A2", 1.0).unwrap();
    engine
        .set_cell_formula("Sheet1", "A1", "=SUM(OFFSET(A2,0,0))")
        .unwrap();

    let stats = engine.bytecode_compile_stats();
    assert_eq!(stats.total_formula_cells, 1);
    assert_eq!(stats.compiled, 1);
    assert_eq!(stats.fallback, 0);
    assert!(
        stats.fallback_reasons.is_empty(),
        "unexpected bytecode fallback reasons: {:?}",
        stats.fallback_reasons
    );

    let report = engine.bytecode_compile_report(10);
    assert!(report.is_empty(), "unexpected bytecode fallback report: {report:?}");

    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(1.0));
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
fn bytecode_compile_diagnostics_reports_unsupported_function_reason() {
    let mut engine = Engine::new();
    // Register a test-only function so this test doesn't become stale if/when the bytecode backend
    // adds more built-in function implementations (e.g. SIN).
    engine
        .set_cell_formula("Sheet1", "A1", "=BYTECODE_UNSUPPORTED_TEST()")
        .unwrap();

    let stats = engine.bytecode_compile_stats();
    assert_eq!(stats.total_formula_cells, 1);
    assert_eq!(stats.compiled, 0);
    assert_eq!(stats.fallback, 1);
    assert_eq!(
        stats
            .fallback_reasons
            .get(&BytecodeCompileReason::UnsupportedFunction(Arc::from(
                "BYTECODE_UNSUPPORTED_TEST",
            )))
            .copied()
            .unwrap_or(0),
        1
    );

    let report = engine.bytecode_compile_report(10);
    assert_eq!(report.len(), 1);
    assert_eq!(report[0].sheet, "Sheet1");
    assert_eq!(report[0].addr, parse_a1("A1").unwrap());
    assert_eq!(
        report[0].reason,
        BytecodeCompileReason::UnsupportedFunction(Arc::from("BYTECODE_UNSUPPORTED_TEST"))
    );
}

#[test]
fn bytecode_compile_diagnostics_reports_grid_and_range_limits() {
    let mut engine = Engine::new();

    // Out-of-bounds cell reference (column exceeds XFD).
    engine.set_cell_formula("Sheet1", "A1", "=XFE1").unwrap();

    // Huge range: this should still be bytecode-eligible now that aggregates can iterate sparsely
    // without materializing a full range buffer.
    engine
        .set_cell_formula("Sheet1", "B1", "=SUM(A1:XFD1048576)")
        .unwrap();

    let stats = engine.bytecode_compile_stats();
    assert_eq!(stats.total_formula_cells, 2);
    assert_eq!(stats.compiled, 1);
    assert_eq!(stats.fallback, 1);
    assert_eq!(
        stats
            .fallback_reasons
            .get(&BytecodeCompileReason::ExceedsGridLimits)
            .copied()
            .unwrap_or(0),
        1
    );

    let report = engine.bytecode_compile_report(usize::MAX);
    assert_eq!(report.len(), 1);

    let a1 = report
        .iter()
        .find(|e| e.sheet == "Sheet1" && e.addr == parse_a1("A1").unwrap())
        .map(|e| e.reason.clone());
    assert_eq!(a1, Some(BytecodeCompileReason::ExceedsGridLimits));
}

#[test]
fn bytecode_compile_diagnostics_reports_unknown_sheet_reason() {
    let mut engine = Engine::new();

    // Reference to a sheet that does not exist in the workbook.
    engine
        .set_cell_formula("Sheet1", "A1", "=MissingSheet!A1")
        .unwrap();

    let stats = engine.bytecode_compile_stats();
    assert_eq!(stats.total_formula_cells, 1);
    assert_eq!(stats.compiled, 0);
    assert_eq!(stats.fallback, 1);
    assert_eq!(
        stats
            .fallback_reasons
            .get(&BytecodeCompileReason::LowerError(
                bytecode::LowerError::UnknownSheet,
            ))
            .copied()
            .unwrap_or(0),
        1
    );
}

#[test]
fn bytecode_compile_diagnostics_reports_external_reference_reason() {
    let mut engine = Engine::new();

    // External workbook references are supported by the bytecode backend, but external 3D sheet
    // spans (`[Book]Sheet1:Sheet3!A1`) cannot be represented via `ExternalValueProvider`, so they
    // should still fall back to the AST evaluator with an ExternalReference lowering error.
    engine
        .set_cell_formula("Sheet1", "A1", "=[Book.xlsx]Sheet1:Sheet3!A1")
        .unwrap();

    let stats = engine.bytecode_compile_stats();
    assert_eq!(stats.total_formula_cells, 1);
    assert_eq!(stats.compiled, 0);
    assert_eq!(stats.fallback, 1);
    assert_eq!(
        stats
            .fallback_reasons
            .get(&BytecodeCompileReason::LowerError(
                bytecode::LowerError::ExternalReference,
            ))
            .copied()
            .unwrap_or(0),
        1
    );
}

#[test]
fn bytecode_compile_diagnostics_allows_external_reference_through_defined_name() {
    let mut engine = Engine::new();
    engine
        .define_name(
            "EXT",
            NameScope::Workbook,
            NameDefinition::Reference("[Book.xlsx]Sheet1!A1".to_string()),
        )
        .unwrap();

    engine.set_cell_formula("Sheet1", "A1", "=EXT+1").unwrap();

    let stats = engine.bytecode_compile_stats();
    assert_eq!(stats.total_formula_cells, 1);
    assert_eq!(
        stats.compiled,
        1,
        "expected defined-name external workbook refs to compile to bytecode (report={:?})",
        engine.bytecode_compile_report(32)
    );
    assert_eq!(stats.fallback, 0);
}

#[test]
fn bytecode_compile_diagnostics_allows_external_reference_through_defined_name_in_field_access() {
    let mut engine = Engine::new();
    engine
        .define_name(
            "EXT",
            NameScope::Workbook,
            NameDefinition::Reference("[Book.xlsx]Sheet1!A1".to_string()),
        )
        .unwrap();

    // Field access expressions like `(EXT).Price` should still compile to bytecode when the base
    // defined name expands to an external workbook reference.
    //
    // Note: Without parentheses, `EXT.Price` is currently tokenized as a single dotted identifier,
    // so use `(EXT).Price` to ensure we exercise the field access operator.
    engine
        .set_cell_formula("Sheet1", "A1", "=(EXT).Price")
        .unwrap();

    let stats = engine.bytecode_compile_stats();
    assert_eq!(stats.total_formula_cells, 1);
    assert_eq!(
        stats.compiled,
        1,
        "expected field access over external workbook refs to compile to bytecode (report={:?})",
        engine.bytecode_compile_report(32)
    );
    assert_eq!(stats.fallback, 0);
}

#[test]
fn bytecode_compile_diagnostics_reports_unknown_sheet_through_defined_name_formula() {
    let mut engine = Engine::new();

    engine
        .define_name(
            "UNK",
            NameScope::Workbook,
            NameDefinition::Formula("=MissingSheet!A1+1".to_string()),
        )
        .unwrap();

    engine.set_cell_formula("Sheet1", "A1", "=UNK").unwrap();

    let stats = engine.bytecode_compile_stats();
    assert_eq!(stats.total_formula_cells, 1);
    assert_eq!(stats.compiled, 0);
    assert_eq!(stats.fallback, 1);
    assert_eq!(
        stats
            .fallback_reasons
            .get(&BytecodeCompileReason::LowerError(
                bytecode::LowerError::UnknownSheet
            ))
            .copied()
            .unwrap_or(0),
        1
    );
}

#[test]
fn bytecode_compile_diagnostics_allows_cross_sheet_reference_through_defined_name_formula() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet2", "A1", 42.0).unwrap();

    engine
        .define_name(
            "CROSS",
            NameScope::Workbook,
            NameDefinition::Formula("=Sheet2!A1+1".to_string()),
        )
        .unwrap();

    engine.set_cell_formula("Sheet1", "A1", "=CROSS").unwrap();

    let stats = engine.bytecode_compile_stats();
    assert_eq!(stats.total_formula_cells, 1);
    assert_eq!(stats.compiled, 1);
    assert_eq!(stats.fallback, 0);

    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(43.0));
}

#[test]
fn bytecode_compile_diagnostics_reports_not_thread_safe_reason() {
    let mut engine = Engine::new();

    engine
        .set_cell_formula("Sheet1", "A1", "=NOT_THREAD_SAFE_TEST()")
        .unwrap();

    let stats = engine.bytecode_compile_stats();
    assert_eq!(stats.total_formula_cells, 1);
    assert_eq!(stats.compiled, 0);
    assert_eq!(stats.fallback, 1);
    assert_eq!(
        stats
            .fallback_reasons
            .get(&BytecodeCompileReason::NotThreadSafe)
            .copied()
            .unwrap_or(0),
        1
    );
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
    engine_bc
        .set_cell_formula("Sheet1", "A1", "=RATE*2")
        .unwrap();
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
    engine_ast
        .set_cell_formula("Sheet1", "A1", "=RATE*2")
        .unwrap();
    engine_ast.recalculate_single_threaded();

    assert_eq!(
        engine_bc.get_cell_value("Sheet1", "A1"),
        engine_ast.get_cell_value("Sheet1", "A1")
    );
}

#[test]
fn bytecode_backend_inlines_constant_defined_names_inside_let() {
    // LET should still be bytecode-eligible when binding values reference constant defined names.
    // This requires inlining the constant into the lowered bytecode AST, while still respecting
    // lexical scoping rules for LET locals.
    let mut engine = Engine::new();
    engine
        .define_name(
            "RATE",
            NameScope::Workbook,
            NameDefinition::Constant(Value::Number(0.1)),
        )
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A1", "=LET(x, RATE, x*2)")
        .unwrap();

    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();
    let via_bytecode = engine.get_cell_value("Sheet1", "A1");
    assert_eq!(via_bytecode, Value::Number(0.2));

    // Compare against the AST backend (which can resolve defined names at runtime).
    engine.set_bytecode_enabled(false);
    engine.recalculate_single_threaded();
    let via_ast = engine.get_cell_value("Sheet1", "A1");

    assert_eq!(via_bytecode, via_ast);
}

#[test]
fn bytecode_backend_inlines_constant_defined_names_in_let_body() {
    // Constant defined names referenced in the LET body should still be inlined so the formula
    // stays bytecode-eligible.
    let mut engine = Engine::new();
    engine
        .define_name(
            "RATE",
            NameScope::Workbook,
            NameDefinition::Constant(Value::Number(0.1)),
        )
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A1", "=LET(x, 1, RATE*2+x)")
        .unwrap();

    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();
    let via_bytecode = engine.get_cell_value("Sheet1", "A1");
    assert_eq!(via_bytecode, Value::Number(1.2));

    engine.set_bytecode_enabled(false);
    engine.recalculate_single_threaded();
    let via_ast = engine.get_cell_value("Sheet1", "A1");

    assert_eq!(via_bytecode, via_ast);
}

#[test]
fn bytecode_backend_let_binding_value_can_reference_defined_name_it_shadows() {
    // LET bindings should not shadow themselves while their value expression is being evaluated.
    //
    // If the binding name collides with a constant defined name, the RHS should resolve to the
    // defined name (not the local), and then the local should shadow the defined name in the body.
    let mut engine = Engine::new();
    engine
        .define_name(
            "X",
            NameScope::Workbook,
            NameDefinition::Constant(Value::Number(10.0)),
        )
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A1", "=LET(x, x+1, x*2)")
        .unwrap();

    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();
    let via_bytecode = engine.get_cell_value("Sheet1", "A1");
    assert_eq!(via_bytecode, Value::Number(22.0));

    engine.set_bytecode_enabled(false);
    engine.recalculate_single_threaded();
    let via_ast = engine.get_cell_value("Sheet1", "A1");

    assert_eq!(via_bytecode, via_ast);
}

#[test]
fn bytecode_backend_sheet_qualified_defined_name_constant_bypasses_let_locals() {
    // LET locals should only shadow unqualified identifiers. A sheet-qualified defined name should
    // bypass the LET local scope and resolve as a defined name.
    let mut engine = Engine::new();
    engine
        .define_name(
            "RATE",
            NameScope::Workbook,
            NameDefinition::Constant(Value::Number(0.1)),
        )
        .unwrap();
    engine
        .define_name(
            "RATE",
            NameScope::Sheet("Sheet1"),
            NameDefinition::Constant(Value::Number(0.2)),
        )
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A1", "=LET(RATE, 1, Sheet1!RATE*2+RATE)")
        .unwrap();

    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();
    let via_bytecode = engine.get_cell_value("Sheet1", "A1");
    assert_eq!(via_bytecode, Value::Number(1.4));

    engine.set_bytecode_enabled(false);
    engine.recalculate_single_threaded();
    let via_ast = engine.get_cell_value("Sheet1", "A1");

    assert_eq!(via_bytecode, via_ast);
}

#[test]
fn bytecode_backend_sheet_qualified_defined_name_constant_on_other_sheet_bypasses_let_locals() {
    // Even when the name is sheet-qualified to a *different* worksheet, the sheet-qualified
    // identifier should bypass LET locals and still resolve as a defined name.
    //
    // For constant defined names, this should inline to a literal value so the formula remains
    // bytecode-eligible (since the bytecode lowering layer does not support sheet-qualified name
    // references directly).
    let mut engine = Engine::new();
    engine
        .define_name(
            "RATE",
            NameScope::Workbook,
            NameDefinition::Constant(Value::Number(0.1)),
        )
        .unwrap();
    engine
        .define_name(
            "RATE",
            NameScope::Sheet("Sheet2"),
            NameDefinition::Constant(Value::Number(0.2)),
        )
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A1", "=LET(RATE, 1, Sheet2!RATE*2+RATE)")
        .unwrap();

    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();
    let via_bytecode = engine.get_cell_value("Sheet1", "A1");
    assert_eq!(via_bytecode, Value::Number(1.4));

    engine.set_bytecode_enabled(false);
    engine.recalculate_single_threaded();
    let via_ast = engine.get_cell_value("Sheet1", "A1");

    assert_eq!(via_bytecode, via_ast);
}

#[test]
fn bytecode_backend_sheet_qualified_defined_name_falls_back_to_workbook_constant() {
    // `Sheet2!RATE` should resolve to the workbook-scoped defined name when no sheet-scoped name
    // exists for that sheet. This should still bypass LET locals and be bytecode-eligible via
    // constant inlining.
    let mut engine = Engine::new();
    engine.ensure_sheet("Sheet2");
    engine
        .define_name(
            "RATE",
            NameScope::Workbook,
            NameDefinition::Constant(Value::Number(0.1)),
        )
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A1", "=LET(RATE, 1, Sheet2!RATE*2+RATE)")
        .unwrap();

    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();
    let via_bytecode = engine.get_cell_value("Sheet1", "A1");
    assert_eq!(via_bytecode, Value::Number(1.2));

    engine.set_bytecode_enabled(false);
    engine.recalculate_single_threaded();
    let via_ast = engine.get_cell_value("Sheet1", "A1");

    assert_eq!(via_bytecode, via_ast);
}

#[test]
fn bytecode_backend_sheet_qualified_defined_name_static_ref_from_other_sheet_can_inline() {
    // Even though `Sheet2!RATE` is sheet-qualified, it can still be bytecode-eligible when it
    // ultimately resolves to a workbook defined name that is a static reference on the current
    // sheet. The name must still bypass LET locals.
    let mut engine = Engine::new();
    engine.ensure_sheet("Sheet2");
    engine.set_cell_value("Sheet1", "B1", 100.0).unwrap();
    engine
        .define_name(
            "RATE",
            NameScope::Workbook,
            NameDefinition::Reference("Sheet1!$B$1".to_string()),
        )
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A1", "=LET(RATE, 1, SUM(Sheet2!RATE)+RATE)")
        .unwrap();

    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();
    let via_bytecode = engine.get_cell_value("Sheet1", "A1");
    assert_eq!(via_bytecode, Value::Number(101.0));

    engine.set_bytecode_enabled(false);
    engine.recalculate_single_threaded();
    let via_ast = engine.get_cell_value("Sheet1", "A1");

    assert_eq!(via_bytecode, via_ast);
}

#[test]
fn bytecode_backend_sheet_qualified_defined_name_static_ref_can_inline_without_let() {
    // When `Sheet2!RATE` resolves to a workbook defined name that is a static reference on the
    // *current* sheet, the reference can be inlined and compiled to bytecode even without LET.
    //
    // This is a regression test ensuring static defined names are inlined before bytecode
    // compilation, so sheet-qualified uses like `Sheet2!RATE` don't force an AST fallback.
    let mut engine = Engine::new();
    engine.ensure_sheet("Sheet2");
    engine.set_cell_value("Sheet1", "B1", 100.0).unwrap();
    engine
        .define_name(
            "RATE",
            NameScope::Workbook,
            NameDefinition::Reference("Sheet1!$B$1".to_string()),
        )
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A1", "=Sheet2!RATE+1")
        .unwrap();

    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();
    let via_bytecode = engine.get_cell_value("Sheet1", "A1");
    assert_eq!(via_bytecode, Value::Number(101.0));

    engine.set_bytecode_enabled(false);
    engine.recalculate_single_threaded();
    let via_ast = engine.get_cell_value("Sheet1", "A1");

    assert_eq!(via_bytecode, via_ast);
}

#[test]
fn bytecode_backend_sheet_qualified_defined_name_static_ref_bypasses_let_locals() {
    // Same as the constant case, but for reference defined names that are inlined to static
    // cell/range references for bytecode eligibility.
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "B1", 100.0).unwrap();
    engine
        .define_name(
            "RATE",
            NameScope::Workbook,
            NameDefinition::Reference("Sheet1!$B$1".to_string()),
        )
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A1", "=LET(RATE, 1, SUM(Sheet1!RATE)+RATE)")
        .unwrap();

    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();
    let via_bytecode = engine.get_cell_value("Sheet1", "A1");
    assert_eq!(via_bytecode, Value::Number(101.0));

    engine.set_bytecode_enabled(false);
    engine.recalculate_single_threaded();
    let via_ast = engine.get_cell_value("Sheet1", "A1");

    assert_eq!(via_bytecode, via_ast);
}

#[test]
fn bytecode_backend_inlines_constant_defined_names_case_insensitive_and_recompiles() {
    // Exercise case-insensitive name matching and ensure bytecode programs are recompiled when
    // a constant defined name changes (since the constant is inlined into the bytecode AST).
    let mut engine_bc = Engine::new();
    engine_bc
        .define_name(
            "X",
            NameScope::Workbook,
            NameDefinition::Constant(Value::Number(2.0)),
        )
        .unwrap();
    engine_bc.set_cell_formula("Sheet1", "A1", "=x+1").unwrap();
    assert_eq!(engine_bc.bytecode_program_count(), 1);
    engine_bc.recalculate_single_threaded();
    assert_eq!(engine_bc.get_cell_value("Sheet1", "A1"), Value::Number(3.0));

    let mut engine_ast = Engine::new();
    engine_ast.set_bytecode_enabled(false);
    engine_ast
        .define_name(
            "X",
            NameScope::Workbook,
            NameDefinition::Constant(Value::Number(2.0)),
        )
        .unwrap();
    engine_ast.set_cell_formula("Sheet1", "A1", "=x+1").unwrap();
    engine_ast.recalculate_single_threaded();
    assert_eq!(
        engine_bc.get_cell_value("Sheet1", "A1"),
        engine_ast.get_cell_value("Sheet1", "A1")
    );

    engine_bc
        .define_name(
            "X",
            NameScope::Workbook,
            NameDefinition::Constant(Value::Number(10.0)),
        )
        .unwrap();
    assert!(engine_bc.is_dirty("Sheet1", "A1"));
    engine_bc.recalculate_single_threaded();
    assert_eq!(
        engine_bc.get_cell_value("Sheet1", "A1"),
        Value::Number(11.0)
    );
    assert!(
        engine_bc.bytecode_program_count() >= 2,
        "expected name change to trigger bytecode recompilation"
    );

    engine_ast
        .define_name(
            "X",
            NameScope::Workbook,
            NameDefinition::Constant(Value::Number(10.0)),
        )
        .unwrap();
    engine_ast.recalculate_single_threaded();
    assert_eq!(
        engine_bc.get_cell_value("Sheet1", "A1"),
        engine_ast.get_cell_value("Sheet1", "A1")
    );
}

#[test]
fn bytecode_backend_inlines_sheet_scoped_constant_defined_names() {
    let mut engine = Engine::new();
    engine
        .define_name(
            "X",
            NameScope::Workbook,
            NameDefinition::Constant(Value::Number(1.0)),
        )
        .unwrap();
    engine
        .define_name(
            "X",
            NameScope::Sheet("Sheet1"),
            NameDefinition::Constant(Value::Number(2.0)),
        )
        .unwrap();
    engine
        .define_name(
            "X",
            NameScope::Sheet("Sheet2"),
            NameDefinition::Constant(Value::Number(3.0)),
        )
        .unwrap();

    engine.set_cell_formula("Sheet1", "A1", "=x+1").unwrap();
    engine.set_cell_formula("Sheet2", "A1", "=X+1").unwrap();

    // The sheet-scoped names should be resolved and inlined to different literals, producing
    // distinct bytecode programs.
    assert_eq!(engine.bytecode_program_count(), 2);

    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(3.0));
    assert_eq!(engine.get_cell_value("Sheet2", "A1"), Value::Number(4.0));
}

#[test]
fn bytecode_backend_inlines_constant_defined_names_under_percent_postfix() {
    let mut engine = Engine::new();
    engine
        .define_name(
            "X",
            NameScope::Workbook,
            NameDefinition::Constant(Value::Number(2.0)),
        )
        .unwrap();
    engine.set_cell_formula("Sheet1", "A1", "=X%").unwrap();

    // Ensure the constant is inlined so the bytecode backend can compile the percent postfix.
    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();
    let via_bytecode = engine.get_cell_value("Sheet1", "A1");
    assert_eq!(via_bytecode, Value::Number(0.02));

    engine.set_bytecode_enabled(false);
    engine.recalculate_single_threaded();
    let via_ast = engine.get_cell_value("Sheet1", "A1");
    assert_eq!(via_ast, via_bytecode);
}

#[test]
fn bytecode_backend_inlines_constant_defined_names_inside_let_bindings() {
    let mut engine = Engine::new();
    engine
        .define_name(
            "X",
            NameScope::Workbook,
            NameDefinition::Constant(Value::Number(2.0)),
        )
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A1", "=LET(y, X, y+1)")
        .unwrap();

    // The defined name should be inlined so the LET formula can compile to bytecode.
    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();
    let via_bytecode = engine.get_cell_value("Sheet1", "A1");
    assert_eq!(via_bytecode, Value::Number(3.0));

    engine.set_bytecode_enabled(false);
    engine.recalculate_single_threaded();
    let via_ast = engine.get_cell_value("Sheet1", "A1");
    assert_eq!(via_ast, via_bytecode);
}

#[test]
fn bytecode_backend_inlines_constant_defined_names_inside_array_literals() {
    let mut engine = Engine::new();
    engine
        .define_name(
            "X",
            NameScope::Workbook,
            NameDefinition::Constant(Value::Number(2.0)),
        )
        .unwrap();
    engine.set_cell_formula("Sheet1", "A1", "=SUM({1,X})").unwrap();

    // Ensure the constant is inlined so numeric-only array-literal lowering can proceed.
    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();
    let via_bytecode = engine.get_cell_value("Sheet1", "A1");
    assert_eq!(via_bytecode, Value::Number(3.0));

    engine.set_bytecode_enabled(false);
    engine.recalculate_single_threaded();
    let via_ast = engine.get_cell_value("Sheet1", "A1");
    assert_eq!(via_ast, via_bytecode);
}

#[test]
fn bytecode_backend_inlines_constant_defined_name_error_values() {
    let mut engine = Engine::new();
    engine
        .define_name(
            "MyNa",
            NameScope::Workbook,
            NameDefinition::Constant(Value::Error(ErrorKind::NA)),
        )
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A1", "=IFNA(MyNa, 7)")
        .unwrap();

    // Ensure the constant error value was inlined and compiled to bytecode.
    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();
    let via_bytecode = engine.get_cell_value("Sheet1", "A1");

    engine.set_bytecode_enabled(false);
    engine.recalculate_single_threaded();
    let via_ast = engine.get_cell_value("Sheet1", "A1");

    assert_eq!(via_bytecode, Value::Number(7.0));
    assert_eq!(via_bytecode, via_ast);
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
        ("=#NULL!", Value::Error(ErrorKind::Null)),
        ("=#DIV/0!", Value::Error(ErrorKind::Div0)),
        ("=#VALUE!", Value::Error(ErrorKind::Value)),
        ("=#REF!", Value::Error(ErrorKind::Ref)),
        ("=#NAME?", Value::Error(ErrorKind::Name)),
        ("=#NUM!", Value::Error(ErrorKind::Num)),
        ("=#N/A", Value::Error(ErrorKind::NA)),
        ("=#GETTING_DATA", Value::Error(ErrorKind::GettingData)),
        ("=#SPILL!", Value::Error(ErrorKind::Spill)),
        ("=#CALC!", Value::Error(ErrorKind::Calc)),
        ("=#FIELD!", Value::Error(ErrorKind::Field)),
        ("=#CONNECT!", Value::Error(ErrorKind::Connect)),
        ("=#BLOCKED!", Value::Error(ErrorKind::Blocked)),
        ("=#UNKNOWN!", Value::Error(ErrorKind::Unknown)),
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
fn bytecode_backend_propagates_error_literals_through_ops_and_functions() {
    let mut engine = Engine::new();

    let cases = [
        ("A1", "=1+#DIV/0!", Value::Error(ErrorKind::Div0)),
        ("A2", "=#DIV/0!+1", Value::Error(ErrorKind::Div0)),
        ("A3", "=#N/A=0", Value::Error(ErrorKind::NA)),
        ("A4", "=IFERROR(1+#DIV/0!,7)", Value::Number(7.0)),
        ("A5", "=SUM(1,#DIV/0!,2)", Value::Error(ErrorKind::Div0)),
        (
            "A6",
            r#"=CONCAT("x",#DIV/0!)"#,
            Value::Error(ErrorKind::Div0),
        ),
        ("A7", "=NOT(#DIV/0!)", Value::Error(ErrorKind::Div0)),
        ("A8", "=ABS(#DIV/0!)", Value::Error(ErrorKind::Div0)),
        ("A9", "=-#DIV/0!", Value::Error(ErrorKind::Div0)),
    ];

    for (cell, formula, _) in &cases {
        engine.set_cell_formula("Sheet1", cell, formula).unwrap();
    }

    let stats = engine.bytecode_compile_stats();
    assert_eq!(
        stats.fallback,
        0,
        "expected all formulas to compile to bytecode (report={:?})",
        engine.bytecode_compile_report(100)
    );
    assert_eq!(stats.compiled, cases.len());

    engine.recalculate_single_threaded();

    for (cell, formula, expected) in &cases {
        assert_eq!(
            engine.get_cell_value("Sheet1", cell),
            expected.clone(),
            "mismatched value for {cell}: {formula}"
        );
        assert_engine_matches_ast(&engine, formula, cell);
    }
}

#[test]
fn bytecode_backend_handles_extended_error_literals_inside_expressions() {
    let mut engine = Engine::new();

    let cases = [
        ("#GETTING_DATA", ErrorKind::GettingData, 8.0),
        ("#FIELD!", ErrorKind::Field, 11.0),
        ("#CONNECT!", ErrorKind::Connect, 12.0),
        ("#BLOCKED!", ErrorKind::Blocked, 13.0),
        ("#UNKNOWN!", ErrorKind::Unknown, 14.0),
    ];

    for (idx, (lit, _kind, _code)) in cases.iter().enumerate() {
        let row = idx + 1;
        engine
            .set_cell_formula("Sheet1", &format!("A{row}"), &format!("={lit}+1"))
            .unwrap();
        engine
            .set_cell_formula("Sheet1", &format!("B{row}"), &format!("=IFERROR({lit},7)"))
            .unwrap();
        engine
            .set_cell_formula("Sheet1", &format!("C{row}"), &format!("=ERROR.TYPE({lit})"))
            .unwrap();
        engine
            .set_cell_formula("Sheet1", &format!("D{row}"), &format!("=ISERROR({lit})"))
            .unwrap();
    }

    let stats = engine.bytecode_compile_stats();
    assert_eq!(
        stats.fallback,
        0,
        "expected all formulas to compile to bytecode (report={:?})",
        engine.bytecode_compile_report(100)
    );
    assert_eq!(stats.total_formula_cells, cases.len() * 4);
    assert_eq!(stats.compiled, cases.len() * 4);

    engine.recalculate_single_threaded();

    for (idx, (lit, kind, code)) in cases.iter().enumerate() {
        let row = idx + 1;

        let a_cell = format!("A{row}");
        let a_formula = format!("={lit}+1");
        assert_eq!(
            engine.get_cell_value("Sheet1", &a_cell),
            Value::Error(*kind)
        );
        assert_engine_matches_ast(&engine, &a_formula, &a_cell);

        let b_cell = format!("B{row}");
        let b_formula = format!("=IFERROR({lit},7)");
        assert_eq!(engine.get_cell_value("Sheet1", &b_cell), Value::Number(7.0));
        assert_engine_matches_ast(&engine, &b_formula, &b_cell);

        let c_cell = format!("C{row}");
        let c_formula = format!("=ERROR.TYPE({lit})");
        assert_eq!(
            engine.get_cell_value("Sheet1", &c_cell),
            Value::Number(*code)
        );
        assert_engine_matches_ast(&engine, &c_formula, &c_cell);

        let d_cell = format!("D{row}");
        let d_formula = format!("=ISERROR({lit})");
        assert_eq!(engine.get_cell_value("Sheet1", &d_cell), Value::Bool(true));
        assert_engine_matches_ast(&engine, &d_formula, &d_cell);
    }
}

#[test]
fn bytecode_backend_compiles_criteria_functions_with_error_literal_criteria_args() {
    let mut engine = Engine::new();

    // Criteria range.
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", 3.0).unwrap();

    // Aggregate range.
    engine.set_cell_value("Sheet1", "B1", 10.0).unwrap();
    engine.set_cell_value("Sheet1", "B2", 20.0).unwrap();
    engine.set_cell_value("Sheet1", "B3", 30.0).unwrap();

    let cases = [
        // Criteria errors should propagate (even when written as error literals).
        (
            "C1",
            "=SUMIF(A1:A3,#DIV/0!,B1:B3)",
            Value::Error(ErrorKind::Div0),
        ),
        (
            "C2",
            "=SUMIFS(B1:B3,A1:A3,#DIV/0!)",
            Value::Error(ErrorKind::Div0),
        ),
        (
            "C3",
            "=COUNTIFS(A1:A3,#DIV/0!)",
            Value::Error(ErrorKind::Div0),
        ),
        (
            "C4",
            "=AVERAGEIF(A1:A3,#DIV/0!,B1:B3)",
            Value::Error(ErrorKind::Div0),
        ),
    ];

    for (cell, formula, _) in &cases {
        engine.set_cell_formula("Sheet1", cell, formula).unwrap();
    }

    let stats = engine.bytecode_compile_stats();
    assert_eq!(
        stats.fallback,
        0,
        "expected all criteria formulas to compile to bytecode (report={:?})",
        engine.bytecode_compile_report(100)
    );
    assert_eq!(stats.total_formula_cells, cases.len());
    assert_eq!(stats.compiled, cases.len());

    engine.recalculate_single_threaded();

    for (cell, formula, expected) in &cases {
        assert_eq!(
            engine.get_cell_value("Sheet1", cell),
            expected.clone(),
            "mismatched value for {cell}: {formula}"
        );
        assert_engine_matches_ast(&engine, formula, cell);
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
fn bytecode_backend_matches_ast_for_conditional_aggregates_over_array_literals() {
    let mut engine = Engine::new();

    // Populate some sheet values so we can cover mixed range+array cases too (A1:D1 and A2:D2 are
    // 1x4 horizontal ranges to match `{...}` array literals).
    for (cell, value) in [
        ("A1", 1.0),
        ("B1", 2.0),
        ("C1", 3.0),
        ("D1", 4.0),
        ("A2", 10.0),
        ("B2", 20.0),
        ("C2", 30.0),
        ("D2", 40.0),
    ] {
        engine.set_cell_value("Sheet1", cell, value).unwrap();
    }

    let cases = [
        (
            "E1",
            r#"=SUMIF({1,2,3,4},">2",{10,20,30,40})"#,
            Value::Number(70.0),
        ),
        (
            "E2",
            r#"=SUMIFS({10,20,30,40},{"A","A","B","B"},"A",{1,2,3,4},">1")"#,
            Value::Number(20.0),
        ),
        (
            "E3",
            r#"=COUNTIFS({"A","A","B","B"},"A",{1,2,3,4},">1")"#,
            Value::Number(1.0),
        ),
        (
            "E4",
            r#"=AVERAGEIF({1,2,3,4},">2",{10,20,30,40})"#,
            Value::Number(35.0),
        ),
        (
            "E5",
            r#"=AVERAGEIFS({10,20,30,40},{"A","A","B","B"},"A",{1,2,3,4},">1")"#,
            Value::Number(20.0),
        ),
        (
            "E6",
            r#"=MAXIFS({10,20,30,40},{1,2,3,4},">2")"#,
            Value::Number(40.0),
        ),
        (
            "E7",
            r#"=MINIFS({10,20,30,40},{1,2,3,4},">2")"#,
            Value::Number(30.0),
        ),
        // Mixed range+array cases.
        (
            "E8",
            r#"=SUMIF(A1:D1,">2",{10,20,30,40})"#,
            Value::Number(70.0),
        ),
        ("E9", r#"=SUMIF({1,2,3,4},">2",A2:D2)"#, Value::Number(70.0)),
    ];

    for (cell, formula, _) in &cases {
        engine.set_cell_formula("Sheet1", cell, formula).unwrap();
    }

    assert_eq!(engine.bytecode_program_count(), cases.len());
    engine.recalculate_single_threaded();

    for (cell, formula, expected) in &cases {
        assert_eq!(
            engine.get_cell_value("Sheet1", cell),
            expected.clone(),
            "mismatched value for {cell}: {formula}"
        );
        assert_engine_matches_ast(&engine, formula, cell);
    }
}

#[test]
fn bytecode_backend_conditional_aggregates_error_precedence_is_row_major_for_2d_ranges() {
    let mut engine = Engine::new();

    // Criteria range (2x2) matches all cells.
    for addr in ["A1", "A2", "B1", "B2"] {
        engine.set_cell_value("Sheet1", addr, 1.0).unwrap();
    }

    // Sum/average range (2x2) contains two different errors. Excel (and the AST evaluator) return
    // the first included error in row-major range order: C1, D1, C2, D2.
    engine.set_cell_value("Sheet1", "C1", 10.0).unwrap();
    engine
        .set_cell_value("Sheet1", "D1", Value::Error(ErrorKind::Ref))
        .unwrap();
    engine
        .set_cell_value("Sheet1", "C2", Value::Error(ErrorKind::Div0))
        .unwrap();
    engine.set_cell_value("Sheet1", "D2", 20.0).unwrap();

    engine
        .set_cell_formula("Sheet1", "E1", r#"=SUMIFS(C1:D2,A1:B2,">0")"#)
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "E2", r#"=AVERAGEIFS(C1:D2,A1:B2,">0")"#)
        .unwrap();

    assert_eq!(engine.bytecode_program_count(), 2);
    engine.recalculate_single_threaded();

    assert_eq!(
        engine.get_cell_value("Sheet1", "E1"),
        Value::Error(ErrorKind::Ref)
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "E2"),
        Value::Error(ErrorKind::Ref)
    );

    assert_engine_matches_ast(&engine, r#"=SUMIFS(C1:D2,A1:B2,">0")"#, "E1");
    assert_engine_matches_ast(&engine, r#"=AVERAGEIFS(C1:D2,A1:B2,">0")"#, "E2");
}

#[test]
fn bytecode_backend_basic_aggregates_error_precedence_is_row_major_for_2d_ranges() {
    let mut engine = Engine::new();

    // 2x2 range with two different errors. Excel (and the AST evaluator) return the first error in
    // row-major range order: A1, B1, A2, B2.
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine
        .set_cell_value("Sheet1", "B1", Value::Error(ErrorKind::Div0))
        .unwrap();
    engine
        .set_cell_value("Sheet1", "A2", Value::Error(ErrorKind::Ref))
        .unwrap();
    engine.set_cell_value("Sheet1", "B2", 2.0).unwrap();

    let cases = [
        ("C1", "=SUM(A1:B2)"),
        ("C2", "=AVERAGE(A1:B2)"),
        ("C3", "=MIN(A1:B2)"),
        ("C4", "=MAX(A1:B2)"),
    ];
    for (cell, formula) in cases {
        engine.set_cell_formula("Sheet1", cell, formula).unwrap();
    }

    assert_eq!(engine.bytecode_program_count(), 4);
    engine.recalculate_single_threaded();

    for (cell, formula) in cases {
        assert_eq!(
            engine.get_cell_value("Sheet1", cell),
            Value::Error(ErrorKind::Div0),
            "mismatched value for {cell}: {formula}",
        );
        assert_engine_matches_ast(&engine, formula, cell);
    }
}

#[test]
fn bytecode_backend_compiles_xlfn_prefixed_minmaxifs() {
    let mut engine = Engine::new();

    engine.set_cell_value("Sheet1", "A1", 5.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 3.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", 7.0).unwrap();

    engine.set_cell_value("Sheet1", "B1", "A").unwrap();
    engine.set_cell_value("Sheet1", "B2", "B").unwrap();
    engine.set_cell_value("Sheet1", "B3", "A").unwrap();

    engine
        .set_cell_formula("Sheet1", "C1", r#"=_xlfn.MINIFS(A1:A3,B1:B3,"A")"#)
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "C2", r#"=_xlfn.MAXIFS(A1:A3,B1:B3,"A")"#)
        .unwrap();

    // Ensure the prefixed functions compile to bytecode (no AST fallback).
    assert_eq!(engine.bytecode_program_count(), 2);

    engine.recalculate_single_threaded();

    assert_engine_matches_ast(&engine, r#"=_xlfn.MINIFS(A1:A3,B1:B3,"A")"#, "C1");
    assert_engine_matches_ast(&engine, r#"=_xlfn.MAXIFS(A1:A3,B1:B3,"A")"#, "C2");
}

#[test]
fn bytecode_backend_propagates_error_literals_through_supported_ops() {
    let mut engine = Engine::new();
    engine.set_cell_formula("Sheet1", "A1", "=#N/A+1").unwrap();

    // Ensure we're exercising the bytecode path.
    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();

    assert_eq!(
        engine.get_cell_value("Sheet1", "A1"),
        Value::Error(ErrorKind::NA)
    );
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
        0,
        bytecode::CellCoord::new(0, 0),
        &formula_engine::LocaleConfig::en_us(),
    );
    assert_eq!(v, bytecode::Value::Error(bytecode::ErrorKind::Ref));
}

#[test]
fn bytecode_backend_matches_ast_for_sumif_error_criteria_and_optional_sum_range() {
    let mut engine = Engine::new();

    // Criteria range contains a numeric, an error, and another numeric.
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine
        .set_cell_value("Sheet1", "A2", Value::Error(ErrorKind::Div0))
        .unwrap();
    engine.set_cell_value("Sheet1", "A3", 3.0).unwrap();

    // Sum range.
    engine.set_cell_value("Sheet1", "B1", 10.0).unwrap();
    engine.set_cell_value("Sheet1", "B2", 20.0).unwrap();
    engine.set_cell_value("Sheet1", "B3", 30.0).unwrap();

    // Error criteria should match the error cell without propagating it.
    engine
        .set_cell_formula("Sheet1", "C1", r##"=SUMIF(A1:A3,"#DIV/0!",B1:B3)"##)
        .unwrap();
    // Numeric criteria should ignore the error in the criteria range.
    engine
        .set_cell_formula("Sheet1", "C2", r#"=SUMIF(A1:A3,">1",B1:B3)"#)
        .unwrap();
    // Errors in the criteria argument always propagate.
    engine
        .set_cell_formula("Sheet1", "C3", r#"=SUMIF(A1:A3,A2,B1:B3)"#)
        .unwrap();
    // Trailing blank sum_range behaves like omitting the argument.
    engine
        .set_cell_formula("Sheet1", "C4", r#"=SUMIF(B1:B3,">15",)"#)
        .unwrap();

    // Ensure these compile to bytecode (no AST fallback).
    assert_eq!(engine.bytecode_program_count(), 4);

    engine.recalculate_single_threaded();

    for (formula, cell) in [
        (r##"=SUMIF(A1:A3,"#DIV/0!",B1:B3)"##, "C1"),
        (r#"=SUMIF(A1:A3,">1",B1:B3)"#, "C2"),
        (r#"=SUMIF(A1:A3,A2,B1:B3)"#, "C3"),
        (r#"=SUMIF(B1:B3,">15",)"#, "C4"),
    ] {
        assert_engine_matches_ast(&engine, formula, cell);
    }
}

#[test]
fn bytecode_backend_supports_array_ranges_for_sumif_and_averageif() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", r#"=SUMIF({1,2,3,4},">2")"#)
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A2", r#"=AVERAGEIF({1,2,3,4},">2")"#)
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A3", r#"=SUMIF({1,2,3,4},">2",{10,20,30,40})"#)
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "A4",
            r#"=AVERAGEIF({1,2,3,4},">2",{10,20,30,40})"#,
        )
        .unwrap();

    // Array expression used as the criteria_range.
    engine.set_cell_value("Sheet1", "D1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "D2", -2.0).unwrap();
    engine.set_cell_value("Sheet1", "D3", 3.0).unwrap();
    engine.set_cell_value("Sheet1", "E1", 10.0).unwrap();
    engine.set_cell_value("Sheet1", "E2", 20.0).unwrap();
    engine.set_cell_value("Sheet1", "E3", 30.0).unwrap();
    engine
        .set_cell_formula("Sheet1", "A5", "=SUMIF(D1:D3>0,TRUE,E1:E3)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A6", "=AVERAGEIF(D1:D3>0,TRUE,E1:E3)")
        .unwrap();

    assert_eq!(
        engine.bytecode_program_count(),
        6,
        "expected SUMIF/AVERAGEIF array-range formulas to compile to bytecode"
    );

    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(7.0));
    assert_eq!(engine.get_cell_value("Sheet1", "A2"), Value::Number(3.5));
    assert_eq!(engine.get_cell_value("Sheet1", "A3"), Value::Number(70.0));
    assert_eq!(engine.get_cell_value("Sheet1", "A4"), Value::Number(35.0));
    assert_eq!(engine.get_cell_value("Sheet1", "A5"), Value::Number(40.0));
    assert_eq!(engine.get_cell_value("Sheet1", "A6"), Value::Number(20.0));

    for (formula, cell) in [
        (r#"=SUMIF({1,2,3,4},">2")"#, "A1"),
        (r#"=AVERAGEIF({1,2,3,4},">2")"#, "A2"),
        (r#"=SUMIF({1,2,3,4},">2",{10,20,30,40})"#, "A3"),
        (r#"=AVERAGEIF({1,2,3,4},">2",{10,20,30,40})"#, "A4"),
        ("=SUMIF(D1:D3>0,TRUE,E1:E3)", "A5"),
        ("=AVERAGEIF(D1:D3>0,TRUE,E1:E3)", "A6"),
    ] {
        assert_engine_matches_ast(&engine, formula, cell);
    }
}

#[test]
fn bytecode_backend_matches_ast_for_sumifs_countifs_and_averageifs() {
    let mut engine = Engine::new();

    for (row, (cat, n, v)) in [
        (1, ("a", 1.0, 10.0)),
        (2, ("b", 2.0, 20.0)),
        (3, ("a", 3.0, 30.0)),
        (4, ("b", 4.0, 40.0)),
        (5, ("a", 5.0, 50.0)),
    ] {
        engine
            .set_cell_value("Sheet1", &format!("A{row}"), cat)
            .unwrap();
        engine
            .set_cell_value("Sheet1", &format!("B{row}"), n)
            .unwrap();
        engine
            .set_cell_value("Sheet1", &format!("C{row}"), v)
            .unwrap();
    }

    engine
        .set_cell_formula("Sheet1", "D1", r#"=SUMIFS(C1:C5,A1:A5,"a",B1:B5,">2")"#)
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "D2", r#"=COUNTIFS(A1:A5,"a",B1:B5,">2")"#)
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "D3", r#"=AVERAGEIFS(C1:C5,A1:A5,"a",B1:B5,">2")"#)
        .unwrap();

    assert_eq!(engine.bytecode_program_count(), 3);
    engine.recalculate_single_threaded();

    for (formula, cell) in [
        (r#"=SUMIFS(C1:C5,A1:A5,"a",B1:B5,">2")"#, "D1"),
        (r#"=COUNTIFS(A1:A5,"a",B1:B5,">2")"#, "D2"),
        (r#"=AVERAGEIFS(C1:C5,A1:A5,"a",B1:B5,">2")"#, "D3"),
    ] {
        assert_engine_matches_ast(&engine, formula, cell);
    }
}

#[test]
fn bytecode_backend_supports_array_ranges_for_ifs_family() {
    let mut engine = Engine::new();

    engine
        .set_cell_formula(
            "Sheet1",
            "A1",
            r#"=SUMIFS({10,20,30,40},{"A","A","B","B"},"A",{1,2,3,4},">1")"#,
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "A2",
            r#"=AVERAGEIFS({10,20,30,40},{"A","A","B","B"},"A",{1,2,3,4},">1")"#,
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "A3",
            r#"=COUNTIFS({"A","A","B","B"},"A",{1,2,3,4},">1")"#,
        )
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A4", r#"=MINIFS({10,20,30,40},{1,2,3,4},">2")"#)
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A5", r#"=MAXIFS({10,20,30,40},{1,2,3,4},">2")"#)
        .unwrap();

    // Array expression used as the criteria_range / criteria_range1.
    engine.set_cell_value("Sheet1", "D1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "D2", -2.0).unwrap();
    engine.set_cell_value("Sheet1", "D3", 3.0).unwrap();
    engine.set_cell_value("Sheet1", "E1", 10.0).unwrap();
    engine.set_cell_value("Sheet1", "E2", 20.0).unwrap();
    engine.set_cell_value("Sheet1", "E3", 30.0).unwrap();
    engine
        .set_cell_formula("Sheet1", "B1", "=SUMIFS(E1:E3,D1:D3>0,TRUE)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B2", "=AVERAGEIFS(E1:E3,D1:D3>0,TRUE)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B3", "=COUNTIFS(D1:D3>0,TRUE)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B4", "=COUNTIF(D1:D3>0,TRUE)")
        .unwrap();

    assert_eq!(
        engine.bytecode_program_count(),
        9,
        "expected IFS family array-range formulas to compile to bytecode"
    );

    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(20.0));
    assert_eq!(engine.get_cell_value("Sheet1", "A2"), Value::Number(20.0));
    assert_eq!(engine.get_cell_value("Sheet1", "A3"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "A4"), Value::Number(30.0));
    assert_eq!(engine.get_cell_value("Sheet1", "A5"), Value::Number(40.0));
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(40.0));
    assert_eq!(engine.get_cell_value("Sheet1", "B2"), Value::Number(20.0));
    assert_eq!(engine.get_cell_value("Sheet1", "B3"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "B4"), Value::Number(2.0));

    for (formula, cell) in [
        (
            r#"=SUMIFS({10,20,30,40},{"A","A","B","B"},"A",{1,2,3,4},">1")"#,
            "A1",
        ),
        (
            r#"=AVERAGEIFS({10,20,30,40},{"A","A","B","B"},"A",{1,2,3,4},">1")"#,
            "A2",
        ),
        (r#"=COUNTIFS({"A","A","B","B"},"A",{1,2,3,4},">1")"#, "A3"),
        (r#"=MINIFS({10,20,30,40},{1,2,3,4},">2")"#, "A4"),
        (r#"=MAXIFS({10,20,30,40},{1,2,3,4},">2")"#, "A5"),
        ("=SUMIFS(E1:E3,D1:D3>0,TRUE)", "B1"),
        ("=AVERAGEIFS(E1:E3,D1:D3>0,TRUE)", "B2"),
        ("=COUNTIFS(D1:D3>0,TRUE)", "B3"),
        ("=COUNTIF(D1:D3>0,TRUE)", "B4"),
    ] {
        assert_engine_matches_ast(&engine, formula, cell);
    }
}

#[test]
fn bytecode_backend_matches_ast_for_averageif_minifs_and_maxifs() {
    let mut engine = Engine::new();

    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", 3.0).unwrap();

    // Aggregate ranges include an error that should only propagate when selected.
    engine
        .set_cell_value("Sheet1", "B1", Value::Error(ErrorKind::Div0))
        .unwrap();
    engine.set_cell_value("Sheet1", "B2", 5.0).unwrap();
    engine.set_cell_value("Sheet1", "B3", 1.0).unwrap();

    engine
        .set_cell_formula("Sheet1", "C1", r#"=AVERAGEIF(A1:A3,">1",B1:B3)"#)
        .unwrap();
    engine
        // Error in B1 is excluded by the criteria (A1 is not > 1), so MINIFS should succeed.
        .set_cell_formula("Sheet1", "C2", r#"=MINIFS(B1:B3,A1:A3,">1")"#)
        .unwrap();
    engine
        // MAXIFS includes the same rows and should return 5.
        .set_cell_formula("Sheet1", "C3", r#"=MAXIFS(B1:B3,A1:A3,">1")"#)
        .unwrap();

    assert_eq!(engine.bytecode_program_count(), 3);
    engine.recalculate_single_threaded();

    for (formula, cell) in [
        (r#"=AVERAGEIF(A1:A3,">1",B1:B3)"#, "C1"),
        (r#"=MINIFS(B1:B3,A1:A3,">1")"#, "C2"),
        (r#"=MAXIFS(B1:B3,A1:A3,">1")"#, "C3"),
    ] {
        assert_engine_matches_ast(&engine, formula, cell);
    }
}

#[test]
fn bytecode_backend_sumifs_date_criteria_respects_engine_value_locale() {
    let mut engine = Engine::new();
    engine.set_value_locale(ValueLocaleConfig::de_de());

    let system = engine.date_system();
    let jan_2 = ymd_to_serial(ExcelDate::new(2020, 1, 2), system).unwrap();
    let feb_1 = ymd_to_serial(ExcelDate::new(2020, 2, 1), system).unwrap();

    engine.set_cell_value("Sheet1", "A1", jan_2 as f64).unwrap();
    engine.set_cell_value("Sheet1", "A2", feb_1 as f64).unwrap();
    engine.set_cell_value("Sheet1", "B1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "B2", 2.0).unwrap();

    // In de-DE (DMY), "1/2/2020" is Feb 1, 2020.
    engine
        .set_cell_value("Sheet1", "C1", Value::Text(">=1/2/2020".to_string()))
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "D1", "=SUMIFS(B1:B2, A1:A2, C1)")
        .unwrap();

    assert_eq!(engine.bytecode_program_count(), 1);
    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Sheet1", "D1"), Value::Number(2.0));
    assert_engine_matches_ast(&engine, "=SUMIFS(B1:B2, A1:A2, C1)", "D1");
}

#[test]
fn bytecode_backend_compiles_large_sumifs_to_bytecode() {
    let mut engine = Engine::new();

    for row in 1..=10_000 {
        engine
            .set_cell_value("Sheet1", &format!("A{row}"), row as f64)
            .unwrap();
        engine
            .set_cell_value("Sheet1", &format!("B{row}"), 1.0)
            .unwrap();
    }

    engine
        .set_cell_formula("Sheet1", "C1", r#"=SUMIFS(B1:B10000,A1:A10000,">5000")"#)
        .unwrap();

    // Ensure bytecode eligibility for a larger range.
    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();
    assert_engine_matches_ast(&engine, r#"=SUMIFS(B1:B10000,A1:A10000,">5000")"#, "C1");
}
