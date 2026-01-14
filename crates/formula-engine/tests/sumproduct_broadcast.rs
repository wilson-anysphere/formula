use formula_engine::eval::{
    parse_a1, EvalContext, Evaluator, Parser, SheetReference, ValueResolver,
};
use formula_engine::value::NumberLocale;
use formula_engine::{Engine, ErrorKind, Value};

fn set_value(engine: &mut Engine, addr: &str, value: impl Into<Value>) {
    engine
        .set_cell_value("Sheet1", addr, value)
        .expect("set cell value");
}

struct EngineResolver<'a> {
    engine: &'a Engine,
}

impl ValueResolver for EngineResolver<'_> {
    fn sheet_exists(&self, sheet_id: usize) -> bool {
        sheet_id == 0
    }

    fn sheet_dimensions(&self, sheet_id: usize) -> (u32, u32) {
        match sheet_id {
            // Keep the test grid small so whole-column references like `A:A` don't require
            // iterating Excel's full 1,048,576-row default.
            0 => (100, 26),
            _ => (0, 0),
        }
    }

    fn get_cell_value(&self, sheet_id: usize, addr: formula_engine::eval::CellAddr) -> Value {
        let sheet = match sheet_id {
            0 => "Sheet1",
            _ => return Value::Blank,
        };
        self.engine.get_cell_value(sheet, &addr.to_a1())
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
}

fn eval_via_ast(engine: &Engine, formula: &str, current_cell: &str) -> Value {
    let resolver = EngineResolver { engine };
    let mut recalc_ctx = formula_engine::eval::RecalcContext::new(0);
    let separators = engine.value_locale().separators;
    recalc_ctx.number_locale =
        NumberLocale::new(separators.decimal_sep, Some(separators.thousands_sep));

    let parsed = Parser::parse(formula).unwrap();
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

#[test]
fn sumproduct_broadcasts_single_cell_references() {
    let mut engine = Engine::new();
    set_value(&mut engine, "A1", 2.0);
    set_value(&mut engine, "B1", 1.0);
    set_value(&mut engine, "B2", 2.0);
    set_value(&mut engine, "B3", 3.0);

    // A1 should be broadcast to match B1:B3 length.
    assert_eq!(
        eval_via_ast(&engine, "=SUMPRODUCT(A1,B1:B3)", "Z1"),
        Value::Number(12.0)
    );
    // Broadcast should work regardless of whether the single-cell reference is the first or
    // second argument.
    assert_eq!(
        eval_via_ast(&engine, "=SUMPRODUCT(B1:B3,A1)", "Z1"),
        Value::Number(12.0)
    );
}

#[test]
fn sumproduct_preserves_error_precedence_for_broadcast_references() {
    let mut engine = Engine::new();
    set_value(&mut engine, "A1", Value::Error(ErrorKind::Value));
    set_value(&mut engine, "B1", Value::Error(ErrorKind::Div0));
    set_value(&mut engine, "B2", 2.0);
    set_value(&mut engine, "B3", 3.0);

    // With broadcast, error precedence should follow per-index coercion order: coerce the first
    // argument before the second for each element.
    //
    // For idx=0 here, B1 is coerced before A1, so B1's error should win.
    assert_eq!(
        eval_via_ast(&engine, "=SUMPRODUCT(B1:B3,A1)", "Z1"),
        Value::Error(ErrorKind::Div0)
    );

    // When the single-cell reference is the first argument, its error should win immediately.
    assert_eq!(
        eval_via_ast(&engine, "=SUMPRODUCT(A1,B1:B3)", "Z1"),
        Value::Error(ErrorKind::Value)
    );
}

#[test]
fn sumproduct_accepts_whole_column_references() {
    let mut engine = Engine::new();
    set_value(&mut engine, "A1", 2.0);
    set_value(&mut engine, "A2", 3.0);
    set_value(&mut engine, "A3", 4.0);
    set_value(&mut engine, "B1", 1.0);
    set_value(&mut engine, "B2", 2.0);
    set_value(&mut engine, "B3", 3.0);

    // `A:A` / `B:B` use sheet-end sentinels at parse time. The evaluator resolves those against
    // `ValueResolver::sheet_dimensions`, so this test also covers that integration.
    assert_eq!(
        eval_via_ast(&engine, "=SUMPRODUCT(A:A,B:B)", "Z1"),
        Value::Number(20.0)
    );
}
