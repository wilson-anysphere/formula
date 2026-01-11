use formula_engine::eval::{parse_a1, EvalContext, Evaluator, SheetReference, ValueResolver};
use formula_engine::{Engine, ExternalValueProvider, Value};
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

#[test]
fn bytecode_backend_matches_ast_for_sum_and_countif() {
    let mut engine = Engine::new();

    for row in 1..=1000 {
        engine
            .set_cell_value("Sheet1", &format!("A{row}"), row as f64)
            .unwrap();
    }

    engine
        .set_cell_formula("Sheet1", "B1", "=SUM(A1:A1000)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B2", "=COUNTIF(A1:A1000, \">500\")")
        .unwrap();

    engine.recalculate_single_threaded();

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
        ) -> Option<(
            usize,
            formula_engine::eval::CellAddr,
            formula_engine::eval::CellAddr,
        )> {
            None
        }
    }

    let resolver = EngineResolver { engine: &engine };

    let parsed_sum = formula_engine::eval::Parser::parse("=SUM(A1:A1000)").unwrap();
    let compiled_sum = {
        let mut map = |sref: &SheetReference<String>| match sref {
            SheetReference::Current => SheetReference::Current,
            SheetReference::Sheet(_name) => SheetReference::Sheet(0),
            SheetReference::External(wb) => SheetReference::External(wb.clone()),
        };
        parsed_sum.map_sheets(&mut map)
    };
    let ctx_sum = EvalContext {
        current_sheet: 0,
        current_cell: parse_a1("B1").unwrap(),
    };
    let eval_sum = Evaluator::new(&resolver, ctx_sum).eval_formula(&compiled_sum);
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), eval_sum);

    let parsed_countif =
        formula_engine::eval::Parser::parse("=COUNTIF(A1:A1000, \">500\")").unwrap();
    let compiled_countif = {
        let mut map = |sref: &SheetReference<String>| match sref {
            SheetReference::Current => SheetReference::Current,
            SheetReference::Sheet(_name) => SheetReference::Sheet(0),
            SheetReference::External(wb) => SheetReference::External(wb.clone()),
        };
        parsed_countif.map_sheets(&mut map)
    };
    let ctx_countif = EvalContext {
        current_sheet: 0,
        current_cell: parse_a1("B2").unwrap(),
    };
    let eval_countif = Evaluator::new(&resolver, ctx_countif).eval_formula(&compiled_countif);
    assert_eq!(engine.get_cell_value("Sheet1", "B2"), eval_countif);

    assert!(engine.bytecode_program_count() > 0);
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
    engine.set_cell_formula("Sheet1", "B1", "=SUM(A1:A3)").unwrap();

    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(6.0));
    assert!(engine.bytecode_program_count() > 0);
}
