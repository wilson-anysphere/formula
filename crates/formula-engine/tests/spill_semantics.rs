use std::sync::Arc;

use formula_engine::eval::parse_a1;
use formula_engine::{Engine, ErrorKind, ExternalValueProvider, Value};

#[test]
fn reference_spill_spills_values() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", 3.0).unwrap();
    engine.set_cell_formula("Sheet1", "C1", "=A1:A3").unwrap();
    engine.recalculate_single_threaded();

    let (start, end) = engine.spill_range("Sheet1", "C1").expect("spill range");
    assert_eq!(start, parse_a1("C1").unwrap());
    assert_eq!(end, parse_a1("C3").unwrap());

    assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "C2"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "C3"), Value::Number(3.0));

    let (origin_sheet, origin_addr) = engine.spill_origin("Sheet1", "C2").expect("spill origin");
    assert_eq!(origin_sheet, 0);
    assert_eq!(origin_addr, parse_a1("C1").unwrap());
}

#[test]
fn external_values_block_spills() {
    struct Provider;

    impl ExternalValueProvider for Provider {
        fn get(&self, sheet: &str, addr: formula_engine::eval::CellAddr) -> Option<Value> {
            if sheet == "Sheet1" && addr == parse_a1("C2").unwrap() {
                return Some(Value::Number(99.0));
            }
            None
        }
    }

    let mut engine = Engine::new();
    engine.set_external_value_provider(Some(Arc::new(Provider)));
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", 3.0).unwrap();
    engine.set_cell_formula("Sheet1", "C1", "=A1:A3").unwrap();
    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::Error(ErrorKind::Spill));
    assert_eq!(engine.get_cell_value("Sheet1", "C2"), Value::Number(99.0));
    assert!(engine.spill_range("Sheet1", "C1").is_none());
}

#[test]
fn spill_blocking_produces_spill_error() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", 3.0).unwrap();
    engine.set_cell_formula("Sheet1", "C1", "=A1:A3").unwrap();
    engine.recalculate_single_threaded();

    // Block the middle spill cell with a user value.
    engine.set_cell_value("Sheet1", "C2", 99.0).unwrap();
    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::Error(ErrorKind::Spill));
    assert_eq!(engine.get_cell_value("Sheet1", "C2"), Value::Number(99.0));
    assert_eq!(engine.get_cell_value("Sheet1", "C3"), Value::Blank);
    assert!(engine.spill_range("Sheet1", "C1").is_none());
}

#[test]
fn spill_resolves_after_blocker_cleared() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", 3.0).unwrap();
    engine.set_cell_formula("Sheet1", "C1", "=A1:A3").unwrap();
    engine.recalculate_single_threaded();

    // Block the spill, producing #SPILL! at the origin.
    engine.set_cell_value("Sheet1", "C2", 99.0).unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::Error(ErrorKind::Spill));
    assert!(engine.spill_range("Sheet1", "C1").is_none());

    // Clearing the blocker should allow the origin to spill again.
    engine
        .set_cell_value("Sheet1", "C2", Value::Blank)
        .unwrap();
    engine.recalculate_single_threaded();

    let (start, end) = engine.spill_range("Sheet1", "C1").expect("spill range");
    assert_eq!(start, parse_a1("C1").unwrap());
    assert_eq!(end, parse_a1("C3").unwrap());

    assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "C2"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "C3"), Value::Number(3.0));
}

#[test]
fn spill_resolves_after_overlapping_spill_shrinks() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 3.0).unwrap();

    // This spill range (D1:E3) overlaps the spill range we want at C2:D4 without covering its origin.
    engine
        .set_cell_formula("Sheet1", "D1", "=SEQUENCE(A1,2,1,1)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "C2", "=SEQUENCE(3,2,1,1)")
        .unwrap();
    engine.recalculate_single_threaded();

    // The lower spill is blocked by the upper spill's occupied cells.
    assert_eq!(engine.get_cell_value("Sheet1", "C2"), Value::Error(ErrorKind::Spill));
    assert!(engine.spill_range("Sheet1", "C2").is_none());

    // Shrink the upper spill so the overlap is cleared; the blocked spill should now succeed.
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.recalculate_single_threaded();

    let (start, end) = engine.spill_range("Sheet1", "C2").expect("spill range");
    assert_eq!(start, parse_a1("C2").unwrap());
    assert_eq!(end, parse_a1("D4").unwrap());

    assert_eq!(engine.get_cell_value("Sheet1", "C2"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "D2"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "C3"), Value::Number(3.0));
    assert_eq!(engine.get_cell_value("Sheet1", "D3"), Value::Number(4.0));
    assert_eq!(engine.get_cell_value("Sheet1", "C4"), Value::Number(5.0));
    assert_eq!(engine.get_cell_value("Sheet1", "D4"), Value::Number(6.0));
}

#[test]
fn transpose_spills_down() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "B1", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "C1", 3.0).unwrap();
    engine
        .set_cell_formula("Sheet1", "E1", "=TRANSPOSE(A1:C1)")
        .unwrap();
    engine.recalculate_single_threaded();

    let (start, end) = engine.spill_range("Sheet1", "E1").expect("spill range");
    assert_eq!(start, parse_a1("E1").unwrap());
    assert_eq!(end, parse_a1("E3").unwrap());

    assert_eq!(engine.get_cell_value("Sheet1", "E1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "E2"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "E3"), Value::Number(3.0));
}

#[test]
fn sequence_spills_matrix() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "C1", "=SEQUENCE(2,2,1,1)")
        .unwrap();
    engine.recalculate_single_threaded();

    let (start, end) = engine.spill_range("Sheet1", "C1").expect("spill range");
    assert_eq!(start, parse_a1("C1").unwrap());
    assert_eq!(end, parse_a1("D2").unwrap());

    assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "D1"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "C2"), Value::Number(3.0));
    assert_eq!(engine.get_cell_value("Sheet1", "D2"), Value::Number(4.0));
}

#[test]
fn dependents_of_spill_cells_recalculate() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", 3.0).unwrap();
    engine.set_cell_formula("Sheet1", "C1", "=A1:A3").unwrap();
    engine.set_cell_formula("Sheet1", "D1", "=C2*10").unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "D1"), Value::Number(20.0));

    engine.set_cell_value("Sheet1", "A2", 5.0).unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "C2"), Value::Number(5.0));
    assert_eq!(engine.get_cell_value("Sheet1", "D1"), Value::Number(50.0));
}
