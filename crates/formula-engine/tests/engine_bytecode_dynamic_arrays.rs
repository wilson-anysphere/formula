use formula_engine::eval::parse_a1;
use formula_engine::functions::{Reference, SheetId};
use formula_engine::value::Array;
use formula_engine::{Engine, ErrorKind, Value};
use formula_model::CellRef;

#[test]
fn bytecode_backend_spills_range_reference() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", 3.0).unwrap();

    engine.set_cell_formula("Sheet1", "C1", "=A1:A3").unwrap();

    // Ensure we're exercising the bytecode backend.
    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();

    let (start, end) = engine.spill_range("Sheet1", "C1").expect("spill range");
    assert_eq!(start, parse_a1("C1").unwrap());
    assert_eq!(end, parse_a1("C3").unwrap());

    assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "C2"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "C3"), Value::Number(3.0));
}

#[test]
fn bytecode_backend_spills_range_reference_with_mixed_types() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", "a").unwrap();
    engine.set_cell_value("Sheet1", "A2", true).unwrap();
    engine
        .set_cell_value("Sheet1", "A3", Value::Error(ErrorKind::Div0))
        .unwrap();

    engine.set_cell_formula("Sheet1", "C1", "=A1:A3").unwrap();

    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();

    let (start, end) = engine.spill_range("Sheet1", "C1").expect("spill range");
    assert_eq!(start, parse_a1("C1").unwrap());
    assert_eq!(end, parse_a1("C3").unwrap());

    assert_eq!(
        engine.get_cell_value("Sheet1", "C1"),
        Value::Text("a".to_string())
    );
    assert_eq!(engine.get_cell_value("Sheet1", "C2"), Value::Bool(true));
    assert_eq!(
        engine.get_cell_value("Sheet1", "C3"),
        Value::Error(ErrorKind::Div0)
    );
}

#[test]
fn bytecode_backend_spills_range_plus_scalar() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", 3.0).unwrap();

    engine.set_cell_formula("Sheet1", "C1", "=A1:A3+1").unwrap();
    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();

    let (start, end) = engine.spill_range("Sheet1", "C1").expect("spill range");
    assert_eq!(start, parse_a1("C1").unwrap());
    assert_eq!(end, parse_a1("C3").unwrap());

    assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "C2"), Value::Number(3.0));
    assert_eq!(engine.get_cell_value("Sheet1", "C3"), Value::Number(4.0));
}

#[test]
fn bytecode_backend_broadcasts_row_and_column_ranges() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", 3.0).unwrap();

    engine.set_cell_value("Sheet1", "B1", 10.0).unwrap();
    engine.set_cell_value("Sheet1", "C1", 20.0).unwrap();
    engine.set_cell_value("Sheet1", "D1", 30.0).unwrap();

    engine
        .set_cell_formula("Sheet1", "F1", "=A1:A3+B1:D1")
        .unwrap();
    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();

    let (start, end) = engine.spill_range("Sheet1", "F1").expect("spill range");
    assert_eq!(start, parse_a1("F1").unwrap());
    assert_eq!(end, parse_a1("H3").unwrap());

    assert_eq!(engine.get_cell_value("Sheet1", "F1"), Value::Number(11.0));
    assert_eq!(engine.get_cell_value("Sheet1", "G1"), Value::Number(21.0));
    assert_eq!(engine.get_cell_value("Sheet1", "H1"), Value::Number(31.0));
    assert_eq!(engine.get_cell_value("Sheet1", "F2"), Value::Number(12.0));
    assert_eq!(engine.get_cell_value("Sheet1", "G2"), Value::Number(22.0));
    assert_eq!(engine.get_cell_value("Sheet1", "H2"), Value::Number(32.0));
    assert_eq!(engine.get_cell_value("Sheet1", "F3"), Value::Number(13.0));
    assert_eq!(engine.get_cell_value("Sheet1", "G3"), Value::Number(23.0));
    assert_eq!(engine.get_cell_value("Sheet1", "H3"), Value::Number(33.0));
}

#[test]
fn bytecode_backend_spills_comparison_results() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", -2.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", 0.0).unwrap();

    engine.set_cell_formula("Sheet1", "B1", "=A1:A3>0").unwrap();
    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();

    let (start, end) = engine.spill_range("Sheet1", "B1").expect("spill range");
    assert_eq!(start, parse_a1("B1").unwrap());
    assert_eq!(end, parse_a1("B3").unwrap());

    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Bool(true));
    assert_eq!(engine.get_cell_value("Sheet1", "B2"), Value::Bool(false));
    assert_eq!(engine.get_cell_value("Sheet1", "B3"), Value::Bool(false));
}

#[test]
fn bytecode_backend_spills_array_value_loaded_from_cell() {
    let mut engine = Engine::new();
    engine
        .set_cell_value(
            "Sheet1",
            "A1",
            Value::Array(Array::new(
                1,
                2,
                vec![Value::Number(1.0), Value::Number(2.0)],
            )),
        )
        .unwrap();

    engine.set_cell_formula("Sheet1", "C1", "=A1").unwrap();
    assert_eq!(
        engine.bytecode_program_count(),
        1,
        "expected simple cell reference to compile to bytecode"
    );

    engine.recalculate_single_threaded();

    let (start, end) = engine.spill_range("Sheet1", "C1").expect("spill range");
    assert_eq!(start, parse_a1("C1").unwrap());
    assert_eq!(end, parse_a1("D1").unwrap());

    assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "D1"), Value::Number(2.0));
}

#[test]
fn bytecode_backend_degrades_illegal_array_elements_to_scalar_errors() {
    let mut engine = Engine::new();

    let nested = Value::Array(Array::new(1, 1, vec![Value::Number(99.0)]));
    let ref_a1 = Reference {
        sheet_id: SheetId::Local(0),
        start: parse_a1("A1").unwrap(),
        end: parse_a1("A1").unwrap(),
    };
    let ref_union = Value::ReferenceUnion(vec![
        ref_a1.clone(),
        Reference {
            sheet_id: SheetId::Local(0),
            start: parse_a1("B1").unwrap(),
            end: parse_a1("B1").unwrap(),
        },
    ]);

    engine
        .set_cell_value(
            "Sheet1",
            "A1",
            Value::Array(Array::new(
                1,
                4,
                vec![
                    nested,
                    Value::Reference(ref_a1),
                    ref_union,
                    Value::Spill {
                        origin: CellRef::new(0, 0),
                    },
                ],
            )),
        )
        .unwrap();

    // Returning an array value should spill it; illegal element types should be degraded to scalar
    // errors instead of panicking or producing nested arrays/ranges.
    engine.set_cell_formula("Sheet1", "C1", "=A1").unwrap();
    assert_eq!(engine.bytecode_program_count(), 1);
    engine.recalculate_single_threaded();

    let (start, end) = engine.spill_range("Sheet1", "C1").expect("spill range");
    assert_eq!(start, parse_a1("C1").unwrap());
    assert_eq!(end, parse_a1("F1").unwrap());

    // nested arrays / references / unions are degraded to #VALUE!
    assert_eq!(
        engine.get_cell_value("Sheet1", "C1"),
        Value::Error(ErrorKind::Value)
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "D1"),
        Value::Error(ErrorKind::Value)
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "E1"),
        Value::Error(ErrorKind::Value)
    );
    // spill markers are degraded to #SPILL!
    assert_eq!(
        engine.get_cell_value("Sheet1", "F1"),
        Value::Error(ErrorKind::Spill)
    );
}

#[test]
fn bytecode_backend_spills_array_literal() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "C1", "={1,2;3,4}")
        .unwrap();

    // Ensure we're exercising the bytecode backend.
    assert_eq!(engine.bytecode_program_count(), 1);

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
fn bytecode_backend_spills_array_literal_plus_scalar() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "C1", "={1,2;3,4}+1")
        .unwrap();

    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();

    let (start, end) = engine.spill_range("Sheet1", "C1").expect("spill range");
    assert_eq!(start, parse_a1("C1").unwrap());
    assert_eq!(end, parse_a1("D2").unwrap());

    assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "D1"), Value::Number(3.0));
    assert_eq!(engine.get_cell_value("Sheet1", "C2"), Value::Number(4.0));
    assert_eq!(engine.get_cell_value("Sheet1", "D2"), Value::Number(5.0));
}

#[test]
fn bytecode_backend_spills_let_bound_array_literal() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "C1", "=LET(a,{1,2;3,4},a+1)")
        .unwrap();

    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();

    let (start, end) = engine.spill_range("Sheet1", "C1").expect("spill range");
    assert_eq!(start, parse_a1("C1").unwrap());
    assert_eq!(end, parse_a1("D2").unwrap());

    assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "D1"), Value::Number(3.0));
    assert_eq!(engine.get_cell_value("Sheet1", "C2"), Value::Number(4.0));
    assert_eq!(engine.get_cell_value("Sheet1", "D2"), Value::Number(5.0));
}

#[test]
fn bytecode_backend_xlookup_spills_row_result() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "B1", "a").unwrap();
    engine.set_cell_value("Sheet1", "B2", "b").unwrap();
    engine.set_cell_value("Sheet1", "B3", "c").unwrap();
    engine.set_cell_value("Sheet1", "B4", "d").unwrap();

    engine.set_cell_value("Sheet1", "C1", 10.0).unwrap();
    engine.set_cell_value("Sheet1", "C2", 20.0).unwrap();
    engine.set_cell_value("Sheet1", "C3", 30.0).unwrap();
    engine.set_cell_value("Sheet1", "C4", 40.0).unwrap();

    engine.set_cell_value("Sheet1", "D1", 100.0).unwrap();
    engine.set_cell_value("Sheet1", "D2", 200.0).unwrap();
    engine.set_cell_value("Sheet1", "D3", 300.0).unwrap();
    engine.set_cell_value("Sheet1", "D4", 400.0).unwrap();

    // Return array is 4x2, so XLOOKUP should spill a 1x2 row slice.
    engine
        .set_cell_formula("Sheet1", "F1", r#"=XLOOKUP("b",B1:B4,C1:D4)"#)
        .unwrap();

    assert_eq!(engine.bytecode_program_count(), 1);
    engine.recalculate_single_threaded();

    let (start, end) = engine.spill_range("Sheet1", "F1").expect("spill range");
    assert_eq!(start, parse_a1("F1").unwrap());
    assert_eq!(end, parse_a1("G1").unwrap());

    assert_eq!(engine.get_cell_value("Sheet1", "F1"), Value::Number(20.0));
    assert_eq!(engine.get_cell_value("Sheet1", "G1"), Value::Number(200.0));
}

#[test]
fn bytecode_backend_xlookup_spills_column_result() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "B6", "a").unwrap();
    engine.set_cell_value("Sheet1", "C6", "b").unwrap();
    engine.set_cell_value("Sheet1", "D6", "c").unwrap();
    engine.set_cell_value("Sheet1", "E6", "d").unwrap();

    engine.set_cell_value("Sheet1", "B7", 10.0).unwrap();
    engine.set_cell_value("Sheet1", "C7", 20.0).unwrap();
    engine.set_cell_value("Sheet1", "D7", 30.0).unwrap();
    engine.set_cell_value("Sheet1", "E7", 40.0).unwrap();

    engine.set_cell_value("Sheet1", "B8", 100.0).unwrap();
    engine.set_cell_value("Sheet1", "C8", 200.0).unwrap();
    engine.set_cell_value("Sheet1", "D8", 300.0).unwrap();
    engine.set_cell_value("Sheet1", "E8", 400.0).unwrap();

    // Return array is 2x4, so XLOOKUP should spill a 2x1 column slice.
    engine
        .set_cell_formula("Sheet1", "G6", r#"=XLOOKUP("c",B6:E6,B7:E8)"#)
        .unwrap();

    assert_eq!(engine.bytecode_program_count(), 1);
    engine.recalculate_single_threaded();

    let (start, end) = engine.spill_range("Sheet1", "G6").expect("spill range");
    assert_eq!(start, parse_a1("G6").unwrap());
    assert_eq!(end, parse_a1("G7").unwrap());

    assert_eq!(engine.get_cell_value("Sheet1", "G6"), Value::Number(30.0));
    assert_eq!(engine.get_cell_value("Sheet1", "G7"), Value::Number(300.0));
}
