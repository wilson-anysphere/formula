use formula_engine::eval::parse_a1;
use formula_engine::locale::ValueLocaleConfig;
use formula_engine::{Engine, ErrorKind, Value};

fn assert_number_close(value: Value, expected: f64) {
    match value {
        Value::Number(n) => {
            assert!(
                (n - expected).abs() < 1e-9,
                "expected {expected}, got {n}"
            );
        }
        other => panic!("expected number {expected}, got {other:?}"),
    }
}

#[test]
fn linest_and_trend_simple_1d() {
    let mut engine = Engine::new();

    // y = 2x + 1 for x=1..5
    for (i, (x, y)) in [(1.0, 3.0), (2.0, 5.0), (3.0, 7.0), (4.0, 9.0), (5.0, 11.0)]
        .into_iter()
        .enumerate()
    {
        let row = i + 1;
        engine
            .set_cell_value("Sheet1", &format!("A{row}"), y)
            .unwrap();
        engine
            .set_cell_value("Sheet1", &format!("B{row}"), x)
            .unwrap();
    }

    engine
        .set_cell_formula("Sheet1", "D1", "=LINEST(A1:A5,B1:B5)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "D3", "=TREND(A1:A5,B1:B5,{6;7})")
        .unwrap();

    engine.recalculate_single_threaded();

    let (start, end) = engine.spill_range("Sheet1", "D1").expect("spill range");
    assert_eq!(start, parse_a1("D1").unwrap());
    assert_eq!(end, parse_a1("E1").unwrap());
    assert_number_close(engine.get_cell_value("Sheet1", "D1"), 2.0);
    assert_number_close(engine.get_cell_value("Sheet1", "E1"), 1.0);

    let (start, end) = engine.spill_range("Sheet1", "D3").expect("spill range");
    assert_eq!(start, parse_a1("D3").unwrap());
    assert_eq!(end, parse_a1("D4").unwrap());
    assert_number_close(engine.get_cell_value("Sheet1", "D3"), 13.0);
    assert_number_close(engine.get_cell_value("Sheet1", "D4"), 15.0);
}

#[test]
fn trend_parses_new_x_text_using_value_locale() {
    let mut engine = Engine::new();
    engine.set_value_locale(ValueLocaleConfig::de_de());

    // y = 2x + 1 for x=1..5
    for (i, (x, y)) in [(1.0, 3.0), (2.0, 5.0), (3.0, 7.0), (4.0, 9.0), (5.0, 11.0)]
        .into_iter()
        .enumerate()
    {
        let row = i + 1;
        engine
            .set_cell_value("Sheet1", &format!("A{row}"), y)
            .unwrap();
        engine
            .set_cell_value("Sheet1", &format!("B{row}"), x)
            .unwrap();
    }

    // new_x is provided as locale-formatted numeric text ("6,0" in de-DE).
    engine
        .set_cell_formula("Sheet1", "D1", r#"=TREND(A1:A5,B1:B5,"6,0")"#)
        .unwrap();
    engine.recalculate_single_threaded();
    assert_number_close(engine.get_cell_value("Sheet1", "D1"), 13.0);
}

#[test]
fn logest_and_growth_simple_exponential() {
    let mut engine = Engine::new();

    // y = 3 * 2^x for x=0..4
    for (i, (x, y)) in [(0.0, 3.0), (1.0, 6.0), (2.0, 12.0), (3.0, 24.0), (4.0, 48.0)]
        .into_iter()
        .enumerate()
    {
        let row = i + 1;
        engine
            .set_cell_value("Sheet1", &format!("A{row}"), y)
            .unwrap();
        engine
            .set_cell_value("Sheet1", &format!("B{row}"), x)
            .unwrap();
    }

    engine
        .set_cell_formula("Sheet1", "D1", "=LOGEST(A1:A5,B1:B5)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "D3", "=GROWTH(A1:A5,B1:B5,{5;6})")
        .unwrap();

    engine.recalculate_single_threaded();

    let (start, end) = engine.spill_range("Sheet1", "D1").expect("spill range");
    assert_eq!(start, parse_a1("D1").unwrap());
    assert_eq!(end, parse_a1("E1").unwrap());
    assert_number_close(engine.get_cell_value("Sheet1", "D1"), 2.0);
    assert_number_close(engine.get_cell_value("Sheet1", "E1"), 3.0);

    let (start, end) = engine.spill_range("Sheet1", "D3").expect("spill range");
    assert_eq!(start, parse_a1("D3").unwrap());
    assert_eq!(end, parse_a1("D4").unwrap());
    assert_number_close(engine.get_cell_value("Sheet1", "D3"), 96.0);
    assert_number_close(engine.get_cell_value("Sheet1", "D4"), 192.0);
}

#[test]
fn linest_multi_x_two_predictors() {
    let mut engine = Engine::new();

    // y = 1 + 2*x1 + 3*x2
    // rows: (x1, x2, y)
    let data = [
        (0.0, 0.0, 1.0),
        (1.0, 0.0, 3.0),
        (0.0, 1.0, 4.0),
        (1.0, 1.0, 6.0),
    ];
    for (i, (x1, x2, y)) in data.into_iter().enumerate() {
        let row = i + 1;
        engine
            .set_cell_value("Sheet1", &format!("A{row}"), y)
            .unwrap();
        engine
            .set_cell_value("Sheet1", &format!("B{row}"), x1)
            .unwrap();
        engine
            .set_cell_value("Sheet1", &format!("C{row}"), x2)
            .unwrap();
    }

    // LINEST returns coefficients in reverse X column order: {m_x2, m_x1, b}
    engine
        .set_cell_formula("Sheet1", "E1", "=LINEST(A1:A4,B1:C4)")
        .unwrap();

    engine
        .set_cell_formula("Sheet1", "E3", "=TREND(A1:A4,B1:C4,{2,0;0,2})")
        .unwrap();

    engine.recalculate_single_threaded();

    let (start, end) = engine.spill_range("Sheet1", "E1").expect("spill range");
    assert_eq!(start, parse_a1("E1").unwrap());
    assert_eq!(end, parse_a1("G1").unwrap());
    assert_number_close(engine.get_cell_value("Sheet1", "E1"), 3.0);
    assert_number_close(engine.get_cell_value("Sheet1", "F1"), 2.0);
    assert_number_close(engine.get_cell_value("Sheet1", "G1"), 1.0);

    let (start, end) = engine.spill_range("Sheet1", "E3").expect("spill range");
    assert_eq!(start, parse_a1("E3").unwrap());
    assert_eq!(end, parse_a1("E4").unwrap());
    assert_number_close(engine.get_cell_value("Sheet1", "E3"), 5.0);
    assert_number_close(engine.get_cell_value("Sheet1", "E4"), 7.0);
}

#[test]
fn linest_errors_on_shape_mismatch_and_insufficient_points() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", 3.0).unwrap();
    engine.set_cell_value("Sheet1", "B1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "B2", 2.0).unwrap();

    engine
        .set_cell_formula("Sheet1", "D1", "=LINEST(A1:A3,B1:B2)")
        .unwrap();

    engine
        .set_cell_formula("Sheet1", "D3", "=LINEST(A1:A1,B1:B1)")
        .unwrap();

    engine.recalculate_single_threaded();

    assert_eq!(
        engine.get_cell_value("Sheet1", "D1"),
        Value::Error(ErrorKind::Ref)
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "D3"),
        Value::Error(ErrorKind::Div0)
    );
}
