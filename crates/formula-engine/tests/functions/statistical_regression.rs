use formula_engine::eval::parse_a1;
use formula_engine::locale::ValueLocaleConfig;
use formula_engine::{Engine, ErrorKind, Value};

fn assert_number_close(value: Value, expected: f64) {
    match value {
        Value::Number(n) => {
            assert!((n - expected).abs() < 1e-9, "expected {expected}, got {n}");
        }
        other => panic!("expected number {expected}, got {other:?}"),
    }
}

fn assert_error(value: Value, expected: ErrorKind) {
    assert_eq!(value, Value::Error(expected));
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
fn linest_and_trend_default_known_x_is_sequence() {
    let mut engine = Engine::new();

    // With known_x omitted, Excel defaults to x = 1..n.
    // y = 2x for x=1..3.
    for (i, y) in [2.0, 4.0, 6.0].into_iter().enumerate() {
        let row = i + 1;
        engine
            .set_cell_value("Sheet1", &format!("A{row}"), y)
            .unwrap();
    }

    engine
        .set_cell_formula("Sheet1", "C1", "=LINEST(A1:A3)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "C3", "=TREND(A1:A3)")
        .unwrap();

    engine.recalculate_single_threaded();

    let (start, end) = engine.spill_range("Sheet1", "C1").expect("spill range");
    assert_eq!(start, parse_a1("C1").unwrap());
    assert_eq!(end, parse_a1("D1").unwrap());
    assert_number_close(engine.get_cell_value("Sheet1", "C1"), 2.0);
    assert_number_close(engine.get_cell_value("Sheet1", "D1"), 0.0);

    let (start, end) = engine.spill_range("Sheet1", "C3").expect("spill range");
    assert_eq!(start, parse_a1("C3").unwrap());
    assert_eq!(end, parse_a1("C5").unwrap());
    assert_number_close(engine.get_cell_value("Sheet1", "C3"), 2.0);
    assert_number_close(engine.get_cell_value("Sheet1", "C4"), 4.0);
    assert_number_close(engine.get_cell_value("Sheet1", "C5"), 6.0);
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
fn linest_const_false_forces_intercept_zero() {
    let mut engine = Engine::new();

    // y = 2x + 1 for x=1..5 (intercept would normally be 1, but const=FALSE forces b=0).
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

    // Slope with forced intercept is Σ(xy) / Σ(x^2) = 125/55 = 25/11.
    engine
        .set_cell_formula("Sheet1", "D1", "=LINEST(A1:A5,B1:B5,FALSE)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "D3", "=TREND(A1:A5,B1:B5,{6;7},FALSE)")
        .unwrap();

    engine.recalculate_single_threaded();

    assert_number_close(engine.get_cell_value("Sheet1", "D1"), 25.0 / 11.0);
    assert_number_close(engine.get_cell_value("Sheet1", "E1"), 0.0);

    assert_number_close(engine.get_cell_value("Sheet1", "D3"), 150.0 / 11.0);
    assert_number_close(engine.get_cell_value("Sheet1", "D4"), 175.0 / 11.0);
}

#[test]
fn linest_stats_true_spills_5_rows() {
    let mut engine = Engine::new();

    // Perfect fit: y = 2x + 1 for x=1..5.
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
        .set_cell_formula("Sheet1", "D1", "=LINEST(A1:A5,B1:B5,TRUE,TRUE)")
        .unwrap();

    engine.recalculate_single_threaded();

    let (start, end) = engine.spill_range("Sheet1", "D1").expect("spill range");
    assert_eq!(start, parse_a1("D1").unwrap());
    assert_eq!(end, parse_a1("E5").unwrap());

    // Row 1: slope, intercept.
    assert_number_close(engine.get_cell_value("Sheet1", "D1"), 2.0);
    assert_number_close(engine.get_cell_value("Sheet1", "E1"), 1.0);
    // Row 3: R^2.
    assert_number_close(engine.get_cell_value("Sheet1", "D3"), 1.0);
    // Row 4: df (n - k, where k = p+1).
    assert_number_close(engine.get_cell_value("Sheet1", "E4"), 3.0);
    // Row 5: ssreg == 40, ssresid == 0.
    assert_number_close(engine.get_cell_value("Sheet1", "D5"), 40.0);
    assert_number_close(engine.get_cell_value("Sheet1", "E5"), 0.0);
}

#[test]
fn logest_and_growth_simple_exponential() {
    let mut engine = Engine::new();

    // y = 3 * 2^x for x=0..4
    for (i, (x, y)) in [
        (0.0, 3.0),
        (1.0, 6.0),
        (2.0, 12.0),
        (3.0, 24.0),
        (4.0, 48.0),
    ]
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
fn logest_errors_on_nonpositive_y() {
    let mut engine = Engine::new();

    engine.set_cell_value("Sheet1", "A1", 0.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "B1", 0.0).unwrap();
    engine.set_cell_value("Sheet1", "B2", 1.0).unwrap();

    engine
        .set_cell_formula("Sheet1", "D1", "=LOGEST(A1:A2,B1:B2)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "D2", "=GROWTH(A1:A2,B1:B2,{2;3})")
        .unwrap();

    engine.recalculate_single_threaded();

    assert_error(engine.get_cell_value("Sheet1", "D1"), ErrorKind::Num);
    assert_error(engine.get_cell_value("Sheet1", "D2"), ErrorKind::Num);
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
fn linest_multi_x_rows_are_predictors_orientation() {
    let mut engine = Engine::new();

    // Same dataset as `linest_multi_x_two_predictors`, but with:
    // - y in a row vector
    // - known_x arranged as p rows x n columns (rows are predictors)
    //
    // y = 1 + 2*x1 + 3*x2
    // Observations:
    //   (x1, x2) = (0,0), (1,0), (0,1), (1,1)
    //   y        =   1 ,   3 ,   4 ,   6
    for (col, y) in [1.0, 3.0, 4.0, 6.0].into_iter().enumerate() {
        let addr = format!("{}1", (b'A' + col as u8) as char);
        engine.set_cell_value("Sheet1", &addr, y).unwrap();
    }
    // x1 row.
    for (col, x1) in [0.0, 1.0, 0.0, 1.0].into_iter().enumerate() {
        let addr = format!("{}2", (b'A' + col as u8) as char);
        engine.set_cell_value("Sheet1", &addr, x1).unwrap();
    }
    // x2 row.
    for (col, x2) in [0.0, 0.0, 1.0, 1.0].into_iter().enumerate() {
        let addr = format!("{}3", (b'A' + col as u8) as char);
        engine.set_cell_value("Sheet1", &addr, x2).unwrap();
    }

    engine
        .set_cell_formula("Sheet1", "F1", "=LINEST(A1:D1,A2:D3)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "F3", "=TREND(A1:D1,A2:D3,{2,0;0,2})")
        .unwrap();

    engine.recalculate_single_threaded();

    let (start, end) = engine.spill_range("Sheet1", "F1").expect("spill range");
    assert_eq!(start, parse_a1("F1").unwrap());
    assert_eq!(end, parse_a1("H1").unwrap());
    // Coefficients in reverse predictor order: {m_x2, m_x1, b}.
    assert_number_close(engine.get_cell_value("Sheet1", "F1"), 3.0);
    assert_number_close(engine.get_cell_value("Sheet1", "G1"), 2.0);
    assert_number_close(engine.get_cell_value("Sheet1", "H1"), 1.0);

    // With rows-as-predictors new_x, TREND spills horizontally (1 x n_new).
    let (start, end) = engine.spill_range("Sheet1", "F3").expect("spill range");
    assert_eq!(start, parse_a1("F3").unwrap());
    assert_eq!(end, parse_a1("G3").unwrap());
    assert_number_close(engine.get_cell_value("Sheet1", "F3"), 5.0);
    assert_number_close(engine.get_cell_value("Sheet1", "G3"), 7.0);
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
