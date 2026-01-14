use formula_engine::{Engine, ErrorKind, Value};

fn eval_formula(engine: &mut Engine, formula: &str) -> Value {
    engine
        .set_cell_formula("Sheet1", "A1", formula)
        .expect("set formula");
    engine.recalculate_single_threaded();
    engine.get_cell_value("Sheet1", "A1")
}

fn assert_error(v: Value, expected: ErrorKind) {
    match v {
        Value::Error(err) => assert_eq!(err, expected, "expected #{expected:?}!, got #{err:?}!"),
        other => panic!("expected error {expected:?}, got {other:?}"),
    }
}

fn assert_number_close(v: Value, expected: f64, abs_tol: f64, rel_tol: f64) {
    match v {
        Value::Number(n) => {
            let diff = (n - expected).abs();
            let tol = abs_tol.max(rel_tol * expected.abs());
            assert!(
                diff <= tol,
                "expected {expected}, got {n} (diff={diff}, tol={tol})"
            );
        }
        other => panic!("expected number, got {other:?}"),
    }
}

#[test]
fn odd_coupon_validation_cases_match_excel_oracle() {
    // These cases mirror `tests/compatibility/excel-oracle/cases.json` entries tagged
    // `odd_coupon_validation` (and their pinned results in
    // `tests/compatibility/excel-oracle/datasets/versioned/*`).
    //
    // Keeping a few representative validations in a Rust integration test helps catch regressions
    // without requiring the python excel-oracle gate to run locally.
    let mut engine = Engine::new();

    // Yield domain boundary: yld == -frequency => #DIV/0!
    assert_error(
        eval_formula(
            &mut engine,
            "=ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,-2,100,2,0)",
        ),
        ErrorKind::Div0,
    );
    assert_error(
        eval_formula(
            &mut engine,
            "=ODDLPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,-2,100,2,0)",
        ),
        ErrorKind::Div0,
    );

    // Yield below domain: yld < -frequency => #NUM!
    assert_error(
        eval_formula(
            &mut engine,
            "=ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,-2.5,100,2,0)",
        ),
        ErrorKind::Num,
    );
    assert_error(
        eval_formula(
            &mut engine,
            "=ODDLPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,-2.5,100,2,0)",
        ),
        ErrorKind::Num,
    );

    // Negative coupon rate => #NUM!
    assert_error(
        eval_formula(
            &mut engine,
            "=ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),-0.01,0.0625,100,2,0)",
        ),
        ErrorKind::Num,
    );
    assert_error(
        eval_formula(
            &mut engine,
            "=ODDLPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),-0.01,0.0625,100,2,0)",
        ),
        ErrorKind::Num,
    );
    assert_error(
        eval_formula(
            &mut engine,
            "=ODDFYIELD(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),-0.01,98,100,2,0)",
        ),
        ErrorKind::Num,
    );
    assert_error(
        eval_formula(
            &mut engine,
            "=ODDLYIELD(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),-0.01,98,100,2,0)",
        ),
        ErrorKind::Num,
    );

    // Negative yields are allowed when yld > -frequency.
    assert_number_close(
        eval_formula(
            &mut engine,
            "=ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,-0.01,100,2,0)",
        ),
        216.19670076328038,
        1e-6,
        1e-6,
    );
    assert_number_close(
        eval_formula(
            &mut engine,
            "=ODDLPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,-0.01,100,2,0)",
        ),
        102.71450083338773,
        1e-6,
        1e-6,
    );

    // Large positive prices are possible for long-dated bonds at strongly negative yields.
    assert_number_close(
        eval_formula(
            &mut engine,
            "=ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,-1.5,100,2,0)",
        ),
        6.910646360118339e+16,
        1.0,
        1e-12,
    );
    assert_number_close(
        eval_formula(
            &mut engine,
            "=ODDLPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,-1.5,100,2,0)",
        ),
        239.6576768188426,
        1e-6,
        1e-6,
    );
    assert_number_close(
        eval_formula(
            &mut engine,
            "=ODDLPRICE(DATE(2020,8,1),DATE(2021,3,1),DATE(2020,10,15),0.0785,-1.5,100,2,0)",
        ),
        523.5423945534777,
        1e-6,
        1e-6,
    );

    // High prices can imply negative yields; Excel-oracle pins that negative yields are returned.
    assert_number_close(
        eval_formula(
            &mut engine,
            "=ODDFYIELD(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,300,100,2,0)",
        ),
        -0.043049218577708326,
        1e-6,
        1e-6,
    );
    assert_number_close(
        eval_formula(
            &mut engine,
            "=ODDLYIELD(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,300,100,2,0)",
        ),
        -1.6534921783048486,
        1e-6,
        1e-6,
    );

    // Roundtrip negative yield below -1 through PRICE/YIELD helpers.
    assert_number_close(
        eval_formula(
            &mut engine,
            "=ODDFYIELD(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),DATE(2021,3,1),0.0785,ODDFPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),DATE(2021,3,1),0.0785,-1.5,100,2,0),100,2,0)",
        ),
        -1.4999999998157394,
        1e-6,
        1e-6,
    );
    assert_number_close(
        eval_formula(
            &mut engine,
            "=ODDLYIELD(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,ODDLPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,-1.5,100,2,0),100,2,0)",
        ),
        -1.4999999998157394,
        1e-6,
        1e-6,
    );
    assert_number_close(
        eval_formula(
            &mut engine,
            "=ODDLYIELD(DATE(2020,8,1),DATE(2021,3,1),DATE(2020,10,15),0.0785,ODDLPRICE(DATE(2020,8,1),DATE(2021,3,1),DATE(2020,10,15),0.0785,-1.5,100,2,0),100,2,0)",
        ),
        -1.4999999999903624,
        1e-6,
        1e-6,
    );
}
