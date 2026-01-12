use formula_engine::{ErrorKind, Value};

use super::harness::{assert_number, TestSheet};

fn assert_close(actual: f64, expected: f64, tol: f64) {
    assert!(
        (actual - expected).abs() <= tol,
        "expected {expected}, got {actual}"
    );
}

#[test]
fn builtins_coup_functions_accept_iso_date_text_via_datevalue() {
    let mut sheet = TestSheet::new();

    // COUP* helpers should accept ISO-like date strings via DATEVALUE-style coercion.
    // Compare directly against the DATE(...) version so we don't need to hardcode serial numbers.
    assert_number(
        &sheet.eval(
            r#"=COUPPCD("2024-06-15","2025-01-01",2,0)-COUPPCD(DATE(2024,6,15),DATE(2025,1,1),2,0)"#,
        ),
        0.0,
    );
    assert_number(
        &sheet.eval(
            r#"=COUPNCD("2024-06-15","2025-01-01",2,0)-COUPNCD(DATE(2024,6,15),DATE(2025,1,1),2,0)"#,
        ),
        0.0,
    );

    assert_number(
        &sheet.eval(
            r#"=COUPNUM("2024-06-15","2025-01-01",2,0)-COUPNUM(DATE(2024,6,15),DATE(2025,1,1),2,0)"#,
        ),
        0.0,
    );
    assert_number(
        &sheet.eval(
            r#"=COUPDAYBS("2024-06-15","2025-01-01",2,0)-COUPDAYBS(DATE(2024,6,15),DATE(2025,1,1),2,0)"#,
        ),
        0.0,
    );
    assert_number(
        &sheet.eval(
            r#"=COUPDAYSNC("2024-06-15","2025-01-01",2,0)-COUPDAYSNC(DATE(2024,6,15),DATE(2025,1,1),2,0)"#,
        ),
        0.0,
    );
    assert_number(
        &sheet.eval(
            r#"=COUPDAYS("2024-06-15","2025-01-01",2,0)-COUPDAYS(DATE(2024,6,15),DATE(2025,1,1),2,0)"#,
        ),
        0.0,
    );
}

#[test]
fn builtins_accrint_functions_accept_iso_date_text_via_datevalue() {
    let mut sheet = TestSheet::new();

    assert_number(
        &sheet.eval(
            r#"=ACCRINTM("2020-01-01","2020-07-01",0.1,1000,0)-ACCRINTM(DATE(2020,1,1),DATE(2020,7,1),0.1,1000,0)"#,
        ),
        0.0,
    );

    assert_number(
        &sheet.eval(
            r#"=ACCRINT("2020-02-15","2020-05-15","2020-04-15",0.1,1000,2,0)-ACCRINT(DATE(2020,2,15),DATE(2020,5,15),DATE(2020,4,15),0.1,1000,2,0)"#,
        ),
        0.0,
    );
}

#[test]
fn builtins_yield_duration_mduration_accept_iso_date_text_via_datevalue() {
    let mut sheet = TestSheet::new();

    // Excel docs:
    // YIELD(DATE(2008,2,15), DATE(2017,11,15), 0.0575, 95.04287, 100, 2, 0) ≈ 0.064
    // Compare the date-text and DATE() forms to validate DATEVALUE-style coercion.
    let y_text = match sheet.eval(
        r#"=YIELD("2008-02-15","2017-11-15",0.0575,95.04287,100,2,0)"#,
    ) {
        Value::Number(n) => n,
        other => panic!("expected number from YIELD with date text, got {other:?}"),
    };
    let y_date = match sheet.eval(
        "=YIELD(DATE(2008,2,15),DATE(2017,11,15),0.0575,95.04287,100,2,0)",
    ) {
        Value::Number(n) => n,
        other => panic!("expected number from YIELD with DATE(), got {other:?}"),
    };
    assert_close(y_text, y_date, 0.0);

    assert_number(
        &sheet.eval(
            r#"=DURATION("2008-01-01","2016-01-01",0.08,0.09,2,1)-DURATION(DATE(2008,1,1),DATE(2016,1,1),0.08,0.09,2,1)"#,
        ),
        0.0,
    );
    assert_number(
        &sheet.eval(
            r#"=MDURATION("2008-01-01","2016-01-01",0.08,0.09,2,1)-MDURATION(DATE(2008,1,1),DATE(2016,1,1),0.08,0.09,2,1)"#,
        ),
        0.0,
    );
}

#[test]
fn unparseable_date_text_maps_to_value_error_in_bond_coupon_builtins() {
    let mut sheet = TestSheet::new();

    assert_eq!(
        sheet.eval(r#"=COUPDAYS("nope",DATE(2025,1,1),2,0)"#),
        Value::Error(ErrorKind::Value)
    );

    assert_eq!(
        sheet.eval(r#"=ACCRINTM("nope",DATE(2020,7,1),0.1,1000,0)"#),
        Value::Error(ErrorKind::Value)
    );
}

