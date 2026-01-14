use formula_engine::{ErrorKind, Value};

use super::harness::{assert_number, TestSheet};

#[test]
fn coup_functions_default_basis_to_zero_when_omitted() {
    let mut sheet = TestSheet::new();

    for formula in [
        "=COUPDAYBS(DATE(2024,6,15),DATE(2025,1,1),2)-COUPDAYBS(DATE(2024,6,15),DATE(2025,1,1),2,0)",
        "=COUPDAYS(DATE(2024,6,15),DATE(2025,1,1),2)-COUPDAYS(DATE(2024,6,15),DATE(2025,1,1),2,0)",
        "=COUPDAYSNC(DATE(2024,6,15),DATE(2025,1,1),2)-COUPDAYSNC(DATE(2024,6,15),DATE(2025,1,1),2,0)",
        // COUPNCD/COUPPCD return date serial numbers; the difference should be 0.
        "=COUPNCD(DATE(2024,6,15),DATE(2025,1,1),2)-COUPNCD(DATE(2024,6,15),DATE(2025,1,1),2,0)",
        "=COUPNUM(DATE(2024,6,15),DATE(2025,1,1),2)-COUPNUM(DATE(2024,6,15),DATE(2025,1,1),2,0)",
        "=COUPPCD(DATE(2024,6,15),DATE(2025,1,1),2)-COUPPCD(DATE(2024,6,15),DATE(2025,1,1),2,0)",
    ] {
        assert_number(&sheet.eval(formula), 0.0);
    }
}

#[test]
fn coup_functions_treat_blank_and_missing_basis_as_zero() {
    let mut sheet = TestSheet::new();

    // `Y1` is unset/blank.
    for formula in [
        // Blank basis cell.
        "=COUPDAYBS(DATE(2024,6,15),DATE(2025,1,1),2,Y1)-COUPDAYBS(DATE(2024,6,15),DATE(2025,1,1),2,0)",
        "=COUPDAYS(DATE(2024,6,15),DATE(2025,1,1),2,Y1)-COUPDAYS(DATE(2024,6,15),DATE(2025,1,1),2,0)",
        "=COUPDAYSNC(DATE(2024,6,15),DATE(2025,1,1),2,Y1)-COUPDAYSNC(DATE(2024,6,15),DATE(2025,1,1),2,0)",
        "=COUPNCD(DATE(2024,6,15),DATE(2025,1,1),2,Y1)-COUPNCD(DATE(2024,6,15),DATE(2025,1,1),2,0)",
        "=COUPNUM(DATE(2024,6,15),DATE(2025,1,1),2,Y1)-COUPNUM(DATE(2024,6,15),DATE(2025,1,1),2,0)",
        "=COUPPCD(DATE(2024,6,15),DATE(2025,1,1),2,Y1)-COUPPCD(DATE(2024,6,15),DATE(2025,1,1),2,0)",
        // Explicitly missing basis arg (trailing comma).
        "=COUPDAYBS(DATE(2024,6,15),DATE(2025,1,1),2,)-COUPDAYBS(DATE(2024,6,15),DATE(2025,1,1),2,0)",
        "=COUPDAYS(DATE(2024,6,15),DATE(2025,1,1),2,)-COUPDAYS(DATE(2024,6,15),DATE(2025,1,1),2,0)",
        "=COUPDAYSNC(DATE(2024,6,15),DATE(2025,1,1),2,)-COUPDAYSNC(DATE(2024,6,15),DATE(2025,1,1),2,0)",
        "=COUPNCD(DATE(2024,6,15),DATE(2025,1,1),2,)-COUPNCD(DATE(2024,6,15),DATE(2025,1,1),2,0)",
        "=COUPNUM(DATE(2024,6,15),DATE(2025,1,1),2,)-COUPNUM(DATE(2024,6,15),DATE(2025,1,1),2,0)",
        "=COUPPCD(DATE(2024,6,15),DATE(2025,1,1),2,)-COUPPCD(DATE(2024,6,15),DATE(2025,1,1),2,0)",
    ] {
        assert_number(&sheet.eval(formula), 0.0);
    }
}

#[test]
fn bond_pricing_functions_default_basis_to_zero_when_omitted() {
    let mut sheet = TestSheet::new();

    // PRICE(settlement,maturity,rate,yld,redemption,frequency,[basis])
    assert_number(
        &sheet.eval("=PRICE(DATE(2024,6,15),DATE(2025,1,1),0.0575,0.065,100,2)-PRICE(DATE(2024,6,15),DATE(2025,1,1),0.0575,0.065,100,2,0)"),
        0.0,
    );

    // YIELD(settlement,maturity,rate,pr,redemption,frequency,[basis])
    assert_number(
        &sheet.eval("=YIELD(DATE(2008,2,15),DATE(2017,11,15),0.0575,95.04287,100,2)-YIELD(DATE(2008,2,15),DATE(2017,11,15),0.0575,95.04287,100,2,0)"),
        0.0,
    );

    // DURATION(settlement,maturity,coupon,yld,frequency,[basis])
    assert_number(
        &sheet.eval(
            "=DURATION(DATE(2024,1,1),DATE(2026,1,1),0.08,0.09,2)-DURATION(DATE(2024,1,1),DATE(2026,1,1),0.08,0.09,2,0)",
        ),
        0.0,
    );

    // MDURATION(settlement,maturity,coupon,yld,frequency,[basis])
    assert_number(
        &sheet.eval(
            "=MDURATION(DATE(2024,1,1),DATE(2026,1,1),0.08,0.09,2)-MDURATION(DATE(2024,1,1),DATE(2026,1,1),0.08,0.09,2,0)",
        ),
        0.0,
    );
}

#[test]
fn bond_pricing_functions_treat_blank_and_missing_basis_as_zero() {
    let mut sheet = TestSheet::new();

    // `Y1` is unset/blank.

    // PRICE(settlement,maturity,rate,yld,redemption,frequency,[basis])
    assert_number(
        &sheet.eval("=PRICE(DATE(2024,6,15),DATE(2025,1,1),0.0575,0.065,100,2,Y1)-PRICE(DATE(2024,6,15),DATE(2025,1,1),0.0575,0.065,100,2,0)"),
        0.0,
    );
    assert_number(
        &sheet.eval("=PRICE(DATE(2024,6,15),DATE(2025,1,1),0.0575,0.065,100,2,)-PRICE(DATE(2024,6,15),DATE(2025,1,1),0.0575,0.065,100,2,0)"),
        0.0,
    );

    // YIELD(settlement,maturity,rate,pr,redemption,frequency,[basis])
    assert_number(
        &sheet.eval("=YIELD(DATE(2008,2,15),DATE(2017,11,15),0.0575,95.04287,100,2,Y1)-YIELD(DATE(2008,2,15),DATE(2017,11,15),0.0575,95.04287,100,2,0)"),
        0.0,
    );
    assert_number(
        &sheet.eval("=YIELD(DATE(2008,2,15),DATE(2017,11,15),0.0575,95.04287,100,2,)-YIELD(DATE(2008,2,15),DATE(2017,11,15),0.0575,95.04287,100,2,0)"),
        0.0,
    );

    // DURATION(settlement,maturity,coupon,yld,frequency,[basis])
    assert_number(
        &sheet.eval(
            "=DURATION(DATE(2024,1,1),DATE(2026,1,1),0.08,0.09,2,Y1)-DURATION(DATE(2024,1,1),DATE(2026,1,1),0.08,0.09,2,0)",
        ),
        0.0,
    );
    assert_number(
        &sheet.eval(
            "=DURATION(DATE(2024,1,1),DATE(2026,1,1),0.08,0.09,2,)-DURATION(DATE(2024,1,1),DATE(2026,1,1),0.08,0.09,2,0)",
        ),
        0.0,
    );

    // MDURATION(settlement,maturity,coupon,yld,frequency,[basis])
    assert_number(
        &sheet.eval(
            "=MDURATION(DATE(2024,1,1),DATE(2026,1,1),0.08,0.09,2,Y1)-MDURATION(DATE(2024,1,1),DATE(2026,1,1),0.08,0.09,2,0)",
        ),
        0.0,
    );
    assert_number(
        &sheet.eval(
            "=MDURATION(DATE(2024,1,1),DATE(2026,1,1),0.08,0.09,2,)-MDURATION(DATE(2024,1,1),DATE(2026,1,1),0.08,0.09,2,0)",
        ),
        0.0,
    );
}

#[test]
fn accrintm_defaults_basis_to_zero_when_omitted() {
    let mut sheet = TestSheet::new();

    // ACCRINTM(issue,settlement,rate,par,[basis])
    assert_number(
        &sheet.eval(
            "=ACCRINTM(DATE(2020,1,1),DATE(2020,7,1),0.1,1000)-ACCRINTM(DATE(2020,1,1),DATE(2020,7,1),0.1,1000,0)",
        ),
        0.0,
    );
}

#[test]
fn accrintm_treats_blank_and_missing_basis_as_zero() {
    let mut sheet = TestSheet::new();

    // `Y1` is unset/blank.
    assert_number(
        &sheet.eval(
            "=ACCRINTM(DATE(2020,1,1),DATE(2020,7,1),0.1,1000,Y1)-ACCRINTM(DATE(2020,1,1),DATE(2020,7,1),0.1,1000,0)",
        ),
        0.0,
    );
    assert_number(
        &sheet.eval(
            "=ACCRINTM(DATE(2020,1,1),DATE(2020,7,1),0.1,1000,)-ACCRINTM(DATE(2020,1,1),DATE(2020,7,1),0.1,1000,0)",
        ),
        0.0,
    );
}

#[test]
fn accrint_defaults_basis_to_zero_and_calc_method_to_false_when_omitted() {
    let mut sheet = TestSheet::new();

    // ACCRINT(issue,first_interest,settlement,rate,par,frequency,[basis],[calc_method])
    assert_number(
        &sheet.eval(
            "=ACCRINT(DATE(2020,2,15),DATE(2020,5,15),DATE(2020,4,15),0.1,1000,2)-ACCRINT(DATE(2020,2,15),DATE(2020,5,15),DATE(2020,4,15),0.1,1000,2,0,FALSE)",
        ),
        0.0,
    );
}

#[test]
fn accrint_defaults_calc_method_to_false_when_omitted_even_if_basis_provided() {
    let mut sheet = TestSheet::new();

    assert_number(
        &sheet.eval(
            "=ACCRINT(DATE(2020,2,15),DATE(2020,5,15),DATE(2020,4,15),0.1,1000,2,0)-ACCRINT(DATE(2020,2,15),DATE(2020,5,15),DATE(2020,4,15),0.1,1000,2,0,FALSE)",
        ),
        0.0,
    );
}

#[test]
fn accrint_treats_blank_and_missing_optional_args_as_defaults() {
    let mut sheet = TestSheet::new();

    // `Y1` is unset/blank.
    assert_number(
        &sheet.eval(
            "=ACCRINT(DATE(2020,2,15),DATE(2020,5,15),DATE(2020,4,15),0.1,1000,2,Y1)-ACCRINT(DATE(2020,2,15),DATE(2020,5,15),DATE(2020,4,15),0.1,1000,2,0,FALSE)",
        ),
        0.0,
    );
    assert_number(
        &sheet.eval(
            "=ACCRINT(DATE(2020,2,15),DATE(2020,5,15),DATE(2020,4,15),0.1,1000,2,,)-ACCRINT(DATE(2020,2,15),DATE(2020,5,15),DATE(2020,4,15),0.1,1000,2,0,FALSE)",
        ),
        0.0,
    );
}

#[test]
fn coup_functions_reject_invalid_inputs_with_num_error() {
    let mut sheet = TestSheet::new();

    for func in [
        "COUPDAYBS",
        "COUPDAYS",
        "COUPDAYSNC",
        "COUPNCD",
        "COUPNUM",
        "COUPPCD",
    ] {
        // settlement >= maturity
        assert_eq!(
            sheet.eval(&format!("={func}(DATE(2025,1,1),DATE(2025,1,1),2,0)")),
            Value::Error(ErrorKind::Num)
        );
        // invalid frequency
        assert_eq!(
            sheet.eval(&format!("={func}(DATE(2024,1,1),DATE(2025,1,1),3,0)")),
            Value::Error(ErrorKind::Num)
        );
        // invalid basis
        assert_eq!(
            sheet.eval(&format!("={func}(DATE(2024,1,1),DATE(2025,1,1),2,99)")),
            Value::Error(ErrorKind::Num)
        );
    }
}
