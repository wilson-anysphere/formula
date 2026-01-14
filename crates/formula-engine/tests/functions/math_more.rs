use formula_engine::{ErrorKind, Value};

use super::harness::{assert_number, TestSheet};

#[test]
fn degrees_radians_match_excel_semantics() {
    let mut sheet = TestSheet::new();
    assert_number(&sheet.eval("=RADIANS(180)"), std::f64::consts::PI);
    assert_number(&sheet.eval("=DEGREES(PI())"), 180.0);
}

#[test]
fn hyperbolic_trig_matches_excel_semantics() {
    let mut sheet = TestSheet::new();

    assert_number(&sheet.eval("=SINH(0)"), 0.0);
    assert_number(&sheet.eval("=COSH(0)"), 1.0);
    assert_number(&sheet.eval("=TANH(0)"), 0.0);

    assert_number(&sheet.eval("=ASINH(0)"), 0.0);
    assert_number(&sheet.eval("=ACOSH(1)"), 0.0);
    assert_number(&sheet.eval("=ATANH(0)"), 0.0);

    assert_eq!(sheet.eval("=ACOSH(0.5)"), Value::Error(ErrorKind::Num));
    assert_eq!(sheet.eval("=ATANH(1)"), Value::Error(ErrorKind::Num));

    assert_eq!(sheet.eval("=COTH(0)"), Value::Error(ErrorKind::Div0));
    assert_eq!(sheet.eval("=CSCH(0)"), Value::Error(ErrorKind::Div0));
    assert_number(&sheet.eval("=SECH(0)"), 1.0);

    assert_number(&sheet.eval("=ACOTH(2)"), 0.549_306_144_334_054_9);
    assert_eq!(sheet.eval("=ACOTH(1)"), Value::Error(ErrorKind::Num));
}

#[test]
fn trig_reciprocals_match_excel_semantics() {
    let mut sheet = TestSheet::new();

    assert_eq!(sheet.eval("=COT(0)"), Value::Error(ErrorKind::Div0));
    assert_eq!(sheet.eval("=CSC(0)"), Value::Error(ErrorKind::Div0));
    assert_number(&sheet.eval("=COT(PI()/4)"), 1.0);
    assert_number(&sheet.eval("=CSC(PI()/2)"), 1.0);
    assert_number(&sheet.eval("=SEC(0)"), 1.0);

    assert_number(&sheet.eval("=ACOT(0)"), std::f64::consts::FRAC_PI_2);
    assert_number(&sheet.eval("=ACOT(1)"), std::f64::consts::FRAC_PI_4);
}

#[test]
fn factorials_and_combinatorics_match_excel_semantics() {
    let mut sheet = TestSheet::new();

    assert_number(&sheet.eval("=FACT(5)"), 120.0);
    assert_number(&sheet.eval("=FACT(5.9)"), 120.0);
    assert_eq!(sheet.eval("=FACT(-1)"), Value::Error(ErrorKind::Num));

    assert_number(&sheet.eval("=FACTDOUBLE(6)"), 48.0);
    assert_number(&sheet.eval("=FACTDOUBLE(7)"), 105.0);
    assert_eq!(sheet.eval("=FACTDOUBLE(-1)"), Value::Error(ErrorKind::Num));

    assert_number(&sheet.eval("=COMBIN(5,2)"), 10.0);
    assert_eq!(sheet.eval("=COMBIN(5,7)"), Value::Error(ErrorKind::Num));
    assert_number(&sheet.eval("=PERMUT(5,2)"), 20.0);

    assert_number(&sheet.eval("=COMBINA(3,2)"), 6.0);
    assert_number(&sheet.eval("=PERMUTATIONA(3,2)"), 9.0);
}

#[test]
fn integer_helpers_match_excel_semantics() {
    let mut sheet = TestSheet::new();

    assert_number(&sheet.eval("=GCD(24,36)"), 12.0);
    assert_number(&sheet.eval("=GCD(0,0)"), 0.0);
    assert_eq!(sheet.eval("=GCD(-2,4)"), Value::Error(ErrorKind::Num));

    assert_number(&sheet.eval("=LCM(4,6)"), 12.0);
    assert_number(&sheet.eval("=LCM(0,5)"), 0.0);

    assert_number(&sheet.eval("=MULTINOMIAL(1,2,3)"), 60.0);

    assert_number(&sheet.eval("=MROUND(10,3)"), 9.0);
    assert_eq!(sheet.eval("=MROUND(-10,3)"), Value::Error(ErrorKind::Num));
    assert_number(&sheet.eval("=MROUND(-10,-3)"), -9.0);

    assert_number(&sheet.eval("=EVEN(1)"), 2.0);
    assert_number(&sheet.eval("=EVEN(-1)"), -2.0);
    assert_number(&sheet.eval("=ODD(2)"), 3.0);
    assert_number(&sheet.eval("=ODD(-2)"), -3.0);
    assert_number(&sheet.eval("=ODD(0)"), 1.0);
    // Excel doesn't preserve a negative-zero sign bit; treat `-0` as `0`.
    assert_number(&sheet.eval("=ODD(-0)"), 1.0);
    assert_number(&sheet.eval("=EVEN(-0)"), 0.0);

    assert_eq!(sheet.eval("=ISEVEN(2.5)"), Value::Bool(true));
    assert_eq!(sheet.eval("=ISODD(2.5)"), Value::Bool(false));
    assert_eq!(sheet.eval("=ISODD(3)"), Value::Bool(true));

    assert_number(&sheet.eval("=QUOTIENT(5,2)"), 2.0);
    assert_number(&sheet.eval("=QUOTIENT(-5,2)"), -2.0);
    assert_eq!(sheet.eval("=QUOTIENT(5,0)"), Value::Error(ErrorKind::Div0));

    assert_number(&sheet.eval("=SQRTPI(1)"), std::f64::consts::PI.sqrt());

    assert_number(&sheet.eval("=DELTA(0)"), 1.0);
    assert_number(&sheet.eval("=DELTA(5,4)"), 0.0);

    assert_number(&sheet.eval("=GESTEP(-1)"), 0.0);
    assert_number(&sheet.eval("=GESTEP(5,4)"), 1.0);
}

#[test]
fn series_and_sumx_helpers_match_excel_semantics() {
    let mut sheet = TestSheet::new();

    assert_number(&sheet.eval("=SERIESSUM(2,0,1,{1,2,3})"), 17.0);
    assert_eq!(
        sheet.eval("=SERIESSUM(0,-1,1,{1})"),
        Value::Error(ErrorKind::Div0)
    );

    assert_number(&sheet.eval("=SUMXMY2({1,2},{3,4})"), 8.0);
    assert_number(&sheet.eval("=SUMX2MY2({1,2},{3,4})"), -20.0);
    assert_number(&sheet.eval("=SUMX2PY2({1,2},{3,4})"), 30.0);
    assert_eq!(
        sheet.eval("=SUMXMY2({1,2},{3})"),
        Value::Error(ErrorKind::NA)
    );
    assert_number(&sheet.eval("=SUMXMY2({1,\"x\",3},{1,2,3})"), 0.0);
}

#[test]
fn math_more_functions_support_elementwise_spilling() {
    let mut sheet = TestSheet::new();
    sheet.set_formula("A1", "=RADIANS({0,180})");
    sheet.recalculate();
    assert_number(&sheet.get("A1"), 0.0);
    assert_number(&sheet.get("B1"), std::f64::consts::PI);
}
