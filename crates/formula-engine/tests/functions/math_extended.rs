use formula_engine::date::ExcelDateSystem;
use formula_engine::locale::ValueLocaleConfig;
use formula_engine::{ErrorKind, Value};

use super::harness::{assert_number, TestSheet};

#[test]
fn core_math_functions_match_excel_errors() {
    let mut sheet = TestSheet::new();

    assert_number(&sheet.eval("=PRODUCT(2,3,4)"), 24.0);
    assert_number(&sheet.eval("=POWER(2,3)"), 8.0);

    // POWER domain errors.
    assert_eq!(sheet.eval("=POWER(0,-1)"), Value::Error(ErrorKind::Div0));
    assert_eq!(sheet.eval("=POWER(-1,0.5)"), Value::Error(ErrorKind::Num));

    // LN/LOG domain errors.
    assert_eq!(sheet.eval("=LN(-1)"), Value::Error(ErrorKind::Num));
    assert_eq!(sheet.eval("=LOG(10,1)"), Value::Error(ErrorKind::Num));

    // EXP overflow.
    assert_eq!(sheet.eval("=EXP(1000)"), Value::Error(ErrorKind::Num));

    assert_number(&sheet.eval("=LOG10(100)"), 2.0);
    assert_number(&sheet.eval("=SQRT(9)"), 3.0);
    assert_eq!(sheet.eval("=SQRT(-1)"), Value::Error(ErrorKind::Num));
}

#[test]
fn pi_returns_expected_constant() {
    let mut sheet = TestSheet::new();
    assert_number(&sheet.eval("=PI()"), std::f64::consts::PI);
}

#[test]
fn trig_functions_match_excel_semantics() {
    let mut sheet = TestSheet::new();

    assert_number(&sheet.eval("=SIN(0)"), 0.0);
    assert_number(&sheet.eval("=COS(0)"), 1.0);
    assert_number(&sheet.eval("=TAN(0)"), 0.0);

    assert_number(&sheet.eval("=ASIN(1)"), std::f64::consts::FRAC_PI_2);
    assert_number(&sheet.eval("=ACOS(1)"), 0.0);
    assert_number(&sheet.eval("=ATAN(1)"), std::f64::consts::FRAC_PI_4);

    assert_eq!(sheet.eval("=ASIN(2)"), Value::Error(ErrorKind::Num));
    assert_eq!(sheet.eval("=ACOS(-2)"), Value::Error(ErrorKind::Num));

    // Excel's ATAN2 argument order is (x, y).
    assert_number(&sheet.eval("=ATAN2(1,0)"), 0.0);
    assert_number(&sheet.eval("=ATAN2(0,1)"), std::f64::consts::FRAC_PI_2);
    assert_number(&sheet.eval("=ATAN2(-1,0)"), std::f64::consts::PI);
    assert_number(&sheet.eval("=ATAN2(-1,-0)"), std::f64::consts::PI);
    assert_eq!(sheet.eval("=ATAN2(0,0)"), Value::Error(ErrorKind::Div0));

    // Elementwise spilling.
    sheet.set_formula("A1", "=SIN({0;PI()/2})");
    sheet.recalculate();
    assert_number(&sheet.get("A1"), 0.0);
    assert_number(&sheet.get("A2"), 1.0);
}

#[test]
fn elementwise_math_spills_for_array_inputs() {
    let mut sheet = TestSheet::new();
    sheet.set_formula("A1", "=LN({1;EXP(1)})");
    sheet.recalculate();

    assert_number(&sheet.get("A1"), 0.0);
    assert_number(&sheet.get("A2"), 1.0);
}

#[test]
fn ceiling_and_floor_variants_match_excel_semantics() {
    let mut sheet = TestSheet::new();

    assert_number(&sheet.eval("=CEILING(4.3,2)"), 6.0);
    assert_number(&sheet.eval("=FLOOR(4.3,2)"), 4.0);
    assert_number(&sheet.eval("=CEILING(-4.3,-2)"), -4.0);
    assert_number(&sheet.eval("=FLOOR(-4.3,-2)"), -6.0);
    assert_eq!(sheet.eval("=CEILING(-4.3,2)"), Value::Error(ErrorKind::Num));
    assert_eq!(sheet.eval("=FLOOR(-4.3,2)"), Value::Error(ErrorKind::Num));

    assert_number(&sheet.eval("=CEILING.MATH(-5.5,2)"), -4.0);
    assert_number(&sheet.eval("=CEILING.MATH(-5.5,2,1)"), -6.0);
    assert_number(&sheet.eval("=FLOOR.MATH(-5.5,2)"), -6.0);
    assert_number(&sheet.eval("=FLOOR.MATH(-5.5,2,1)"), -4.0);

    assert_number(&sheet.eval("=CEILING.PRECISE(-4.3)"), -4.0);
    assert_number(&sheet.eval("=FLOOR.PRECISE(-4.3)"), -5.0);
    assert_number(&sheet.eval("=ISO.CEILING(-4.3,-2)"), -4.0);
}

#[test]
fn ceiling_and_floor_spill_elementwise_for_array_inputs() {
    let mut sheet = TestSheet::new();
    sheet.set_formula("A1", "=CEILING({4.3;5.1},2)");
    sheet.recalculate();
    assert_number(&sheet.get("A1"), 6.0);
    assert_number(&sheet.get("A2"), 6.0);

    sheet.set_formula("B1", "=FLOOR({4.3;5.1},2)");
    sheet.recalculate();
    assert_number(&sheet.get("B1"), 4.0);
    assert_number(&sheet.get("B2"), 4.0);
}

#[test]
fn criteria_aggregates_support_ranges_and_arrays() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", 1.0);
    sheet.set("A2", 2.0);
    sheet.set("A3", 3.0);
    sheet.set("A4", 4.0);
    sheet.set("B1", 10.0);
    sheet.set("B2", 20.0);
    sheet.set("B3", 30.0);
    sheet.set("B4", 40.0);

    assert_number(&sheet.eval("=SUMIF(A1:A4,\">2\",B1:B4)"), 70.0);
    assert_number(&sheet.eval("=AVERAGEIF(A1:A4,\">2\",B1:B4)"), 35.0);
    assert_number(&sheet.eval("=MAXIFS(B1:B4,A1:A4,\">2\")"), 40.0);
    assert_number(&sheet.eval("=MINIFS(B1:B4,A1:A4,\">2\")"), 30.0);
    assert_number(&sheet.eval("=MAXIFS(B1:B4,A1:A4,\">10\")"), 0.0);
    assert_number(&sheet.eval("=MINIFS(B1:B4,A1:A4,\">10\")"), 0.0);

    // Array-literal args.
    assert_number(&sheet.eval("=SUMIF({1,2,3,4},\">2\",{10,20,30,40})"), 70.0);
    assert_number(&sheet.eval("=MAXIFS({10,20,30,40},{1,2,3,4},\">2\")"), 40.0);
    assert_number(&sheet.eval("=MINIFS({10,20,30,40},{1,2,3,4},\">2\")"), 30.0);
    assert_number(
        &sheet.eval("=SUMIFS({10,20,30,40},{\"A\",\"A\",\"B\",\"B\"},\"A\",{1,2,3,4},\">1\")"),
        20.0,
    );
    assert_number(
        &sheet.eval("=COUNTIFS({\"A\",\"A\",\"B\",\"B\"},\"A\",{1,2,3,4},\">1\")"),
        1.0,
    );

    assert_number(
        &sheet.eval("=AVERAGEIFS({10,20,30,40},{\"A\",\"A\",\"B\",\"B\"},\"A\",{1,2,3,4},\">1\")"),
        20.0,
    );

    // Wildcards + blank criteria.
    sheet.set("C1", Value::from("apple"));
    sheet.set("C2", Value::from("banana"));
    sheet.set("C3", Value::from("apricot"));
    sheet.set("C4", Value::Blank);
    sheet.set("C5", Value::Text(String::new()));
    sheet.set("D1", 1.0);
    sheet.set("D2", 2.0);
    sheet.set("D3", 3.0);
    sheet.set("D4", 4.0);
    sheet.set("D5", 5.0);

    assert_number(&sheet.eval("=SUMIF(C1:C5,\"ap*\",D1:D5)"), 4.0);
    assert_number(&sheet.eval("=SUMIF(C1:C5,\"\",D1:D5)"), 9.0);
    assert_number(&sheet.eval("=COUNTIFS(C1:C5,\"ap*\")"), 2.0);
}

#[test]
fn countif_counts_implicit_blanks() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", 1.0);
    // A2 is missing (implicit blank).
    sheet.set("A3", Value::Text(String::new()));

    assert_number(&sheet.eval("=COUNTIF(A1:A3,\"\")"), 2.0);
}

#[test]
fn criteria_aggregates_respect_workbook_date_system() {
    let mut sheet = TestSheet::new();
    sheet.set_date_system(ExcelDateSystem::Excel1904);
    sheet.set("A1", 0.0); // 1904-01-01 in the 1904 system.
    sheet.set("A2", 1.0);
    sheet.set("B1", 10.0);
    sheet.set("B2", 20.0);

    assert_number(&sheet.eval("=COUNTIFS(A1:A2,\"1904-01-01\")"), 1.0);
    assert_number(&sheet.eval("=SUMIF(A1:A2,\"1904-01-01\",B1:B2)"), 10.0);
    assert_number(&sheet.eval("=AVERAGEIF(A1:A2,\"1904-01-01\",B1:B2)"), 10.0);
}

#[test]
fn criteria_aggregates_respect_value_locale_for_numeric_and_date_criteria() {
    let mut sheet = TestSheet::new();
    sheet.set_value_locale(ValueLocaleConfig::de_de());

    sheet.set("A1", 1.0);
    sheet.set("A2", 1.5);
    sheet.set("A3", 2.0);
    assert_number(&sheet.eval("=SUMIF(A1:A3,\">1,5\")"), 2.0);

    sheet.set_formula("B1", "=DATE(2020,2,1)");
    sheet.set_formula("B2", "=DATE(2020,1,2)");
    sheet.recalculate();
    assert_number(&sheet.eval("=COUNTIF(B1:B2,\"1.2.2020\")"), 1.0);
}

#[test]
fn subtotal_and_aggregate_cover_common_subtypes() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", 1.0);
    sheet.set("A2", 2.0);
    sheet.set("A3", 3.0);
    sheet.set("A4", Value::from("x"));
    sheet.set("A5", Value::Blank);

    assert_number(&sheet.eval("=SUBTOTAL(9,A1:A5)"), 6.0);
    assert_number(&sheet.eval("=SUBTOTAL(1,A1:A5)"), 2.0);
    assert_number(&sheet.eval("=SUBTOTAL(2,A1:A5)"), 3.0);
    assert_number(&sheet.eval("=SUBTOTAL(3,A1:A5)"), 4.0);
    assert_number(&sheet.eval("=SUBTOTAL(109,A1:A3)"), 6.0);

    sheet.set("E1", 1.0);
    sheet.set("E2", Value::Error(ErrorKind::Div0));
    sheet.set("E3", 2.0);
    assert_number(&sheet.eval("=AGGREGATE(9,2,E1:E3)"), 3.0);
    assert_eq!(
        sheet.eval("=AGGREGATE(9,4,E1:E3)"),
        Value::Error(ErrorKind::Div0)
    );
}
