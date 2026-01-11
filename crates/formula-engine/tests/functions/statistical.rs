use formula_engine::{ErrorKind, Value};

use super::harness::{assert_number, TestSheet};

#[test]
fn stdev_s_matches_known_value() {
    let mut sheet = TestSheet::new();
    assert_number(&sheet.eval("=STDEV.S({1,2,3})"), 1.0);
}

#[test]
fn var_p_all_equal_is_zero() {
    let mut sheet = TestSheet::new();
    assert_number(&sheet.eval("=VAR.P({2,2,2})"), 0.0);
}

#[test]
fn median_selects_middle_value() {
    let mut sheet = TestSheet::new();
    assert_number(&sheet.eval("=MEDIAN({1,100,2})"), 2.0);
}

#[test]
fn median_returns_num_when_no_numeric_values() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", Value::Text("x".to_string()));
    sheet.set("A2", Value::Blank);
    assert_eq!(sheet.eval("=MEDIAN(A1:A2)"), Value::Error(ErrorKind::Num));
}

#[test]
fn mode_sngl_returns_most_frequent_value() {
    let mut sheet = TestSheet::new();
    assert_number(&sheet.eval("=MODE.SNGL({1,2,2,3})"), 2.0);
}

#[test]
fn mode_sngl_returns_na_when_no_duplicates() {
    let mut sheet = TestSheet::new();
    assert_eq!(sheet.eval("=MODE.SNGL({1,2,3})"), Value::Error(ErrorKind::NA));
}

#[test]
fn large_small_return_expected_order_stats() {
    let mut sheet = TestSheet::new();
    assert_number(&sheet.eval("=LARGE({1,5,3},2)"), 3.0);
    assert_number(&sheet.eval("=SMALL({1,5,3},2)"), 3.0);
}

#[test]
fn large_returns_num_for_invalid_k() {
    let mut sheet = TestSheet::new();
    assert_eq!(sheet.eval("=LARGE({1,2,3},0)"), Value::Error(ErrorKind::Num));
    assert_eq!(sheet.eval("=SMALL({1,2,3},4)"), Value::Error(ErrorKind::Num));
}

#[test]
fn rank_eq_defaults_to_descending_order() {
    let mut sheet = TestSheet::new();
    assert_number(&sheet.eval("=RANK.EQ(3,{1,3,5})"), 2.0);
}

#[test]
fn rank_returns_na_when_ref_has_no_numeric_values() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", Value::Text("x".to_string()));
    assert_eq!(sheet.eval("=RANK.EQ(1,A1:A1)"), Value::Error(ErrorKind::NA));
}

#[test]
fn percentile_inc_interpolates_between_points() {
    let mut sheet = TestSheet::new();
    assert_number(&sheet.eval("=PERCENTILE.INC({0,10},0.5)"), 5.0);
}

#[test]
fn correl_matches_perfect_positive_relationship() {
    let mut sheet = TestSheet::new();
    assert_number(&sheet.eval("=CORREL({1,2,3},{1,2,3})"), 1.0);
}

#[test]
fn var_s_ignores_text_and_logicals_in_references() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", 1.0);
    sheet.set("A2", Value::Text("2".to_string()));
    sheet.set("A3", true);

    // In references, text/bools are ignored, leaving a single numeric value.
    assert_eq!(sheet.eval("=VAR.S(A1:A3)"), Value::Error(ErrorKind::Div0));

    // As direct scalar arguments, numeric text/bools are coerced.
    assert_number(&sheet.eval(r#"=VAR.S(1,"2",TRUE)"#), 1.0 / 3.0);
}

