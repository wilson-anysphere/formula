use formula_engine::{ErrorKind, Value};

use super::harness::{assert_number, TestSheet};

#[test]
fn stdev_s_matches_known_value() {
    let mut sheet = TestSheet::new();
    assert_number(&sheet.eval("=STDEV.S({1,2,3})"), 1.0);
}

#[test]
fn statistical_functions_reject_lambda_values() {
    let mut sheet = TestSheet::new();
    assert_eq!(
        sheet.eval("=STDEV.S({1,LAMBDA(x,x)})"),
        Value::Error(ErrorKind::Value)
    );
    assert_eq!(
        sheet.eval("=VARA({1,LAMBDA(x,x)})"),
        Value::Error(ErrorKind::Value)
    );
    assert_eq!(
        sheet.eval("=CORREL({1,2},{LAMBDA(x,x),2})"),
        Value::Error(ErrorKind::Value)
    );
}

#[test]
fn legacy_stat_functions_are_accepted_as_aliases() {
    let mut sheet = TestSheet::new();
    assert_number(&sheet.eval("=STDEV({1,2,3})"), 1.0);
    assert_number(&sheet.eval("=VAR({1,2,3})"), 1.0);
    assert_number(&sheet.eval("=MODE({1,2,2,3})"), 2.0);
    assert_number(&sheet.eval("=PERCENTILE({0,10},0.5)"), 5.0);
    assert_number(&sheet.eval("=QUARTILE({0,10},2)"), 5.0);
    assert_number(&sheet.eval("=RANK(3,{1,3,5})"), 2.0);
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
    assert_eq!(
        sheet.eval("=MODE.SNGL({1,2,3})"),
        Value::Error(ErrorKind::NA)
    );
}

#[test]
fn mode_mult_spills_multiple_modes() {
    let mut sheet = TestSheet::new();
    sheet.set_formula("Z1", "=MODE.MULT({1,1,2,2,3})");
    sheet.recalc();

    assert_eq!(sheet.get("Z1"), Value::Number(1.0));
    assert_eq!(sheet.get("Z2"), Value::Number(2.0));
    assert_eq!(sheet.get("Z3"), Value::Blank);
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
    assert_eq!(
        sheet.eval("=LARGE({1,2,3},0)"),
        Value::Error(ErrorKind::Num)
    );
    assert_eq!(
        sheet.eval("=SMALL({1,2,3},4)"),
        Value::Error(ErrorKind::Num)
    );
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
fn percentile_exc_errors_outside_open_interval() {
    let mut sheet = TestSheet::new();
    assert_eq!(
        sheet.eval("=PERCENTILE.EXC({0,10},0)"),
        Value::Error(ErrorKind::Num)
    );
    assert_eq!(
        sheet.eval("=PERCENTILE.EXC({0,10},1)"),
        Value::Error(ErrorKind::Num)
    );
}

#[test]
fn percentrank_inc_and_exc_match_expected_ranks() {
    let mut sheet = TestSheet::new();

    // Default significance rounds to 3 decimal places.
    assert_number(&sheet.eval("=PERCENTRANK.INC({1,2,3,4},2)"), 0.333);
    assert_number(&sheet.eval("=PERCENTRANK({1,2,3,4},2)"), 0.333);
    assert_number(&sheet.eval("=PERCENTRANK.INC({1,2,3,4},2,2)"), 0.33);

    assert_number(&sheet.eval("=PERCENTRANK.EXC({1,2,3,4},2)"), 0.4);
}

#[test]
fn percentrank_errors_on_out_of_range_x_and_invalid_significance() {
    let mut sheet = TestSheet::new();
    assert_eq!(
        sheet.eval("=PERCENTRANK.INC({1,2,3},0)"),
        Value::Error(ErrorKind::NA)
    );
    assert_eq!(
        sheet.eval("=PERCENTRANK.EXC({1,2,3},0)"),
        Value::Error(ErrorKind::NA)
    );

    assert_eq!(
        sheet.eval("=PERCENTRANK.INC({1,2,3},1,0)"),
        Value::Error(ErrorKind::Num)
    );

    assert_eq!(
        sheet.eval("=PERCENTRANK.INC({1},1)"),
        Value::Error(ErrorKind::Div0)
    );
}

#[test]
fn standardize_returns_expected_z_score() {
    let mut sheet = TestSheet::new();
    assert_number(&sheet.eval("=STANDARDIZE(4,2,2)"), 1.0);
}

#[test]
fn standardize_rejects_non_positive_standard_dev() {
    let mut sheet = TestSheet::new();
    assert_eq!(
        sheet.eval("=STANDARDIZE(1,1,0)"),
        Value::Error(ErrorKind::Num)
    );
}

#[test]
fn correl_matches_perfect_positive_relationship() {
    let mut sheet = TestSheet::new();
    assert_number(&sheet.eval("=CORREL({1,2,3},{1,2,3})"), 1.0);
}

#[test]
fn pearson_is_alias_of_correl() {
    let mut sheet = TestSheet::new();
    assert_number(&sheet.eval("=PEARSON({1,2,3},{1,2,3})"), 1.0);
}

#[test]
fn rsq_slope_and_intercept_match_simple_regression() {
    let mut sheet = TestSheet::new();
    assert_number(&sheet.eval("=RSQ({1,2,3},{1,2,3})"), 1.0);
    assert_number(&sheet.eval("=SLOPE({1,2,3},{1,2,3})"), 1.0);
    assert_number(&sheet.eval("=INTERCEPT({1,2,3},{1,2,3})"), 0.0);
}

#[test]
fn steyx_is_zero_for_perfect_fit() {
    let mut sheet = TestSheet::new();
    assert_number(&sheet.eval("=STEYX({1,2,3},{1,2,3})"), 0.0);
}

#[test]
fn steyx_requires_at_least_three_points() {
    let mut sheet = TestSheet::new();
    assert_eq!(
        sheet.eval("=STEYX({1,2},{1,2})"),
        Value::Error(ErrorKind::Div0)
    );
    assert_eq!(
        sheet.eval("=STEYX({1,2,3},{1,2})"),
        Value::Error(ErrorKind::NA)
    );
}

#[test]
fn forecast_linear_matches_identity_relationship() {
    let mut sheet = TestSheet::new();
    assert_number(&sheet.eval("=FORECAST(4,{1,2,3},{1,2,3})"), 4.0);
    assert_number(&sheet.eval("=FORECAST.LINEAR(4,{1,2,3},{1,2,3})"), 4.0);
    assert_number(
        &sheet.eval("=_xlfn.FORECAST.LINEAR(4,{1,2,3},{1,2,3})"),
        4.0,
    );
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

    // The `*A` variants treat text/bools as 0/1 and include them in the sample size.
    assert_number(&sheet.eval("=VARA(A1:A3)"), 1.0 / 3.0);
}

#[test]
fn vara_and_stdevpa_include_text_and_blanks() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", 1.0);
    sheet.set("A2", true);
    sheet.set("A3", Value::Text("x".to_string()));
    sheet.set("A4", Value::Blank);

    assert_number(&sheet.eval("=VARA(A1:A4)"), 1.0 / 3.0);
    assert_number(&sheet.eval("=VARPA(A1:A4)"), 0.25);
    assert_number(&sheet.eval("=STDEVA(A1:A4)"), (1.0_f64 / 3.0).sqrt());
    assert_number(&sheet.eval("=STDEVPA(A1:A4)"), 0.5);
}

#[test]
fn vara_treats_text_values_as_zero_even_when_numeric() {
    let mut sheet = TestSheet::new();
    assert_number(&sheet.eval(r#"=VARA("2",2)"#), 2.0);
}

#[test]
fn averagea_maxa_and_mina_include_text_and_blanks() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", 1.0);
    sheet.set("A2", Value::Text("x".to_string()));
    sheet.set("A3", true);
    sheet.set("A4", Value::Blank);

    assert_number(&sheet.eval("=AVERAGEA(A1:A4)"), 0.5);
    assert_number(&sheet.eval("=MAXA(A1:A4)"), 1.0);
    assert_number(&sheet.eval("=MINA(A1:A4)"), 0.0);
}

#[test]
fn averagea_treats_text_args_as_zero() {
    let mut sheet = TestSheet::new();
    assert_number(&sheet.eval(r#"=AVERAGEA("2",2)"#), 1.0);
}

#[test]
fn sumsq_devsq_and_avedev_match_known_values() {
    let mut sheet = TestSheet::new();
    assert_number(&sheet.eval("=SUMSQ({1,2,3})"), 14.0);
    assert_number(&sheet.eval("=DEVSQ({1,2,3})"), 2.0);
    assert_number(&sheet.eval("=AVEDEV({1,2,3})"), 2.0 / 3.0);
}

#[test]
fn geomean_and_harmean_match_known_values() {
    let mut sheet = TestSheet::new();
    assert_number(&sheet.eval("=GEOMEAN({1,4})"), 2.0);
    assert_number(&sheet.eval("=HARMEAN({1,2,4})"), 12.0 / 7.0);
}

#[test]
fn geomean_and_harmean_reject_non_positive_values() {
    let mut sheet = TestSheet::new();
    assert_eq!(sheet.eval("=GEOMEAN({0,1})"), Value::Error(ErrorKind::Num));
    assert_eq!(sheet.eval("=HARMEAN({-1,1})"), Value::Error(ErrorKind::Num));
}

#[test]
fn geomean_returns_div0_when_no_numeric_values_in_reference() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", Value::Text("x".to_string()));
    assert_eq!(sheet.eval("=GEOMEAN(A1)"), Value::Error(ErrorKind::Div0));
}

#[test]
fn devsq_returns_div0_when_no_numeric_values_in_reference() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", Value::Text("x".to_string()));
    assert_eq!(sheet.eval("=DEVSQ(A1)"), Value::Error(ErrorKind::Div0));
}

#[test]
fn trimmean_excludes_even_number_of_points_from_tails() {
    let mut sheet = TestSheet::new();
    assert_number(&sheet.eval("=TRIMMEAN({1,2,3,4,5,6,7,8,9,100},0.2)"), 5.5);
    assert_number(&sheet.eval("=TRIMMEAN({1,2,3},0)"), 2.0);
}

#[test]
fn trimmean_rejects_invalid_percent() {
    let mut sheet = TestSheet::new();
    assert_eq!(
        sheet.eval("=TRIMMEAN({1,2,3},-0.1)"),
        Value::Error(ErrorKind::Num)
    );
    assert_eq!(
        sheet.eval("=TRIMMEAN({1,2,3},1.1)"),
        Value::Error(ErrorKind::Num)
    );
}

#[test]
fn binomial_distribution_functions_match_known_values() {
    let mut sheet = TestSheet::new();
    assert_number(&sheet.eval("=BINOM.DIST(2,5,0.5,FALSE)"), 0.3125);
    assert_number(&sheet.eval("=BINOM.DIST(2,5,0.5,TRUE)"), 0.5);

    // Range form.
    assert_number(&sheet.eval("=BINOM.DIST.RANGE(5,0.5,1,3)"), 0.78125);
    assert_number(&sheet.eval("=BINOM.DIST.RANGE(5,0.5,2)"), 0.3125);

    // Inverse.
    assert_number(&sheet.eval("=BINOM.INV(10,0.5,0.5)"), 5.0);
}

#[test]
fn binomial_legacy_aliases_match_modern_functions() {
    let mut sheet = TestSheet::new();
    assert_eq!(
        sheet.eval("=BINOMDIST(2,5,0.5,FALSE)"),
        sheet.eval("=BINOM.DIST(2,5,0.5,FALSE)")
    );
    assert_eq!(
        sheet.eval("=CRITBINOM(10,0.5,0.5)"),
        sheet.eval("=BINOM.INV(10,0.5,0.5)")
    );
}

#[test]
fn poisson_distribution_matches_known_values() {
    let mut sheet = TestSheet::new();
    assert_number(&sheet.eval("=POISSON.DIST(2,3,FALSE)"), 0.22404180765538775);
    assert_number(&sheet.eval("=POISSON.DIST(2,3,TRUE)"), 0.42319008112684353);
    assert_eq!(
        sheet.eval("=POISSON(2,3,TRUE)"),
        sheet.eval("=POISSON.DIST(2,3,TRUE)")
    );
}

#[test]
fn negative_binomial_distribution_matches_known_values() {
    let mut sheet = TestSheet::new();
    assert_number(&sheet.eval("=NEGBINOM.DIST(3,2,0.5,FALSE)"), 0.125);
    assert_number(&sheet.eval("=NEGBINOM.DIST(3,2,0.5,TRUE)"), 0.8125);
    assert_eq!(
        sheet.eval("=NEGBINOMDIST(3,2,0.5)"),
        sheet.eval("=NEGBINOM.DIST(3,2,0.5,FALSE)")
    );
}

#[test]
fn hypergeometric_distribution_matches_known_values() {
    let mut sheet = TestSheet::new();
    assert_number(
        &sheet.eval("=HYPGEOM.DIST(2,5,5,10,FALSE)"),
        0.3968253968253968,
    );
    assert_number(&sheet.eval("=HYPGEOM.DIST(2,5,5,10,TRUE)"), 0.5);
    assert_eq!(
        sheet.eval("=HYPGEOMDIST(2,5,5,10)"),
        sheet.eval("=HYPGEOM.DIST(2,5,5,10,FALSE)")
    );
}

#[test]
fn prob_returns_expected_probability_mass() {
    let mut sheet = TestSheet::new();
    assert_number(&sheet.eval("=PROB({0,1,2},{0.2,0.5,0.3},1)"), 0.5);
    assert_number(&sheet.eval("=PROB({0,1,2},{0.2,0.5,0.3},0,1)"), 0.7);
}

#[test]
fn prob_errors_on_invalid_probability_ranges() {
    let mut sheet = TestSheet::new();
    assert_eq!(
        sheet.eval("=PROB({0,1},{0.1,0.2},0,1)"),
        Value::Error(ErrorKind::Num)
    );
    assert_eq!(
        sheet.eval("=PROB({0,1,2},{0.2,0.5,0.3},2,1)"),
        Value::Error(ErrorKind::Num)
    );
    assert_eq!(
        sheet.eval("=PROB({0,1},{0.5,0.5,0.0},0,1)"),
        Value::Error(ErrorKind::NA)
    );
}

#[test]
fn hypothesis_tests_return_sane_p_values_and_error_codes() {
    let mut sheet = TestSheet::new();

    // Z.TEST: one-tailed, sigma omitted (uses sample stdev).
    assert_number(&sheet.eval("=Z.TEST({1,2,3,4},2)"), 0.21928901304049997);
    assert_eq!(
        sheet.eval("=ZTEST({1,2,3,4},2)"),
        sheet.eval("=Z.TEST({1,2,3,4},2)")
    );
    assert_eq!(sheet.eval("=Z.TEST({1},1)"), Value::Error(ErrorKind::Div0));

    // T.TEST: paired, two-tailed and one-tailed.
    assert_number(&sheet.eval("=T.TEST({1,2,3},{3,2,1},2,1)"), 1.0);
    assert_number(&sheet.eval("=T.TEST({1,2,3},{3,2,1},1,1)"), 0.5);
    assert_eq!(
        sheet.eval("=TTEST({1,2,3},{3,2,1},2,1)"),
        sheet.eval("=T.TEST({1,2,3},{3,2,1},2,1)")
    );
    assert_eq!(
        sheet.eval("=T.TEST({1,2,3},{3,2,1},3,1)"),
        Value::Error(ErrorKind::Num)
    );

    // F.TEST: identical arrays => ratio 1 with equal dfs => p-value 1.
    assert_number(&sheet.eval("=F.TEST({1,2,3},{1,2,3})"), 1.0);
    assert_eq!(
        sheet.eval("=FTEST({1,2,3},{1,2,3})"),
        sheet.eval("=F.TEST({1,2,3},{1,2,3})")
    );

    // CHISQ.TEST: 2x2 contingency table.
    assert_number(
        &sheet.eval("=CHISQ.TEST({10,20;30,40},{12,18;28,42})"),
        0.372998483613487,
    );
    assert_eq!(
        sheet.eval("=CHITEST({10,20;30,40},{12,18;28,42})"),
        sheet.eval("=CHISQ.TEST({10,20;30,40},{12,18;28,42})")
    );
    assert_eq!(
        sheet.eval("=CHISQ.TEST({1,2},{1,2;3,4})"),
        Value::Error(ErrorKind::NA)
    );
    assert_eq!(
        sheet.eval("=CHISQ.TEST({1},{0})"),
        Value::Error(ErrorKind::Num)
    );
}
