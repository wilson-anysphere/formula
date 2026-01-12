use std::f64::consts::PI;

use formula_engine::{ErrorKind, Value};

use super::harness::{assert_number, TestSheet};

#[test]
fn t_distribution_matches_cauchy_known_values_and_aliases() {
    let mut sheet = TestSheet::new();

    // df=1 => Cauchy distribution.
    assert_number(&sheet.eval("=T.DIST(0,1,TRUE)"), 0.5);
    assert_number(&sheet.eval("=T.DIST(0,1,FALSE)"), 1.0 / PI);
    assert_number(&sheet.eval("=T.DIST.RT(1,1)"), 0.25);
    assert_number(&sheet.eval("=T.DIST.2T(1,1)"), 0.5);

    assert_number(&sheet.eval("=T.INV(0.75,1)"), 1.0);
    assert_number(&sheet.eval("=T.INV.2T(0.5,1)"), 1.0);

    // Legacy names.
    assert_number(&sheet.eval("=TDIST(1,1,2)"), 0.5);
    assert_number(&sheet.eval("=TINV(0.5,1)"), 1.0);

    // Array lift.
    sheet.set_formula("A1", "=T.DIST({0,1},1,TRUE)");
    sheet.recalc();
    assert_number(&sheet.get("A1"), 0.5);
    assert_number(&sheet.get("B1"), 0.75);
}

#[test]
fn t_distribution_rejects_invalid_domains() {
    let mut sheet = TestSheet::new();
    assert_eq!(
        sheet.eval("=T.DIST(0,0,TRUE)"),
        Value::Error(ErrorKind::Num)
    );
    assert_eq!(sheet.eval("=T.DIST.RT(0,1)"), Value::Error(ErrorKind::Num));
    assert_eq!(sheet.eval("=T.DIST.2T(-1,1)"), Value::Error(ErrorKind::Num));
    assert_eq!(sheet.eval("=T.INV(-0.1,1)"), Value::Error(ErrorKind::Num));
    assert_eq!(sheet.eval("=TDIST(1,1,3)"), Value::Error(ErrorKind::Num));
}

#[test]
fn chisq_distribution_matches_df2_exponential_known_values_and_aliases() {
    let mut sheet = TestSheet::new();
    let exp_neg_1 = (-1.0_f64).exp();
    let cdf = 1.0 - exp_neg_1;
    let pdf = 0.5 * exp_neg_1;

    assert_number(&sheet.eval("=CHISQ.DIST(2,2,TRUE)"), cdf);
    assert_number(&sheet.eval("=CHISQ.DIST(2,2,FALSE)"), pdf);
    assert_number(&sheet.eval("=CHISQ.DIST.RT(2,2)"), exp_neg_1);
    assert_number(&sheet.eval("=CHIDIST(2,2)"), exp_neg_1);

    assert_number(&sheet.eval(&format!("=CHISQ.INV({cdf},2)")), 2.0);
    assert_number(&sheet.eval(&format!("=CHISQ.INV.RT({exp_neg_1},2)")), 2.0);
    assert_number(&sheet.eval(&format!("=CHIINV({exp_neg_1},2)")), 2.0);

    // Array lift.
    sheet.set_formula("A1", "=CHISQ.DIST({0,2},2,TRUE)");
    sheet.recalc();
    assert_number(&sheet.get("A1"), 0.0);
    assert_number(&sheet.get("B1"), cdf);
}

#[test]
fn chisq_distribution_rejects_invalid_domains() {
    let mut sheet = TestSheet::new();
    assert_eq!(
        sheet.eval("=CHISQ.DIST(-1,2,TRUE)"),
        Value::Error(ErrorKind::Num)
    );
    assert_eq!(
        sheet.eval("=CHISQ.DIST(1,0,TRUE)"),
        Value::Error(ErrorKind::Num)
    );
    assert_eq!(
        sheet.eval("=CHISQ.INV(1.1,2)"),
        Value::Error(ErrorKind::Num)
    );
}

#[test]
fn f_distribution_matches_f11_known_values_and_aliases() {
    let mut sheet = TestSheet::new();

    // For df1=df2=1, CDF has a simple form: CDF(x) = 2/pi * atan(sqrt(x)).
    let cdf_1 = 2.0 * 1.0_f64.atan() / PI;
    let cdf_4 = 2.0 * 2.0_f64.atan() / PI;

    assert_number(&sheet.eval("=F.DIST(1,1,1,TRUE)"), cdf_1);
    assert_number(&sheet.eval("=F.DIST.RT(1,1,1)"), 1.0 - cdf_1);
    assert_number(&sheet.eval("=FDIST(1,1,1)"), 1.0 - cdf_1);

    assert_number(&sheet.eval("=F.INV(0.5,1,1)"), 1.0);
    assert_number(&sheet.eval("=F.INV.RT(0.5,1,1)"), 1.0);
    assert_number(&sheet.eval("=FINV(0.5,1,1)"), 1.0);

    // Array lift.
    sheet.set_formula("A1", "=F.DIST({1,4},1,1,TRUE)");
    sheet.recalc();
    assert_number(&sheet.get("A1"), cdf_1);
    assert_number(&sheet.get("B1"), cdf_4);
}

#[test]
fn beta_distribution_uniform_bounds_known_values_and_aliases() {
    let mut sheet = TestSheet::new();

    assert_number(&sheet.eval("=BETA.DIST(3,1,1,TRUE,2,4)"), 0.5);
    assert_number(&sheet.eval("=BETA.DIST(3,1,1,FALSE,2,4)"), 0.5);
    assert_number(&sheet.eval("=BETA.INV(0.25,1,1,2,4)"), 2.5);

    // Legacy parity.
    assert_number(&sheet.eval("=BETADIST(3,1,1,2,4)"), 0.5);
    assert_number(&sheet.eval("=BETAINV(0.25,1,1,2,4)"), 2.5);

    // Array lift.
    sheet.set_formula("A1", "=BETA.DIST({2,3,4},1,1,TRUE,2,4)");
    sheet.recalc();
    assert_number(&sheet.get("A1"), 0.0);
    assert_number(&sheet.get("B1"), 0.5);
    assert_number(&sheet.get("C1"), 1.0);
}

#[test]
fn beta_distribution_rejects_invalid_domains() {
    let mut sheet = TestSheet::new();
    assert_eq!(
        sheet.eval("=BETA.DIST(3,1,1,TRUE,2,2)"),
        Value::Error(ErrorKind::Num)
    );
    assert_eq!(
        sheet.eval("=BETA.DIST(5,1,1,TRUE,2,4)"),
        Value::Error(ErrorKind::Num)
    );
    assert_eq!(
        sheet.eval("=BETA.INV(-0.1,1,1,0,1)"),
        Value::Error(ErrorKind::Num)
    );
}

#[test]
fn gamma_distribution_matches_alpha1_exponential_known_values_and_aliases() {
    let mut sheet = TestSheet::new();
    let exp_neg_1 = (-1.0_f64).exp();
    let cdf = 1.0 - exp_neg_1;
    let pdf = 0.5 * exp_neg_1;

    assert_number(&sheet.eval("=GAMMA.DIST(2,1,2,TRUE)"), cdf);
    assert_number(&sheet.eval("=GAMMA.DIST(2,1,2,FALSE)"), pdf);
    assert_number(&sheet.eval(&format!("=GAMMA.INV({cdf},1,2)")), 2.0);

    // Legacy parity.
    assert_number(&sheet.eval("=GAMMADIST(2,1,2,TRUE)"), cdf);
    assert_number(&sheet.eval(&format!("=GAMMAINV({cdf},1,2)")), 2.0);

    // Array lift.
    sheet.set_formula("A1", "=GAMMA.DIST({0,2},1,2,TRUE)");
    sheet.recalc();
    assert_number(&sheet.get("A1"), 0.0);
    assert_number(&sheet.get("B1"), cdf);
}

#[test]
fn gamma_special_functions_match_factorial_identities() {
    let mut sheet = TestSheet::new();
    assert_number(&sheet.eval("=GAMMA(5)"), 24.0);
    assert_number(&sheet.eval("=GAMMALN(5)"), 24.0_f64.ln());
    assert_number(&sheet.eval("=GAMMALN.PRECISE(5)"), 24.0_f64.ln());

    // Array lift for GAMMA.
    sheet.set_formula("A1", "=GAMMA({2,3,4})");
    sheet.recalc();
    assert_number(&sheet.get("A1"), 1.0);
    assert_number(&sheet.get("B1"), 2.0);
    assert_number(&sheet.get("C1"), 6.0);
}

#[test]
fn gamma_special_functions_reject_invalid_domains() {
    let mut sheet = TestSheet::new();
    assert_eq!(sheet.eval("=GAMMA(0)"), Value::Error(ErrorKind::Num));
    assert_eq!(sheet.eval("=GAMMALN(0)"), Value::Error(ErrorKind::Num));
}

#[test]
fn lognormal_distribution_matches_standard_params_known_values_and_aliases() {
    let mut sheet = TestSheet::new();
    let sqrt_2pi = (2.0 * PI).sqrt();

    assert_number(&sheet.eval("=LOGNORM.DIST(1,0,1,TRUE)"), 0.5);
    assert_number(&sheet.eval("=LOGNORM.DIST(1,0,1,FALSE)"), 1.0 / sqrt_2pi);
    assert_number(&sheet.eval("=LOGNORM.INV(0.5,0,1)"), 1.0);

    // Legacy parity.
    assert_number(&sheet.eval("=LOGNORMDIST(1,0,1)"), 0.5);
    assert_number(&sheet.eval("=LOGINV(0.5,0,1)"), 1.0);

    // Array lift (x={1,e} -> ln(x)={0,1}).
    sheet.set_formula("A1", "=LOGNORM.DIST({1,2.718281828459045},0,1,TRUE)");
    sheet.recalc();
    assert_number(&sheet.get("A1"), 0.5);
    assert_number(&sheet.get("B1"), 0.8413447460685429);
}

#[test]
fn lognormal_distribution_rejects_invalid_domains() {
    let mut sheet = TestSheet::new();
    assert_eq!(
        sheet.eval("=LOGNORM.DIST(0,0,1,TRUE)"),
        Value::Error(ErrorKind::Num)
    );
    assert_eq!(
        sheet.eval("=LOGNORM.INV(0.5,0,0)"),
        Value::Error(ErrorKind::Num)
    );
}

#[test]
fn exponential_distribution_matches_known_values_and_aliases() {
    let mut sheet = TestSheet::new();
    let exp_neg_1 = (-1.0_f64).exp();
    let cdf = 1.0 - exp_neg_1;
    let pdf = 2.0 * exp_neg_1;

    assert_number(&sheet.eval("=EXPON.DIST(0.5,2,TRUE)"), cdf);
    assert_number(&sheet.eval("=EXPON.DIST(0.5,2,FALSE)"), pdf);
    assert_number(&sheet.eval("=EXPONDIST(0.5,2,TRUE)"), cdf);

    // Array lift.
    sheet.set_formula("A1", "=EXPON.DIST({0,0.5},2,TRUE)");
    sheet.recalc();
    assert_number(&sheet.get("A1"), 0.0);
    assert_number(&sheet.get("B1"), cdf);
}

#[test]
fn weibull_distribution_matches_alpha1_exponential_known_values_and_aliases() {
    let mut sheet = TestSheet::new();
    let exp_neg_1 = (-1.0_f64).exp();
    let cdf = 1.0 - exp_neg_1;
    let pdf = 0.5 * exp_neg_1;

    assert_number(&sheet.eval("=WEIBULL.DIST(2,1,2,TRUE)"), cdf);
    assert_number(&sheet.eval("=WEIBULL.DIST(2,1,2,FALSE)"), pdf);

    // Legacy name.
    assert_number(&sheet.eval("=WEIBULL(2,1,2,TRUE)"), cdf);

    // Array lift.
    sheet.set_formula("A1", "=WEIBULL.DIST({0,2},1,2,TRUE)");
    sheet.recalc();
    assert_number(&sheet.get("A1"), 0.0);
    assert_number(&sheet.get("B1"), cdf);
}

#[test]
fn fisher_transforms_match_known_values_and_domains() {
    let mut sheet = TestSheet::new();
    let fisher_0_5 = 0.5 * 3.0_f64.ln();
    assert_number(&sheet.eval("=FISHER(0.5)"), fisher_0_5);
    assert_number(&sheet.eval(&format!("=FISHERINV({fisher_0_5})")), 0.5);

    assert_eq!(sheet.eval("=FISHER(1)"), Value::Error(ErrorKind::Num));
    assert_eq!(sheet.eval("=FISHER(-1)"), Value::Error(ErrorKind::Num));

    // Array lift.
    sheet.set_formula("A1", "=FISHER({0,0.5})");
    sheet.recalc();
    assert_number(&sheet.get("A1"), 0.0);
    assert_number(&sheet.get("B1"), fisher_0_5);
}

#[test]
fn confidence_functions_match_known_values_and_aliases() {
    let mut sheet = TestSheet::new();

    // CONFIDENCE.NORM(alpha, std_dev, size) = z_{1-alpha/2} * std_dev / sqrt(size)
    let z_0_975 = 1.959_963_984_540_054_f64;
    let expected_0_05 = z_0_975 / 10.0;
    assert_number(&sheet.eval("=CONFIDENCE.NORM(0.05,1,100)"), expected_0_05);
    assert_number(&sheet.eval("=CONFIDENCE(0.05,1,100)"), expected_0_05);

    // Array lift across alpha.
    let z_0_95 = 1.644_853_626_951_472_2_f64;
    sheet.set_formula("A1", "=CONFIDENCE.NORM({0.05,0.1},1,100)");
    sheet.recalc();
    assert_number(&sheet.get("A1"), expected_0_05);
    assert_number(&sheet.get("B1"), z_0_95 / 10.0);

    // CONFIDENCE.T(alpha, std_dev, size): use size=2 => df=1 (Cauchy) so quantile is analytic.
    let t_0_975_df1 = (PI * (0.975 - 0.5)).tan();
    let expected_t = t_0_975_df1 / 2.0_f64.sqrt();
    assert_number(&sheet.eval("=CONFIDENCE.T(0.05,1,2)"), expected_t);
}

#[test]
fn confidence_functions_reject_invalid_domains() {
    let mut sheet = TestSheet::new();
    assert_eq!(
        sheet.eval("=CONFIDENCE.NORM(0,1,10)"),
        Value::Error(ErrorKind::Num)
    );
    assert_eq!(
        sheet.eval("=CONFIDENCE.T(0.05,1,1)"),
        Value::Error(ErrorKind::Num)
    );
    assert_eq!(
        sheet.eval("=CONFIDENCE.NORM(0.05,0,10)"),
        Value::Error(ErrorKind::Num)
    );
}
