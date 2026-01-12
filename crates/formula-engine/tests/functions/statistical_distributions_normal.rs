use formula_engine::{ErrorKind, Value};

use super::harness::{assert_number, TestSheet};

fn as_number(v: Value) -> f64 {
    match v {
        Value::Number(n) => n,
        other => panic!("expected number, got {other:?}"),
    }
}

#[test]
fn norm_dist_pdf_and_cdf_match_known_values() {
    let mut sheet = TestSheet::new();

    // Standard normal at 0.
    assert_number(&sheet.eval("=NORM.DIST(0,0,1,FALSE)"), 0.3989422804014327);
    assert_number(&sheet.eval("=NORM.DIST(0,0,1,TRUE)"), 0.5);

    // Standard normal at 1.
    assert_number(
        &sheet.eval("=NORM.S.DIST(1,FALSE)"),
        0.24197072451914337,
    );
    assert_number(&sheet.eval("=NORM.S.DIST(1,TRUE)"), 0.8413447460685429);
}

#[test]
fn norm_inv_and_norm_s_inv_match_known_values() {
    let mut sheet = TestSheet::new();

    // Median is the mean.
    assert_number(&sheet.eval("=NORM.INV(0.5,1,2)"), 1.0);
    assert_number(&sheet.eval("=NORM.S.INV(0.5)"), 0.0);

    // Common z-score threshold.
    assert_number(&sheet.eval("=NORM.S.INV(0.975)"), 1.959963984540054);
}

#[test]
fn phi_and_gauss_match_expected_transforms() {
    let mut sheet = TestSheet::new();

    assert_number(&sheet.eval("=PHI(0)"), 0.3989422804014327);
    assert_number(&sheet.eval("=GAUSS(1)"), 0.3413447460685429);
}

#[test]
fn legacy_aliases_match_modern_names() {
    let mut sheet = TestSheet::new();

    let modern = as_number(sheet.eval("=NORM.DIST(0,0,1,TRUE)"));
    let legacy = as_number(sheet.eval("=NORMDIST(0,0,1,TRUE)"));
    assert!((modern - legacy).abs() < 1e-9);

    let modern = as_number(sheet.eval("=NORM.S.DIST(1,TRUE)"));
    let legacy = as_number(sheet.eval("=NORMSDIST(1)"));
    assert!((modern - legacy).abs() < 1e-9);

    let modern = as_number(sheet.eval("=NORM.INV(0.5,1,2)"));
    let legacy = as_number(sheet.eval("=NORMINV(0.5,1,2)"));
    assert!((modern - legacy).abs() < 1e-9);

    let modern = as_number(sheet.eval("=NORM.S.INV(0.5)"));
    let legacy = as_number(sheet.eval("=NORMSINV(0.5)"));
    assert!((modern - legacy).abs() < 1e-9);
}

#[test]
fn normal_distribution_domain_errors_match_excel() {
    let mut sheet = TestSheet::new();

    assert_eq!(
        sheet.eval("=NORM.DIST(0,0,0,TRUE)"),
        Value::Error(ErrorKind::Num)
    );
    assert_eq!(
        sheet.eval("=NORM.INV(0.5,0,0)"),
        Value::Error(ErrorKind::Num)
    );

    assert_eq!(
        sheet.eval("=NORM.S.INV(0)"),
        Value::Error(ErrorKind::Num)
    );
    assert_eq!(
        sheet.eval("=NORM.S.INV(1)"),
        Value::Error(ErrorKind::Num)
    );
}

#[test]
fn norm_s_dist_array_lift_spills() {
    let mut sheet = TestSheet::new();
    sheet.set_formula("Z1", "=NORM.S.DIST({-1,0,1},TRUE)");
    sheet.recalc();

    assert_number(&sheet.get("Z1"), 0.15865525393145707);
    assert_number(&sheet.get("AA1"), 0.5);
    assert_number(&sheet.get("AB1"), 0.8413447460685429);
}

