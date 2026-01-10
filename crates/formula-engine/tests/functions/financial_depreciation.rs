use formula_engine::functions::financial::{ddb, sln, syd};

fn assert_close(actual: f64, expected: f64, tol: f64) {
    assert!(
        (actual - expected).abs() <= tol,
        "expected {expected}, got {actual}"
    );
}

#[test]
fn sln_basic() {
    let dep = sln(30_000.0, 7_500.0, 5.0).unwrap();
    assert_close(dep, 4_500.0, 1e-12);
}

#[test]
fn syd_basic() {
    // Excel: SYD(30000, 7500, 5, 1) = 7500
    let dep = syd(30_000.0, 7_500.0, 5.0, 1.0).unwrap();
    assert_close(dep, 7_500.0, 1e-12);
}

#[test]
fn ddb_basic() {
    // Example from Excel docs: DDB(2400, 300, 10, 1, 2) = 480
    let dep = ddb(2_400.0, 300.0, 10.0, 1.0, Some(2.0)).unwrap();
    assert_close(dep, 480.0, 1e-12);
}
