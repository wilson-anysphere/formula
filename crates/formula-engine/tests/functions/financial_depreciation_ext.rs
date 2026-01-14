use formula_engine::functions::financial::{db, vdb};
use formula_engine::ExcelError;

use super::harness::{assert_number, TestSheet};

fn assert_close(actual: f64, expected: f64, tol: f64) {
    assert!(
        (actual - expected).abs() <= tol,
        "expected {expected}, got {actual}"
    );
}

#[test]
fn db_examples_from_excel_docs() {
    // Example from Excel docs: DB(10000, 1000, 5, 1) = 3690
    let dep1 = db(10_000.0, 1_000.0, 5.0, 1.0, None).unwrap();
    assert_close(dep1, 3_690.0, 1e-12);

    // Example from Excel docs: DB(10000, 1000, 5, 2) = 2328.39
    let dep2 = db(10_000.0, 1_000.0, 5.0, 2.0, None).unwrap();
    assert_close(dep2, 2_328.39, 1e-12);
}

#[test]
fn db_month_proration() {
    // First period depreciation prorated by `month/12`.
    // DB(10000, 1000, 5, 1, 7) = 2152.5
    let dep = db(10_000.0, 1_000.0, 5.0, 1.0, Some(7.0)).unwrap();
    assert_close(dep, 2_152.5, 1e-12);
}

#[test]
fn db_extra_period_when_month_is_not_12() {
    // When `month` is not 12, Excel includes an extra (final) period for the remaining months.
    let dep = db(10_000.0, 1_000.0, 5.0, 6.0, Some(7.0)).unwrap();
    assert_close(dep, 191.27749950985103, 1e-9);
}

#[test]
fn db_errors() {
    assert_eq!(
        db(1_000.0, 0.0, 5.0, 1.0, Some(13.0)).unwrap_err(),
        ExcelError::Num
    );
    assert_eq!(
        db(1_000.0, 0.0, 5.0, 0.0, None).unwrap_err(),
        ExcelError::Num
    );
    assert_eq!(
        db(1_000.0, 0.0, 5.0, 6.0, None).unwrap_err(),
        ExcelError::Num
    );
    assert_eq!(
        db(-1_000.0, 0.0, 5.0, 1.0, None).unwrap_err(),
        ExcelError::Num
    );
}

#[test]
fn vdb_examples_from_excel_docs() {
    // Example from Excel docs: VDB(2400, 300, 10, 0, 1, 2, FALSE) = 480
    let dep = vdb(2_400.0, 300.0, 10.0, 0.0, 1.0, Some(2.0), Some(0.0)).unwrap();
    assert_close(dep, 480.0, 1e-12);

    // Fractional periods are prorated.
    let dep_half = vdb(2_400.0, 300.0, 10.0, 0.0, 0.5, Some(2.0), Some(0.0)).unwrap();
    assert_close(dep_half, 240.0, 1e-12);
}

#[test]
fn vdb_no_switch_changes_behavior() {
    // With salvage=0 and factor=2, the variable declining balance switches to straight-line
    // for the tail of the schedule unless no_switch=TRUE.
    let switched = vdb(2_400.0, 0.0, 10.0, 6.0, 10.0, Some(2.0), Some(0.0)).unwrap();
    assert_close(switched, 629.1456, 1e-9);

    let no_switch = vdb(2_400.0, 0.0, 10.0, 6.0, 10.0, Some(2.0), Some(1.0)).unwrap();
    assert_close(no_switch, 371.44756224, 1e-9);
}

#[test]
fn vdb_errors() {
    assert_eq!(
        vdb(2_400.0, 300.0, 10.0, -1.0, 1.0, None, None).unwrap_err(),
        ExcelError::Num
    );
    assert_eq!(
        vdb(2_400.0, 300.0, 10.0, 1.0, 1.0, None, None).unwrap_err(),
        ExcelError::Num
    );
    assert_eq!(
        vdb(2_400.0, 300.0, 10.0, 0.0, 1.0, Some(0.0), None).unwrap_err(),
        ExcelError::Num
    );
    assert_eq!(
        vdb(2_400.0, 300.0, 10.0, 0.0, 11.0, None, None).unwrap_err(),
        ExcelError::Num
    );
}

#[test]
fn builtins_db_and_vdb_wiring() {
    let mut sheet = TestSheet::new();

    assert_number(&sheet.eval("=DB(10000,1000,5,1)"), 3690.0);
    assert_number(&sheet.eval("=DB(10000,1000,5,6,7)"), 191.27749950985103);
    assert_number(&sheet.eval("=VDB(2400,0,10,6,10,2,TRUE)"), 371.44756224);
}
