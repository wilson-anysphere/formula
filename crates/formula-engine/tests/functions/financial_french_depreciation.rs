use formula_engine::date::{ymd_to_serial, ExcelDate, ExcelDateSystem};
use formula_engine::functions::financial::{amordegrec, amorlinc};
use formula_engine::{ErrorKind, Value};

use super::harness::{assert_number, TestSheet};

fn assert_close(actual: f64, expected: f64, tol: f64) {
    assert!(
        (actual - expected).abs() <= tol,
        "expected {expected}, got {actual}"
    );
}

#[test]
fn amorlinc_example_from_excel_docs() {
    // Excel docs example (parameters commonly used across AMOR* docs):
    // AMORLINC(2400, DATE(2008,8,19), DATE(2008,12,31), 300, 1, 0.15, 1) = 360
    let system = ExcelDateSystem::EXCEL_1900;
    let purchased = ymd_to_serial(ExcelDate::new(2008, 8, 19), system).unwrap();
    let first_period = ymd_to_serial(ExcelDate::new(2008, 12, 31), system).unwrap();
    let dep = amorlinc(
        2400.0,
        purchased,
        first_period,
        300.0,
        1.0,
        0.15,
        Some(1),
        system,
    )
    .unwrap();
    assert_close(dep, 360.0, 1e-12);
}

#[test]
fn amordegrec_example_from_excel_docs() {
    // Excel docs example:
    // AMORDEGRC(2400, DATE(2008,8,19), DATE(2008,12,31), 300, 1, 0.15, 1) = 776
    let system = ExcelDateSystem::EXCEL_1900;
    let purchased = ymd_to_serial(ExcelDate::new(2008, 8, 19), system).unwrap();
    let first_period = ymd_to_serial(ExcelDate::new(2008, 12, 31), system).unwrap();
    let dep = amordegrec(
        2400.0,
        purchased,
        first_period,
        300.0,
        1.0,
        0.15,
        Some(1),
        system,
    )
    .unwrap();
    assert_close(dep, 776.0, 1e-12);
}

#[test]
fn amor_functions_error_on_invalid_basis() {
    let system = ExcelDateSystem::EXCEL_1900;
    let purchased = ymd_to_serial(ExcelDate::new(2020, 1, 1), system).unwrap();
    let first_period = ymd_to_serial(ExcelDate::new(2020, 12, 31), system).unwrap();
    assert!(amorlinc(
        1000.0,
        purchased,
        first_period,
        0.0,
        0.0,
        0.1,
        Some(5),
        system
    )
    .is_err());
    assert!(amordegrec(
        1000.0,
        purchased,
        first_period,
        0.0,
        0.0,
        0.1,
        Some(5),
        system
    )
    .is_err());
}

#[test]
fn amor_functions_error_on_invalid_chronology() {
    let system = ExcelDateSystem::EXCEL_1900;
    let purchased = ymd_to_serial(ExcelDate::new(2020, 12, 31), system).unwrap();
    let first_period = ymd_to_serial(ExcelDate::new(2020, 1, 1), system).unwrap();
    assert!(matches!(
        amorlinc(
            1000.0,
            purchased,
            first_period,
            0.0,
            0.0,
            0.1,
            Some(0),
            system
        ),
        Err(formula_engine::ExcelError::Num)
    ));
}

#[test]
fn builtins_parse_date_strings() {
    let mut sheet = TestSheet::new();
    // Uses ISO date strings to exercise the DATEVALUE-style coercion in the builtins layer.
    assert_number(
        &sheet.eval("=AMORLINC(2400,\"2008-08-19\",\"2008-12-31\",300,1,0.15,1)"),
        360.0,
    );

    let v = sheet.eval("=AMORDEGRC(2400,\"2008-08-19\",\"2008-12-31\",300,1,0.15,1)");
    assert_number(&v, 776.0);
}

#[test]
fn builtins_error_invalid_basis() {
    let mut sheet = TestSheet::new();
    assert_eq!(
        sheet.eval("=AMORLINC(1000,\"2020-01-01\",\"2020-12-31\",0,0,0.1,5)"),
        Value::Error(ErrorKind::Num)
    );
}
