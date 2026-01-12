use formula_engine::date::{ymd_to_serial, ExcelDate, ExcelDateSystem};
use formula_engine::error::ExcelError;
use formula_engine::functions::financial::{
    coupdaybs, coupdays, coupdaysnc, coupncd, coupnum, couppcd,
};

use super::harness::{assert_number, TestSheet};

#[test]
fn coup_functions_known_values_basis_0_semiannual() {
    let system = ExcelDateSystem::EXCEL_1900;
    let settlement = ymd_to_serial(ExcelDate::new(2024, 4, 1), system).unwrap();
    let maturity = ymd_to_serial(ExcelDate::new(2025, 1, 1), system).unwrap();

    let expected_pcd = ymd_to_serial(ExcelDate::new(2024, 1, 1), system).unwrap();
    let expected_ncd = ymd_to_serial(ExcelDate::new(2024, 7, 1), system).unwrap();

    assert_eq!(
        couppcd(settlement, maturity, 2, 0, system).unwrap(),
        expected_pcd
    );
    assert_eq!(
        coupncd(settlement, maturity, 2, 0, system).unwrap(),
        expected_ncd
    );

    assert_eq!(coupdaybs(settlement, maturity, 2, 0, system).unwrap(), 90.0);
    assert_eq!(
        coupdaysnc(settlement, maturity, 2, 0, system).unwrap(),
        90.0
    );
    assert_eq!(coupdays(settlement, maturity, 2, 0, system).unwrap(), 180.0);
    assert_eq!(coupnum(settlement, maturity, 2, 0, system).unwrap(), 2.0);
}

#[test]
fn coup_functions_known_values_basis_1_quarterly() {
    let system = ExcelDateSystem::EXCEL_1900;
    let settlement = ymd_to_serial(ExcelDate::new(2024, 2, 1), system).unwrap();
    let maturity = ymd_to_serial(ExcelDate::new(2024, 12, 15), system).unwrap();

    let expected_pcd = ymd_to_serial(ExcelDate::new(2023, 12, 15), system).unwrap();
    let expected_ncd = ymd_to_serial(ExcelDate::new(2024, 3, 15), system).unwrap();

    assert_eq!(
        couppcd(settlement, maturity, 4, 1, system).unwrap(),
        expected_pcd
    );
    assert_eq!(
        coupncd(settlement, maturity, 4, 1, system).unwrap(),
        expected_ncd
    );

    assert_eq!(coupdaybs(settlement, maturity, 4, 1, system).unwrap(), 48.0);
    assert_eq!(
        coupdaysnc(settlement, maturity, 4, 1, system).unwrap(),
        43.0
    );
    assert_eq!(coupdays(settlement, maturity, 4, 1, system).unwrap(), 91.0);
    assert_eq!(coupnum(settlement, maturity, 4, 1, system).unwrap(), 4.0);
}

#[test]
fn coup_days_consistency_properties() {
    let system = ExcelDateSystem::EXCEL_1900;
    let settlement = ymd_to_serial(ExcelDate::new(2024, 2, 1), system).unwrap();
    let maturity = ymd_to_serial(ExcelDate::new(2024, 12, 15), system).unwrap();
    let basis = 1;

    let days = coupdays(settlement, maturity, 4, basis, system).unwrap();
    let daybs = coupdaybs(settlement, maturity, 4, basis, system).unwrap();
    let daysnc = coupdaysnc(settlement, maturity, 4, basis, system).unwrap();
    assert_eq!(days, daybs + daysnc);

    let settlement_on_coupon = couppcd(settlement, maturity, 4, basis, system).unwrap();
    assert_eq!(
        coupdaybs(settlement_on_coupon, maturity, 4, basis, system).unwrap(),
        0.0
    );
}

#[test]
fn coup_error_cases() {
    let system = ExcelDateSystem::EXCEL_1900;
    let settlement = ymd_to_serial(ExcelDate::new(2024, 4, 1), system).unwrap();
    let maturity = ymd_to_serial(ExcelDate::new(2025, 1, 1), system).unwrap();

    assert_eq!(
        coupdaybs(settlement, maturity, 3, 0, system).unwrap_err(),
        ExcelError::Num
    );
    assert_eq!(
        coupdaybs(settlement, maturity, 2, 5, system).unwrap_err(),
        ExcelError::Num
    );
    assert_eq!(
        coupdaybs(maturity, maturity, 2, 0, system).unwrap_err(),
        ExcelError::Num
    );
    assert_eq!(
        coupdaybs(maturity, settlement, 2, 0, system).unwrap_err(),
        ExcelError::Num
    );
}

#[test]
fn builtins_support_date_strings_and_default_basis() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", "2024-04-01");
    sheet.set("A2", "2025-01-01");

    // Basis omitted -> defaults to 0.
    let v = sheet.eval("=COUPDAYBS(A1,A2,2)");
    assert_number(&v, 90.0);

    // Date-returning functions should also accept text and return serial numbers.
    let pcd = sheet.eval("=COUPPCD(A1,A2,2)");
    let expected_pcd =
        ymd_to_serial(ExcelDate::new(2024, 1, 1), ExcelDateSystem::EXCEL_1900).unwrap() as f64;
    assert_number(&pcd, expected_pcd);
}

#[test]
fn coup_functions_apply_end_of_month_schedule_when_maturity_is_month_end_basis_1() {
    let system = ExcelDateSystem::EXCEL_1900;

    // Maturity at month-end but not the 31st: Excel pins coupon dates to month-end when maturity
    // is EOM. This affects basis=1 because COUPDAYS uses the actual day-count between coupon dates.
    //
    // Quarterly schedule, maturity=2020-04-30 => PCD=2020-01-31, NCD=2020-04-30.
    let settlement = ymd_to_serial(ExcelDate::new(2020, 2, 15), system).unwrap();
    let maturity = ymd_to_serial(ExcelDate::new(2020, 4, 30), system).unwrap();
    let expected_pcd = ymd_to_serial(ExcelDate::new(2020, 1, 31), system).unwrap();
    let expected_ncd = maturity;

    assert_eq!(couppcd(settlement, maturity, 4, 1, system).unwrap(), expected_pcd);
    assert_eq!(coupncd(settlement, maturity, 4, 1, system).unwrap(), expected_ncd);
    assert_eq!(coupnum(settlement, maturity, 4, 1, system).unwrap(), 1.0);

    assert_eq!(coupdaybs(settlement, maturity, 4, 1, system).unwrap(), 15.0);
    assert_eq!(coupdaysnc(settlement, maturity, 4, 1, system).unwrap(), 75.0);
    assert_eq!(coupdays(settlement, maturity, 4, 1, system).unwrap(), 90.0);

    // Semiannual schedule, maturity=2021-02-28 => PCD=2020-08-31, NCD=2021-02-28.
    let settlement = ymd_to_serial(ExcelDate::new(2020, 11, 15), system).unwrap();
    let maturity = ymd_to_serial(ExcelDate::new(2021, 2, 28), system).unwrap();
    let expected_pcd = ymd_to_serial(ExcelDate::new(2020, 8, 31), system).unwrap();
    let expected_ncd = maturity;

    assert_eq!(couppcd(settlement, maturity, 2, 1, system).unwrap(), expected_pcd);
    assert_eq!(coupncd(settlement, maturity, 2, 1, system).unwrap(), expected_ncd);
    assert_eq!(coupnum(settlement, maturity, 2, 1, system).unwrap(), 1.0);

    assert_eq!(coupdaybs(settlement, maturity, 2, 1, system).unwrap(), 76.0);
    assert_eq!(coupdaysnc(settlement, maturity, 2, 1, system).unwrap(), 105.0);
    assert_eq!(coupdays(settlement, maturity, 2, 1, system).unwrap(), 181.0);
}

#[test]
fn coupdays_basis_4_uses_fixed_360_over_frequency_even_when_days360_is_not_180() {
    let system = ExcelDateSystem::EXCEL_1900;

    // Semiannual schedule with an end-of-month February maturity. For European 30/360 (basis=4),
    // `DAYS360(2020-08-31, 2021-02-28, TRUE) = 178`, but Excel still models the coupon period length
    // E as `360/frequency = 180` for COUP* and bond pricing helpers.
    let settlement = ymd_to_serial(ExcelDate::new(2020, 11, 15), system).unwrap();
    let maturity = ymd_to_serial(ExcelDate::new(2021, 2, 28), system).unwrap();
    let expected_pcd = ymd_to_serial(ExcelDate::new(2020, 8, 31), system).unwrap();
    let expected_ncd = maturity;

    assert_eq!(couppcd(settlement, maturity, 2, 4, system).unwrap(), expected_pcd);
    assert_eq!(coupncd(settlement, maturity, 2, 4, system).unwrap(), expected_ncd);
    assert_eq!(coupnum(settlement, maturity, 2, 4, system).unwrap(), 1.0);

    assert_eq!(coupdaybs(settlement, maturity, 2, 4, system).unwrap(), 75.0);
    assert_eq!(coupdaysnc(settlement, maturity, 2, 4, system).unwrap(), 105.0);
    assert_eq!(coupdays(settlement, maturity, 2, 4, system).unwrap(), 180.0);
}
