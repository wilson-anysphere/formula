use formula_engine::date::{ymd_to_serial, ExcelDate, ExcelDateSystem};
use formula_engine::error::ExcelError;
use formula_engine::functions::date_time;
use formula_engine::functions::financial::{oddfprice, oddfyield, oddlprice, oddlyield};
use formula_engine::locale::ValueLocaleConfig;
use formula_engine::{ErrorKind, Value};

use super::financial_bonds_invariants::eval_number_or_skip;
use super::harness::TestSheet;

fn assert_close(actual: f64, expected: f64, tol: f64) {
    assert!(
        (actual - expected).abs() <= tol,
        "expected {expected}, got {actual}"
    );
}

fn serial(year: i32, month: u8, day: u8, system: ExcelDateSystem) -> i32 {
    ymd_to_serial(ExcelDate::new(year, month, day), system).expect("valid excel serial")
}

fn cell_number_or_skip(sheet: &TestSheet, addr: &str) -> Option<f64> {
    match sheet.get(addr) {
        Value::Number(n) => Some(n),
        Value::Error(ErrorKind::Name) => None,
        other => panic!("expected number, got {other:?} from cell {addr}"),
    }
}
fn eval_value_or_skip(sheet: &mut TestSheet, formula: &str) -> Option<Value> {
    match sheet.eval(formula) {
        Value::Error(ErrorKind::Name) => None,
        other => Some(other),
    }
}

fn is_end_of_month(date: i32, system: ExcelDateSystem) -> bool {
    date_time::eomonth(date, 0, system).unwrap() == date
}

fn coupon_date_with_eom(anchor: i32, months: i32, eom: bool, system: ExcelDateSystem) -> i32 {
    if eom {
        date_time::eomonth(anchor, months, system).unwrap()
    } else {
        date_time::edate(anchor, months, system).unwrap()
    }
}

fn days_between(start: i32, end: i32, basis: i32, system: ExcelDateSystem) -> f64 {
    match basis {
        0 => date_time::days360(start, end, false, system).unwrap() as f64,
        4 => date_time::days360(start, end, true, system).unwrap() as f64,
        1 | 2 | 3 => (end - start) as f64,
        _ => panic!("invalid basis {basis}"),
    }
}

fn coupon_period_e(
    pcd: i32,
    ncd: i32,
    basis: i32,
    frequency: i32,
    _system: ExcelDateSystem,
) -> f64 {
    let freq = frequency as f64;
    // Keep in sync with the engine's odd-coupon helper:
    // `crates/formula-engine/src/functions/financial/odd_coupon.rs::coupon_period_e`.
    //
    // Note: for odd-coupon bond functions, bases with a 360-day year (0=US 30/360, 2=Actual/360,
    // 4=European 30E/360) use the same coupon-period length convention as COUPDAYS:
    // `E = 360/frequency` (fixed). For basis=4 this can differ from `DAYS360(PCD,NCD,TRUE)` for
    // some EOM schedules around February, but Excel still uses the fixed `E`.
    match basis {
        1 => (ncd - pcd) as f64,
        0 | 2 | 4 => 360.0 / freq,
        3 => 365.0 / freq,
        _ => panic!("invalid basis {basis}"),
    }
}

fn oddf_coupon_schedule(
    first_coupon: i32,
    maturity: i32,
    frequency: i32,
    system: ExcelDateSystem,
) -> Vec<i32> {
    let months_per_period = 12 / frequency;
    let eom = is_end_of_month(maturity, system);

    let mut dates_rev = Vec::new();
    for k in 0..1000 {
        let offset = -(k as i32) * months_per_period;
        let d = coupon_date_with_eom(maturity, offset, eom, system);
        if d < first_coupon {
            break;
        }
        dates_rev.push(d);
        if d == first_coupon {
            break;
        }
    }
    dates_rev.reverse();
    dates_rev
}

fn oddf_price_excel_model(
    settlement: i32,
    maturity: i32,
    issue: i32,
    first_coupon: i32,
    rate: f64,
    yld: f64,
    redemption: f64,
    frequency: i32,
    basis: i32,
    system: ExcelDateSystem,
) -> f64 {
    let freq = frequency as f64;
    // Excel's odd-coupon bond functions are priced per $100 face value, and coupon payments are
    // based on the $100 face value (not on the `redemption` amount).
    let c = 100.0 * rate / freq;

    let a = days_between(issue, settlement, basis, system);
    let dfc = days_between(issue, first_coupon, basis, system);
    let dsc = days_between(settlement, first_coupon, basis, system);

    let months_per_period = 12 / frequency;
    let coupon_dates = oddf_coupon_schedule(first_coupon, maturity, frequency, system);
    assert_eq!(
        coupon_dates[0], first_coupon,
        "schedule must start at first_coupon"
    );
    assert_eq!(
        *coupon_dates.last().unwrap(),
        maturity,
        "schedule must end at maturity"
    );

    let eom = is_end_of_month(maturity, system);
    // Keep in sync with `odd_coupon::oddf_equation`:
    // - For basis=4, Excel determines the coupon period by stepping back from the first coupon.
    // - Otherwise, derive PCD from the maturity-anchored cashflow schedule.
    let prev_coupon = if basis == 4 {
        coupon_date_with_eom(first_coupon, -months_per_period, eom, system)
    } else {
        let n = coupon_dates.len() as i32;
        coupon_date_with_eom(maturity, -(n * months_per_period), eom, system)
    };
    let e = coupon_period_e(prev_coupon, first_coupon, basis, frequency, system);

    let odd_first_coupon = c * (dfc / e);
    let accrued_interest = c * (a / e);

    let base = 1.0 + yld / freq;
    let t0 = dsc / e;

    let mut pv = 0.0;
    for (idx, date) in coupon_dates.iter().copied().enumerate() {
        let t = t0 + idx as f64;
        let amount = if date == maturity {
            if idx == 0 {
                redemption + odd_first_coupon
            } else {
                redemption + c
            }
        } else if idx == 0 {
            odd_first_coupon
        } else {
            c
        };
        pv += amount / base.powf(t);
    }

    pv - accrued_interest
}

fn oddl_price_excel_model(
    settlement: i32,
    maturity: i32,
    last_interest: i32,
    rate: f64,
    yld: f64,
    redemption: f64,
    frequency: i32,
    basis: i32,
    system: ExcelDateSystem,
) -> f64 {
    let freq = frequency as f64;
    // Excel's odd-coupon bond functions are priced per $100 face value, and coupon payments are
    // based on the $100 face value (not on the `redemption` amount).
    let c = 100.0 * rate / freq;

    let a = days_between(last_interest, settlement, basis, system);
    let dlm = days_between(last_interest, maturity, basis, system);
    let dsm = days_between(settlement, maturity, basis, system);

    let months_per_period = 12 / frequency;
    let eom = is_end_of_month(last_interest, system);
    let prev_coupon = coupon_date_with_eom(last_interest, -months_per_period, eom, system);
    let e = coupon_period_e(prev_coupon, last_interest, basis, frequency, system);

    let accrued_interest = c * (a / e);
    let odd_last_coupon = c * (dlm / e);
    let amount = redemption + odd_last_coupon;

    let base = 1.0 + yld / freq;
    let t = dsm / e;
    (amount / base.powf(t)) - accrued_interest
}

fn assert_num_error_or_skip(sheet: &mut TestSheet, formula: &str) -> bool {
    match sheet.eval(formula) {
        Value::Error(ErrorKind::Name) => false,
        Value::Error(ErrorKind::Num) => true,
        other => panic!("expected #NUM! from {formula}, got {other:?}"),
    }
}

#[test]
fn odd_coupon_boundary_date_validations_match_engine_behavior() {
    let mut sheet = TestSheet::new();

    // ODDF* boundaries
    // issue == settlement is allowed (zero accrued interest).
    let Some(pr) = eval_number_or_skip(
        &mut sheet,
        "=ODDFPRICE(DATE(2020,1,1),DATE(2021,7,1),DATE(2020,1,1),DATE(2020,7,1),0.05,0.06,100,2,0)",
    ) else {
        return;
    };
    assert!(pr.is_finite(), "expected finite ODDFPRICE, got {pr}");

    let Some(yld) = eval_number_or_skip(
        &mut sheet,
        "=ODDFYIELD(DATE(2020,1,1),DATE(2021,7,1),DATE(2020,1,1),DATE(2020,7,1),0.05,99,100,2,0)",
    ) else {
        return;
    };
    assert!(yld.is_finite(), "expected finite ODDFYIELD, got {yld}");

    // settlement == first_coupon is allowed (settlement on coupon date).
    let Some(pr) = eval_number_or_skip(
        &mut sheet,
        "=ODDFPRICE(DATE(2020,7,1),DATE(2021,7,1),DATE(2019,10,1),DATE(2020,7,1),0.05,0.06,100,2,0)",
    ) else {
        return;
    };
    assert!(pr.is_finite(), "expected finite ODDFPRICE, got {pr}");

    let Some(yld) = eval_number_or_skip(
        &mut sheet,
        "=ODDFYIELD(DATE(2020,7,1),DATE(2021,7,1),DATE(2019,10,1),DATE(2020,7,1),0.05,99,100,2,0)",
    ) else {
        return;
    };
    assert!(yld.is_finite(), "expected finite ODDFYIELD, got {yld}");

    // issue == first_coupon is invalid.
    if !assert_num_error_or_skip(
        &mut sheet,
        "=ODDFPRICE(DATE(2020,7,1),DATE(2025,1,1),DATE(2020,7,1),DATE(2020,7,1),0.05,0.06,100,2,0)",
    ) {
        return;
    }
    if !assert_num_error_or_skip(
        &mut sheet,
        "=ODDFYIELD(DATE(2020,7,1),DATE(2025,1,1),DATE(2020,7,1),DATE(2020,7,1),0.05,99,100,2,0)",
    ) {
        return;
    }

    // first_coupon > maturity
    if !assert_num_error_or_skip(&mut sheet, "=ODDFPRICE(DATE(2020,1,1),DATE(2021,7,1),DATE(2019,10,1),DATE(2021,8,1),0.05,0.06,100,2,0)") {
        return;
    }
    if !assert_num_error_or_skip(
        &mut sheet,
        "=ODDFYIELD(DATE(2020,1,1),DATE(2021,7,1),DATE(2019,10,1),DATE(2021,8,1),0.05,98,100,2,0)",
    ) {
        return;
    }

    // settlement >= maturity
    if !assert_num_error_or_skip(&mut sheet, "=ODDFPRICE(DATE(2021,8,1),DATE(2021,7,1),DATE(2019,10,1),DATE(2020,7,1),0.05,0.06,100,2,0)") {
        return;
    }
    if !assert_num_error_or_skip(
        &mut sheet,
        "=ODDFYIELD(DATE(2021,8,1),DATE(2021,7,1),DATE(2019,10,1),DATE(2020,7,1),0.05,98,100,2,0)",
    ) {
        return;
    }

    // ODDL* boundaries
    // settlement == last_interest is allowed (zero accrued interest).
    // (Also covered by `odd_coupon_date_boundaries.rs`.)
    match sheet
        .eval("=ODDLPRICE(DATE(2020,10,15),DATE(2021,3,1),DATE(2020,10,15),0.05,0.06,100,2,0)")
    {
        Value::Error(ErrorKind::Name) => return,
        Value::Number(n) => assert!(n.is_finite(), "expected finite number, got {n}"),
        other => panic!("expected number from ODDLPRICE boundary case, got {other:?}"),
    };
    match sheet.eval("=ODDLYIELD(DATE(2020,10,15),DATE(2021,3,1),DATE(2020,10,15),0.05,ODDLPRICE(DATE(2020,10,15),DATE(2021,3,1),DATE(2020,10,15),0.05,0.06,100,2,0),100,2,0)") {
        Value::Error(ErrorKind::Name) => return,
        Value::Number(n) => assert!(n.is_finite(), "expected finite number, got {n}"),
        other => panic!("expected number from ODDLYIELD boundary case, got {other:?}"),
    };

    // settlement == maturity
    if !assert_num_error_or_skip(
        &mut sheet,
        "=ODDLPRICE(DATE(2021,3,1),DATE(2021,3,1),DATE(2020,10,15),0.05,0.06,100,2,0)",
    ) {
        return;
    }
    if !assert_num_error_or_skip(
        &mut sheet,
        "=ODDLYIELD(DATE(2021,3,1),DATE(2021,3,1),DATE(2020,10,15),0.05,98,100,2,0)",
    ) {
        return;
    }

    // last_interest >= maturity
    if !assert_num_error_or_skip(
        &mut sheet,
        "=ODDLPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2021,3,1),0.05,0.06,100,2,0)",
    ) {
        return;
    }
    if !assert_num_error_or_skip(
        &mut sheet,
        "=ODDLYIELD(DATE(2020,11,11),DATE(2021,3,1),DATE(2021,3,1),0.05,98,100,2,0)",
    ) {
        return;
    }
}

#[test]
fn odd_coupon_date_serials_are_floored_like_excel() {
    let mut sheet = TestSheet::new();

    // ODDFPRICE: all date arguments accept time fractions and are floored.
    let baseline = match eval_number_or_skip(
        &mut sheet,
        "=ODDFPRICE(DATE(2020,1,15),DATE(2021,7,1),DATE(2019,10,1),DATE(2020,7,1),0.05,0.06,100,2,0)",
    ) {
        Some(v) => v,
        None => return,
    };

    let with_time = eval_number_or_skip(
        &mut sheet,
        "=ODDFPRICE(DATE(2020,1,15)+0.9,DATE(2021,7,1)+0.9,DATE(2019,10,1)+0.9,DATE(2020,7,1)+0.9,0.05,0.06,100,2,0)",
    )
    .expect("ODDFPRICE should evaluate");
    assert_close(with_time, baseline, 1e-10);

    // ODDLYIELD: date arguments accept time fractions and are floored.
    sheet.set_formula(
        "A1",
        "=ODDLPRICE(DATE(2021,2,1),DATE(2021,5,1),DATE(2021,1,1),0.05,0.06,100,2,0)",
    );
    sheet.recalc();
    let Some(_price) = cell_number_or_skip(&sheet, "A1") else {
        return;
    };

    let y_baseline = eval_number_or_skip(
        &mut sheet,
        "=ODDLYIELD(DATE(2021,2,1),DATE(2021,5,1),DATE(2021,1,1),0.05,A1,100,2,0)",
    )
    .expect("ODDLYIELD should evaluate");
    let y_with_time = eval_number_or_skip(
        &mut sheet,
        "=ODDLYIELD(DATE(2021,2,1)+0.9,DATE(2021,5,1)+0.9,DATE(2021,1,1)+0.9,0.05,A1,100,2,0)",
    )
    .expect("ODDLYIELD should evaluate with time fractions in dates");
    assert_close(y_with_time, y_baseline, 1e-10);
}

#[test]
fn odd_coupon_price_optional_basis_defaults_to_zero() {
    let mut sheet = TestSheet::new();

    // ODDFPRICE
    let baseline = match eval_number_or_skip(
        &mut sheet,
        "=ODDFPRICE(DATE(2020,1,15),DATE(2021,7,1),DATE(2019,10,1),DATE(2020,7,1),0.05,0.06,100,2,0)",
    ) {
        Some(v) => v,
        None => return,
    };

    let omitted = eval_number_or_skip(
        &mut sheet,
        "=ODDFPRICE(DATE(2020,1,15),DATE(2021,7,1),DATE(2019,10,1),DATE(2020,7,1),0.05,0.06,100,2)",
    )
    .expect("ODDFPRICE with omitted basis should evaluate");
    let blank_arg = eval_number_or_skip(
        &mut sheet,
        "=ODDFPRICE(DATE(2020,1,15),DATE(2021,7,1),DATE(2019,10,1),DATE(2020,7,1),0.05,0.06,100,2,)",
    )
    .expect("ODDFPRICE with blank basis arg should evaluate");
    let blank_cell = eval_number_or_skip(
        &mut sheet,
        "=ODDFPRICE(DATE(2020,1,15),DATE(2021,7,1),DATE(2019,10,1),DATE(2020,7,1),0.05,0.06,100,2,B1)",
    )
    .expect("ODDFPRICE with blank-cell basis should evaluate");
    assert_close(omitted, baseline, 1e-10);
    assert_close(blank_arg, baseline, 1e-10);
    assert_close(blank_cell, baseline, 1e-10);

    // ODDLPRICE
    let baseline = match eval_number_or_skip(
        &mut sheet,
        "=ODDLPRICE(DATE(2021,2,1),DATE(2021,5,1),DATE(2021,1,1),0.05,0.06,100,2,0)",
    ) {
        Some(v) => v,
        None => return,
    };
    let omitted = eval_number_or_skip(
        &mut sheet,
        "=ODDLPRICE(DATE(2021,2,1),DATE(2021,5,1),DATE(2021,1,1),0.05,0.06,100,2)",
    )
    .expect("ODDLPRICE with omitted basis should evaluate");
    let blank_arg = eval_number_or_skip(
        &mut sheet,
        "=ODDLPRICE(DATE(2021,2,1),DATE(2021,5,1),DATE(2021,1,1),0.05,0.06,100,2,)",
    )
    .expect("ODDLPRICE with blank basis arg should evaluate");
    let blank_cell = eval_number_or_skip(
        &mut sheet,
        "=ODDLPRICE(DATE(2021,2,1),DATE(2021,5,1),DATE(2021,1,1),0.05,0.06,100,2,B1)",
    )
    .expect("ODDLPRICE with blank-cell basis should evaluate");
    assert_close(omitted, baseline, 1e-10);
    assert_close(blank_arg, baseline, 1e-10);
    assert_close(blank_cell, baseline, 1e-10);
}

#[test]
fn odd_coupon_yield_optional_basis_defaults_to_zero() {
    let mut sheet = TestSheet::new();

    // ODDFYIELD
    sheet.set_formula(
        "A1",
        "=ODDFPRICE(DATE(2020,1,15),DATE(2021,7,1),DATE(2019,10,1),DATE(2020,7,1),0.05,0.06,100,2,0)",
    );
    sheet.recalc();
    let Some(_price) = cell_number_or_skip(&sheet, "A1") else {
        return;
    };

    let baseline = eval_number_or_skip(
        &mut sheet,
        "=ODDFYIELD(DATE(2020,1,15),DATE(2021,7,1),DATE(2019,10,1),DATE(2020,7,1),0.05,A1,100,2,0)",
    )
    .expect("ODDFYIELD should evaluate");
    let omitted = eval_number_or_skip(
        &mut sheet,
        "=ODDFYIELD(DATE(2020,1,15),DATE(2021,7,1),DATE(2019,10,1),DATE(2020,7,1),0.05,A1,100,2)",
    )
    .expect("ODDFYIELD with omitted basis should evaluate");
    let blank_arg = eval_number_or_skip(
        &mut sheet,
        "=ODDFYIELD(DATE(2020,1,15),DATE(2021,7,1),DATE(2019,10,1),DATE(2020,7,1),0.05,A1,100,2,)",
    )
    .expect("ODDFYIELD with blank basis arg should evaluate");
    let blank_cell = eval_number_or_skip(
        &mut sheet,
        "=ODDFYIELD(DATE(2020,1,15),DATE(2021,7,1),DATE(2019,10,1),DATE(2020,7,1),0.05,A1,100,2,B1)",
    )
    .expect("ODDFYIELD with blank-cell basis should evaluate");
    assert_close(omitted, baseline, 1e-10);
    assert_close(blank_arg, baseline, 1e-10);
    assert_close(blank_cell, baseline, 1e-10);

    // ODDLYIELD
    sheet.set_formula(
        "A2",
        "=ODDLPRICE(DATE(2021,2,1),DATE(2021,5,1),DATE(2021,1,1),0.05,0.06,100,2,0)",
    );
    sheet.recalc();
    let Some(_price) = cell_number_or_skip(&sheet, "A2") else {
        return;
    };

    let baseline = eval_number_or_skip(
        &mut sheet,
        "=ODDLYIELD(DATE(2021,2,1),DATE(2021,5,1),DATE(2021,1,1),0.05,A2,100,2,0)",
    )
    .expect("ODDLYIELD should evaluate");
    let omitted = eval_number_or_skip(
        &mut sheet,
        "=ODDLYIELD(DATE(2021,2,1),DATE(2021,5,1),DATE(2021,1,1),0.05,A2,100,2)",
    )
    .expect("ODDLYIELD with omitted basis should evaluate");
    let blank_arg = eval_number_or_skip(
        &mut sheet,
        "=ODDLYIELD(DATE(2021,2,1),DATE(2021,5,1),DATE(2021,1,1),0.05,A2,100,2,)",
    )
    .expect("ODDLYIELD with blank basis arg should evaluate");
    let blank_cell = eval_number_or_skip(
        &mut sheet,
        "=ODDLYIELD(DATE(2021,2,1),DATE(2021,5,1),DATE(2021,1,1),0.05,A2,100,2,B1)",
    )
    .expect("ODDLYIELD with blank-cell basis should evaluate");
    assert_close(omitted, baseline, 1e-10);
    assert_close(blank_arg, baseline, 1e-10);
    assert_close(blank_cell, baseline, 1e-10);
}

#[test]
fn odd_coupon_price_and_yield_handle_zero_yield() {
    let system = ExcelDateSystem::EXCEL_1900;

    // ODDF*: long odd first coupon period, then regular semiannual coupons.
    let issue = serial(2019, 10, 1, system);
    let settlement = serial(2020, 1, 1, system);
    let first_coupon = serial(2020, 7, 1, system);
    let maturity = serial(2021, 7, 1, system);
    let rate = 0.05;
    let redemption = 100.0;
    let frequency = 2;
    let basis = 0;

    // With yld=0, all discount factors are 1 and price becomes:
    // price = (sum of cashflows) - accrued_interest.
    // For this schedule/basis:
    // - regular coupon C = 100 * rate / frequency
    // - odd first coupon = 1.5 * C (issue->first_coupon is 270 days vs E=180)
    // - accrued_interest = 0.5 * C (issue->settlement is 90 days vs E=180)
    // - total cashflows = odd_first_coupon + C + (redemption + C)
    // => price = redemption + 3*C
    let c = 100.0 * rate / (frequency as f64);
    let expected_price = redemption + 3.0 * c;

    let price = oddfprice(
        settlement,
        maturity,
        issue,
        first_coupon,
        rate,
        0.0,
        redemption,
        frequency,
        basis,
        system,
    )
    .expect("oddfprice should succeed for yld=0");
    assert!(price.is_finite());
    assert_close(price, expected_price, 1e-12);

    let recovered = oddfyield(
        settlement,
        maturity,
        issue,
        first_coupon,
        rate,
        price,
        redemption,
        frequency,
        basis,
        system,
    )
    .expect("oddfyield should invert yld=0 price");
    assert_close(recovered, 0.0, 1e-7);

    // ODDL*: short odd last coupon.
    let last_interest = serial(2021, 1, 1, system);
    let settlement_last = serial(2021, 2, 1, system);
    let maturity_last = serial(2021, 5, 1, system);

    // For this setup/basis:
    // - DLM/E = 120/180 = 2/3, A/E = 30/180 = 1/6
    // - price = redemption + C*(2/3 - 1/6) = redemption + C/2
    let expected_price_last = redemption + 0.5 * c;

    let price_last = oddlprice(
        settlement_last,
        maturity_last,
        last_interest,
        rate,
        0.0,
        redemption,
        frequency,
        basis,
        system,
    )
    .expect("oddlprice should succeed for yld=0");
    assert!(price_last.is_finite());
    assert_close(price_last, expected_price_last, 1e-12);

    let recovered_last = oddlyield(
        settlement_last,
        maturity_last,
        last_interest,
        rate,
        price_last,
        redemption,
        frequency,
        basis,
        system,
    )
    .expect("oddlyield should invert yld=0 price");
    assert_close(recovered_last, 0.0, 1e-7);
}

#[test]
fn oddfprice_invalid_date_text_returns_value_error() {
    let mut sheet = TestSheet::new();
    let formula = r#"=ODDFPRICE("not a date",DATE(2021,7,1),DATE(2020,1,1),DATE(2020,7,1),0.05,0.06,100,2,0)"#;
    match sheet.eval(formula) {
        Value::Error(ErrorKind::Name) => return,
        Value::Error(ErrorKind::Value) => {}
        other => panic!("expected #VALUE! for invalid date text in ODDFPRICE, got {other:?}"),
    }
}

#[test]
fn oddlprice_invalid_date_text_returns_value_error() {
    let mut sheet = TestSheet::new();

    let invalid_maturity =
        r#"=ODDLPRICE(DATE(2020,11,11),"not a date",DATE(2020,10,15),0.0785,0.0625,100,2,0)"#;
    match sheet.eval(invalid_maturity) {
        Value::Error(ErrorKind::Name) => return,
        Value::Error(ErrorKind::Value) => {}
        other => {
            panic!("expected #VALUE! for invalid maturity date text in ODDLPRICE, got {other:?}")
        }
    }

    let invalid_settlement =
        r#"=ODDLPRICE("not a date",DATE(2021,3,1),DATE(2020,10,15),0.0785,0.0625,100,2,0)"#;
    match sheet.eval(invalid_settlement) {
        Value::Error(ErrorKind::Name) => return,
        Value::Error(ErrorKind::Value) => {}
        other => {
            panic!("expected #VALUE! for invalid settlement date text in ODDLPRICE, got {other:?}")
        }
    }
}

#[test]
fn oddfprice_accepts_locale_stable_date_text() {
    let mut sheet = TestSheet::new();

    let baseline =
        "=ODDFPRICE(DATE(2020,3,1),DATE(2023,7,1),DATE(2020,1,1),DATE(2020,7,1),0.06,0.05,100,1,0)";
    let baseline_value = match eval_number_or_skip(&mut sheet, baseline) {
        Some(v) => v,
        None => return,
    };

    // Use ISO-8601-ish year-month-day, which should be locale-stable.
    let text_settlement = r#"=ODDFPRICE("2020-03-01",DATE(2023,7,1),DATE(2020,1,1),DATE(2020,7,1),0.06,0.05,100,1,0)"#;
    let text_value = eval_number_or_skip(&mut sheet, text_settlement)
        .expect("ODDFPRICE should accept settlement supplied as ISO date text");

    assert_close(text_value, baseline_value, 1e-9);
}

#[test]
fn oddfprice_zero_coupon_rate_reduces_to_discounted_redemption() {
    let system = ExcelDateSystem::EXCEL_1900;

    // Long first coupon period: issue -> first_coupon spans 9 months, then regular semiannual.
    let issue = ymd_to_serial(ExcelDate::new(2019, 10, 1), system).unwrap();
    let settlement = ymd_to_serial(ExcelDate::new(2020, 1, 1), system).unwrap();
    let first_coupon = ymd_to_serial(ExcelDate::new(2020, 7, 1), system).unwrap();
    let maturity = ymd_to_serial(ExcelDate::new(2021, 7, 1), system).unwrap();

    let rate = 0.0;
    let yld = 0.1;
    let redemption = 100.0;
    let frequency = 2;
    let basis = 0;

    let price = oddfprice(
        settlement,
        maturity,
        issue,
        first_coupon,
        rate,
        yld,
        redemption,
        frequency,
        basis,
        system,
    )
    .unwrap();
    assert!(price.is_finite());

    // With rate=0, coupons and accrued interest are 0, so the price reduces to a discounted redemption:
    // P = redemption / (1 + yld/frequency)^(n-1 + DSC/E)
    //
    // Here (basis 0, 30/360):
    // - Coupon dates: 2020-07-01, 2021-01-01, 2021-07-01 => n = 3
    // - E = 360/frequency = 180, DSC = 180 => DSC/E = 1
    // - exponent = 3
    let y = yld / (frequency as f64);
    let expected = redemption / (1.0 + y).powi(3);
    assert_close(price, expected, 1e-12);
}

#[test]
fn oddfprice_eom_schedule_does_not_drift_off_maturity_basis1() {
    let system = ExcelDateSystem::EXCEL_1900;

    // Regression: naive EDATE chaining from `first_coupon` drifts off maturity:
    // 2020-04-30 + 3 months = 2020-07-30 (never hits 2020-07-31).
    //
    // Excel anchors the schedule on maturity and uses EOM stepping, producing:
    // 2020-04-30, 2020-07-31.
    let issue = serial(2020, 1, 15, system);
    let settlement = serial(2020, 2, 15, system);
    let first_coupon = serial(2020, 4, 30, system);
    let maturity = serial(2020, 7, 31, system);

    let rate = 0.0;
    let yld = 0.1;
    let redemption = 100.0;
    let frequency = 4;
    let basis = 1;

    let price = oddfprice(
        settlement,
        maturity,
        issue,
        first_coupon,
        rate,
        yld,
        redemption,
        frequency,
        basis,
        system,
    )
    .expect("ODDFPRICE should accept EOM schedules without drifting");
    assert!(price.is_finite());

    // With rate=0, coupons and accrued interest are 0:
    // P = redemption / (1 + yld/frequency)^((N-1) + DSC/E)
    //
    // For this schedule:
    // - Coupon dates: 2020-04-30, 2020-07-31 => N = 2
    // - prev_coupon(F) under EOM rule is 2020-01-31 => E = 90 (basis=1)
    let prev_coupon = serial(2020, 1, 31, system);
    let e = (first_coupon - prev_coupon) as f64;
    assert_eq!(e, 90.0);
    let dsc = (first_coupon - settlement) as f64;
    let n = 2.0;
    let exponent = (n - 1.0) + dsc / e;
    let y = yld / (frequency as f64);
    let expected = redemption / (1.0 + y).powf(exponent);
    assert_close(price, expected, 1e-12);
}

#[test]
fn oddfprice_eom_schedule_month_end_maturity_not_31st_basis1() {
    let system = ExcelDateSystem::EXCEL_1900;

    // Maturity is month-end but not the 31st (Apr 30). Excel treats this as an end-of-month coupon
    // schedule and pins coupon dates to month-end when stepping from maturity.
    //
    // Quarterly schedule anchored at maturity=2020-04-30 yields coupon dates:
    // 2019-10-31, 2020-01-31, 2020-04-30.
    let issue = serial(2019, 12, 15, system);
    let settlement = serial(2020, 1, 15, system);
    let first_coupon = serial(2020, 1, 31, system);
    let maturity = serial(2020, 4, 30, system);

    let rate = 0.0;
    let yld = 0.1;
    let redemption = 100.0;
    let frequency = 4;
    let basis = 1;

    let price = oddfprice(
        settlement,
        maturity,
        issue,
        first_coupon,
        rate,
        yld,
        redemption,
        frequency,
        basis,
        system,
    )
    .expect("ODDFPRICE should accept month-end schedules where maturity is not the 31st");
    assert!(price.is_finite());

    // With rate=0, coupons and accrued interest are 0:
    // P = redemption / (1 + yld/frequency)^((N-1) + DSC/E)
    //
    // Under EOM schedule:
    // - prev_coupon(F) = 2019-10-31
    // - E = 92 (basis=1)
    // - DSC = 16 (2020-01-15 -> 2020-01-31)
    // - N = 2 (2020-01-31, 2020-04-30)
    let prev_coupon = serial(2019, 10, 31, system);
    let e = (first_coupon - prev_coupon) as f64;
    assert_eq!(e, 92.0);
    let dsc = (first_coupon - settlement) as f64;
    assert_eq!(dsc, 16.0);
    let exponent = 1.0 + dsc / e;

    let y = yld / (frequency as f64);
    let expected = redemption / (1.0 + y).powf(exponent);
    assert_close(price, expected, 1e-12);
}

#[test]
fn oddfprice_eom_schedule_month_end_maturity_feb28_basis1() {
    let system = ExcelDateSystem::EXCEL_1900;

    // Maturity is Feb 28 (month-end). Excel's EOM schedule rule implies the prior semiannual coupon
    // date is Aug 31 (not Aug 28), and the regular period before that is Feb 29 (leap day).
    let issue = serial(2020, 6, 15, system);
    let settlement = serial(2020, 7, 15, system);
    let first_coupon = serial(2020, 8, 31, system);
    let maturity = serial(2021, 2, 28, system);

    let rate = 0.0;
    let yld = 0.1;
    let redemption = 100.0;
    let frequency = 2;
    let basis = 1;

    let price = oddfprice(
        settlement,
        maturity,
        issue,
        first_coupon,
        rate,
        yld,
        redemption,
        frequency,
        basis,
        system,
    )
    .expect("ODDFPRICE should accept month-end maturity schedules (Feb 28)");
    assert!(price.is_finite());

    // With rate=0, coupons and accrued interest are 0:
    // P = redemption / (1 + yld/frequency)^((N-1) + DSC/E)
    //
    // Under EOM schedule:
    // - prev_coupon(F) = 2020-02-29
    // - E = 184 (basis=1)
    // - DSC = 47 (2020-07-15 -> 2020-08-31)
    // - N = 2 (2020-08-31, 2021-02-28)
    let prev_coupon = serial(2020, 2, 29, system);
    let e = (first_coupon - prev_coupon) as f64;
    assert_eq!(e, 184.0);
    let dsc = (first_coupon - settlement) as f64;
    assert_eq!(dsc, 47.0);
    let exponent = 1.0 + dsc / e;

    let y = yld / (frequency as f64);
    let expected = redemption / (1.0 + y).powf(exponent);
    assert_close(price, expected, 1e-12);
}

#[test]
fn oddlprice_zero_coupon_rate_reduces_to_discounted_redemption() {
    let system = ExcelDateSystem::EXCEL_1900;

    // Short odd last period inside an otherwise regular semiannual schedule.
    let last_interest = ymd_to_serial(ExcelDate::new(2021, 1, 1), system).unwrap();
    let settlement = ymd_to_serial(ExcelDate::new(2021, 2, 1), system).unwrap();
    let maturity = ymd_to_serial(ExcelDate::new(2021, 5, 1), system).unwrap();

    let rate = 0.0;
    let yld = 0.1;
    let redemption = 100.0;
    let frequency = 2;
    let basis = 0;

    let price = oddlprice(
        settlement,
        maturity,
        last_interest,
        rate,
        yld,
        redemption,
        frequency,
        basis,
        system,
    )
    .unwrap();
    assert!(price.is_finite());

    // With rate=0, coupons and accrued interest are 0:
    // P = redemption / (1 + yld/frequency)^(DSC/E)
    //
    // Basis 0 (30/360), frequency=2: E=180, DSC=90 => exponent=0.5.
    let y = yld / (frequency as f64);
    let expected = redemption / (1.0 + y).powf(0.5);
    assert_close(price, expected, 1e-12);
}

#[test]
fn oddlprice_eom_prev_coupon_basis1_uses_month_end() {
    let system = ExcelDateSystem::EXCEL_1900;

    // Regression: prev coupon for 2020-04-30 on a quarterly schedule should be 2020-01-31 under
    // Excel's EOM rule. Naive EDATE stepping yields 2020-01-30, changing E for basis=1.
    let last_interest = serial(2020, 4, 30, system);
    let settlement = serial(2020, 5, 15, system);
    let maturity = serial(2020, 7, 31, system);

    let rate = 0.0;
    let yld = 0.1;
    let redemption = 100.0;
    let frequency = 4;
    let basis = 1;

    let price = oddlprice(
        settlement,
        maturity,
        last_interest,
        rate,
        yld,
        redemption,
        frequency,
        basis,
        system,
    )
    .expect("ODDLPRICE should use EOM-aware prev_coupon for basis=1");
    assert!(price.is_finite());

    // With rate=0, coupons and accrued interest are 0:
    // P = redemption / (1 + yld/frequency)^(DSM/E)
    //
    // Under EOM rule:
    // - prev_coupon(L) = 2020-01-31
    // - E = 90 days (basis=1)
    let prev_coupon = serial(2020, 1, 31, system);
    let e = (last_interest - prev_coupon) as f64;
    assert_eq!(e, 90.0);
    let dsm = (maturity - settlement) as f64;
    let exponent = dsm / e;
    let y = yld / (frequency as f64);
    let expected = redemption / (1.0 + y).powf(exponent);
    assert_close(price, expected, 1e-12);
}

#[test]
fn oddfprice_zero_coupon_basis2_uses_fixed_360_over_frequency_for_e() {
    let system = ExcelDateSystem::EXCEL_1900;

    // Semiannual schedule aligned from `first_coupon` by 6 months:
    // 2020-07-31, 2021-01-31, 2021-07-31 (maturity) => N = 3.
    let issue = serial(2020, 1, 1, system);
    let settlement = serial(2020, 2, 15, system);
    let first_coupon = serial(2020, 7, 31, system);
    let maturity = serial(2021, 7, 31, system);

    let rate = 0.0;
    let yld = 0.1;
    let redemption = 100.0;
    let frequency = 2;
    let basis = 2; // Actual/360

    let price = oddfprice(
        settlement,
        maturity,
        issue,
        first_coupon,
        rate,
        yld,
        redemption,
        frequency,
        basis,
        system,
    )
    .unwrap();
    assert!(price.is_finite());

    // With rate=0, ODDFPRICE reduces to a discounted redemption:
    // P = redemption / (1 + yld/frequency)^((N-1) + DSC/E)
    //
    // For basis=2 (Actual/360), Excel treats E as fixed 360/frequency and DSC as actual days.
    let e = 360.0 / (frequency as f64);
    let dsc = (first_coupon - settlement) as f64;
    let n = 3.0;
    let exponent = (n - 1.0) + dsc / e;

    let y = yld / (frequency as f64);
    let expected = redemption / (1.0 + y).powf(exponent);
    assert_close(price, expected, 1e-12);
}

#[test]
fn oddlprice_zero_coupon_basis2_uses_fixed_360_over_frequency_for_e() {
    let system = ExcelDateSystem::EXCEL_1900;

    let last_interest = serial(2020, 8, 31, system);
    let settlement = serial(2020, 10, 15, system);
    let maturity = serial(2021, 2, 15, system);

    let rate = 0.0;
    let yld = 0.1;
    let redemption = 100.0;
    let frequency = 2;
    let basis = 2; // Actual/360

    let price = oddlprice(
        settlement,
        maturity,
        last_interest,
        rate,
        yld,
        redemption,
        frequency,
        basis,
        system,
    )
    .unwrap();
    assert!(price.is_finite());

    // With rate=0, ODDLPRICE reduces to a discounted redemption:
    // P = redemption / (1 + yld/frequency)^(DSM/E)
    //
    // For basis=2 (Actual/360), Excel treats E as fixed 360/frequency and DSM as actual days.
    let e = 360.0 / (frequency as f64);
    let dsm = (maturity - settlement) as f64;
    let exponent = dsm / e;

    let y = yld / (frequency as f64);
    let expected = redemption / (1.0 + y).powf(exponent);
    assert_close(price, expected, 1e-12);
}

#[test]
fn oddfprice_zero_coupon_basis3_uses_fixed_365_over_frequency_for_e() {
    let system = ExcelDateSystem::EXCEL_1900;

    // Reuse the basis=2 ODDF schedule for easier comparison.
    let issue = serial(2020, 1, 1, system);
    let settlement = serial(2020, 2, 15, system);
    let first_coupon = serial(2020, 7, 31, system);
    let maturity = serial(2021, 7, 31, system);

    let rate = 0.0;
    let yld = 0.1;
    let redemption = 100.0;
    let frequency = 2;
    let basis = 3; // Actual/365

    let price = oddfprice(
        settlement,
        maturity,
        issue,
        first_coupon,
        rate,
        yld,
        redemption,
        frequency,
        basis,
        system,
    )
    .unwrap();
    assert!(price.is_finite());

    // With rate=0:
    // P = redemption / (1 + yld/frequency)^((N-1) + DSC/E)
    //
    // For basis=3 (Actual/365), Excel treats E as fixed 365/frequency and DSC as actual days.
    let e = 365.0 / (frequency as f64);
    let dsc = (first_coupon - settlement) as f64;
    let n = 3.0;
    let exponent = (n - 1.0) + dsc / e;

    let y = yld / (frequency as f64);
    let expected = redemption / (1.0 + y).powf(exponent);
    assert_close(price, expected, 1e-12);
}

#[test]
fn oddlprice_zero_coupon_basis3_uses_fixed_365_over_frequency_for_e() {
    let system = ExcelDateSystem::EXCEL_1900;

    let last_interest = serial(2020, 8, 31, system);
    let settlement = serial(2020, 10, 15, system);
    let maturity = serial(2021, 2, 15, system);

    let rate = 0.0;
    let yld = 0.1;
    let redemption = 100.0;
    let frequency = 2;
    let basis = 3; // Actual/365

    let price = oddlprice(
        settlement,
        maturity,
        last_interest,
        rate,
        yld,
        redemption,
        frequency,
        basis,
        system,
    )
    .unwrap();
    assert!(price.is_finite());

    // With rate=0:
    // P = redemption / (1 + yld/frequency)^(DSM/E)
    //
    // For basis=3 (Actual/365), Excel treats E as fixed 365/frequency and DSM as actual days.
    let e = 365.0 / (frequency as f64);
    let dsm = (maturity - settlement) as f64;
    let exponent = dsm / e;

    let y = yld / (frequency as f64);
    let expected = redemption / (1.0 + y).powf(exponent);
    assert_close(price, expected, 1e-12);
}

#[test]
fn oddfprice_basis0_vs_basis4_diverge_on_february_eom_day_count() {
    let system = ExcelDateSystem::EXCEL_1900;

    // US 30/360 (basis=0) treats "end of February" as day 30, while European 30/360 (basis=4)
    // only adjusts 31st-of-month dates. This creates a measurable difference in DSC for:
    //   settlement = 2021-02-28 (end-of-February, non-leap year)
    //   first_coupon = 2021-03-31 (31st-of-month)
    //
    // Keep rate=0 so the price depends only on discounting the redemption.
    let issue = serial(2021, 1, 31, system);
    let settlement = serial(2021, 2, 28, system);
    let first_coupon = serial(2021, 3, 31, system);
    let maturity = serial(2021, 9, 30, system); // aligns with EDATE stepping from first_coupon

    let rate = 0.0;
    let yld = 0.1;
    let redemption = 100.0;
    let frequency = 2;

    let price_us = oddfprice(
        settlement,
        maturity,
        issue,
        first_coupon,
        rate,
        yld,
        redemption,
        frequency,
        0,
        system,
    )
    .unwrap();
    let price_eu = oddfprice(
        settlement,
        maturity,
        issue,
        first_coupon,
        rate,
        yld,
        redemption,
        frequency,
        4,
        system,
    )
    .unwrap();

    assert!(price_us.is_finite());
    assert!(price_eu.is_finite());

    // With rate=0, ODDFPRICE = redemption / (1 + yld/frequency)^((N-1) + DSC/E).
    // For frequency=2, E=180 and N=2 (first_coupon and maturity).
    let y = yld / (frequency as f64);
    let e = 360.0 / (frequency as f64);
    let n = 2.0;

    // DSC under 30/360 US vs European:
    // - US (basis 0): Feb 28 (EOM) -> day 30, Mar 31 (EOM) -> day 30 => 30 days
    // - EU (basis 4): Feb 28 stays day 28, Mar 31 -> day 30 => 32 days
    let dsc_us = 30.0;
    let dsc_eu = 32.0;

    let expected_us = redemption / (1.0 + y).powf((n - 1.0) + dsc_us / e);
    let expected_eu = redemption / (1.0 + y).powf((n - 1.0) + dsc_eu / e);

    assert_close(price_us, expected_us, 1e-12);
    assert_close(price_eu, expected_eu, 1e-12);
    assert!(
        (price_us - price_eu).abs() > 1e-6,
        "expected basis=0 and basis=4 to diverge (basis0={price_us}, basis4={price_eu})"
    );
}

#[test]
fn odd_coupon_yield_inverts_zero_coupon_prices() {
    let system = ExcelDateSystem::EXCEL_1900;

    // ODDF*
    let issue = ymd_to_serial(ExcelDate::new(2019, 10, 1), system).unwrap();
    let settlement = ymd_to_serial(ExcelDate::new(2020, 1, 1), system).unwrap();
    let first_coupon = ymd_to_serial(ExcelDate::new(2020, 7, 1), system).unwrap();
    let maturity = ymd_to_serial(ExcelDate::new(2021, 7, 1), system).unwrap();

    let rate = 0.0;
    let yld = 0.1;
    let redemption = 100.0;
    let frequency = 2;

    for basis in [0, 2] {
        let price = oddfprice(
            settlement,
            maturity,
            issue,
            first_coupon,
            rate,
            yld,
            redemption,
            frequency,
            basis,
            system,
        )
        .unwrap();

        let solved = oddfyield(
            settlement,
            maturity,
            issue,
            first_coupon,
            rate,
            price,
            redemption,
            frequency,
            basis,
            system,
        )
        .unwrap();
        assert_close(solved, yld, 1e-10);
    }

    // ODDL*
    let last_interest = ymd_to_serial(ExcelDate::new(2021, 1, 1), system).unwrap();
    let settlement2 = ymd_to_serial(ExcelDate::new(2021, 2, 1), system).unwrap();
    let maturity2 = ymd_to_serial(ExcelDate::new(2021, 5, 1), system).unwrap();

    for basis in [0, 2] {
        let price2 = oddlprice(
            settlement2,
            maturity2,
            last_interest,
            rate,
            yld,
            redemption,
            frequency,
            basis,
            system,
        )
        .unwrap();

        let solved2 = oddlyield(
            settlement2,
            maturity2,
            last_interest,
            rate,
            price2,
            redemption,
            frequency,
            basis,
            system,
        )
        .unwrap();
        assert_close(solved2, yld, 1e-10);
    }
}

#[test]
fn builtins_odd_coupon_zero_coupon_rate_oracle_cases() {
    let mut sheet = TestSheet::new();

    // Pinned by current engine behavior (Excel 1900 date system); verify against real Excel via
    // tools/excel-oracle/run-excel-oracle.ps1 (Task 393).
    let price = match eval_number_or_skip(
        &mut sheet,
        "=ODDFPRICE(DATE(2020,1,1),DATE(2021,7,1),DATE(2019,10,1),DATE(2020,7,1),0,0.1,100,2,0)",
    ) {
        Some(v) => v,
        None => return,
    };
    assert_close(price, 86.3837598531476, 1e-9);

    let price2 = eval_number_or_skip(
        &mut sheet,
        "=ODDLPRICE(DATE(2021,2,1),DATE(2021,5,1),DATE(2021,1,1),0,0.1,100,2,0)",
    )
    .expect("ODDLPRICE should evaluate");
    assert_close(price2, 97.59000729485331, 1e-9);

    let yld = eval_number_or_skip(
        &mut sheet,
        &format!(
            "=ODDFYIELD(DATE(2020,1,1),DATE(2021,7,1),DATE(2019,10,1),DATE(2020,7,1),0,{price},100,2,0)"
        ),
    )
    .expect("ODDFYIELD should evaluate");
    assert_close(yld, 0.1, 1e-10);

    let yld2 = eval_number_or_skip(
        &mut sheet,
        &format!("=ODDLYIELD(DATE(2021,2,1),DATE(2021,5,1),DATE(2021,1,1),0,{price2},100,2,0)"),
    )
    .expect("ODDLYIELD should evaluate");
    assert_close(yld2, 0.1, 1e-10);
}

#[test]
fn oddfprice_matches_known_example_basis0() {
    let mut sheet = TestSheet::new();
    let formula = "=ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,0)";
    let v = sheet.eval(formula);
    match v {
        Value::Number(n) => assert_close(n, 113.59920582823823, 1e-9),
        other => panic!("expected number, got {other:?} from {formula}"),
    }
}

#[test]
fn oddfprice_eom_coupon_schedule_returns_finite_number() {
    let system = ExcelDateSystem::EXCEL_1900;
    let settlement = serial(2020, 2, 15, system);
    let maturity = serial(2020, 12, 31, system);
    let issue = serial(2020, 1, 31, system);
    let first_coupon = serial(2020, 6, 30, system);

    let price_basis0 = oddfprice(
        settlement,
        maturity,
        issue,
        first_coupon,
        0.05,
        0.04,
        100.0,
        2,
        0,
        system,
    )
    .expect("oddfprice basis=0 should succeed for end-of-month coupon schedules");
    assert!(price_basis0.is_finite());

    // Spot-check Actual/Actual (basis=1) as well, since it computes `E` from adjacent coupon dates.
    let price_basis1 = oddfprice(
        settlement,
        maturity,
        issue,
        first_coupon,
        0.05,
        0.04,
        100.0,
        2,
        1,
        system,
    )
    .expect("oddfprice basis=1 should succeed for end-of-month coupon schedules");
    assert!(price_basis1.is_finite());
}

#[test]
fn oddfyield_extreme_prices_roundtrip() {
    let system = ExcelDateSystem::EXCEL_1900;

    // Odd first coupon: issue -> first_coupon is a short stub, followed by a regular period.
    let issue = ymd_to_serial(ExcelDate::new(2020, 1, 1), system).unwrap();
    let settlement = ymd_to_serial(ExcelDate::new(2020, 1, 15), system).unwrap();
    let first_coupon = ymd_to_serial(ExcelDate::new(2020, 2, 15), system).unwrap();
    let maturity = ymd_to_serial(ExcelDate::new(2020, 8, 15), system).unwrap();

    let rate = 0.05;
    let redemption = 100.0;
    let frequency = 2;
    let basis = 0;

    // Cover both moderately extreme prices and stress cases that push yields toward the
    // domain boundary (-frequency) or into very high-yield regions.
    for pr in [0.1, 50.0, 200.0, 10_000.0] {
        let yld = oddfyield(
            settlement,
            maturity,
            issue,
            first_coupon,
            rate,
            pr,
            redemption,
            frequency,
            basis,
            system,
        )
        .expect("ODDFYIELD should converge");

        assert!(yld.is_finite(), "yield should be finite, got {yld}");
        assert!(
            yld > -(frequency as f64),
            "yield should be > -frequency, got {yld}"
        );

        let price = oddfprice(
            settlement,
            maturity,
            issue,
            first_coupon,
            rate,
            yld,
            redemption,
            frequency,
            basis,
            system,
        )
        .expect("ODDFPRICE should succeed");

        assert_close(price, pr, 1e-6);
    }
}

#[test]
fn oddlyield_extreme_prices_roundtrip() {
    let system = ExcelDateSystem::EXCEL_1900;

    // Odd last coupon: settlement is after the last interest date, with a long stub to maturity.
    let last_interest = ymd_to_serial(ExcelDate::new(2020, 6, 30), system).unwrap();
    let settlement = ymd_to_serial(ExcelDate::new(2020, 9, 15), system).unwrap();
    let maturity = ymd_to_serial(ExcelDate::new(2021, 1, 15), system).unwrap();

    let rate = 0.05;
    let redemption = 100.0;
    let frequency = 2;
    let basis = 0;

    // Cover both moderately extreme prices and stress cases that push yields toward the
    // domain boundary (-frequency) or into very high-yield regions.
    for pr in [0.1, 50.0, 200.0, 10_000.0] {
        let yld = oddlyield(
            settlement,
            maturity,
            last_interest,
            rate,
            pr,
            redemption,
            frequency,
            basis,
            system,
        )
        .expect("ODDLYIELD should converge");

        assert!(yld.is_finite(), "yield should be finite, got {yld}");
        assert!(
            yld > -(frequency as f64),
            "yield should be > -frequency, got {yld}"
        );

        let price = oddlprice(
            settlement,
            maturity,
            last_interest,
            rate,
            yld,
            redemption,
            frequency,
            basis,
            system,
        )
        .expect("ODDLPRICE should succeed");

        assert_close(price, pr, 1e-6);
    }
}

#[test]
fn odd_last_coupon_supports_settlement_before_last_interest() {
    let system = ExcelDateSystem::EXCEL_1900;

    // Semiannual schedule with an odd last stub from last_interest -> maturity.
    // Settlement is well before last_interest, so multiple regular coupons remain.
    let settlement = ymd_to_serial(ExcelDate::new(2023, 1, 15), system).unwrap();
    let last_interest = ymd_to_serial(ExcelDate::new(2024, 8, 31), system).unwrap();
    let maturity = ymd_to_serial(ExcelDate::new(2024, 11, 15), system).unwrap();

    let rate = 0.05;
    let redemption = 100.0;
    let frequency = 2;
    let basis = 1;
    let yld = 0.06;

    let price = oddlprice(
        settlement,
        maturity,
        last_interest,
        rate,
        yld,
        redemption,
        frequency,
        basis,
        system,
    )
    .expect("ODDLPRICE should accept settlement before last_interest");
    assert!(
        price.is_finite() && price > 0.0,
        "expected finite positive price, got {price}"
    );

    let recovered_yield = oddlyield(
        settlement,
        maturity,
        last_interest,
        rate,
        price,
        redemption,
        frequency,
        basis,
        system,
    )
    .expect("ODDLYIELD should converge for settlement before last_interest");
    assert_close(recovered_yield, yld, 1e-10);
}

#[test]
fn odd_last_coupon_supports_settlement_before_last_interest_for_other_bases() {
    let system = ExcelDateSystem::EXCEL_1900;

    // End-of-month schedule with last_interest at Feb 28 (EOM). For basis=4, the European
    // DAYS360 between coupon dates differs from `360/frequency` (e.g. Aug 31 -> Feb 28 is
    // 178 days, not 180), but Excel still uses a fixed `E = 360/frequency`.
    //
    // Settlement is before last_interest, so at least one regular coupon remains and the pricing
    // logic must include those regular coupons plus the final odd-stub cashflow at maturity.
    let settlement = ymd_to_serial(ExcelDate::new(2018, 9, 15), system).unwrap();
    let last_interest = ymd_to_serial(ExcelDate::new(2019, 2, 28), system).unwrap();
    let maturity = ymd_to_serial(ExcelDate::new(2019, 3, 31), system).unwrap();

    let rate = 0.05;
    let redemption = 100.0;
    let frequency = 2;
    let yld = 0.06;

    for basis in [3, 4] {
        let price = oddlprice(
            settlement,
            maturity,
            last_interest,
            rate,
            yld,
            redemption,
            frequency,
            basis,
            system,
        )
        .unwrap_or_else(|e| {
            panic!(
                "ODDLPRICE should accept settlement before last_interest for basis={basis}: {e:?}"
            )
        });
        assert!(
            price.is_finite() && price > 0.0,
            "expected finite positive price for basis={basis}, got {price}"
        );

        let recovered_yield = oddlyield(
            settlement,
            maturity,
            last_interest,
            rate,
            price,
            redemption,
            frequency,
            basis,
            system,
        )
        .unwrap_or_else(|e| panic!("ODDLYIELD should converge for basis={basis}: {e:?}"));
        assert_close(recovered_yield, yld, 1e-10);
    }
}

#[test]
fn odd_coupon_functions_coerce_frequency_like_excel() {
    let mut sheet = TestSheet::new();
    // Task: coercion edge cases for `frequency`.
    //
    // Excel coerces:
    // - numeric text: "2" -> 2
    // - TRUE/FALSE -> 1/0
    //
    // `frequency` must be one of {1,2,4}. So FALSE (0) should produce #NUM!.

    // Baseline semiannual (frequency=2) example (Task 56).
    let baseline_semiannual = "=ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,0)";
    let baseline_semiannual_value = match eval_number_or_skip(&mut sheet, baseline_semiannual) {
        Some(v) => v,
        None => return,
    };

    let semiannual_text_freq = r#"=ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,"2",0)"#;
    let semiannual_text_freq_value = eval_number_or_skip(&mut sheet, semiannual_text_freq)
        .expect("ODDFPRICE should accept frequency supplied as numeric text");
    assert_close(semiannual_text_freq_value, baseline_semiannual_value, 1e-9);

    // Annual schedule (frequency=1) example.
    let baseline_annual =
        "=ODDFPRICE(DATE(2020,3,1),DATE(2023,7,1),DATE(2020,1,1),DATE(2020,7,1),0.06,0.05,100,1,0)";
    let baseline_annual_value = eval_number_or_skip(&mut sheet, baseline_annual)
        .expect("ODDFPRICE should accept explicit annual frequency");

    let annual_true_freq = "=ODDFPRICE(DATE(2020,3,1),DATE(2023,7,1),DATE(2020,1,1),DATE(2020,7,1),0.06,0.05,100,TRUE,0)";
    let annual_true_freq_value = eval_number_or_skip(&mut sheet, annual_true_freq)
        .expect("ODDFPRICE should accept TRUE frequency (TRUE->1)");
    assert_close(annual_true_freq_value, baseline_annual_value, 1e-9);

    let annual_false_freq = "=ODDFPRICE(DATE(2020,3,1),DATE(2023,7,1),DATE(2020,1,1),DATE(2020,7,1),0.06,0.05,100,FALSE,0)";
    match sheet.eval(annual_false_freq) {
        Value::Error(ErrorKind::Name) => return,
        Value::Error(ErrorKind::Num) => {}
        other => panic!("expected #NUM! for frequency=FALSE (0), got {other:?}"),
    }

    // Blank is coerced to 0 for numeric args.
    let annual_blank_cell_freq = "=ODDFPRICE(DATE(2020,3,1),DATE(2023,7,1),DATE(2020,1,1),DATE(2020,7,1),0.06,0.05,100,A1,0)";
    match sheet.eval(annual_blank_cell_freq) {
        Value::Error(ErrorKind::Name) => return,
        Value::Error(ErrorKind::Num) => {}
        other => panic!("expected #NUM! for frequency=<blank> (0), got {other:?}"),
    }
    let annual_blank_arg_freq =
        "=ODDFPRICE(DATE(2020,3,1),DATE(2023,7,1),DATE(2020,1,1),DATE(2020,7,1),0.06,0.05,100,,0)";
    match sheet.eval(annual_blank_arg_freq) {
        Value::Error(ErrorKind::Name) => return,
        Value::Error(ErrorKind::Num) => {}
        other => panic!("expected #NUM! for frequency=<explicit blank arg> (0), got {other:?}"),
    }

    // Empty-string text coerces to 0 in numeric contexts.
    let annual_empty_string_freq = r#"=ODDFPRICE(DATE(2020,3,1),DATE(2023,7,1),DATE(2020,1,1),DATE(2020,7,1),0.06,0.05,100,"",0)"#;
    match sheet.eval(annual_empty_string_freq) {
        Value::Error(ErrorKind::Name) => return,
        Value::Error(ErrorKind::Num) => {}
        other => panic!("expected #NUM! for frequency=\"\" (0), got {other:?}"),
    }

    // Spot-check ODDLPRICE as well (odd last coupon) to ensure the coercion behavior is consistent
    // across ODDF*/ODDL*.
    let oddl_baseline_semiannual =
        "=ODDLPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,0.0625,100,2,0)";
    let oddl_baseline_semiannual_value =
        match eval_number_or_skip(&mut sheet, oddl_baseline_semiannual) {
            Some(v) => v,
            None => return,
        };
    let oddl_text_freq =
        r#"=ODDLPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,0.0625,100,"2",0)"#;
    let oddl_text_freq_value = eval_number_or_skip(&mut sheet, oddl_text_freq)
        .expect("ODDLPRICE should accept frequency supplied as numeric text");
    assert_close(oddl_text_freq_value, oddl_baseline_semiannual_value, 1e-9);

    // TRUE/FALSE should also coerce to 1/0 for ODDLPRICE.
    let oddl_baseline_annual =
        "=ODDLPRICE(DATE(2022,11,1),DATE(2023,3,1),DATE(2022,7,1),0.06,0.05,100,1,0)";
    let oddl_baseline_annual_value = eval_number_or_skip(&mut sheet, oddl_baseline_annual)
        .expect("ODDLPRICE should accept explicit annual frequency");
    let oddl_true_freq =
        "=ODDLPRICE(DATE(2022,11,1),DATE(2023,3,1),DATE(2022,7,1),0.06,0.05,100,TRUE,0)";
    let oddl_true_freq_value = eval_number_or_skip(&mut sheet, oddl_true_freq)
        .expect("ODDLPRICE should accept TRUE frequency (TRUE->1)");
    assert_close(oddl_true_freq_value, oddl_baseline_annual_value, 1e-9);

    let oddl_false_freq =
        "=ODDLPRICE(DATE(2022,11,1),DATE(2023,3,1),DATE(2022,7,1),0.06,0.05,100,FALSE,0)";
    match sheet.eval(oddl_false_freq) {
        Value::Error(ErrorKind::Name) => return,
        Value::Error(ErrorKind::Num) => {}
        other => panic!("expected #NUM! for ODDLPRICE frequency=FALSE (0), got {other:?}"),
    }

    let oddl_blank_cell_freq =
        "=ODDLPRICE(DATE(2022,11,1),DATE(2023,3,1),DATE(2022,7,1),0.06,0.05,100,A1,0)";
    match sheet.eval(oddl_blank_cell_freq) {
        Value::Error(ErrorKind::Name) => return,
        Value::Error(ErrorKind::Num) => {}
        other => panic!("expected #NUM! for ODDLPRICE frequency=<blank> (0), got {other:?}"),
    }
    let oddl_blank_arg_freq =
        "=ODDLPRICE(DATE(2022,11,1),DATE(2023,3,1),DATE(2022,7,1),0.06,0.05,100,,0)";
    match sheet.eval(oddl_blank_arg_freq) {
        Value::Error(ErrorKind::Name) => return,
        Value::Error(ErrorKind::Num) => {}
        other => {
            panic!("expected #NUM! for ODDLPRICE frequency=<explicit blank arg> (0), got {other:?}")
        }
    }

    let oddl_empty_string_freq =
        r#"=ODDLPRICE(DATE(2022,11,1),DATE(2023,3,1),DATE(2022,7,1),0.06,0.05,100,"",0)"#;
    match sheet.eval(oddl_empty_string_freq) {
        Value::Error(ErrorKind::Name) => return,
        Value::Error(ErrorKind::Num) => {}
        other => panic!("expected #NUM! for ODDLPRICE frequency=\"\" (0), got {other:?}"),
    }
}

#[test]
fn odd_coupon_functions_truncate_frequency_like_excel() {
    let mut sheet = TestSheet::new();
    // Task: Excel-like truncation/coercion for non-integer `frequency` inputs.
    //
    // Excel truncates `frequency` to an integer (towards zero) before validating membership in
    // {1, 2, 4}.

    // ODDFPRICE: 2.9 truncates to 2.
    let oddf_baseline_2 = "=ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,0)";
    let oddf_baseline_2_value = match eval_number_or_skip(&mut sheet, oddf_baseline_2) {
        Some(v) => v,
        None => return,
    };
    let oddf_2_9 =
        "=ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2.9,0)";
    let oddf_2_9_value =
        eval_number_or_skip(&mut sheet, oddf_2_9).expect("ODDFPRICE should truncate frequency");
    assert_close(oddf_2_9_value, oddf_baseline_2_value, 1e-9);
    let oddf_text_2_9 = r#"=ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,"2.9",0)"#;
    let oddf_text_2_9_value = eval_number_or_skip(&mut sheet, oddf_text_2_9)
        .expect("ODDFPRICE should truncate frequency supplied as numeric text");
    assert_close(oddf_text_2_9_value, oddf_baseline_2_value, 1e-9);

    // ODDFPRICE: 1.9 truncates to 1.
    let oddf_baseline_1 =
        "=ODDFPRICE(DATE(2020,3,1),DATE(2023,7,1),DATE(2020,1,1),DATE(2020,7,1),0.06,0.05,100,1,0)";
    let oddf_baseline_1_value = eval_number_or_skip(&mut sheet, oddf_baseline_1)
        .expect("ODDFPRICE should accept explicit annual frequency");
    let oddf_1_9 =
        "=ODDFPRICE(DATE(2020,3,1),DATE(2023,7,1),DATE(2020,1,1),DATE(2020,7,1),0.06,0.05,100,1.9,0)";
    let oddf_1_9_value =
        eval_number_or_skip(&mut sheet, oddf_1_9).expect("ODDFPRICE should truncate frequency");
    assert_close(oddf_1_9_value, oddf_baseline_1_value, 1e-9);

    // ODDFPRICE: 1.1 truncates to 1.
    let oddf_1_1 =
        "=ODDFPRICE(DATE(2020,3,1),DATE(2023,7,1),DATE(2020,1,1),DATE(2020,7,1),0.06,0.05,100,1.1,0)";
    let oddf_1_1_value =
        eval_number_or_skip(&mut sheet, oddf_1_1).expect("ODDFPRICE should truncate frequency");
    assert_close(oddf_1_1_value, oddf_baseline_1_value, 1e-9);

    // Confirm truncation (not rounding): a value just below 2 should still truncate to 1.
    let oddf_1_999999999 =
        "=ODDFPRICE(DATE(2020,3,1),DATE(2023,7,1),DATE(2020,1,1),DATE(2020,7,1),0.06,0.05,100,1.999999999,0)";
    let oddf_1_999999999_value = eval_number_or_skip(&mut sheet, oddf_1_999999999)
        .expect("ODDFPRICE should truncate frequency");
    assert_close(oddf_1_999999999_value, oddf_baseline_1_value, 1e-9);

    // ODDFPRICE: 4.1 truncates to 4.
    let oddf_baseline_4 =
        "=ODDFPRICE(DATE(2020,1,20),DATE(2021,8,15),DATE(2020,1,1),DATE(2020,2,15),0.08,0.07,100,4,0)";
    let oddf_baseline_4_value = eval_number_or_skip(&mut sheet, oddf_baseline_4)
        .expect("ODDFPRICE should accept quarterly frequency");
    let oddf_4_1 =
        "=ODDFPRICE(DATE(2020,1,20),DATE(2021,8,15),DATE(2020,1,1),DATE(2020,2,15),0.08,0.07,100,4.1,0)";
    let oddf_4_1_value =
        eval_number_or_skip(&mut sheet, oddf_4_1).expect("ODDFPRICE should truncate frequency");
    assert_close(oddf_4_1_value, oddf_baseline_4_value, 1e-9);

    // frequency=0.9 truncates to 0 and should return #NUM!.
    let oddf_0_9 =
        "=ODDFPRICE(DATE(2020,3,1),DATE(2023,7,1),DATE(2020,1,1),DATE(2020,7,1),0.06,0.05,100,0.9,0)";
    match sheet.eval(oddf_0_9) {
        Value::Error(ErrorKind::Name) => return,
        Value::Error(ErrorKind::Num) => {}
        other => panic!("expected #NUM! for ODDFPRICE frequency=0.9 (trunc->0), got {other:?}"),
    }
    let oddf_text_0_9 = r#"=ODDFPRICE(DATE(2020,3,1),DATE(2023,7,1),DATE(2020,1,1),DATE(2020,7,1),0.06,0.05,100,"0.9",0)"#;
    match sheet.eval(oddf_text_0_9) {
        Value::Error(ErrorKind::Name) => return,
        Value::Error(ErrorKind::Num) => {}
        other => {
            panic!("expected #NUM! for ODDFPRICE frequency=\"0.9\" (trunc->0), got {other:?}")
        }
    }

    // Repeat key cases for ODDLPRICE as well.
    let oddl_baseline_2 =
        "=ODDLPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,0.0625,100,2,0)";
    let oddl_baseline_2_value = match eval_number_or_skip(&mut sheet, oddl_baseline_2) {
        Some(v) => v,
        None => return,
    };
    let oddl_2_9 =
        "=ODDLPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,0.0625,100,2.9,0)";
    let oddl_2_9_value =
        eval_number_or_skip(&mut sheet, oddl_2_9).expect("ODDLPRICE should truncate frequency");
    assert_close(oddl_2_9_value, oddl_baseline_2_value, 1e-9);
    let oddl_text_2_9 =
        r#"=ODDLPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,0.0625,100,"2.9",0)"#;
    let oddl_text_2_9_value = eval_number_or_skip(&mut sheet, oddl_text_2_9)
        .expect("ODDLPRICE should truncate frequency supplied as numeric text");
    assert_close(oddl_text_2_9_value, oddl_baseline_2_value, 1e-9);

    let oddl_baseline_1 =
        "=ODDLPRICE(DATE(2022,11,1),DATE(2023,3,1),DATE(2022,7,1),0.06,0.05,100,1,0)";
    let oddl_baseline_1_value = eval_number_or_skip(&mut sheet, oddl_baseline_1)
        .expect("ODDLPRICE should accept explicit annual frequency");
    let oddl_1_9 = "=ODDLPRICE(DATE(2022,11,1),DATE(2023,3,1),DATE(2022,7,1),0.06,0.05,100,1.9,0)";
    let oddl_1_9_value =
        eval_number_or_skip(&mut sheet, oddl_1_9).expect("ODDLPRICE should truncate frequency");
    assert_close(oddl_1_9_value, oddl_baseline_1_value, 1e-9);

    let oddl_baseline_4 =
        "=ODDLPRICE(DATE(2021,7,1),DATE(2021,8,15),DATE(2021,6,15),0.08,0.07,100,4,0)";
    let oddl_baseline_4_value = eval_number_or_skip(&mut sheet, oddl_baseline_4)
        .expect("ODDLPRICE should accept quarterly frequency");
    let oddl_4_1 = "=ODDLPRICE(DATE(2021,7,1),DATE(2021,8,15),DATE(2021,6,15),0.08,0.07,100,4.1,0)";
    let oddl_4_1_value =
        eval_number_or_skip(&mut sheet, oddl_4_1).expect("ODDLPRICE should truncate frequency");
    assert_close(oddl_4_1_value, oddl_baseline_4_value, 1e-9);

    let oddl_0_9 = "=ODDLPRICE(DATE(2022,11,1),DATE(2023,3,1),DATE(2022,7,1),0.06,0.05,100,0.9,0)";
    match sheet.eval(oddl_0_9) {
        Value::Error(ErrorKind::Name) => return,
        Value::Error(ErrorKind::Num) => {}
        other => panic!("expected #NUM! for ODDLPRICE frequency=0.9 (trunc->0), got {other:?}"),
    }
    let oddl_text_0_9 =
        r#"=ODDLPRICE(DATE(2022,11,1),DATE(2023,3,1),DATE(2022,7,1),0.06,0.05,100,"0.9",0)"#;
    match sheet.eval(oddl_text_0_9) {
        Value::Error(ErrorKind::Name) => return,
        Value::Error(ErrorKind::Num) => {}
        other => {
            panic!("expected #NUM! for ODDLPRICE frequency=\"0.9\" (trunc->0), got {other:?}")
        }
    }
}

#[test]
fn odd_coupon_functions_coerce_basis_like_excel() {
    let mut sheet = TestSheet::new();
    // Task: coercion edge cases for `basis`.
    //
    // Excel coerces:
    // - TRUE/FALSE -> 1/0
    // - blank -> 0 (same as default)
    //
    // Use an ODDFPRICE example with basis=0/1 to confirm.

    let baseline_basis_0 = "=ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,0)";
    let baseline_basis_0_value = match eval_number_or_skip(&mut sheet, baseline_basis_0) {
        Some(v) => v,
        None => return,
    };

    // Omitting an optional parameter in Excel behaves the same as the default value (basis=0).
    let omitted_basis = "=ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2)";
    let omitted_basis_value = eval_number_or_skip(&mut sheet, omitted_basis)
        .expect("ODDFPRICE should accept omitted basis and default it to 0");
    assert_close(omitted_basis_value, baseline_basis_0_value, 1e-9);

    // Basis passed as a blank cell should behave like basis=0.
    // (A1 is unset/blank by default.)
    let blank_basis = "=ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,A1)";
    let blank_basis_value = eval_number_or_skip(&mut sheet, blank_basis)
        .expect("ODDFPRICE should accept blank basis and treat it as 0");
    assert_close(blank_basis_value, baseline_basis_0_value, 1e-9);

    // Passing an explicit blank argument for an optional parameter behaves like 0 in Excel.
    let blank_basis_arg = "=ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,)";
    let blank_basis_arg_value = eval_number_or_skip(&mut sheet, blank_basis_arg)
        .expect("ODDFPRICE should accept blank basis argument and treat it as 0");
    assert_close(blank_basis_arg_value, baseline_basis_0_value, 1e-9);

    let false_basis = "=ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,FALSE)";
    let false_basis_value = eval_number_or_skip(&mut sheet, false_basis)
        .expect("ODDFPRICE should accept FALSE basis (FALSE->0)");
    assert_close(false_basis_value, baseline_basis_0_value, 1e-9);

    let empty_string_basis = r#"=ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,"")"#;
    let empty_string_basis_value = eval_number_or_skip(&mut sheet, empty_string_basis)
        .expect("ODDFPRICE should accept basis=\"\" and treat it as 0");
    assert_close(empty_string_basis_value, baseline_basis_0_value, 1e-9);

    let text_basis_0 = r#"=ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,"0")"#;
    let text_basis_0_value = eval_number_or_skip(&mut sheet, text_basis_0)
        .expect("ODDFPRICE should accept basis supplied as numeric text");
    assert_close(text_basis_0_value, baseline_basis_0_value, 1e-9);

    let baseline_basis_1 = "=ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,1)";
    let baseline_basis_1_value = eval_number_or_skip(&mut sheet, baseline_basis_1)
        .expect("ODDFPRICE should accept explicit basis=1");

    let true_basis = "=ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,TRUE)";
    let true_basis_value = eval_number_or_skip(&mut sheet, true_basis)
        .expect("ODDFPRICE should accept TRUE basis (TRUE->1)");
    assert_close(true_basis_value, baseline_basis_1_value, 1e-9);

    let text_basis_1 = r#"=ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,"1")"#;
    let text_basis_1_value = eval_number_or_skip(&mut sheet, text_basis_1)
        .expect("ODDFPRICE should accept basis=1 supplied as numeric text");
    assert_close(text_basis_1_value, baseline_basis_1_value, 1e-9);

    // Spot-check ODDLPRICE for blank/boolean basis coercions.
    let oddl_baseline_basis_0 =
        "=ODDLPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,0.0625,100,2,0)";
    let oddl_baseline_basis_0_value = match eval_number_or_skip(&mut sheet, oddl_baseline_basis_0) {
        Some(v) => v,
        None => return,
    };

    let oddl_omitted_basis =
        "=ODDLPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,0.0625,100,2)";
    let oddl_omitted_basis_value = eval_number_or_skip(&mut sheet, oddl_omitted_basis)
        .expect("ODDLPRICE should accept omitted basis and default it to 0");
    assert_close(oddl_omitted_basis_value, oddl_baseline_basis_0_value, 1e-9);

    let oddl_blank_basis =
        "=ODDLPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,0.0625,100,2,A1)";
    let oddl_blank_basis_value = eval_number_or_skip(&mut sheet, oddl_blank_basis)
        .expect("ODDLPRICE should treat blank basis as 0");
    assert_close(oddl_blank_basis_value, oddl_baseline_basis_0_value, 1e-9);

    let oddl_blank_basis_arg =
        "=ODDLPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,0.0625,100,2,)";
    let oddl_blank_basis_arg_value = eval_number_or_skip(&mut sheet, oddl_blank_basis_arg)
        .expect("ODDLPRICE should treat blank basis argument as 0");
    assert_close(
        oddl_blank_basis_arg_value,
        oddl_baseline_basis_0_value,
        1e-9,
    );

    let oddl_false_basis =
        "=ODDLPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,0.0625,100,2,FALSE)";
    let oddl_false_basis_value = eval_number_or_skip(&mut sheet, oddl_false_basis)
        .expect("ODDLPRICE should accept FALSE basis (FALSE->0)");
    assert_close(oddl_false_basis_value, oddl_baseline_basis_0_value, 1e-9);

    let oddl_empty_string_basis =
        r#"=ODDLPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,0.0625,100,2,"")"#;
    let oddl_empty_string_basis_value = eval_number_or_skip(&mut sheet, oddl_empty_string_basis)
        .expect("ODDLPRICE should accept basis=\"\" and treat it as 0");
    assert_close(
        oddl_empty_string_basis_value,
        oddl_baseline_basis_0_value,
        1e-9,
    );

    let oddl_text_basis_0 =
        r#"=ODDLPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,0.0625,100,2,"0")"#;
    let oddl_text_basis_0_value = eval_number_or_skip(&mut sheet, oddl_text_basis_0)
        .expect("ODDLPRICE should accept basis supplied as numeric text");
    assert_close(oddl_text_basis_0_value, oddl_baseline_basis_0_value, 1e-9);

    let oddl_baseline_basis_1 =
        "=ODDLPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,0.0625,100,2,1)";
    let oddl_baseline_basis_1_value = eval_number_or_skip(&mut sheet, oddl_baseline_basis_1)
        .expect("ODDLPRICE should accept explicit basis=1");
    let oddl_true_basis =
        "=ODDLPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,0.0625,100,2,TRUE)";
    let oddl_true_basis_value = eval_number_or_skip(&mut sheet, oddl_true_basis)
        .expect("ODDLPRICE should accept TRUE basis (TRUE->1)");
    assert_close(oddl_true_basis_value, oddl_baseline_basis_1_value, 1e-9);

    let oddl_text_basis_1 =
        r#"=ODDLPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,0.0625,100,2,"1")"#;
    let oddl_text_basis_1_value = eval_number_or_skip(&mut sheet, oddl_text_basis_1)
        .expect("ODDLPRICE should accept basis=1 supplied as numeric text");
    assert_close(oddl_text_basis_1_value, oddl_baseline_basis_1_value, 1e-9);
}

#[test]
fn odd_coupon_yield_functions_coerce_frequency_like_excel() {
    let mut sheet = TestSheet::new();

    // Task: make sure coercions behave consistently across ODDF*/ODDL*, including the yield
    // (solver) functions.

    // Semiannual (frequency=2) odd first coupon: text "2" should match numeric 2.
    let oddf_baseline =
        "=LET(pr,ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,0),ODDFYIELD(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,pr,100,2,0))";
    let oddf_baseline_value = match eval_number_or_skip(&mut sheet, oddf_baseline) {
        Some(v) => v,
        None => return,
    };
    let oddf_text_freq = r#"=LET(pr,ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,0),ODDFYIELD(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,pr,100,"2",0))"#;
    let oddf_text_freq_value = eval_number_or_skip(&mut sheet, oddf_text_freq)
        .expect("ODDFYIELD should accept frequency supplied as numeric text");
    assert_close(oddf_text_freq_value, oddf_baseline_value, 1e-9);

    // Annual (frequency=1): TRUE->1, FALSE->0 => #NUM!, blank->0 => #NUM!.
    let oddf_baseline_annual =
        "=LET(pr,ODDFPRICE(DATE(2020,3,1),DATE(2023,7,1),DATE(2020,1,1),DATE(2020,7,1),0.06,0.05,100,1,0),ODDFYIELD(DATE(2020,3,1),DATE(2023,7,1),DATE(2020,1,1),DATE(2020,7,1),0.06,pr,100,1,0))";
    let oddf_baseline_annual_value = eval_number_or_skip(&mut sheet, oddf_baseline_annual)
        .expect("ODDFYIELD should accept explicit annual frequency");
    let oddf_true_freq =
        "=LET(pr,ODDFPRICE(DATE(2020,3,1),DATE(2023,7,1),DATE(2020,1,1),DATE(2020,7,1),0.06,0.05,100,1,0),ODDFYIELD(DATE(2020,3,1),DATE(2023,7,1),DATE(2020,1,1),DATE(2020,7,1),0.06,pr,100,TRUE,0))";
    let oddf_true_freq_value = eval_number_or_skip(&mut sheet, oddf_true_freq)
        .expect("ODDFYIELD should accept TRUE frequency (TRUE->1)");
    assert_close(oddf_true_freq_value, oddf_baseline_annual_value, 1e-9);

    let oddf_false_freq =
        "=LET(pr,ODDFPRICE(DATE(2020,3,1),DATE(2023,7,1),DATE(2020,1,1),DATE(2020,7,1),0.06,0.05,100,1,0),ODDFYIELD(DATE(2020,3,1),DATE(2023,7,1),DATE(2020,1,1),DATE(2020,7,1),0.06,pr,100,FALSE,0))";
    match sheet.eval(oddf_false_freq) {
        Value::Error(ErrorKind::Name) => return,
        Value::Error(ErrorKind::Num) => {}
        other => panic!("expected #NUM! for ODDFYIELD frequency=FALSE (0), got {other:?}"),
    }
    let oddf_blank_cell_freq =
        "=LET(pr,ODDFPRICE(DATE(2020,3,1),DATE(2023,7,1),DATE(2020,1,1),DATE(2020,7,1),0.06,0.05,100,1,0),ODDFYIELD(DATE(2020,3,1),DATE(2023,7,1),DATE(2020,1,1),DATE(2020,7,1),0.06,pr,100,A1,0))";
    match sheet.eval(oddf_blank_cell_freq) {
        Value::Error(ErrorKind::Name) => return,
        Value::Error(ErrorKind::Num) => {}
        other => panic!("expected #NUM! for ODDFYIELD frequency=<blank> (0), got {other:?}"),
    }
    let oddf_blank_arg_freq =
        "=LET(pr,ODDFPRICE(DATE(2020,3,1),DATE(2023,7,1),DATE(2020,1,1),DATE(2020,7,1),0.06,0.05,100,1,0),ODDFYIELD(DATE(2020,3,1),DATE(2023,7,1),DATE(2020,1,1),DATE(2020,7,1),0.06,pr,100,,0))";
    match sheet.eval(oddf_blank_arg_freq) {
        Value::Error(ErrorKind::Name) => return,
        Value::Error(ErrorKind::Num) => {}
        other => {
            panic!("expected #NUM! for ODDFYIELD frequency=<explicit blank arg> (0), got {other:?}")
        }
    }
    let oddf_empty_string_freq = r#"=LET(pr,ODDFPRICE(DATE(2020,3,1),DATE(2023,7,1),DATE(2020,1,1),DATE(2020,7,1),0.06,0.05,100,1,0),ODDFYIELD(DATE(2020,3,1),DATE(2023,7,1),DATE(2020,1,1),DATE(2020,7,1),0.06,pr,100,"",0))"#;
    match sheet.eval(oddf_empty_string_freq) {
        Value::Error(ErrorKind::Name) => return,
        Value::Error(ErrorKind::Num) => {}
        other => panic!("expected #NUM! for ODDFYIELD frequency=\"\" (0), got {other:?}"),
    }

    // Repeat the same checks for odd last coupon yield.
    let oddl_baseline =
        "=LET(pr,ODDLPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,0.0625,100,2,0),ODDLYIELD(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,pr,100,2,0))";
    let oddl_baseline_value =
        eval_number_or_skip(&mut sheet, oddl_baseline).expect("ODDLYIELD baseline should evaluate");
    let oddl_text_freq = r#"=LET(pr,ODDLPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,0.0625,100,2,0),ODDLYIELD(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,pr,100,"2",0))"#;
    let oddl_text_freq_value = eval_number_or_skip(&mut sheet, oddl_text_freq)
        .expect("ODDLYIELD should accept frequency supplied as numeric text");
    assert_close(oddl_text_freq_value, oddl_baseline_value, 1e-9);

    let oddl_baseline_annual =
        "=LET(pr,ODDLPRICE(DATE(2022,11,1),DATE(2023,3,1),DATE(2022,7,1),0.06,0.05,100,1,0),ODDLYIELD(DATE(2022,11,1),DATE(2023,3,1),DATE(2022,7,1),0.06,pr,100,1,0))";
    let oddl_baseline_annual_value = eval_number_or_skip(&mut sheet, oddl_baseline_annual)
        .expect("ODDLYIELD should accept explicit annual frequency");
    let oddl_true_freq =
        "=LET(pr,ODDLPRICE(DATE(2022,11,1),DATE(2023,3,1),DATE(2022,7,1),0.06,0.05,100,1,0),ODDLYIELD(DATE(2022,11,1),DATE(2023,3,1),DATE(2022,7,1),0.06,pr,100,TRUE,0))";
    let oddl_true_freq_value = eval_number_or_skip(&mut sheet, oddl_true_freq)
        .expect("ODDLYIELD should accept TRUE frequency (TRUE->1)");
    assert_close(oddl_true_freq_value, oddl_baseline_annual_value, 1e-9);

    let oddl_false_freq =
        "=LET(pr,ODDLPRICE(DATE(2022,11,1),DATE(2023,3,1),DATE(2022,7,1),0.06,0.05,100,1,0),ODDLYIELD(DATE(2022,11,1),DATE(2023,3,1),DATE(2022,7,1),0.06,pr,100,FALSE,0))";
    match sheet.eval(oddl_false_freq) {
        Value::Error(ErrorKind::Name) => return,
        Value::Error(ErrorKind::Num) => {}
        other => panic!("expected #NUM! for ODDLYIELD frequency=FALSE (0), got {other:?}"),
    }
    let oddl_blank_cell_freq =
        "=LET(pr,ODDLPRICE(DATE(2022,11,1),DATE(2023,3,1),DATE(2022,7,1),0.06,0.05,100,1,0),ODDLYIELD(DATE(2022,11,1),DATE(2023,3,1),DATE(2022,7,1),0.06,pr,100,A1,0))";
    match sheet.eval(oddl_blank_cell_freq) {
        Value::Error(ErrorKind::Name) => return,
        Value::Error(ErrorKind::Num) => {}
        other => panic!("expected #NUM! for ODDLYIELD frequency=<blank> (0), got {other:?}"),
    }
    let oddl_blank_arg_freq =
        "=LET(pr,ODDLPRICE(DATE(2022,11,1),DATE(2023,3,1),DATE(2022,7,1),0.06,0.05,100,1,0),ODDLYIELD(DATE(2022,11,1),DATE(2023,3,1),DATE(2022,7,1),0.06,pr,100,,0))";
    match sheet.eval(oddl_blank_arg_freq) {
        Value::Error(ErrorKind::Name) => return,
        Value::Error(ErrorKind::Num) => {}
        other => {
            panic!("expected #NUM! for ODDLYIELD frequency=<explicit blank arg> (0), got {other:?}")
        }
    }
    let oddl_empty_string_freq = r#"=LET(pr,ODDLPRICE(DATE(2022,11,1),DATE(2023,3,1),DATE(2022,7,1),0.06,0.05,100,1,0),ODDLYIELD(DATE(2022,11,1),DATE(2023,3,1),DATE(2022,7,1),0.06,pr,100,"",0))"#;
    match sheet.eval(oddl_empty_string_freq) {
        Value::Error(ErrorKind::Name) => return,
        Value::Error(ErrorKind::Num) => {}
        other => panic!("expected #NUM! for ODDLYIELD frequency=\"\" (0), got {other:?}"),
    }
}

#[test]
fn odd_coupon_yield_functions_coerce_basis_like_excel() {
    let mut sheet = TestSheet::new();

    // Baseline odd first coupon yield roundtrip under basis=0.
    let oddf_baseline_basis_0 =
        "=LET(pr,ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,0),ODDFYIELD(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,pr,100,2,0))";
    let oddf_baseline_basis_0_value = match eval_number_or_skip(&mut sheet, oddf_baseline_basis_0) {
        Some(v) => v,
        None => return,
    };

    let oddf_omitted_basis =
        "=LET(pr,ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,0),ODDFYIELD(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,pr,100,2))";
    let oddf_omitted_basis_value = eval_number_or_skip(&mut sheet, oddf_omitted_basis)
        .expect("ODDFYIELD should accept omitted basis and default it to 0");
    assert_close(oddf_omitted_basis_value, oddf_baseline_basis_0_value, 1e-9);

    let oddf_blank_cell_basis =
        "=LET(pr,ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,0),ODDFYIELD(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,pr,100,2,A1))";
    let oddf_blank_cell_basis_value = eval_number_or_skip(&mut sheet, oddf_blank_cell_basis)
        .expect("ODDFYIELD should treat blank basis cell as 0");
    assert_close(
        oddf_blank_cell_basis_value,
        oddf_baseline_basis_0_value,
        1e-9,
    );

    let oddf_blank_arg_basis =
        "=LET(pr,ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,0),ODDFYIELD(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,pr,100,2,))";
    let oddf_blank_arg_basis_value = eval_number_or_skip(&mut sheet, oddf_blank_arg_basis)
        .expect("ODDFYIELD should treat blank basis argument as 0");
    assert_close(
        oddf_blank_arg_basis_value,
        oddf_baseline_basis_0_value,
        1e-9,
    );

    let oddf_false_basis =
        "=LET(pr,ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,0),ODDFYIELD(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,pr,100,2,FALSE))";
    let oddf_false_basis_value = eval_number_or_skip(&mut sheet, oddf_false_basis)
        .expect("ODDFYIELD should accept FALSE basis (FALSE->0)");
    assert_close(oddf_false_basis_value, oddf_baseline_basis_0_value, 1e-9);

    let oddf_empty_string_basis = r#"=LET(pr,ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,0),ODDFYIELD(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,pr,100,2,""))"#;
    let oddf_empty_string_basis_value = eval_number_or_skip(&mut sheet, oddf_empty_string_basis)
        .expect("ODDFYIELD should treat basis=\"\" as 0");
    assert_close(
        oddf_empty_string_basis_value,
        oddf_baseline_basis_0_value,
        1e-9,
    );

    let oddf_text_basis_0 = r#"=LET(pr,ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,0),ODDFYIELD(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,pr,100,2,"0"))"#;
    let oddf_text_basis_0_value = eval_number_or_skip(&mut sheet, oddf_text_basis_0)
        .expect("ODDFYIELD should accept basis supplied as numeric text");
    assert_close(oddf_text_basis_0_value, oddf_baseline_basis_0_value, 1e-9);

    // Basis=1 / TRUE -> 1.
    let oddf_baseline_basis_1 =
        "=LET(pr,ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,1),ODDFYIELD(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,pr,100,2,1))";
    let oddf_baseline_basis_1_value = eval_number_or_skip(&mut sheet, oddf_baseline_basis_1)
        .expect("ODDFYIELD should accept explicit basis=1");
    let oddf_true_basis =
        "=LET(pr,ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,1),ODDFYIELD(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,pr,100,2,TRUE))";
    let oddf_true_basis_value = eval_number_or_skip(&mut sheet, oddf_true_basis)
        .expect("ODDFYIELD should accept TRUE basis (TRUE->1)");
    assert_close(oddf_true_basis_value, oddf_baseline_basis_1_value, 1e-9);
    let oddf_text_basis_1 = r#"=LET(pr,ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,1),ODDFYIELD(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,pr,100,2,"1"))"#;
    let oddf_text_basis_1_value = eval_number_or_skip(&mut sheet, oddf_text_basis_1)
        .expect("ODDFYIELD should accept basis=1 supplied as numeric text");
    assert_close(oddf_text_basis_1_value, oddf_baseline_basis_1_value, 1e-9);

    // Repeat for odd last coupon yield.
    let oddl_baseline_basis_0 =
        "=LET(pr,ODDLPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,0.0625,100,2,0),ODDLYIELD(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,pr,100,2,0))";
    let oddl_baseline_basis_0_value = eval_number_or_skip(&mut sheet, oddl_baseline_basis_0)
        .expect("ODDLYIELD baseline should evaluate");

    let oddl_omitted_basis =
        "=LET(pr,ODDLPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,0.0625,100,2,0),ODDLYIELD(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,pr,100,2))";
    let oddl_omitted_basis_value = eval_number_or_skip(&mut sheet, oddl_omitted_basis)
        .expect("ODDLYIELD should accept omitted basis and default it to 0");
    assert_close(oddl_omitted_basis_value, oddl_baseline_basis_0_value, 1e-9);

    let oddl_blank_cell_basis =
        "=LET(pr,ODDLPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,0.0625,100,2,0),ODDLYIELD(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,pr,100,2,A1))";
    let oddl_blank_cell_basis_value = eval_number_or_skip(&mut sheet, oddl_blank_cell_basis)
        .expect("ODDLYIELD should treat blank basis cell as 0");
    assert_close(
        oddl_blank_cell_basis_value,
        oddl_baseline_basis_0_value,
        1e-9,
    );
    let oddl_blank_arg_basis =
        "=LET(pr,ODDLPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,0.0625,100,2,0),ODDLYIELD(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,pr,100,2,))";
    let oddl_blank_arg_basis_value = eval_number_or_skip(&mut sheet, oddl_blank_arg_basis)
        .expect("ODDLYIELD should treat blank basis argument as 0");
    assert_close(
        oddl_blank_arg_basis_value,
        oddl_baseline_basis_0_value,
        1e-9,
    );
    let oddl_false_basis =
        "=LET(pr,ODDLPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,0.0625,100,2,0),ODDLYIELD(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,pr,100,2,FALSE))";
    let oddl_false_basis_value = eval_number_or_skip(&mut sheet, oddl_false_basis)
        .expect("ODDLYIELD should accept FALSE basis (FALSE->0)");
    assert_close(oddl_false_basis_value, oddl_baseline_basis_0_value, 1e-9);

    let oddl_empty_string_basis = r#"=LET(pr,ODDLPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,0.0625,100,2,0),ODDLYIELD(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,pr,100,2,""))"#;
    let oddl_empty_string_basis_value = eval_number_or_skip(&mut sheet, oddl_empty_string_basis)
        .expect("ODDLYIELD should treat basis=\"\" as 0");
    assert_close(
        oddl_empty_string_basis_value,
        oddl_baseline_basis_0_value,
        1e-9,
    );
    let oddl_text_basis_0 = r#"=LET(pr,ODDLPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,0.0625,100,2,0),ODDLYIELD(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,pr,100,2,"0"))"#;
    let oddl_text_basis_0_value = eval_number_or_skip(&mut sheet, oddl_text_basis_0)
        .expect("ODDLYIELD should accept basis supplied as numeric text");
    assert_close(oddl_text_basis_0_value, oddl_baseline_basis_0_value, 1e-9);

    let oddl_baseline_basis_1 =
        "=LET(pr,ODDLPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,0.0625,100,2,1),ODDLYIELD(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,pr,100,2,1))";
    let oddl_baseline_basis_1_value = eval_number_or_skip(&mut sheet, oddl_baseline_basis_1)
        .expect("ODDLYIELD should accept explicit basis=1");
    let oddl_true_basis =
        "=LET(pr,ODDLPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,0.0625,100,2,1),ODDLYIELD(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,pr,100,2,TRUE))";
    let oddl_true_basis_value = eval_number_or_skip(&mut sheet, oddl_true_basis)
        .expect("ODDLYIELD should accept TRUE basis (TRUE->1)");
    assert_close(oddl_true_basis_value, oddl_baseline_basis_1_value, 1e-9);
    let oddl_text_basis_1 = r#"=LET(pr,ODDLPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,0.0625,100,2,1),ODDLYIELD(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,pr,100,2,"1"))"#;
    let oddl_text_basis_1_value = eval_number_or_skip(&mut sheet, oddl_text_basis_1)
        .expect("ODDLYIELD should accept basis=1 supplied as numeric text");
    assert_close(oddl_text_basis_1_value, oddl_baseline_basis_1_value, 1e-9);
}

#[test]
fn odd_coupon_yield_functions_truncate_frequency_like_excel() {
    let mut sheet = TestSheet::new();
    // Task: Excel-like truncation/coercion for non-integer `frequency` inputs in the yield
    // functions.
    //
    // Excel truncates `frequency` to an integer (towards zero) before validating membership in
    // {1, 2, 4}.

    // ODDFYIELD: 2.9 truncates to 2.
    let oddf_baseline_2 =
        "=LET(pr,ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,0),ODDFYIELD(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,pr,100,2,0))";
    let oddf_baseline_2_value = match eval_number_or_skip(&mut sheet, oddf_baseline_2) {
        Some(v) => v,
        None => return,
    };
    let oddf_2_9 =
        "=LET(pr,ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,0),ODDFYIELD(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,pr,100,2.9,0))";
    let oddf_2_9_value =
        eval_number_or_skip(&mut sheet, oddf_2_9).expect("ODDFYIELD should truncate frequency");
    assert_close(oddf_2_9_value, oddf_baseline_2_value, 1e-9);
    let oddf_text_2_9 = r#"=LET(pr,ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,0),ODDFYIELD(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,pr,100,"2.9",0))"#;
    let oddf_text_2_9_value = eval_number_or_skip(&mut sheet, oddf_text_2_9)
        .expect("ODDFYIELD should truncate frequency supplied as numeric text");
    assert_close(oddf_text_2_9_value, oddf_baseline_2_value, 1e-9);

    // ODDFYIELD: 1.9 truncates to 1.
    let oddf_baseline_1 =
        "=LET(pr,ODDFPRICE(DATE(2020,3,1),DATE(2023,7,1),DATE(2020,1,1),DATE(2020,7,1),0.06,0.05,100,1,0),ODDFYIELD(DATE(2020,3,1),DATE(2023,7,1),DATE(2020,1,1),DATE(2020,7,1),0.06,pr,100,1,0))";
    let oddf_baseline_1_value = eval_number_or_skip(&mut sheet, oddf_baseline_1)
        .expect("ODDFYIELD should accept explicit annual frequency");
    let oddf_1_9 =
        "=LET(pr,ODDFPRICE(DATE(2020,3,1),DATE(2023,7,1),DATE(2020,1,1),DATE(2020,7,1),0.06,0.05,100,1,0),ODDFYIELD(DATE(2020,3,1),DATE(2023,7,1),DATE(2020,1,1),DATE(2020,7,1),0.06,pr,100,1.9,0))";
    let oddf_1_9_value =
        eval_number_or_skip(&mut sheet, oddf_1_9).expect("ODDFYIELD should truncate frequency");
    assert_close(oddf_1_9_value, oddf_baseline_1_value, 1e-9);

    // ODDFYIELD: 4.1 truncates to 4.
    let oddf_baseline_4 =
        "=LET(pr,ODDFPRICE(DATE(2020,1,20),DATE(2021,8,15),DATE(2020,1,1),DATE(2020,2,15),0.08,0.07,100,4,0),ODDFYIELD(DATE(2020,1,20),DATE(2021,8,15),DATE(2020,1,1),DATE(2020,2,15),0.08,pr,100,4,0))";
    let oddf_baseline_4_value = eval_number_or_skip(&mut sheet, oddf_baseline_4)
        .expect("ODDFYIELD should accept quarterly frequency");
    let oddf_4_1 =
        "=LET(pr,ODDFPRICE(DATE(2020,1,20),DATE(2021,8,15),DATE(2020,1,1),DATE(2020,2,15),0.08,0.07,100,4,0),ODDFYIELD(DATE(2020,1,20),DATE(2021,8,15),DATE(2020,1,1),DATE(2020,2,15),0.08,pr,100,4.1,0))";
    let oddf_4_1_value =
        eval_number_or_skip(&mut sheet, oddf_4_1).expect("ODDFYIELD should truncate frequency");
    assert_close(oddf_4_1_value, oddf_baseline_4_value, 1e-9);

    // frequency=0.9 truncates to 0 and should return #NUM!.
    let oddf_0_9 =
        "=LET(pr,ODDFPRICE(DATE(2020,3,1),DATE(2023,7,1),DATE(2020,1,1),DATE(2020,7,1),0.06,0.05,100,1,0),ODDFYIELD(DATE(2020,3,1),DATE(2023,7,1),DATE(2020,1,1),DATE(2020,7,1),0.06,pr,100,0.9,0))";
    match sheet.eval(oddf_0_9) {
        Value::Error(ErrorKind::Name) => return,
        Value::Error(ErrorKind::Num) => {}
        other => panic!("expected #NUM! for ODDFYIELD frequency=0.9 (trunc->0), got {other:?}"),
    }
    let oddf_text_0_9 = r#"=LET(pr,ODDFPRICE(DATE(2020,3,1),DATE(2023,7,1),DATE(2020,1,1),DATE(2020,7,1),0.06,0.05,100,1,0),ODDFYIELD(DATE(2020,3,1),DATE(2023,7,1),DATE(2020,1,1),DATE(2020,7,1),0.06,pr,100,"0.9",0))"#;
    match sheet.eval(oddf_text_0_9) {
        Value::Error(ErrorKind::Name) => return,
        Value::Error(ErrorKind::Num) => {}
        other => {
            panic!("expected #NUM! for ODDFYIELD frequency=\"0.9\" (trunc->0), got {other:?}")
        }
    }

    // Repeat key cases for ODDLYIELD as well.
    let oddl_baseline_2 =
        "=LET(pr,ODDLPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,0.0625,100,2,0),ODDLYIELD(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,pr,100,2,0))";
    let oddl_baseline_2_value = eval_number_or_skip(&mut sheet, oddl_baseline_2)
        .expect("ODDLYIELD baseline should evaluate");
    let oddl_2_9 =
        "=LET(pr,ODDLPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,0.0625,100,2,0),ODDLYIELD(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,pr,100,2.9,0))";
    let oddl_2_9_value =
        eval_number_or_skip(&mut sheet, oddl_2_9).expect("ODDLYIELD should truncate frequency");
    assert_close(oddl_2_9_value, oddl_baseline_2_value, 1e-9);
    let oddl_text_2_9 = r#"=LET(pr,ODDLPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,0.0625,100,2,0),ODDLYIELD(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,pr,100,"2.9",0))"#;
    let oddl_text_2_9_value = eval_number_or_skip(&mut sheet, oddl_text_2_9)
        .expect("ODDLYIELD should truncate frequency supplied as numeric text");
    assert_close(oddl_text_2_9_value, oddl_baseline_2_value, 1e-9);

    let oddl_baseline_1 =
        "=LET(pr,ODDLPRICE(DATE(2022,11,1),DATE(2023,3,1),DATE(2022,7,1),0.06,0.05,100,1,0),ODDLYIELD(DATE(2022,11,1),DATE(2023,3,1),DATE(2022,7,1),0.06,pr,100,1,0))";
    let oddl_baseline_1_value = eval_number_or_skip(&mut sheet, oddl_baseline_1)
        .expect("ODDLYIELD should accept explicit annual frequency");
    let oddl_1_9 =
        "=LET(pr,ODDLPRICE(DATE(2022,11,1),DATE(2023,3,1),DATE(2022,7,1),0.06,0.05,100,1,0),ODDLYIELD(DATE(2022,11,1),DATE(2023,3,1),DATE(2022,7,1),0.06,pr,100,1.9,0))";
    let oddl_1_9_value =
        eval_number_or_skip(&mut sheet, oddl_1_9).expect("ODDLYIELD should truncate frequency");
    assert_close(oddl_1_9_value, oddl_baseline_1_value, 1e-9);

    let oddl_baseline_4 =
        "=LET(pr,ODDLPRICE(DATE(2021,7,1),DATE(2021,8,15),DATE(2021,6,15),0.08,0.07,100,4,0),ODDLYIELD(DATE(2021,7,1),DATE(2021,8,15),DATE(2021,6,15),0.08,pr,100,4,0))";
    let oddl_baseline_4_value = eval_number_or_skip(&mut sheet, oddl_baseline_4)
        .expect("ODDLYIELD should accept quarterly frequency");
    let oddl_4_1 =
        "=LET(pr,ODDLPRICE(DATE(2021,7,1),DATE(2021,8,15),DATE(2021,6,15),0.08,0.07,100,4,0),ODDLYIELD(DATE(2021,7,1),DATE(2021,8,15),DATE(2021,6,15),0.08,pr,100,4.1,0))";
    let oddl_4_1_value =
        eval_number_or_skip(&mut sheet, oddl_4_1).expect("ODDLYIELD should truncate frequency");
    assert_close(oddl_4_1_value, oddl_baseline_4_value, 1e-9);

    let oddl_0_9 =
        "=LET(pr,ODDLPRICE(DATE(2022,11,1),DATE(2023,3,1),DATE(2022,7,1),0.06,0.05,100,1,0),ODDLYIELD(DATE(2022,11,1),DATE(2023,3,1),DATE(2022,7,1),0.06,pr,100,0.9,0))";
    match sheet.eval(oddl_0_9) {
        Value::Error(ErrorKind::Name) => return,
        Value::Error(ErrorKind::Num) => {}
        other => panic!("expected #NUM! for ODDLYIELD frequency=0.9 (trunc->0), got {other:?}"),
    }
    let oddl_text_0_9 = r#"=LET(pr,ODDLPRICE(DATE(2022,11,1),DATE(2023,3,1),DATE(2022,7,1),0.06,0.05,100,1,0),ODDLYIELD(DATE(2022,11,1),DATE(2023,3,1),DATE(2022,7,1),0.06,pr,100,"0.9",0))"#;
    match sheet.eval(oddl_text_0_9) {
        Value::Error(ErrorKind::Name) => return,
        Value::Error(ErrorKind::Num) => {}
        other => {
            panic!("expected #NUM! for ODDLYIELD frequency=\"0.9\" (trunc->0), got {other:?}")
        }
    }
}

#[test]
fn odd_coupon_yield_functions_truncate_basis_like_excel() {
    let mut sheet = TestSheet::new();
    // Task: Excel-like truncation/coercion for non-integer `basis` inputs in the yield functions.
    //
    // Excel truncates `basis` to an integer (towards zero) before validating membership in
    // {0, 1, 2, 3, 4}.

    // ODDFYIELD: 0.9 truncates to 0.
    let oddf_basis_0 =
        "=LET(pr,ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,0),ODDFYIELD(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,pr,100,2,0))";
    let oddf_basis_0_value = match eval_number_or_skip(&mut sheet, oddf_basis_0) {
        Some(v) => v,
        None => return,
    };
    let oddf_basis_0_9 =
        "=LET(pr,ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,0),ODDFYIELD(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,pr,100,2,0.9))";
    let oddf_basis_0_9_value =
        eval_number_or_skip(&mut sheet, oddf_basis_0_9).expect("ODDFYIELD should truncate basis");
    assert_close(oddf_basis_0_9_value, oddf_basis_0_value, 1e-9);
    let oddf_basis_text_0_9 = r#"=LET(pr,ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,0),ODDFYIELD(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,pr,100,2,"0.9"))"#;
    let oddf_basis_text_0_9_value = eval_number_or_skip(&mut sheet, oddf_basis_text_0_9)
        .expect("ODDFYIELD should truncate basis supplied as numeric text");
    assert_close(oddf_basis_text_0_9_value, oddf_basis_0_value, 1e-9);
    let oddf_basis_neg_0_9 =
        "=LET(pr,ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,0),ODDFYIELD(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,pr,100,2,-0.9))";
    let oddf_basis_neg_0_9_value = eval_number_or_skip(&mut sheet, oddf_basis_neg_0_9)
        .expect("ODDFYIELD should truncate basis toward zero");
    assert_close(oddf_basis_neg_0_9_value, oddf_basis_0_value, 1e-9);
    let oddf_basis_text_neg_0_9 = r#"=LET(pr,ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,0),ODDFYIELD(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,pr,100,2,"-0.9"))"#;
    let oddf_basis_text_neg_0_9_value = eval_number_or_skip(&mut sheet, oddf_basis_text_neg_0_9)
        .expect("ODDFYIELD should truncate basis supplied as numeric text toward zero");
    assert_close(oddf_basis_text_neg_0_9_value, oddf_basis_0_value, 1e-9);

    // ODDFYIELD: 1.9 truncates to 1.
    let oddf_basis_1 =
        "=LET(pr,ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,1),ODDFYIELD(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,pr,100,2,1))";
    let oddf_basis_1_value =
        eval_number_or_skip(&mut sheet, oddf_basis_1).expect("ODDFYIELD should accept basis=1");
    let oddf_basis_1_9 =
        "=LET(pr,ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,1),ODDFYIELD(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,pr,100,2,1.9))";
    let oddf_basis_1_9_value =
        eval_number_or_skip(&mut sheet, oddf_basis_1_9).expect("ODDFYIELD should truncate basis");
    assert_close(oddf_basis_1_9_value, oddf_basis_1_value, 1e-9);

    // ODDFYIELD: 4.9 truncates to 4.
    let oddf_basis_4 =
        "=LET(pr,ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,4),ODDFYIELD(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,pr,100,2,4))";
    let oddf_basis_4_value =
        eval_number_or_skip(&mut sheet, oddf_basis_4).expect("ODDFYIELD should accept basis=4");
    let oddf_basis_4_9 =
        "=LET(pr,ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,4),ODDFYIELD(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,pr,100,2,4.9))";
    let oddf_basis_4_9_value =
        eval_number_or_skip(&mut sheet, oddf_basis_4_9).expect("ODDFYIELD should truncate basis");
    assert_close(oddf_basis_4_9_value, oddf_basis_4_value, 1e-9);

    // basis=5.1 truncates to 5 and should return #NUM!.
    let oddf_basis_5_1 =
        "=LET(pr,ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,0),ODDFYIELD(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,pr,100,2,5.1))";
    match sheet.eval(oddf_basis_5_1) {
        Value::Error(ErrorKind::Name) => return,
        Value::Error(ErrorKind::Num) => {}
        other => panic!("expected #NUM! for ODDFYIELD basis=5.1 (trunc->5), got {other:?}"),
    }
    let oddf_basis_text_5_1 = r#"=LET(pr,ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,0),ODDFYIELD(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,pr,100,2,"5.1"))"#;
    match sheet.eval(oddf_basis_text_5_1) {
        Value::Error(ErrorKind::Name) => return,
        Value::Error(ErrorKind::Num) => {}
        other => {
            panic!("expected #NUM! for ODDFYIELD basis=\"5.1\" (trunc->5), got {other:?}")
        }
    }

    // Spot-check ODDLYIELD as well.
    let oddl_basis_0 =
        "=LET(pr,ODDLPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,0.0625,100,2,0),ODDLYIELD(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,pr,100,2,0))";
    let oddl_basis_0_value =
        eval_number_or_skip(&mut sheet, oddl_basis_0).expect("ODDLYIELD baseline should evaluate");
    let oddl_basis_0_9 =
        "=LET(pr,ODDLPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,0.0625,100,2,0),ODDLYIELD(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,pr,100,2,0.9))";
    let oddl_basis_0_9_value =
        eval_number_or_skip(&mut sheet, oddl_basis_0_9).expect("ODDLYIELD should truncate basis");
    assert_close(oddl_basis_0_9_value, oddl_basis_0_value, 1e-9);
    let oddl_basis_text_0_9 = r#"=LET(pr,ODDLPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,0.0625,100,2,0),ODDLYIELD(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,pr,100,2,"0.9"))"#;
    let oddl_basis_text_0_9_value = eval_number_or_skip(&mut sheet, oddl_basis_text_0_9)
        .expect("ODDLYIELD should truncate basis supplied as numeric text");
    assert_close(oddl_basis_text_0_9_value, oddl_basis_0_value, 1e-9);
    let oddl_basis_neg_0_9 =
        "=LET(pr,ODDLPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,0.0625,100,2,0),ODDLYIELD(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,pr,100,2,-0.9))";
    let oddl_basis_neg_0_9_value = eval_number_or_skip(&mut sheet, oddl_basis_neg_0_9)
        .expect("ODDLYIELD should truncate basis toward zero");
    assert_close(oddl_basis_neg_0_9_value, oddl_basis_0_value, 1e-9);
    let oddl_basis_text_neg_0_9 = r#"=LET(pr,ODDLPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,0.0625,100,2,0),ODDLYIELD(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,pr,100,2,"-0.9"))"#;
    let oddl_basis_text_neg_0_9_value = eval_number_or_skip(&mut sheet, oddl_basis_text_neg_0_9)
        .expect("ODDLYIELD should truncate basis supplied as numeric text toward zero");
    assert_close(oddl_basis_text_neg_0_9_value, oddl_basis_0_value, 1e-9);

    let oddl_basis_1 =
        "=LET(pr,ODDLPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,0.0625,100,2,1),ODDLYIELD(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,pr,100,2,1))";
    let oddl_basis_1_value =
        eval_number_or_skip(&mut sheet, oddl_basis_1).expect("ODDLYIELD should accept basis=1");
    let oddl_basis_1_9 =
        "=LET(pr,ODDLPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,0.0625,100,2,1),ODDLYIELD(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,pr,100,2,1.9))";
    let oddl_basis_1_9_value =
        eval_number_or_skip(&mut sheet, oddl_basis_1_9).expect("ODDLYIELD should truncate basis");
    assert_close(oddl_basis_1_9_value, oddl_basis_1_value, 1e-9);

    let oddl_basis_4 =
        "=LET(pr,ODDLPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,0.0625,100,2,4),ODDLYIELD(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,pr,100,2,4))";
    let oddl_basis_4_value =
        eval_number_or_skip(&mut sheet, oddl_basis_4).expect("ODDLYIELD should accept basis=4");
    let oddl_basis_4_9 =
        "=LET(pr,ODDLPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,0.0625,100,2,4),ODDLYIELD(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,pr,100,2,4.9))";
    let oddl_basis_4_9_value =
        eval_number_or_skip(&mut sheet, oddl_basis_4_9).expect("ODDLYIELD should truncate basis");
    assert_close(oddl_basis_4_9_value, oddl_basis_4_value, 1e-9);

    let oddl_basis_5_1 =
        "=LET(pr,ODDLPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,0.0625,100,2,0),ODDLYIELD(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,pr,100,2,5.1))";
    match sheet.eval(oddl_basis_5_1) {
        Value::Error(ErrorKind::Name) => return,
        Value::Error(ErrorKind::Num) => {}
        other => panic!("expected #NUM! for ODDLYIELD basis=5.1 (trunc->5), got {other:?}"),
    }
    let oddl_basis_text_5_1 = r#"=LET(pr,ODDLPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,0.0625,100,2,0),ODDLYIELD(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,pr,100,2,"5.1"))"#;
    match sheet.eval(oddl_basis_text_5_1) {
        Value::Error(ErrorKind::Name) => return,
        Value::Error(ErrorKind::Num) => {}
        other => {
            panic!("expected #NUM! for ODDLYIELD basis=\"5.1\" (trunc->5), got {other:?}")
        }
    }
}

#[test]
fn odd_coupon_functions_accept_iso_date_text_arguments() {
    let mut sheet = TestSheet::new();
    // Excel date coercion: ISO-like text should be parsed as a date serial.
    // Ensure odd coupon functions accept text dates and produce the same result
    // as DATE()-based inputs.

    // Baseline case (Task 56): odd first coupon period.
    let baseline_oddfprice = "=ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,0)";
    let baseline_oddfprice_value = match eval_number_or_skip(&mut sheet, baseline_oddfprice) {
        Some(v) => v,
        None => return,
    };
    let iso_oddfprice =
        r#"=ODDFPRICE("2008-11-11","2021-03-01","2008-10-15","2009-03-01",0.0785,0.0625,100,2,0)"#;
    let iso_oddfprice_value = eval_number_or_skip(&mut sheet, iso_oddfprice)
        .expect("ODDFPRICE should accept ISO date text arguments");
    assert_close(iso_oddfprice_value, baseline_oddfprice_value, 1e-9);

    // Baseline case (Task 56): odd last coupon period.
    let baseline_oddlprice =
        "=ODDLPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,0.0625,100,2,0)";
    let baseline_oddlprice_value = eval_number_or_skip(&mut sheet, baseline_oddlprice)
        .expect("ODDLPRICE should return a number for the baseline");
    let iso_oddlprice =
        r#"=ODDLPRICE("2020-11-11","2021-03-01","2020-10-15",0.0785,0.0625,100,2,0)"#;
    let iso_oddlprice_value = eval_number_or_skip(&mut sheet, iso_oddlprice)
        .expect("ODDLPRICE should accept ISO date text arguments");
    assert_close(iso_oddlprice_value, baseline_oddlprice_value, 1e-9);
}

#[test]
fn odd_coupon_functions_floor_time_fractions_in_date_arguments() {
    let mut sheet = TestSheet::new();
    // Excel date serials can contain a fractional time component. For these bond functions,
    // Excel behaves as though date arguments are floored to the day.

    let baseline_oddfprice = "=ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,0)";
    let baseline_oddfprice_value = match eval_number_or_skip(&mut sheet, baseline_oddfprice) {
        Some(v) => v,
        None => return,
    };
    let fractional_oddfprice = "=ODDFPRICE(DATE(2008,11,11)+0.75,DATE(2021,3,1)+0.1,DATE(2008,10,15)+0.9,DATE(2009,3,1)+0.5,0.0785,0.0625,100,2,0)";
    let fractional_oddfprice_value = eval_number_or_skip(&mut sheet, fractional_oddfprice)
        .expect("ODDFPRICE should floor fractional date serials");
    assert_close(fractional_oddfprice_value, baseline_oddfprice_value, 1e-9);

    let baseline_oddlprice =
        "=ODDLPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,0.0625,100,2,0)";
    let baseline_oddlprice_value = eval_number_or_skip(&mut sheet, baseline_oddlprice)
        .expect("ODDLPRICE should return a number for the baseline");
    let fractional_oddlprice = "=ODDLPRICE(DATE(2020,11,11)+0.75,DATE(2021,3,1)+0.1,DATE(2020,10,15)+0.9,0.0785,0.0625,100,2,0)";
    let fractional_oddlprice_value = eval_number_or_skip(&mut sheet, fractional_oddlprice)
        .expect("ODDLPRICE should floor fractional date serials");
    assert_close(fractional_oddlprice_value, baseline_oddlprice_value, 1e-9);
}

#[test]
fn oddfyield_roundtrips_price_with_text_dates() {
    let mut sheet = TestSheet::new();
    // Ensure ODDFYIELD accepts date arguments supplied as ISO-like text, and that the
    // ODDFPRICE/ODDFYIELD pair roundtrips the yield.

    sheet.set_formula(
        "A1",
        "=ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,0)",
    );
    sheet.recalc();

    let _price = match cell_number_or_skip(&sheet, "A1") {
        Some(v) => v,
        None => return,
    };

    let recovered_yield = match eval_number_or_skip(
        &mut sheet,
        r#"=ODDFYIELD("2008-11-11","2021-03-01","2008-10-15","2009-03-01",0.0785,A1,100,2,0)"#,
    ) {
        Some(v) => v,
        None => return,
    };
    assert_close(recovered_yield, 0.0625, 1e-10);
}

#[test]
fn oddlyield_roundtrips_price_with_text_dates() {
    let mut sheet = TestSheet::new();
    // Ensure ODDLYIELD accepts date arguments supplied as ISO-like text, and that the
    // ODDLPRICE/ODDLYIELD pair roundtrips the yield.

    sheet.set_formula(
        "A1",
        "=ODDLPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,0.0625,100,2,0)",
    );
    sheet.recalc();

    let _price = match cell_number_or_skip(&sheet, "A1") {
        Some(v) => v,
        None => return,
    };

    let recovered_yield = match eval_number_or_skip(
        &mut sheet,
        r#"=ODDLYIELD("2020-11-11","2021-03-01","2020-10-15",0.0785,A1,100,2,0)"#,
    ) {
        Some(v) => v,
        None => return,
    };
    assert_close(recovered_yield, 0.0625, 1e-10);
}

#[test]
fn odd_coupon_yield_functions_floor_time_fractions_in_date_arguments() {
    let mut sheet = TestSheet::new();
    // Like ODDFPRICE/ODDLPRICE, Excel behaves as though the date args to ODDFYIELD/ODDLYIELD are
    // floored to whole days.

    let baseline_oddfyield =
        "=ODDFYIELD(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,98,100,2,0)";
    let baseline_oddfyield_value = match eval_number_or_skip(&mut sheet, baseline_oddfyield) {
        Some(v) => v,
        None => return,
    };

    let fractional_oddfyield = "=ODDFYIELD(DATE(2008,11,11)+0.75,DATE(2021,3,1)+0.1,DATE(2008,10,15)+0.9,DATE(2009,3,1)+0.5,0.0785,98,100,2,0)";
    let fractional_oddfyield_value = eval_number_or_skip(&mut sheet, fractional_oddfyield)
        .expect("ODDFYIELD should floor fractional date serials");
    assert_close(fractional_oddfyield_value, baseline_oddfyield_value, 1e-10);

    let baseline_oddlyield =
        "=ODDLYIELD(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,98,100,2,0)";
    let baseline_oddlyield_value = match eval_number_or_skip(&mut sheet, baseline_oddlyield) {
        Some(v) => v,
        None => return,
    };

    let fractional_oddlyield =
        "=ODDLYIELD(DATE(2020,11,11)+0.75,DATE(2021,3,1)+0.1,DATE(2020,10,15)+0.9,0.0785,98,100,2,0)";
    let fractional_oddlyield_value = eval_number_or_skip(&mut sheet, fractional_oddlyield)
        .expect("ODDLYIELD should floor fractional date serials");
    assert_close(fractional_oddlyield_value, baseline_oddlyield_value, 1e-10);
}

#[test]
fn odd_coupon_functions_truncate_basis_like_excel() {
    let mut sheet = TestSheet::new();
    // Task: Excel-like truncation/coercion for non-integer `basis` inputs.
    //
    // Excel truncates `basis` to an integer (towards zero) before validating membership in
    // {0, 1, 2, 3, 4}.

    let oddf_basis_0 = "=ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,0)";
    let oddf_basis_0_value = match eval_number_or_skip(&mut sheet, oddf_basis_0) {
        Some(v) => v,
        None => return,
    };
    let oddf_basis_0_9 =
        "=ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,0.9)";
    let oddf_basis_0_9_value =
        eval_number_or_skip(&mut sheet, oddf_basis_0_9).expect("ODDFPRICE should truncate basis");
    assert_close(oddf_basis_0_9_value, oddf_basis_0_value, 1e-9);
    let oddf_basis_text_0_9 = r#"=ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,"0.9")"#;
    let oddf_basis_text_0_9_value = eval_number_or_skip(&mut sheet, oddf_basis_text_0_9)
        .expect("ODDFPRICE should truncate basis supplied as numeric text");
    assert_close(oddf_basis_text_0_9_value, oddf_basis_0_value, 1e-9);
    // Truncation is toward zero (not floor), so -0.9 becomes 0.
    let oddf_basis_neg_0_9 =
        "=ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,-0.9)";
    let oddf_basis_neg_0_9_value = eval_number_or_skip(&mut sheet, oddf_basis_neg_0_9)
        .expect("ODDFPRICE should truncate basis toward zero");
    assert_close(oddf_basis_neg_0_9_value, oddf_basis_0_value, 1e-9);
    let oddf_basis_text_neg_0_9 = r#"=ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,"-0.9")"#;
    let oddf_basis_text_neg_0_9_value = eval_number_or_skip(&mut sheet, oddf_basis_text_neg_0_9)
        .expect("ODDFPRICE should truncate basis supplied as numeric text toward zero");
    assert_close(oddf_basis_text_neg_0_9_value, oddf_basis_0_value, 1e-9);

    // basis=-0.1 truncates to 0 (towards zero).
    let oddf_basis_neg_0_1 =
        "=ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,-0.1)";
    let oddf_basis_neg_0_1_value = eval_number_or_skip(&mut sheet, oddf_basis_neg_0_1)
        .expect("ODDFPRICE should truncate basis");
    assert_close(oddf_basis_neg_0_1_value, oddf_basis_0_value, 1e-9);

    let oddf_basis_1 =
        "=ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,1)";
    let oddf_basis_1_value =
        eval_number_or_skip(&mut sheet, oddf_basis_1).expect("ODDFPRICE should accept basis=1");
    let oddf_basis_1_9 =
        "=ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,1.9)";
    let oddf_basis_1_9_value =
        eval_number_or_skip(&mut sheet, oddf_basis_1_9).expect("ODDFPRICE should truncate basis");
    assert_close(oddf_basis_1_9_value, oddf_basis_1_value, 1e-9);

    let oddf_basis_4 =
        "=ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,4)";
    let oddf_basis_4_value =
        eval_number_or_skip(&mut sheet, oddf_basis_4).expect("ODDFPRICE should accept basis=4");
    let oddf_basis_4_9 =
        "=ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,4.9)";
    let oddf_basis_4_9_value =
        eval_number_or_skip(&mut sheet, oddf_basis_4_9).expect("ODDFPRICE should truncate basis");
    assert_close(oddf_basis_4_9_value, oddf_basis_4_value, 1e-9);

    // basis=5.1 truncates to 5 and should return #NUM!.
    let oddf_basis_5_1 =
        "=ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,5.1)";
    match sheet.eval(oddf_basis_5_1) {
        Value::Error(ErrorKind::Name) => return,
        Value::Error(ErrorKind::Num) => {}
        other => panic!("expected #NUM! for ODDFPRICE basis=5.1 (trunc->5), got {other:?}"),
    }
    let oddf_basis_text_5_1 = r#"=ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,"5.1")"#;
    match sheet.eval(oddf_basis_text_5_1) {
        Value::Error(ErrorKind::Name) => return,
        Value::Error(ErrorKind::Num) => {}
        other => {
            panic!("expected #NUM! for ODDFPRICE basis=\"5.1\" (trunc->5), got {other:?}")
        }
    }

    // Spot-check ODDLPRICE as well.
    let oddl_basis_0 =
        "=ODDLPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,0.0625,100,2,0)";
    let oddl_basis_0_value = match eval_number_or_skip(&mut sheet, oddl_basis_0) {
        Some(v) => v,
        None => return,
    };
    let oddl_basis_0_9 =
        "=ODDLPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,0.0625,100,2,0.9)";
    let oddl_basis_0_9_value =
        eval_number_or_skip(&mut sheet, oddl_basis_0_9).expect("ODDLPRICE should truncate basis");
    assert_close(oddl_basis_0_9_value, oddl_basis_0_value, 1e-9);
    let oddl_basis_text_0_9 =
        r#"=ODDLPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,0.0625,100,2,"0.9")"#;
    let oddl_basis_text_0_9_value = eval_number_or_skip(&mut sheet, oddl_basis_text_0_9)
        .expect("ODDLPRICE should truncate basis supplied as numeric text");
    assert_close(oddl_basis_text_0_9_value, oddl_basis_0_value, 1e-9);
    let oddl_basis_neg_0_9 =
        "=ODDLPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,0.0625,100,2,-0.9)";
    let oddl_basis_neg_0_9_value = eval_number_or_skip(&mut sheet, oddl_basis_neg_0_9)
        .expect("ODDLPRICE should truncate basis toward zero");
    assert_close(oddl_basis_neg_0_9_value, oddl_basis_0_value, 1e-9);
    let oddl_basis_text_neg_0_9 = r#"=ODDLPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,0.0625,100,2,"-0.9")"#;
    let oddl_basis_text_neg_0_9_value = eval_number_or_skip(&mut sheet, oddl_basis_text_neg_0_9)
        .expect("ODDLPRICE should truncate basis supplied as numeric text toward zero");
    assert_close(oddl_basis_text_neg_0_9_value, oddl_basis_0_value, 1e-9);

    let oddl_basis_1 =
        "=ODDLPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,0.0625,100,2,1)";
    let oddl_basis_1_value =
        eval_number_or_skip(&mut sheet, oddl_basis_1).expect("ODDLPRICE should accept basis=1");
    let oddl_basis_1_9 =
        "=ODDLPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,0.0625,100,2,1.9)";
    let oddl_basis_1_9_value =
        eval_number_or_skip(&mut sheet, oddl_basis_1_9).expect("ODDLPRICE should truncate basis");
    assert_close(oddl_basis_1_9_value, oddl_basis_1_value, 1e-9);

    let oddl_basis_4 =
        "=ODDLPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,0.0625,100,2,4)";
    let oddl_basis_4_value =
        eval_number_or_skip(&mut sheet, oddl_basis_4).expect("ODDLPRICE should accept basis=4");
    let oddl_basis_4_9 =
        "=ODDLPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,0.0625,100,2,4.9)";
    let oddl_basis_4_9_value =
        eval_number_or_skip(&mut sheet, oddl_basis_4_9).expect("ODDLPRICE should truncate basis");
    assert_close(oddl_basis_4_9_value, oddl_basis_4_value, 1e-9);

    let oddl_basis_5_1 =
        "=ODDLPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,0.0625,100,2,5.1)";
    match sheet.eval(oddl_basis_5_1) {
        Value::Error(ErrorKind::Name) => return,
        Value::Error(ErrorKind::Num) => {}
        other => panic!("expected #NUM! for ODDLPRICE basis=5.1 (trunc->5), got {other:?}"),
    }
    let oddl_basis_text_5_1 =
        r#"=ODDLPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,0.0625,100,2,"5.1")"#;
    match sheet.eval(oddl_basis_text_5_1) {
        Value::Error(ErrorKind::Name) => return,
        Value::Error(ErrorKind::Num) => {}
        other => {
            panic!("expected #NUM! for ODDLPRICE basis=\"5.1\" (trunc->5), got {other:?}")
        }
    }
}

#[test]
fn odd_first_coupon_bond_functions_respect_workbook_date_system() {
    let system_1900 = ExcelDateSystem::EXCEL_1900;
    let system_1904 = ExcelDateSystem::Excel1904;

    // Baseline case (Task 56): odd first coupon period.
    let settlement_1900 = serial(2008, 11, 11, system_1900);
    let maturity_1900 = serial(2021, 3, 1, system_1900);
    let issue_1900 = serial(2008, 10, 15, system_1900);
    let first_coupon_1900 = serial(2009, 3, 1, system_1900);

    let settlement_1904 = serial(2008, 11, 11, system_1904);
    let maturity_1904 = serial(2021, 3, 1, system_1904);
    let issue_1904 = serial(2008, 10, 15, system_1904);
    let first_coupon_1904 = serial(2009, 3, 1, system_1904);

    let price_1900 = oddfprice(
        settlement_1900,
        maturity_1900,
        issue_1900,
        first_coupon_1900,
        0.0785,
        0.0625,
        100.0,
        2,
        0,
        system_1900,
    )
    .expect("oddfprice should succeed under Excel1900");
    let yield_1900 = oddfyield(
        settlement_1900,
        maturity_1900,
        issue_1900,
        first_coupon_1900,
        0.0785,
        98.0,
        100.0,
        2,
        0,
        system_1900,
    )
    .expect("oddfyield should succeed under Excel1900");

    let price_1904 = oddfprice(
        settlement_1904,
        maturity_1904,
        issue_1904,
        first_coupon_1904,
        0.0785,
        0.0625,
        100.0,
        2,
        0,
        system_1904,
    )
    .expect("oddfprice should succeed under Excel1904");
    let yield_1904 = oddfyield(
        settlement_1904,
        maturity_1904,
        issue_1904,
        first_coupon_1904,
        0.0785,
        98.0,
        100.0,
        2,
        0,
        system_1904,
    )
    .expect("oddfyield should succeed under Excel1904");

    assert_close(price_1904, price_1900, 1e-9);
    assert_close(yield_1904, yield_1900, 1e-10);
}

#[test]
fn odd_first_coupon_bond_functions_round_trip_long_stub() {
    let mut sheet = TestSheet::new();

    // Long odd-first coupon period:
    // - issue is far before first_coupon so DFC/E > 1 (long stub)
    // - settlement is between issue and first_coupon
    // - maturity is aligned with the regular schedule after first_coupon
    //
    // Also includes a basis=1 variant that crosses the 2020 leap day, exercising
    // E computation for actual/actual.
    //
    // (Pinned by current engine behavior; verify against real Excel via
    // tools/excel-oracle/run-excel-oracle.ps1 (Task 393). This unit test asserts
    // ODDFPRICE/ODDFYIELD are internally consistent.)
    let yield_target = 0.0625;
    let rate = 0.0785;
    let system = ExcelDateSystem::EXCEL_1900;

    let settlement = serial(2019, 6, 1, system);
    let maturity = serial(2022, 3, 1, system);
    let issue = serial(2019, 1, 1, system);
    let first_coupon = serial(2020, 3, 1, system);

    for basis in [0, 1] {
        // Ensure this is actually a *long* first coupon period for this basis.
        // (DFC/E > 1 and DSC/E > 1).
        let dfc = days_between(issue, first_coupon, basis, system);
        let dsc = days_between(settlement, first_coupon, basis, system);
        let months_per_period = 12 / 2;
        let eom = is_end_of_month(maturity, system);
        let prev_coupon = coupon_date_with_eom(first_coupon, -months_per_period, eom, system);
        let e = coupon_period_e(prev_coupon, first_coupon, basis, 2, system);
        assert!(
            dfc / e > 1.0,
            "expected long first coupon (DFC/E > 1), got DFC={dfc} E={e}"
        );
        assert!(
            dsc / e > 1.0,
            "expected settlement to be more than one coupon period before first coupon (DSC/E > 1), got DSC={dsc} E={e}"
        );

        // Independent price model: compute what Excel's published ODDF* formula should produce
        // for this schedule (long first coupon amount scaled by DFC/E).
        let expected_price = oddf_price_excel_model(
            settlement,
            maturity,
            issue,
            first_coupon,
            rate,
            yield_target,
            100.0,
            2,
            basis,
            system,
        );

        let price_formula = format!(
            "=ODDFPRICE(DATE(2019,6,1),DATE(2022,3,1),DATE(2019,1,1),DATE(2020,3,1),{rate},{yield_target},100,2,{basis})"
        );
        sheet.set_formula("A1", &price_formula);
        sheet.recalc();

        let Some(price) = cell_number_or_skip(&sheet, "A1") else {
            return;
        };

        assert_close(price, expected_price, 1e-10);

        assert!(
            price.is_finite() && price > 0.0,
            "expected positive finite price, got {price}"
        );

        // Round-trip yield using the model price (so ODDFYIELD is validated independently
        // of the ODDFPRICE implementation).
        sheet.set("B1", expected_price);
        let yield_formula = format!(
            "=ODDFYIELD(DATE(2019,6,1),DATE(2022,3,1),DATE(2019,1,1),DATE(2020,3,1),{rate},B1,100,2,{basis})"
        );
        let Some(y) = eval_number_or_skip(&mut sheet, &yield_formula) else {
            return;
        };
        assert_close(y, yield_target, 1e-9);
    }
}

#[test]
fn odd_last_coupon_bond_functions_respect_workbook_date_system() {
    let system_1900 = ExcelDateSystem::EXCEL_1900;
    let system_1904 = ExcelDateSystem::Excel1904;

    // Baseline case (Task 56): odd last coupon period.
    let settlement_1900 = serial(2020, 11, 11, system_1900);
    let maturity_1900 = serial(2021, 3, 1, system_1900);
    let last_interest_1900 = serial(2020, 10, 15, system_1900);

    let settlement_1904 = serial(2020, 11, 11, system_1904);
    let maturity_1904 = serial(2021, 3, 1, system_1904);
    let last_interest_1904 = serial(2020, 10, 15, system_1904);

    let price_1900 = oddlprice(
        settlement_1900,
        maturity_1900,
        last_interest_1900,
        0.0785,
        0.0625,
        100.0,
        2,
        0,
        system_1900,
    )
    .expect("oddlprice should succeed under Excel1900");
    let yield_1900 = oddlyield(
        settlement_1900,
        maturity_1900,
        last_interest_1900,
        0.0785,
        98.0,
        100.0,
        2,
        0,
        system_1900,
    )
    .expect("oddlyield should succeed under Excel1900");

    let price_1904 = oddlprice(
        settlement_1904,
        maturity_1904,
        last_interest_1904,
        0.0785,
        0.0625,
        100.0,
        2,
        0,
        system_1904,
    )
    .expect("oddlprice should succeed under Excel1904");
    let yield_1904 = oddlyield(
        settlement_1904,
        maturity_1904,
        last_interest_1904,
        0.0785,
        98.0,
        100.0,
        2,
        0,
        system_1904,
    )
    .expect("oddlyield should succeed under Excel1904");

    assert_close(price_1904, price_1900, 1e-9);
    assert_close(yield_1904, yield_1900, 1e-10);
}

#[test]
fn odd_first_coupon_roundtrips_yield_with_annual_frequency() {
    // Aligned annual schedule from `first_coupon` by 12 months:
    // 2020-07-01, 2021-07-01, 2022-07-01, 2023-07-01 (maturity).
    let system = ExcelDateSystem::EXCEL_1900;
    let settlement = serial(2020, 3, 1, system);
    let maturity = serial(2023, 7, 1, system);
    let issue = serial(2020, 1, 1, system);
    let first_coupon = serial(2020, 7, 1, system);

    let yld = 0.05;
    let price = oddfprice(
        settlement,
        maturity,
        issue,
        first_coupon,
        0.06,
        yld,
        100.0,
        1,
        0,
        system,
    )
    .expect("oddfprice should succeed");

    let recovered_yield = oddfyield(
        settlement,
        maturity,
        issue,
        first_coupon,
        0.06,
        price,
        100.0,
        1,
        0,
        system,
    )
    .expect("oddfyield should succeed");

    assert_close(recovered_yield, yld, 1e-7);
}

#[test]
fn odd_first_coupon_roundtrips_yield_with_quarterly_frequency_and_non_100_redemption() {
    // Aligned quarterly schedule from `first_coupon` by 3 months:
    // 2020-02-15, 2020-05-15, 2020-08-15, 2020-11-15, 2021-02-15, 2021-05-15, 2021-08-15.
    let system = ExcelDateSystem::EXCEL_1900;
    let settlement = serial(2020, 1, 20, system);
    let maturity = serial(2021, 8, 15, system);
    let issue = serial(2020, 1, 1, system);
    let first_coupon = serial(2020, 2, 15, system);

    let yld = 0.07;
    let price_100 = oddfprice(
        settlement,
        maturity,
        issue,
        first_coupon,
        0.08,
        yld,
        100.0,
        4,
        0,
        system,
    )
    .expect("oddfprice redemption=100 should succeed");
    let price_105 = oddfprice(
        settlement,
        maturity,
        issue,
        first_coupon,
        0.08,
        yld,
        105.0,
        4,
        0,
        system,
    )
    .expect("oddfprice redemption=105 should succeed");

    assert!(
        (price_105 - price_100).abs() > 1e-9,
        "expected redemption to affect price (redemption=100 => {price_100}, redemption=105 => {price_105})"
    );
    assert!(
        price_105 > price_100,
        "expected higher redemption to increase price (redemption=100 => {price_100}, redemption=105 => {price_105})"
    );

    let recovered_yield_100 = oddfyield(
        settlement,
        maturity,
        issue,
        first_coupon,
        0.08,
        price_100,
        100.0,
        4,
        0,
        system,
    )
    .expect("oddfyield redemption=100 should succeed");
    let recovered_yield_105 = oddfyield(
        settlement,
        maturity,
        issue,
        first_coupon,
        0.08,
        price_105,
        105.0,
        4,
        0,
        system,
    )
    .expect("oddfyield redemption=105 should succeed");

    assert_close(recovered_yield_100, yld, 1e-7);
    assert_close(recovered_yield_105, yld, 1e-7);
}

#[test]
fn odd_last_coupon_roundtrips_yield_with_annual_frequency() {
    // `last_interest` is a coupon date on an annual schedule (12 month stepping). Maturity
    // occurs 8 months later, making this an odd last coupon period.
    let system = ExcelDateSystem::EXCEL_1900;
    let settlement = serial(2022, 11, 1, system);
    let maturity = serial(2023, 3, 1, system);
    let last_interest = serial(2022, 7, 1, system);

    let yld = 0.05;
    let price = oddlprice(
        settlement,
        maturity,
        last_interest,
        0.06,
        yld,
        100.0,
        1,
        0,
        system,
    )
    .expect("oddlprice should succeed");
    let recovered_yield = oddlyield(
        settlement,
        maturity,
        last_interest,
        0.06,
        price,
        100.0,
        1,
        0,
        system,
    )
    .expect("oddlyield should succeed");

    assert_close(recovered_yield, yld, 1e-7);
}

#[test]
fn odd_last_coupon_roundtrips_yield_with_quarterly_frequency_and_non_100_redemption() {
    // `last_interest` is a coupon date on a quarterly schedule. Maturity occurs 2 months later
    // (shorter than the regular 3 month period), making this an odd last coupon period.
    let system = ExcelDateSystem::EXCEL_1900;
    let settlement = serial(2021, 7, 1, system);
    let maturity = serial(2021, 8, 15, system);
    let last_interest = serial(2021, 6, 15, system);

    let yld = 0.07;
    let price_100 = oddlprice(
        settlement,
        maturity,
        last_interest,
        0.08,
        yld,
        100.0,
        4,
        0,
        system,
    )
    .expect("oddlprice redemption=100 should succeed");
    let price_105 = oddlprice(
        settlement,
        maturity,
        last_interest,
        0.08,
        yld,
        105.0,
        4,
        0,
        system,
    )
    .expect("oddlprice redemption=105 should succeed");

    assert!(
        (price_105 - price_100).abs() > 1e-9,
        "expected redemption to affect price (redemption=100 => {price_100}, redemption=105 => {price_105})"
    );
    assert!(
        price_105 > price_100,
        "expected higher redemption to increase price (redemption=100 => {price_100}, redemption=105 => {price_105})"
    );

    let recovered_yield_100 = oddlyield(
        settlement,
        maturity,
        last_interest,
        0.08,
        price_100,
        100.0,
        4,
        0,
        system,
    )
    .expect("oddlyield redemption=100 should succeed");
    let recovered_yield_105 = oddlyield(
        settlement,
        maturity,
        last_interest,
        0.08,
        price_105,
        105.0,
        4,
        0,
        system,
    )
    .expect("oddlyield redemption=105 should succeed");

    assert_close(recovered_yield_100, yld, 1e-7);
    assert_close(recovered_yield_105, yld, 1e-7);
}

#[test]
fn oddfprice_rejects_yield_at_or_below_negative_frequency() {
    let system = ExcelDateSystem::EXCEL_1900;

    let issue = ymd_to_serial(ExcelDate::new(2020, 1, 1), system).unwrap();
    let settlement = ymd_to_serial(ExcelDate::new(2020, 1, 15), system).unwrap();
    let first_coupon = ymd_to_serial(ExcelDate::new(2020, 7, 15), system).unwrap();
    let maturity = ymd_to_serial(ExcelDate::new(2033, 1, 15), system).unwrap();

    let rate = 0.05;
    let redemption = 100.0;
    let frequency = 2;
    let basis = 0;

    // A yield below -1.0 can still be valid when `frequency > 1`, as long as `1 + yld/frequency > 0`.
    let ok_yld = -(frequency as f64) + 0.5; // e.g. -1.5 when frequency=2
    let price = oddfprice(
        settlement,
        maturity,
        issue,
        first_coupon,
        rate,
        ok_yld,
        redemption,
        frequency,
        basis,
        system,
    )
    .expect("ODDFPRICE should allow yields in (-frequency, ∞)");
    assert!(price.is_finite());

    let boundary = -(frequency as f64);
    let result = oddfprice(
        settlement,
        maturity,
        issue,
        first_coupon,
        rate,
        boundary,
        redemption,
        frequency,
        basis,
        system,
    );
    assert!(
        matches!(result, Err(ExcelError::Div0)),
        "expected #DIV/0! for yld=-frequency, got {result:?}"
    );

    let result = oddfprice(
        settlement,
        maturity,
        issue,
        first_coupon,
        rate,
        boundary - 0.5,
        redemption,
        frequency,
        basis,
        system,
    );
    assert!(
        matches!(result, Err(ExcelError::Num)),
        "expected #NUM! for yld < -frequency, got {result:?}"
    );
}

#[test]
fn odd_coupon_internal_api_rejects_non_finite_numeric_inputs() {
    let system = ExcelDateSystem::EXCEL_1900;

    // ODDF* setup.
    let issue = ymd_to_serial(ExcelDate::new(2019, 10, 1), system).unwrap();
    let settlement = ymd_to_serial(ExcelDate::new(2020, 1, 1), system).unwrap();
    let first_coupon = ymd_to_serial(ExcelDate::new(2020, 7, 1), system).unwrap();
    let maturity = ymd_to_serial(ExcelDate::new(2021, 7, 1), system).unwrap();

    // ODDL* setup.
    let last_interest = ymd_to_serial(ExcelDate::new(2021, 1, 1), system).unwrap();
    let settlement2 = ymd_to_serial(ExcelDate::new(2021, 2, 1), system).unwrap();
    let maturity2 = ymd_to_serial(ExcelDate::new(2021, 5, 1), system).unwrap();

    let rate = 0.05;
    let yld = 0.06;
    let pr = 98.0;
    let redemption = 100.0;
    let frequency = 2;
    let basis = 0;

    for bad in [f64::NAN, f64::INFINITY] {
        let result = oddfprice(
            settlement,
            maturity,
            issue,
            first_coupon,
            bad,
            yld,
            redemption,
            frequency,
            basis,
            system,
        );
        assert!(
            matches!(result, Err(ExcelError::Num)),
            "expected #NUM! for ODDFPRICE non-finite rate={bad}, got {result:?}"
        );

        let result = oddfprice(
            settlement,
            maturity,
            issue,
            first_coupon,
            rate,
            bad,
            redemption,
            frequency,
            basis,
            system,
        );
        assert!(
            matches!(result, Err(ExcelError::Num)),
            "expected #NUM! for ODDFPRICE non-finite yld={bad}, got {result:?}"
        );

        let result = oddfprice(
            settlement,
            maturity,
            issue,
            first_coupon,
            rate,
            yld,
            bad,
            frequency,
            basis,
            system,
        );
        assert!(
            matches!(result, Err(ExcelError::Num)),
            "expected #NUM! for ODDFPRICE non-finite redemption={bad}, got {result:?}"
        );

        let result = oddfyield(
            settlement,
            maturity,
            issue,
            first_coupon,
            bad,
            pr,
            redemption,
            frequency,
            basis,
            system,
        );
        assert!(
            matches!(result, Err(ExcelError::Num)),
            "expected #NUM! for ODDFYIELD non-finite rate={bad}, got {result:?}"
        );

        let result = oddfyield(
            settlement,
            maturity,
            issue,
            first_coupon,
            rate,
            bad,
            redemption,
            frequency,
            basis,
            system,
        );
        assert!(
            matches!(result, Err(ExcelError::Num)),
            "expected #NUM! for ODDFYIELD non-finite pr={bad}, got {result:?}"
        );

        let result = oddfyield(
            settlement,
            maturity,
            issue,
            first_coupon,
            rate,
            pr,
            bad,
            frequency,
            basis,
            system,
        );
        assert!(
            matches!(result, Err(ExcelError::Num)),
            "expected #NUM! for ODDFYIELD non-finite redemption={bad}, got {result:?}"
        );

        let result = oddlprice(
            settlement2,
            maturity2,
            last_interest,
            bad,
            yld,
            redemption,
            frequency,
            basis,
            system,
        );
        assert!(
            matches!(result, Err(ExcelError::Num)),
            "expected #NUM! for ODDLPRICE non-finite rate={bad}, got {result:?}"
        );

        let result = oddlprice(
            settlement2,
            maturity2,
            last_interest,
            rate,
            bad,
            redemption,
            frequency,
            basis,
            system,
        );
        assert!(
            matches!(result, Err(ExcelError::Num)),
            "expected #NUM! for ODDLPRICE non-finite yld={bad}, got {result:?}"
        );

        let result = oddlprice(
            settlement2,
            maturity2,
            last_interest,
            rate,
            yld,
            bad,
            frequency,
            basis,
            system,
        );
        assert!(
            matches!(result, Err(ExcelError::Num)),
            "expected #NUM! for ODDLPRICE non-finite redemption={bad}, got {result:?}"
        );

        let result = oddlyield(
            settlement2,
            maturity2,
            last_interest,
            bad,
            pr,
            redemption,
            frequency,
            basis,
            system,
        );
        assert!(
            matches!(result, Err(ExcelError::Num)),
            "expected #NUM! for ODDLYIELD non-finite rate={bad}, got {result:?}"
        );

        let result = oddlyield(
            settlement2,
            maturity2,
            last_interest,
            rate,
            bad,
            redemption,
            frequency,
            basis,
            system,
        );
        assert!(
            matches!(result, Err(ExcelError::Num)),
            "expected #NUM! for ODDLYIELD non-finite pr={bad}, got {result:?}"
        );

        let result = oddlyield(
            settlement2,
            maturity2,
            last_interest,
            rate,
            pr,
            bad,
            frequency,
            basis,
            system,
        );
        assert!(
            matches!(result, Err(ExcelError::Num)),
            "expected #NUM! for ODDLYIELD non-finite redemption={bad}, got {result:?}"
        );
    }
}

#[test]
fn oddlprice_rejects_yield_at_or_below_negative_frequency() {
    let system = ExcelDateSystem::EXCEL_1900;

    let last_interest = ymd_to_serial(ExcelDate::new(2020, 1, 15), system).unwrap();
    let settlement = ymd_to_serial(ExcelDate::new(2020, 7, 15), system).unwrap();
    let maturity = ymd_to_serial(ExcelDate::new(2033, 1, 15), system).unwrap();

    let rate = 0.05;
    let redemption = 100.0;
    let frequency = 2;
    let basis = 0;

    let ok_yld = -(frequency as f64) + 0.5; // e.g. -1.5 when frequency=2
    let price = oddlprice(
        settlement,
        maturity,
        last_interest,
        rate,
        ok_yld,
        redemption,
        frequency,
        basis,
        system,
    )
    .expect("ODDLPRICE should allow yields in (-frequency, ∞)");
    assert!(price.is_finite());

    let boundary = -(frequency as f64);
    let result = oddlprice(
        settlement,
        maturity,
        last_interest,
        rate,
        boundary,
        redemption,
        frequency,
        basis,
        system,
    );
    assert!(
        matches!(result, Err(ExcelError::Div0)),
        "expected #DIV/0! for yld=-frequency, got {result:?}"
    );

    let result = oddlprice(
        settlement,
        maturity,
        last_interest,
        rate,
        boundary - 0.5,
        redemption,
        frequency,
        basis,
        system,
    );
    assert!(
        matches!(result, Err(ExcelError::Num)),
        "expected #NUM! for yld < -frequency, got {result:?}"
    );
}

#[test]
fn odd_coupon_worksheet_price_rejects_yield_at_or_below_negative_frequency() {
    let mut sheet = TestSheet::new();

    // Use frequency=2 (so the domain boundary is yld=-2).
    let oddf_boundary = "=ODDFPRICE(DATE(2020,1,15),DATE(2033,1,15),DATE(2020,1,1),DATE(2020,7,15),0.05,-2,100,2,0)";
    let oddf_below = "=ODDFPRICE(DATE(2020,1,15),DATE(2033,1,15),DATE(2020,1,1),DATE(2020,7,15),0.05,-2.5,100,2,0)";

    let Some(out) = eval_value_or_skip(&mut sheet, oddf_boundary) else {
        return;
    };
    assert!(
        matches!(out, Value::Error(ErrorKind::Div0)),
        "expected #DIV/0! for worksheet ODDFPRICE at yld=-frequency, got {out:?}"
    );

    let Some(out) = eval_value_or_skip(&mut sheet, oddf_below) else {
        return;
    };
    assert!(
        matches!(out, Value::Error(ErrorKind::Num)),
        "expected #NUM! for worksheet ODDFPRICE when yld < -frequency, got {out:?}"
    );

    let oddl_boundary =
        "=ODDLPRICE(DATE(2020,7,15),DATE(2033,1,15),DATE(2020,1,15),0.05,-2,100,2,0)";
    let oddl_below =
        "=ODDLPRICE(DATE(2020,7,15),DATE(2033,1,15),DATE(2020,1,15),0.05,-2.5,100,2,0)";

    let Some(out) = eval_value_or_skip(&mut sheet, oddl_boundary) else {
        return;
    };
    assert!(
        matches!(out, Value::Error(ErrorKind::Div0)),
        "expected #DIV/0! for worksheet ODDLPRICE at yld=-frequency, got {out:?}"
    );

    let Some(out) = eval_value_or_skip(&mut sheet, oddl_below) else {
        return;
    };
    assert!(
        matches!(out, Value::Error(ErrorKind::Num)),
        "expected #NUM! for worksheet ODDLPRICE when yld < -frequency, got {out:?}"
    );
}

#[test]
fn odd_coupon_bond_price_allows_negative_yield() {
    let mut sheet = TestSheet::new();

    // Ensure worksheet wrappers accept negative yields below -1 (but still within the valid per-period
    // domain yld > -frequency when frequency=2).
    for yld in [-0.5, -1.5] {
        let oddf = format!(
            "=ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,{yld},100,2,0)"
        );
        let oddl = format!(
            "=ODDLPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,{yld},100,2,0)"
        );

        let oddf_price = match eval_number_or_skip(&mut sheet, &oddf) {
            Some(v) => v,
            None => return,
        };
        let oddl_price = eval_number_or_skip(&mut sheet, &oddl)
            .expect("ODDLPRICE should return a number for negative yld within (-frequency, ∞)");

        assert!(
            oddf_price.is_finite(),
            "expected finite price, got {oddf_price}"
        );
        assert!(
            oddl_price.is_finite(),
            "expected finite price, got {oddl_price}"
        );
    }
}

#[test]
fn odd_coupon_yield_roundtrips_negative_yield() {
    let system = ExcelDateSystem::EXCEL_1900;

    // ODDF* (use a zero-coupon case for numerical stability).
    let issue = ymd_to_serial(ExcelDate::new(2019, 10, 1), system).unwrap();
    let settlement = ymd_to_serial(ExcelDate::new(2020, 1, 1), system).unwrap();
    let first_coupon = ymd_to_serial(ExcelDate::new(2020, 7, 1), system).unwrap();
    let maturity = ymd_to_serial(ExcelDate::new(2021, 7, 1), system).unwrap();

    let rate = 0.0;
    let redemption = 100.0;
    let frequency = 2;
    let basis = 0;

    // Exercise both "mild" and "very" negative yields (the latter is < -1.0 but still valid when
    // `frequency > 1`).
    for yld in [-0.5, -1.5] {
        let price = oddfprice(
            settlement,
            maturity,
            issue,
            first_coupon,
            rate,
            yld,
            redemption,
            frequency,
            basis,
            system,
        )
        .unwrap();
        assert!(price.is_finite(), "expected finite ODDFPRICE, got {price}");

        let solved = oddfyield(
            settlement,
            maturity,
            issue,
            first_coupon,
            rate,
            price,
            redemption,
            frequency,
            basis,
            system,
        )
        .unwrap();
        assert_close(solved, yld, 1e-6);

        // ODDL*
        let last_interest = ymd_to_serial(ExcelDate::new(2021, 1, 1), system).unwrap();
        let settlement2 = ymd_to_serial(ExcelDate::new(2021, 2, 1), system).unwrap();
        let maturity2 = ymd_to_serial(ExcelDate::new(2021, 5, 1), system).unwrap();

        let price2 = oddlprice(
            settlement2,
            maturity2,
            last_interest,
            rate,
            yld,
            redemption,
            frequency,
            basis,
            system,
        )
        .unwrap();
        assert!(
            price2.is_finite(),
            "expected finite ODDLPRICE, got {price2}"
        );

        let solved2 = oddlyield(
            settlement2,
            maturity2,
            last_interest,
            rate,
            price2,
            redemption,
            frequency,
            basis,
            system,
        )
        .unwrap();
        assert_close(solved2, yld, 1e-6);
    }
}

#[test]
fn odd_coupon_prices_are_finite_for_large_redemption_values() {
    let system = ExcelDateSystem::EXCEL_1900;

    // Reuse the existing odd first coupon setup.
    let issue = ymd_to_serial(ExcelDate::new(2020, 1, 1), system).unwrap();
    let settlement = ymd_to_serial(ExcelDate::new(2020, 1, 15), system).unwrap();
    let first_coupon = ymd_to_serial(ExcelDate::new(2020, 7, 15), system).unwrap();
    let maturity = ymd_to_serial(ExcelDate::new(2023, 1, 15), system).unwrap();

    let rate = 0.05;
    let yld = 0.06;
    let redemption = 1e12;
    let frequency = 2;
    let basis = 0;

    let price = oddfprice(
        settlement,
        maturity,
        issue,
        first_coupon,
        rate,
        yld,
        redemption,
        frequency,
        basis,
        system,
    )
    .expect("ODDFPRICE should succeed for large finite redemption");
    assert!(price.is_finite(), "expected finite price, got {price}");

    // Odd last coupon setup.
    let last_interest = ymd_to_serial(ExcelDate::new(2022, 7, 15), system).unwrap();
    let settlement_last = ymd_to_serial(ExcelDate::new(2022, 10, 15), system).unwrap();
    let maturity_last = ymd_to_serial(ExcelDate::new(2023, 1, 15), system).unwrap();

    let price_last = oddlprice(
        settlement_last,
        maturity_last,
        last_interest,
        rate,
        yld,
        redemption,
        frequency,
        basis,
        system,
    )
    .expect("ODDLPRICE should succeed for large finite redemption");
    assert!(
        price_last.is_finite(),
        "expected finite price, got {price_last}"
    );
}

#[test]
fn odd_coupon_bond_functions_reject_negative_coupon_rate() {
    let mut sheet = TestSheet::new();

    let oddf = "=ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),-0.01,0.0625,100,2,0)";
    let oddl = "=ODDLPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),-0.01,0.0625,100,2,0)";
    let oddfy = "=ODDFYIELD(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),-0.01,98,100,2,0)";
    let oddly = "=ODDLYIELD(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),-0.01,98,100,2,0)";

    let Some(out) = eval_value_or_skip(&mut sheet, oddf) else {
        return;
    };
    assert!(
        matches!(out, Value::Error(ErrorKind::Num)),
        "expected #NUM! for negative rate in ODDFPRICE, got {out:?}"
    );

    let Some(out) = eval_value_or_skip(&mut sheet, oddl) else {
        return;
    };
    assert!(
        matches!(out, Value::Error(ErrorKind::Num)),
        "expected #NUM! for negative rate in ODDLPRICE, got {out:?}"
    );

    let Some(out) = eval_value_or_skip(&mut sheet, oddfy) else {
        return;
    };
    assert!(
        matches!(out, Value::Error(ErrorKind::Num)),
        "expected #NUM! for negative rate in ODDFYIELD, got {out:?}"
    );

    let Some(out) = eval_value_or_skip(&mut sheet, oddly) else {
        return;
    };
    assert!(
        matches!(out, Value::Error(ErrorKind::Num)),
        "expected #NUM! for negative rate in ODDLYIELD, got {out:?}"
    );
}

#[test]
fn odd_yield_solver_falls_back_when_derivative_is_non_finite() {
    let system = ExcelDateSystem::EXCEL_1900;

    // Construct a case where the Newton step fails because the analytic derivative overflows at the
    // default guess (0.1), but the price itself remains finite.
    let issue = ymd_to_serial(ExcelDate::new(2020, 1, 1), system).unwrap();
    let settlement = ymd_to_serial(ExcelDate::new(2020, 1, 15), system).unwrap();
    let first_coupon = ymd_to_serial(ExcelDate::new(2020, 7, 15), system).unwrap();
    let maturity = ymd_to_serial(ExcelDate::new(2033, 1, 15), system).unwrap();

    let rate = 0.05;
    let frequency = 2;
    let basis = 0;
    let redemption = 1e308;
    let target_yield = 0.1;

    let pr = oddfprice(
        settlement,
        maturity,
        issue,
        first_coupon,
        rate,
        target_yield,
        redemption,
        frequency,
        basis,
        system,
    )
    .expect("ODDFPRICE should be finite for the target yield");

    let recovered = oddfyield(
        settlement,
        maturity,
        issue,
        first_coupon,
        rate,
        pr,
        redemption,
        frequency,
        basis,
        system,
    )
    .expect("ODDFYIELD should converge via bisection fallback");

    assert_close(recovered, target_yield, 1e-6);

    // Repeat for the odd last coupon solver.
    let last_interest = ymd_to_serial(ExcelDate::new(2020, 1, 15), system).unwrap();
    let settlement_last = ymd_to_serial(ExcelDate::new(2020, 7, 15), system).unwrap();
    let maturity_last = ymd_to_serial(ExcelDate::new(2033, 1, 15), system).unwrap();

    let pr_last = oddlprice(
        settlement_last,
        maturity_last,
        last_interest,
        rate,
        target_yield,
        redemption,
        frequency,
        basis,
        system,
    )
    .expect("ODDLPRICE should be finite for the target yield");

    let recovered_last = oddlyield(
        settlement_last,
        maturity_last,
        last_interest,
        rate,
        pr_last,
        redemption,
        frequency,
        basis,
        system,
    )
    .expect("ODDLYIELD should converge via bisection fallback");

    assert_close(recovered_last, target_yield, 1e-6);
}

#[test]
fn odd_coupon_bond_yield_can_be_negative() {
    let mut sheet = TestSheet::new();

    // A price above the undiscounted cashflows implies a negative yield when yields are allowed
    // below 0. The yield domain matches the per-period discount base (`1 + yld/frequency > 0`),
    // i.e. `yld > -frequency`.
    let oddf = "=ODDFYIELD(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,300,100,2,0)";
    let oddl = "=ODDLYIELD(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,300,100,2,0)";

    let oddf_yld = match eval_number_or_skip(&mut sheet, oddf) {
        Some(v) => v,
        None => return,
    };
    let oddl_yld = eval_number_or_skip(&mut sheet, oddl).expect("ODDLYIELD should return a number");

    assert!(
        oddf_yld < 0.0 && oddf_yld > -2.0,
        "expected ODDFYIELD to return a negative yield in (-2, 0), got {oddf_yld}"
    );
    assert!(
        oddl_yld < 0.0 && oddl_yld > -2.0,
        "expected ODDLYIELD to return a negative yield in (-2, 0), got {oddl_yld}"
    );
    assert!(
        oddl_yld < -1.0,
        "expected ODDLYIELD to exercise yields below -1.0 for frequency=2; got {oddl_yld}"
    );
}

#[test]
fn odd_coupon_bond_price_allows_zero_coupon_rate() {
    let mut sheet = TestSheet::new();

    // Zero-coupon odd-first/odd-last cases should still be valid.
    let oddf = "=ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0,0.0625,100,2,0)";
    let oddl = "=ODDLPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0,0.0625,100,2,0)";

    let Some(out) = eval_value_or_skip(&mut sheet, oddf) else {
        return;
    };
    assert!(
        matches!(out, Value::Number(_)),
        "expected a numeric price for ODDFPRICE with rate=0, got {out:?}"
    );

    let Some(out) = eval_value_or_skip(&mut sheet, oddl) else {
        return;
    };
    assert!(
        matches!(out, Value::Number(_)),
        "expected a numeric price for ODDLPRICE with rate=0, got {out:?}"
    );
}

#[test]
fn odd_last_coupon_bond_functions_round_trip_long_stub() {
    let mut sheet = TestSheet::new();

    // Long odd-last coupon period:
    // - last_interest is far before maturity so DSM/E > 1 (long stub)
    // - settlement is between last_interest and maturity
    //
    // We keep the schedule simple: there are no regular coupon payments between settlement
    // and maturity, so the functions must correctly scale the final coupon amount.
    let yield_target = 0.0625;
    let rate = 0.0785;
    let system = ExcelDateSystem::EXCEL_1900;

    let settlement = serial(2021, 2, 1, system);
    let maturity = serial(2022, 3, 1, system);
    let last_interest = serial(2020, 10, 15, system);

    for basis in [0, 1] {
        // Ensure this is actually a *long* last coupon period for this basis.
        // (DLM/E > 1 and DSM/E > 1).
        let dlm = days_between(last_interest, maturity, basis, system);
        let dsm = days_between(settlement, maturity, basis, system);
        let months_per_period = 12 / 2;
        let eom = is_end_of_month(last_interest, system);
        let prev_coupon = coupon_date_with_eom(last_interest, -months_per_period, eom, system);
        let e = coupon_period_e(prev_coupon, last_interest, basis, 2, system);
        assert!(
            dlm / e > 1.0,
            "expected long last coupon (DLM/E > 1), got DLM={dlm} E={e}"
        );
        assert!(
            dsm / e > 1.0,
            "expected settlement to be more than one coupon period before maturity (DSM/E > 1), got DSM={dsm} E={e}"
        );

        let expected_price = oddl_price_excel_model(
            settlement,
            maturity,
            last_interest,
            rate,
            yield_target,
            100.0,
            2,
            basis,
            system,
        );

        let price_formula = format!(
            "=ODDLPRICE(DATE(2021,2,1),DATE(2022,3,1),DATE(2020,10,15),{rate},{yield_target},100,2,{basis})"
        );
        sheet.set_formula("A1", &price_formula);
        sheet.recalc();

        let Some(price) = cell_number_or_skip(&sheet, "A1") else {
            return;
        };

        assert_close(price, expected_price, 1e-10);

        assert!(
            price.is_finite() && price > 0.0,
            "expected positive finite price, got {price}"
        );

        sheet.set("B1", expected_price);
        let yield_formula = format!(
            "=ODDLYIELD(DATE(2021,2,1),DATE(2022,3,1),DATE(2020,10,15),{rate},B1,100,2,{basis})"
        );
        let Some(y) = eval_number_or_skip(&mut sheet, &yield_formula) else {
            return;
        };
        assert_close(y, yield_target, 1e-9);
    }
}

#[test]
fn oddfprice_matches_excel_model_for_actual_day_bases() {
    let system = ExcelDateSystem::EXCEL_1900;

    // Maturity on the 30th, with the first coupon clamped in February (leap year).
    // This case exercises Excel's maturity-anchored schedule stepping and day-count
    // conventions for bases 1/2/3.
    let issue = ymd_to_serial(ExcelDate::new(2020, 1, 15), system).unwrap();
    let settlement = ymd_to_serial(ExcelDate::new(2020, 1, 20), system).unwrap();
    let first_coupon = ymd_to_serial(ExcelDate::new(2020, 2, 29), system).unwrap();
    let maturity = ymd_to_serial(ExcelDate::new(2020, 8, 30), system).unwrap();

    let rate = 0.08;
    let yld = 0.075;
    let redemption = 100.0;
    let frequency = 2;

    for basis in [1, 2, 3] {
        let expected = oddf_price_excel_model(
            settlement,
            maturity,
            issue,
            first_coupon,
            rate,
            yld,
            redemption,
            frequency,
            basis,
            system,
        );

        let actual = oddfprice(
            settlement,
            maturity,
            issue,
            first_coupon,
            rate,
            yld,
            redemption,
            frequency,
            basis,
            system,
        )
        .unwrap();

        assert_close(actual, expected, 1e-10);

        let recovered = oddfyield(
            settlement,
            maturity,
            issue,
            first_coupon,
            rate,
            expected,
            redemption,
            frequency,
            basis,
            system,
        )
        .unwrap();
        assert_close(recovered, yld, 1e-10);
    }
}

#[test]
fn oddlprice_matches_excel_model_for_actual_day_bases_with_eom_last_interest() {
    let system = ExcelDateSystem::EXCEL_1900;

    // Last interest date is an end-of-month date (April 30). Excel treats this as an EOM
    // coupon schedule, so the prior regular coupon date is January 31 (not January 30).
    let last_interest = ymd_to_serial(ExcelDate::new(2021, 4, 30), system).unwrap();
    let settlement = ymd_to_serial(ExcelDate::new(2021, 5, 15), system).unwrap();
    let maturity = ymd_to_serial(ExcelDate::new(2021, 6, 15), system).unwrap();

    let rate = 0.06;
    let yld = 0.055;
    let redemption = 100.0;
    let frequency = 4;

    for basis in [1, 2, 3] {
        let expected = oddl_price_excel_model(
            settlement,
            maturity,
            last_interest,
            rate,
            yld,
            redemption,
            frequency,
            basis,
            system,
        );

        let actual = oddlprice(
            settlement,
            maturity,
            last_interest,
            rate,
            yld,
            redemption,
            frequency,
            basis,
            system,
        )
        .unwrap();

        assert_close(actual, expected, 1e-10);

        let recovered = oddlyield(
            settlement,
            maturity,
            last_interest,
            rate,
            expected,
            redemption,
            frequency,
            basis,
            system,
        )
        .unwrap();
        assert_close(recovered, yld, 1e-10);
    }
}

#[test]
fn oddfprice_matches_excel_model_for_30_360_bases() {
    let system = ExcelDateSystem::EXCEL_1900;
    // Month-end / Feb dates so 30/360 US (basis=0) vs European (basis=4) diverge.
    //
    // Under the odd-coupon conventions:
    // - basis=0 uses a fixed `E = 360/frequency`.
    // - basis=4 uses European 30E/360 (`DAYS360(..., TRUE)`) for day counts like `A` and `DSC`, but
    //   still uses a fixed `E = 360/frequency` (matching Excel's COUPDAYS behavior). This means `E`
    //   can differ from `DAYS360(PCD, NCD, TRUE)` for some end-of-month schedules involving
    //   February.
    //
    // Include two scenarios:
    // - one where European DAYS360 between coupon dates matches `360/frequency`
    // - one where `DAYS360(PCD, NCD, TRUE) != 360/frequency` (to guard that basis=4 does *not* use
    //   European DAYS360 for `E`)
    let scenarios = [
        // issue=2019-01-31, settlement=2019-02-28, first_coupon=2019-03-31, maturity=2019-09-30
        (
            ExcelDate::new(2019, 1, 31),
            ExcelDate::new(2019, 2, 28),
            ExcelDate::new(2019, 3, 31),
            ExcelDate::new(2019, 9, 30),
        ),
        // maturity=2019-02-28 is EOM, so the maturity-anchored schedule is EOM-pinned:
        // prev_coupon=2018-02-28, first_coupon=2018-08-31.
        // For basis=4, DAYS360(2018-02-28, 2018-08-31, method=true) = 182 days (not 180), but `E`
        // is still modeled as a fixed `360/frequency` (180 for semiannual coupons).
        (
            ExcelDate::new(2018, 1, 31),
            ExcelDate::new(2018, 2, 15),
            ExcelDate::new(2018, 8, 31),
            ExcelDate::new(2019, 2, 28),
        ),
    ];

    let rate = 0.05;
    let yld = 0.06;
    let redemption = 100.0;
    let frequency = 2;

    for (issue, settlement, first_coupon, maturity) in scenarios {
        let issue = ymd_to_serial(issue, system).unwrap();
        let settlement = ymd_to_serial(settlement, system).unwrap();
        let first_coupon = ymd_to_serial(first_coupon, system).unwrap();
        let maturity = ymd_to_serial(maturity, system).unwrap();

        // Guard: ensure we cover a basis=4 EOM schedule where `DAYS360(PCD, NCD, TRUE) != 360/frequency`.
        if maturity == ymd_to_serial(ExcelDate::new(2019, 2, 28), system).unwrap() {
            let months_per_period = 12 / frequency;
            let eom = is_end_of_month(maturity, system);
            assert!(eom, "expected maturity to be EOM for this scenario");
            let prev_coupon = coupon_date_with_eom(first_coupon, -months_per_period, eom, system);
            let days360_eu =
                date_time::days360(prev_coupon, first_coupon, true, system).unwrap() as f64;
            assert_close(days360_eu, 182.0, 0.0);
            let e_fixed = 360.0 / (frequency as f64);
            let e4 = coupon_period_e(prev_coupon, first_coupon, 4, frequency, system);
            assert_close(e4, e_fixed, 0.0);
            assert!(
                (days360_eu - e_fixed).abs() > 0.0,
                "expected DAYS360(PCD, NCD, TRUE) != 360/frequency for this scenario"
            );
        }
        // Guard: ensure US vs EU DAYS360 behavior actually diverges in this scenario.
        let a0 = days_between(issue, settlement, 0, system);
        let a4 = days_between(issue, settlement, 4, system);
        let dsc0 = days_between(settlement, first_coupon, 0, system);
        let dsc4 = days_between(settlement, first_coupon, 4, system);
        assert!(
            a0 != a4 || dsc0 != dsc4,
            "expected DAYS360 method=false vs method=true to diverge for this scenario"
        );

        // (The above guard asserts `E == 360/frequency` and `E != DAYS360(PCD, NCD, TRUE)` for basis=4.)
        for basis in [0, 4] {
            let expected = oddf_price_excel_model(
                settlement,
                maturity,
                issue,
                first_coupon,
                rate,
                yld,
                redemption,
                frequency,
                basis,
                system,
            );

            let actual = oddfprice(
                settlement,
                maturity,
                issue,
                first_coupon,
                rate,
                yld,
                redemption,
                frequency,
                basis,
                system,
            )
            .unwrap();

            assert!(
                (actual - expected).abs() <= 1e-10,
                "basis {basis}: expected {expected}, got {actual}"
            );

            let recovered = oddfyield(
                settlement,
                maturity,
                issue,
                first_coupon,
                rate,
                expected,
                redemption,
                frequency,
                basis,
                system,
            )
            .unwrap();
            assert!(
                (recovered - yld).abs() <= 1e-9,
                "basis {basis}: expected yield {yld}, got {recovered}"
            );
        }
    }
}

#[test]
fn oddlprice_matches_excel_model_for_30_360_bases() {
    let system = ExcelDateSystem::EXCEL_1900;

    // Month-end / Feb dates so 30/360 US vs EU diverge (e.g. Mar 31 handling).
    // last_interest=2019-02-28, settlement=2019-03-15, maturity=2019-03-31
    let last_interest = ymd_to_serial(ExcelDate::new(2019, 2, 28), system).unwrap();
    let settlement = ymd_to_serial(ExcelDate::new(2019, 3, 15), system).unwrap();
    let maturity = ymd_to_serial(ExcelDate::new(2019, 3, 31), system).unwrap();

    let rate = 0.05;
    let yld = 0.06;
    let redemption = 100.0;
    let frequency = 2;

    // Guard: ensure this scenario exercises a schedule where European DAYS360 between coupon
    // dates differs from the fixed `E = 360/frequency`.
    //
    // last_interest is EOM Feb 28, so the EOM-pinned prior coupon date is Aug 31.
    // Under European DAYS360, this period is 178 days (not 180), but the odd-coupon bond functions
    // still use `E = 360/frequency`.
    let months_per_period = 12 / frequency;
    let eom = is_end_of_month(last_interest, system);
    assert!(eom, "expected last_interest to be EOM for this scenario");
    let prev_coupon = coupon_date_with_eom(last_interest, -months_per_period, eom, system);
    let days360_eu = date_time::days360(prev_coupon, last_interest, true, system).unwrap() as f64;
    assert_close(days360_eu, 178.0, 0.0);
    let e_fixed = 360.0 / (frequency as f64);
    let e4 = coupon_period_e(prev_coupon, last_interest, 4, frequency, system);
    assert_close(e4, e_fixed, 0.0);
    assert!(
        (days360_eu - e_fixed).abs() > 0.0,
        "expected DAYS360(PCD, NCD, TRUE) != 360/frequency for this scenario"
    );

    // Guard: ensure this scenario actually exercises a case where DAYS360 US vs European differ.
    let a0 = days_between(last_interest, settlement, 0, system);
    let a4 = days_between(last_interest, settlement, 4, system);
    let dsm0 = days_between(settlement, maturity, 0, system);
    let dsm4 = days_between(settlement, maturity, 4, system);
    assert!(
        a0 != a4 || dsm0 != dsm4,
        "expected DAYS360 method=false vs method=true to diverge for this scenario"
    );

    for basis in [0, 4] {
        let expected = oddl_price_excel_model(
            settlement,
            maturity,
            last_interest,
            rate,
            yld,
            redemption,
            frequency,
            basis,
            system,
        );

        let actual = oddlprice(
            settlement,
            maturity,
            last_interest,
            rate,
            yld,
            redemption,
            frequency,
            basis,
            system,
        )
        .unwrap();

        assert!(
            (actual - expected).abs() <= 1e-10,
            "basis {basis}: expected {expected}, got {actual}"
        );

        let recovered = oddlyield(
            settlement,
            maturity,
            last_interest,
            rate,
            expected,
            redemption,
            frequency,
            basis,
            system,
        )
        .unwrap();
        assert!(
            (recovered - yld).abs() <= 1e-9,
            "basis {basis}: expected yield {yld}, got {recovered}"
        );
    }
}

#[test]
fn odd_coupon_bond_functions_reject_non_finite_numeric_inputs() {
    let mut sheet = TestSheet::new();

    // Use explicit non-finite cell values so the test doesn't depend on the engine's arithmetic
    // overflow/NaN behavior.
    sheet.set("A1", Value::Number(f64::INFINITY));
    sheet.set("A2", Value::Number(f64::NAN));
    sheet.set("A3", Value::Number(f64::NEG_INFINITY));

    // Ensure the test is actually exercising non-finite numbers (not pre-coerced errors).
    match sheet.get("A1") {
        Value::Number(n) => assert!(n.is_infinite(), "expected A1 to be +Inf, got {n:?}"),
        other => panic!("expected A1 to be a number (+Inf), got {other:?}"),
    }
    match sheet.get("A2") {
        Value::Number(n) => assert!(n.is_nan(), "expected A2 to be NaN, got {n:?}"),
        other => panic!("expected A2 to be a number (NaN), got {other:?}"),
    }
    match sheet.get("A3") {
        Value::Number(n) => assert!(
            n.is_infinite() && n.is_sign_negative(),
            "expected A3 to be -Inf, got {n:?}"
        ),
        other => panic!("expected A3 to be a number (-Inf), got {other:?}"),
    }

    // Infinity in rate.
    match sheet.eval(
        "=ODDFPRICE(DATE(2020,3,1),DATE(2023,7,1),DATE(2020,1,1),DATE(2020,7,1),A1,0.05,100,1,0)",
    ) {
        Value::Error(ErrorKind::Num) => {}
        other => panic!("expected #NUM!, got {other:?}"),
    }

    // Infinity in yld.
    match sheet.eval(
        "=ODDFPRICE(DATE(2020,3,1),DATE(2023,7,1),DATE(2020,1,1),DATE(2020,7,1),0.06,A1,100,1,0)",
    ) {
        Value::Error(ErrorKind::Num) => {}
        other => panic!("expected #NUM!, got {other:?}"),
    }

    // NaN in redemption.
    match sheet.eval(
        "=ODDFPRICE(DATE(2020,3,1),DATE(2023,7,1),DATE(2020,1,1),DATE(2020,7,1),0.06,0.05,A2,1,0)",
    ) {
        Value::Error(ErrorKind::Num) => {}
        other => panic!("expected #NUM!, got {other:?}"),
    }

    // NaN in frequency.
    match sheet.eval("=ODDFPRICE(DATE(2020,3,1),DATE(2023,7,1),DATE(2020,1,1),DATE(2020,7,1),0.06,0.05,100,A2,0)") {
        Value::Error(ErrorKind::Num) => {}
        other => panic!("expected #NUM!, got {other:?}"),
    }

    // NaN in basis.
    match sheet.eval("=ODDFPRICE(DATE(2020,3,1),DATE(2023,7,1),DATE(2020,1,1),DATE(2020,7,1),0.06,0.05,100,1,A2)") {
        Value::Error(ErrorKind::Num) => {}
        other => panic!("expected #NUM!, got {other:?}"),
    }

    // Non-finite dates should also be rejected (#NUM!).
    match sheet
        .eval("=ODDFPRICE(A1,DATE(2023,7,1),DATE(2020,1,1),DATE(2020,7,1),0.06,0.05,100,1,0)")
    {
        Value::Error(ErrorKind::Num) => {}
        other => panic!("expected #NUM!, got {other:?}"),
    }
    match sheet
        .eval("=ODDFPRICE(DATE(2020,3,1),A1,DATE(2020,1,1),DATE(2020,7,1),0.06,0.05,100,1,0)")
    {
        Value::Error(ErrorKind::Num) => {}
        other => panic!("expected #NUM!, got {other:?}"),
    }
    match sheet
        .eval("=ODDFPRICE(DATE(2020,3,1),DATE(2023,7,1),A2,DATE(2020,7,1),0.06,0.05,100,1,0)")
    {
        Value::Error(ErrorKind::Num) => {}
        other => panic!("expected #NUM!, got {other:?}"),
    }
    match sheet
        .eval("=ODDFPRICE(DATE(2020,3,1),DATE(2023,7,1),DATE(2020,1,1),A1,0.06,0.05,100,1,0)")
    {
        Value::Error(ErrorKind::Num) => {}
        other => panic!("expected #NUM!, got {other:?}"),
    }

    match sheet.eval("=ODDLPRICE(DATE(2022,11,1),A2,DATE(2022,7,1),0.06,0.05,100,1,0)") {
        Value::Error(ErrorKind::Num) => {}
        other => panic!("expected #NUM!, got {other:?}"),
    }
    match sheet.eval("=ODDLPRICE(A1,DATE(2023,3,1),DATE(2022,7,1),0.06,0.05,100,1,0)") {
        Value::Error(ErrorKind::Num) => {}
        other => panic!("expected #NUM!, got {other:?}"),
    }
    match sheet.eval("=ODDLPRICE(DATE(2022,11,1),DATE(2023,3,1),A1,0.06,0.05,100,1,0)") {
        Value::Error(ErrorKind::Num) => {}
        other => panic!("expected #NUM!, got {other:?}"),
    }
    match sheet.eval("=ODDLPRICE(DATE(2022,11,1),DATE(2023,3,1),DATE(2022,7,1),A1,0.05,100,1,0)") {
        Value::Error(ErrorKind::Num) => {}
        other => panic!("expected #NUM!, got {other:?}"),
    }

    // NaN in yield.
    match sheet.eval("=ODDLPRICE(DATE(2022,11,1),DATE(2023,3,1),DATE(2022,7,1),0.06,A2,100,1,0)") {
        Value::Error(ErrorKind::Num) => {}
        other => panic!("expected #NUM!, got {other:?}"),
    }

    // Infinity in redemption.
    match sheet.eval("=ODDLPRICE(DATE(2022,11,1),DATE(2023,3,1),DATE(2022,7,1),0.06,0.05,A1,1,0)") {
        Value::Error(ErrorKind::Num) => {}
        other => panic!("expected #NUM!, got {other:?}"),
    }
    match sheet.eval("=ODDLPRICE(DATE(2022,11,1),DATE(2023,3,1),DATE(2022,7,1),0.06,0.05,100,A1,0)")
    {
        Value::Error(ErrorKind::Num) => {}
        other => panic!("expected #NUM!, got {other:?}"),
    }
    match sheet.eval("=ODDLPRICE(DATE(2022,11,1),DATE(2023,3,1),DATE(2022,7,1),0.06,0.05,100,1,A2)")
    {
        Value::Error(ErrorKind::Num) => {}
        other => panic!("expected #NUM!, got {other:?}"),
    }

    // Infinity in price.
    match sheet.eval(
        "=ODDFYIELD(DATE(2020,3,1),DATE(2023,7,1),DATE(2020,1,1),DATE(2020,7,1),0.06,A1,100,1,0)",
    ) {
        Value::Error(ErrorKind::Num) => {}
        other => panic!("expected #NUM!, got {other:?}"),
    }

    // NaN in price.
    match sheet.eval(
        "=ODDFYIELD(DATE(2020,3,1),DATE(2023,7,1),DATE(2020,1,1),DATE(2020,7,1),0.06,A2,100,1,0)",
    ) {
        Value::Error(ErrorKind::Num) => {}
        other => panic!("expected #NUM!, got {other:?}"),
    }

    // Non-finite dates in ODDFYIELD should also return #NUM!.
    match sheet.eval("=ODDFYIELD(A1,DATE(2023,7,1),DATE(2020,1,1),DATE(2020,7,1),0.06,95,100,1,0)")
    {
        Value::Error(ErrorKind::Num) => {}
        other => panic!("expected #NUM!, got {other:?}"),
    }
    match sheet.eval("=ODDFYIELD(DATE(2020,3,1),A1,DATE(2020,1,1),DATE(2020,7,1),0.06,95,100,1,0)")
    {
        Value::Error(ErrorKind::Num) => {}
        other => panic!("expected #NUM!, got {other:?}"),
    }
    match sheet.eval("=ODDFYIELD(DATE(2020,3,1),DATE(2023,7,1),A2,DATE(2020,7,1),0.06,95,100,1,0)")
    {
        Value::Error(ErrorKind::Num) => {}
        other => panic!("expected #NUM!, got {other:?}"),
    }
    match sheet.eval("=ODDFYIELD(DATE(2020,3,1),DATE(2023,7,1),DATE(2020,1,1),A1,0.06,95,100,1,0)")
    {
        Value::Error(ErrorKind::Num) => {}
        other => panic!("expected #NUM!, got {other:?}"),
    }

    // Non-finite numeric args in ODDFYIELD should return #NUM!.
    match sheet.eval(
        "=ODDFYIELD(DATE(2020,3,1),DATE(2023,7,1),DATE(2020,1,1),DATE(2020,7,1),A3,95,100,1,0)",
    ) {
        Value::Error(ErrorKind::Num) => {}
        other => panic!("expected #NUM!, got {other:?}"),
    }
    match sheet.eval(
        "=ODDFYIELD(DATE(2020,3,1),DATE(2023,7,1),DATE(2020,1,1),DATE(2020,7,1),0.06,95,A2,1,0)",
    ) {
        Value::Error(ErrorKind::Num) => {}
        other => panic!("expected #NUM!, got {other:?}"),
    }
    match sheet.eval(
        "=ODDFYIELD(DATE(2020,3,1),DATE(2023,7,1),DATE(2020,1,1),DATE(2020,7,1),0.06,95,100,A1,0)",
    ) {
        Value::Error(ErrorKind::Num) => {}
        other => panic!("expected #NUM!, got {other:?}"),
    }
    match sheet.eval(
        "=ODDFYIELD(DATE(2020,3,1),DATE(2023,7,1),DATE(2020,1,1),DATE(2020,7,1),0.06,95,100,1,A2)",
    ) {
        Value::Error(ErrorKind::Num) => {}
        other => panic!("expected #NUM!, got {other:?}"),
    }

    // ODDLYIELD should also reject non-finite numeric args.
    match sheet.eval("=ODDLYIELD(DATE(2022,11,1),DATE(2023,3,1),DATE(2022,7,1),0.06,A1,100,1,0)") {
        Value::Error(ErrorKind::Num) => {}
        other => panic!("expected #NUM!, got {other:?}"),
    }
    match sheet.eval("=ODDLYIELD(DATE(2022,11,1),DATE(2023,3,1),DATE(2022,7,1),0.06,A2,100,1,0)") {
        Value::Error(ErrorKind::Num) => {}
        other => panic!("expected #NUM!, got {other:?}"),
    }

    // Non-finite dates in ODDLYIELD should also return #NUM!.
    match sheet.eval("=ODDLYIELD(A1,DATE(2023,3,1),DATE(2022,7,1),0.06,95,100,1,0)") {
        Value::Error(ErrorKind::Num) => {}
        other => panic!("expected #NUM!, got {other:?}"),
    }
    match sheet.eval("=ODDLYIELD(DATE(2022,11,1),A1,DATE(2022,7,1),0.06,95,100,1,0)") {
        Value::Error(ErrorKind::Num) => {}
        other => panic!("expected #NUM!, got {other:?}"),
    }
    match sheet.eval("=ODDLYIELD(DATE(2022,11,1),DATE(2023,3,1),A2,0.06,95,100,1,0)") {
        Value::Error(ErrorKind::Num) => {}
        other => panic!("expected #NUM!, got {other:?}"),
    }

    // Non-finite numeric args in ODDLYIELD should return #NUM!.
    match sheet.eval("=ODDLYIELD(DATE(2022,11,1),DATE(2023,3,1),DATE(2022,7,1),A3,95,100,1,0)") {
        Value::Error(ErrorKind::Num) => {}
        other => panic!("expected #NUM!, got {other:?}"),
    }
    match sheet.eval("=ODDLYIELD(DATE(2022,11,1),DATE(2023,3,1),DATE(2022,7,1),0.06,95,A1,1,0)") {
        Value::Error(ErrorKind::Num) => {}
        other => panic!("expected #NUM!, got {other:?}"),
    }
    match sheet.eval("=ODDLYIELD(DATE(2022,11,1),DATE(2023,3,1),DATE(2022,7,1),0.06,95,100,A2,0)") {
        Value::Error(ErrorKind::Num) => {}
        other => panic!("expected #NUM!, got {other:?}"),
    }
    match sheet.eval("=ODDLYIELD(DATE(2022,11,1),DATE(2023,3,1),DATE(2022,7,1),0.06,95,100,1,A1)") {
        Value::Error(ErrorKind::Num) => {}
        other => panic!("expected #NUM!, got {other:?}"),
    }
}

#[test]
fn odd_coupon_bond_functions_return_value_for_unparseable_date_text() {
    let mut sheet = TestSheet::new();

    // Unparseable (non-date-like) text.
    match sheet
        .eval(r#"=ODDFPRICE("nope",DATE(2025,1,1),DATE(2019,1,1),DATE(2020,7,1),0.05,0.05,100,2)"#)
    {
        Value::Error(ErrorKind::Value) => {}
        other => panic!("expected #VALUE!, got {other:?}"),
    }

    match sheet.eval(r#"=ODDLPRICE(DATE(2020,1,1),"nope",DATE(2024,7,1),0.05,0.05,100,2)"#) {
        Value::Error(ErrorKind::Value) => {}
        other => panic!("expected #VALUE!, got {other:?}"),
    }

    // Parseable-but-invalid dates should also return #VALUE! (e.g. Feb 30).
    match sheet.eval(
        r#"=ODDFPRICE("2020-02-30",DATE(2025,1,1),DATE(2019,1,1),DATE(2020,7,1),0.05,0.05,100,2)"#,
    ) {
        Value::Error(ErrorKind::Value) => {}
        other => panic!("expected #VALUE!, got {other:?}"),
    }
    match sheet.eval(
        r#"=ODDFYIELD("2020-02-30",DATE(2025,1,1),DATE(2019,1,1),DATE(2020,7,1),0.05,95,100,2)"#,
    ) {
        Value::Error(ErrorKind::Value) => {}
        other => panic!("expected #VALUE!, got {other:?}"),
    }
    match sheet.eval(r#"=ODDLPRICE(DATE(2020,1,1),"2020-02-30",DATE(2024,7,1),0.05,0.05,100,2)"#) {
        Value::Error(ErrorKind::Value) => {}
        other => panic!("expected #VALUE!, got {other:?}"),
    }
    match sheet.eval(r#"=ODDLYIELD(DATE(2020,1,1),"2020-02-30",DATE(2024,7,1),0.05,95,100,2)"#) {
        Value::Error(ErrorKind::Value) => {}
        other => panic!("expected #VALUE!, got {other:?}"),
    }

    // Other date positions.
    match sheet.eval(
        r#"=ODDFPRICE(DATE(2020,3,1),"nope",DATE(2020,1,1),DATE(2020,7,1),0.06,0.05,100,1,0)"#,
    ) {
        Value::Error(ErrorKind::Value) => {}
        other => panic!("expected #VALUE!, got {other:?}"),
    }
    match sheet.eval(
        r#"=ODDFPRICE(DATE(2020,3,1),DATE(2023,7,1),"nope",DATE(2020,7,1),0.06,0.05,100,1,0)"#,
    ) {
        Value::Error(ErrorKind::Value) => {}
        other => panic!("expected #VALUE!, got {other:?}"),
    }
    match sheet.eval(
        r#"=ODDFPRICE(DATE(2020,3,1),DATE(2023,7,1),DATE(2020,1,1),"nope",0.06,0.05,100,1,0)"#,
    ) {
        Value::Error(ErrorKind::Value) => {}
        other => panic!("expected #VALUE!, got {other:?}"),
    }
    match sheet.eval(r#"=ODDLPRICE(DATE(2022,11,1),DATE(2023,3,1),"nope",0.06,0.05,100,1,0)"#) {
        Value::Error(ErrorKind::Value) => {}
        other => panic!("expected #VALUE!, got {other:?}"),
    }
    match sheet.eval(r#"=ODDLPRICE("nope",DATE(2023,3,1),DATE(2022,7,1),0.06,0.05,100,1,0)"#) {
        Value::Error(ErrorKind::Value) => {}
        other => panic!("expected #VALUE!, got {other:?}"),
    }

    match sheet
        .eval(r#"=ODDFYIELD("nope",DATE(2025,1,1),DATE(2019,1,1),DATE(2020,7,1),0.05,95,100,2)"#)
    {
        Value::Error(ErrorKind::Value) => {}
        other => panic!("expected #VALUE!, got {other:?}"),
    }
    match sheet
        .eval(r#"=ODDFYIELD(DATE(2020,3,1),"nope",DATE(2019,1,1),DATE(2020,7,1),0.05,95,100,2)"#)
    {
        Value::Error(ErrorKind::Value) => {}
        other => panic!("expected #VALUE!, got {other:?}"),
    }
    match sheet
        .eval(r#"=ODDFYIELD(DATE(2020,3,1),DATE(2025,1,1),"nope",DATE(2020,7,1),0.05,95,100,2)"#)
    {
        Value::Error(ErrorKind::Value) => {}
        other => panic!("expected #VALUE!, got {other:?}"),
    }
    match sheet
        .eval(r#"=ODDFYIELD(DATE(2020,3,1),DATE(2025,1,1),DATE(2019,1,1),"nope",0.05,95,100,2)"#)
    {
        Value::Error(ErrorKind::Value) => {}
        other => panic!("expected #VALUE!, got {other:?}"),
    }

    match sheet.eval(r#"=ODDLYIELD(DATE(2020,1,1),"nope",DATE(2024,7,1),0.05,95,100,2)"#) {
        Value::Error(ErrorKind::Value) => {}
        other => panic!("expected #VALUE!, got {other:?}"),
    }
    match sheet.eval(r#"=ODDLYIELD("nope",DATE(2024,7,1),DATE(2020,1,1),0.05,95,100,2)"#) {
        Value::Error(ErrorKind::Value) => {}
        other => panic!("expected #VALUE!, got {other:?}"),
    }

    // Invalid "last_interest" for ODDLYIELD.
    match sheet.eval(r#"=ODDLYIELD(DATE(2022,11,1),DATE(2023,3,1),"nope",0.06,95,100,1,0)"#) {
        Value::Error(ErrorKind::Value) => {}
        other => panic!("expected #VALUE!, got {other:?}"),
    }
}

#[test]
fn odd_coupon_bond_functions_return_value_for_unparseable_frequency_and_basis_text() {
    let mut sheet = TestSheet::new();

    // Unparseable frequency text should return #VALUE! (not #NUM!) because the text cannot be
    // coerced to a number at all.
    match sheet.eval(
        r#"=ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,"nope",0)"#,
    ) {
        Value::Error(ErrorKind::Name) => return,
        Value::Error(ErrorKind::Value) => {}
        other => panic!("expected #VALUE!, got {other:?}"),
    }
    match sheet.eval(
        r#"=ODDLPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,0.0625,100,"nope",0)"#,
    ) {
        Value::Error(ErrorKind::Name) => return,
        Value::Error(ErrorKind::Value) => {}
        other => panic!("expected #VALUE!, got {other:?}"),
    }
    match sheet.eval(
        r#"=ODDFYIELD(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,98,100,"nope",0)"#,
    ) {
        Value::Error(ErrorKind::Name) => return,
        Value::Error(ErrorKind::Value) => {}
        other => panic!("expected #VALUE!, got {other:?}"),
    }
    match sheet.eval(
        r#"=ODDLYIELD(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,98,100,"nope",0)"#,
    ) {
        Value::Error(ErrorKind::Name) => return,
        Value::Error(ErrorKind::Value) => {}
        other => panic!("expected #VALUE!, got {other:?}"),
    }

    // Unparseable basis text (optional arg) should also return #VALUE!.
    match sheet.eval(
        r#"=ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,"nope")"#,
    ) {
        Value::Error(ErrorKind::Name) => return,
        Value::Error(ErrorKind::Value) => {}
        other => panic!("expected #VALUE!, got {other:?}"),
    }
    match sheet.eval(
        r#"=ODDLPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,0.0625,100,2,"nope")"#,
    ) {
        Value::Error(ErrorKind::Name) => return,
        Value::Error(ErrorKind::Value) => {}
        other => panic!("expected #VALUE!, got {other:?}"),
    }
    match sheet.eval(
        r#"=ODDFYIELD(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,98,100,2,"nope")"#,
    ) {
        Value::Error(ErrorKind::Name) => return,
        Value::Error(ErrorKind::Value) => {}
        other => panic!("expected #VALUE!, got {other:?}"),
    }
    match sheet.eval(
        r#"=ODDLYIELD(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,98,100,2,"nope")"#,
    ) {
        Value::Error(ErrorKind::Name) => return,
        Value::Error(ErrorKind::Value) => {}
        other => panic!("expected #VALUE!, got {other:?}"),
    }
}

#[test]
fn odd_coupon_bond_functions_coerce_date_text_in_frequency_and_basis_like_value() {
    let mut sheet = TestSheet::new();

    // Excel's numeric coercion (VALUE-like) can interpret some text as a date serial. If so, the
    // odd-coupon bond functions should validate the resulting integer and surface #NUM! (not
    // #VALUE!) when it is out of the allowed domain.
    //
    // Example: "2020-01-01" -> date serial (~43831) -> frequency not in {1,2,4} => #NUM!
    match sheet.eval(
        r#"=ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,"2020-01-01",0)"#,
    ) {
        Value::Error(ErrorKind::Name) => return,
        Value::Error(ErrorKind::Num) => {}
        other => panic!("expected #NUM!, got {other:?}"),
    }
    match sheet.eval(
        r#"=ODDFYIELD(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,98,100,"2020-01-01",0)"#,
    ) {
        Value::Error(ErrorKind::Name) => return,
        Value::Error(ErrorKind::Num) => {}
        other => panic!("expected #NUM!, got {other:?}"),
    }
    match sheet.eval(
        r#"=ODDLPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,0.0625,100,"2020-01-01",0)"#,
    ) {
        Value::Error(ErrorKind::Name) => return,
        Value::Error(ErrorKind::Num) => {}
        other => panic!("expected #NUM!, got {other:?}"),
    }
    match sheet.eval(
        r#"=ODDLYIELD(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,98,100,"2020-01-01",0)"#,
    ) {
        Value::Error(ErrorKind::Name) => return,
        Value::Error(ErrorKind::Num) => {}
        other => panic!("expected #NUM!, got {other:?}"),
    }

    // The same applies to basis: parseable date text yields a large integer which is invalid for
    // basis {0..4} and should return #NUM!.
    match sheet.eval(
        r#"=ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,"2020-01-01")"#,
    ) {
        Value::Error(ErrorKind::Name) => return,
        Value::Error(ErrorKind::Num) => {}
        other => panic!("expected #NUM!, got {other:?}"),
    }
    match sheet.eval(
        r#"=ODDFYIELD(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,98,100,2,"2020-01-01")"#,
    ) {
        Value::Error(ErrorKind::Name) => return,
        Value::Error(ErrorKind::Num) => {}
        other => panic!("expected #NUM!, got {other:?}"),
    }
    match sheet.eval(
        r#"=ODDLPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,0.0625,100,2,"2020-01-01")"#,
    ) {
        Value::Error(ErrorKind::Name) => return,
        Value::Error(ErrorKind::Num) => {}
        other => panic!("expected #NUM!, got {other:?}"),
    }
    match sheet.eval(
        r#"=ODDLYIELD(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,98,100,2,"2020-01-01")"#,
    ) {
        Value::Error(ErrorKind::Name) => return,
        Value::Error(ErrorKind::Num) => {}
        other => panic!("expected #NUM!, got {other:?}"),
    }
}

#[test]
fn odd_coupon_bond_functions_coerce_time_text_in_frequency_and_basis_like_value() {
    let mut sheet = TestSheet::new();

    // Excel's numeric coercion (VALUE-like) can interpret some text as a time serial (a fraction
    // of a day). If so, the odd-coupon bond functions should truncate and then validate the
    // resulting integer.
    //
    // Example: "1:00" -> 1/24 (~0.04166) -> trunc->0.
    // - For frequency, 0 is invalid => #NUM! (not #VALUE!)
    match sheet.eval(
        r#"=ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,"1:00",0)"#,
    ) {
        Value::Error(ErrorKind::Name) => return,
        Value::Error(ErrorKind::Num) => {}
        other => panic!("expected #NUM!, got {other:?}"),
    }
    match sheet.eval(
        r#"=ODDFYIELD(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,98,100,"1:00",0)"#,
    ) {
        Value::Error(ErrorKind::Name) => return,
        Value::Error(ErrorKind::Num) => {}
        other => panic!("expected #NUM!, got {other:?}"),
    }
    match sheet.eval(
        r#"=ODDLPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,0.0625,100,"1:00",0)"#,
    ) {
        Value::Error(ErrorKind::Name) => return,
        Value::Error(ErrorKind::Num) => {}
        other => panic!("expected #NUM!, got {other:?}"),
    }
    match sheet.eval(
        r#"=ODDLYIELD(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,98,100,"1:00",0)"#,
    ) {
        Value::Error(ErrorKind::Name) => return,
        Value::Error(ErrorKind::Num) => {}
        other => panic!("expected #NUM!, got {other:?}"),
    }

    // For basis, the same "1:00" input should coerce to 0 and behave like basis=0.
    let baseline_oddfprice = "=ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,0)";
    let baseline_oddfprice_value = match eval_number_or_skip(&mut sheet, baseline_oddfprice) {
        Some(v) => v,
        None => return,
    };
    let time_basis_oddfprice = r#"=ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,"1:00")"#;
    let time_basis_oddfprice_value = eval_number_or_skip(&mut sheet, time_basis_oddfprice)
        .expect("ODDFPRICE should accept time-like text for basis and truncate it to 0");
    assert_close(time_basis_oddfprice_value, baseline_oddfprice_value, 1e-9);

    let baseline_oddfyield = "=ODDFYIELD(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,98,100,2,0)";
    let baseline_oddfyield_value = eval_number_or_skip(&mut sheet, baseline_oddfyield)
        .expect("ODDFYIELD should return a number for the baseline");
    let time_basis_oddfyield = r#"=ODDFYIELD(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,98,100,2,"1:00")"#;
    let time_basis_oddfyield_value = eval_number_or_skip(&mut sheet, time_basis_oddfyield)
        .expect("ODDFYIELD should accept time-like text for basis and truncate it to 0");
    assert_close(time_basis_oddfyield_value, baseline_oddfyield_value, 1e-9);

    let baseline_oddlprice =
        "=ODDLPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,0.0625,100,2,0)";
    let baseline_oddlprice_value = eval_number_or_skip(&mut sheet, baseline_oddlprice)
        .expect("ODDLPRICE should return a number for the baseline");
    let time_basis_oddlprice = r#"=ODDLPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,0.0625,100,2,"1:00")"#;
    let time_basis_oddlprice_value = eval_number_or_skip(&mut sheet, time_basis_oddlprice)
        .expect("ODDLPRICE should accept time-like text for basis and truncate it to 0");
    assert_close(time_basis_oddlprice_value, baseline_oddlprice_value, 1e-9);

    let baseline_oddlyield =
        "=ODDLYIELD(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,98,100,2,0)";
    let baseline_oddlyield_value = eval_number_or_skip(&mut sheet, baseline_oddlyield)
        .expect("ODDLYIELD should return a number for the baseline");
    let time_basis_oddlyield =
        r#"=ODDLYIELD(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,98,100,2,"1:00")"#;
    let time_basis_oddlyield_value = eval_number_or_skip(&mut sheet, time_basis_oddlyield)
        .expect("ODDLYIELD should accept time-like text for basis and truncate it to 0");
    assert_close(time_basis_oddlyield_value, baseline_oddlyield_value, 1e-9);
}

#[test]
fn odd_coupon_bond_functions_accept_time_text_that_truncates_to_valid_frequency_and_basis() {
    let mut sheet = TestSheet::new();

    // Some time values can exceed one full day and therefore coerce to whole numbers.
    //
    // Example: "24:00" -> 1 day -> 1.0 -> trunc->1.
    //
    // Ensure this behaves the same as passing the integer directly for both `frequency` and
    // `basis`.
    let baseline_oddfprice =
        "=ODDFPRICE(DATE(2020,3,1),DATE(2023,7,1),DATE(2020,1,1),DATE(2020,7,1),0.06,0.05,100,1,0)";
    let baseline_oddfprice_value = match eval_number_or_skip(&mut sheet, baseline_oddfprice) {
        Some(v) => v,
        None => return,
    };
    let time_freq_oddfprice = r#"=ODDFPRICE(DATE(2020,3,1),DATE(2023,7,1),DATE(2020,1,1),DATE(2020,7,1),0.06,0.05,100,"24:00",0)"#;
    let time_freq_oddfprice_value = eval_number_or_skip(&mut sheet, time_freq_oddfprice)
        .expect("ODDFPRICE should accept time-like text for frequency that truncates to 1");
    assert_close(time_freq_oddfprice_value, baseline_oddfprice_value, 1e-9);

    let baseline_oddlprice =
        "=ODDLPRICE(DATE(2022,11,1),DATE(2023,3,1),DATE(2022,7,1),0.06,0.05,100,1,0)";
    let baseline_oddlprice_value = eval_number_or_skip(&mut sheet, baseline_oddlprice)
        .expect("ODDLPRICE should return a number for the baseline");
    let time_freq_oddlprice =
        r#"=ODDLPRICE(DATE(2022,11,1),DATE(2023,3,1),DATE(2022,7,1),0.06,0.05,100,"24:00",0)"#;
    let time_freq_oddlprice_value = eval_number_or_skip(&mut sheet, time_freq_oddlprice)
        .expect("ODDLPRICE should accept time-like text for frequency that truncates to 1");
    assert_close(time_freq_oddlprice_value, baseline_oddlprice_value, 1e-9);

    let baseline_oddfyield =
        "=LET(pr,ODDFPRICE(DATE(2020,3,1),DATE(2023,7,1),DATE(2020,1,1),DATE(2020,7,1),0.06,0.05,100,1,0),ODDFYIELD(DATE(2020,3,1),DATE(2023,7,1),DATE(2020,1,1),DATE(2020,7,1),0.06,pr,100,1,0))";
    let baseline_oddfyield_value = eval_number_or_skip(&mut sheet, baseline_oddfyield)
        .expect("ODDFYIELD should return a number for the baseline");
    let time_freq_oddfyield = r#"=LET(pr,ODDFPRICE(DATE(2020,3,1),DATE(2023,7,1),DATE(2020,1,1),DATE(2020,7,1),0.06,0.05,100,1,0),ODDFYIELD(DATE(2020,3,1),DATE(2023,7,1),DATE(2020,1,1),DATE(2020,7,1),0.06,pr,100,"24:00",0))"#;
    let time_freq_oddfyield_value = eval_number_or_skip(&mut sheet, time_freq_oddfyield)
        .expect("ODDFYIELD should accept time-like text for frequency that truncates to 1");
    assert_close(time_freq_oddfyield_value, baseline_oddfyield_value, 1e-9);

    let baseline_oddlyield =
        "=LET(pr,ODDLPRICE(DATE(2022,11,1),DATE(2023,3,1),DATE(2022,7,1),0.06,0.05,100,1,0),ODDLYIELD(DATE(2022,11,1),DATE(2023,3,1),DATE(2022,7,1),0.06,pr,100,1,0))";
    let baseline_oddlyield_value = eval_number_or_skip(&mut sheet, baseline_oddlyield)
        .expect("ODDLYIELD should return a number for the baseline");
    let time_freq_oddlyield = r#"=LET(pr,ODDLPRICE(DATE(2022,11,1),DATE(2023,3,1),DATE(2022,7,1),0.06,0.05,100,1,0),ODDLYIELD(DATE(2022,11,1),DATE(2023,3,1),DATE(2022,7,1),0.06,pr,100,"24:00",0))"#;
    let time_freq_oddlyield_value = eval_number_or_skip(&mut sheet, time_freq_oddlyield)
        .expect("ODDLYIELD should accept time-like text for frequency that truncates to 1");
    assert_close(time_freq_oddlyield_value, baseline_oddlyield_value, 1e-9);

    // "24:00" also truncates to 1 for `basis`.
    let baseline_oddfprice_basis_1 = "=ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,1)";
    let baseline_oddfprice_basis_1_value =
        eval_number_or_skip(&mut sheet, baseline_oddfprice_basis_1)
            .expect("ODDFPRICE baseline basis=1 should evaluate");
    let time_basis_oddfprice = r#"=ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,"24:00")"#;
    let time_basis_oddfprice_value = eval_number_or_skip(&mut sheet, time_basis_oddfprice)
        .expect("ODDFPRICE should accept time-like text for basis that truncates to 1");
    assert_close(
        time_basis_oddfprice_value,
        baseline_oddfprice_basis_1_value,
        1e-9,
    );

    let baseline_oddlprice_basis_1 =
        "=ODDLPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,0.0625,100,2,1)";
    let baseline_oddlprice_basis_1_value =
        eval_number_or_skip(&mut sheet, baseline_oddlprice_basis_1)
            .expect("ODDLPRICE baseline basis=1 should evaluate");
    let time_basis_oddlprice = r#"=ODDLPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,0.0625,100,2,"24:00")"#;
    let time_basis_oddlprice_value = eval_number_or_skip(&mut sheet, time_basis_oddlprice)
        .expect("ODDLPRICE should accept time-like text for basis that truncates to 1");
    assert_close(
        time_basis_oddlprice_value,
        baseline_oddlprice_basis_1_value,
        1e-9,
    );

    let baseline_oddfyield_basis_1 = "=LET(pr,ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,1),ODDFYIELD(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,pr,100,2,1))";
    let baseline_oddfyield_basis_1_value =
        eval_number_or_skip(&mut sheet, baseline_oddfyield_basis_1)
            .expect("ODDFYIELD baseline basis=1 should evaluate");
    let time_basis_oddfyield = r#"=LET(pr,ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,1),ODDFYIELD(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,pr,100,2,"24:00"))"#;
    let time_basis_oddfyield_value = eval_number_or_skip(&mut sheet, time_basis_oddfyield)
        .expect("ODDFYIELD should accept time-like text for basis that truncates to 1");
    assert_close(
        time_basis_oddfyield_value,
        baseline_oddfyield_basis_1_value,
        1e-9,
    );

    let baseline_oddlyield_basis_1 = "=LET(pr,ODDLPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,0.0625,100,2,1),ODDLYIELD(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,pr,100,2,1))";
    let baseline_oddlyield_basis_1_value =
        eval_number_or_skip(&mut sheet, baseline_oddlyield_basis_1)
            .expect("ODDLYIELD baseline basis=1 should evaluate");
    let time_basis_oddlyield = r#"=LET(pr,ODDLPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,0.0625,100,2,1),ODDLYIELD(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,pr,100,2,"24:00"))"#;
    let time_basis_oddlyield_value = eval_number_or_skip(&mut sheet, time_basis_oddlyield)
        .expect("ODDLYIELD should accept time-like text for basis that truncates to 1");
    assert_close(
        time_basis_oddlyield_value,
        baseline_oddlyield_basis_1_value,
        1e-9,
    );
}

#[test]
fn odd_coupon_bond_functions_respect_value_locale_for_numeric_text_in_frequency_and_basis() {
    let mut sheet = TestSheet::new();
    sheet.set_value_locale(ValueLocaleConfig::de_de());

    // In locales where the decimal separator is a comma, numeric text such as "2,9" should be
    // parsed as 2.9 (VALUE-style), then truncated to an integer and validated.

    // frequency="2,9" -> 2.9 -> trunc->2
    let baseline_oddfprice =
        "=ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,0)";
    let baseline_oddfprice_value = match eval_number_or_skip(&mut sheet, baseline_oddfprice) {
        Some(v) => v,
        None => return,
    };
    let oddf_text_freq = r#"=ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,"2,9",0)"#;
    let oddf_text_freq_value = eval_number_or_skip(&mut sheet, oddf_text_freq)
        .expect("ODDFPRICE should parse frequency=\"2,9\" under de-DE locale and truncate");
    assert_close(oddf_text_freq_value, baseline_oddfprice_value, 1e-9);

    let baseline_oddlprice =
        "=ODDLPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,0.0625,100,2,0)";
    let baseline_oddlprice_value = eval_number_or_skip(&mut sheet, baseline_oddlprice)
        .expect("ODDLPRICE baseline should evaluate");
    let oddl_text_freq =
        r#"=ODDLPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,0.0625,100,"2,9",0)"#;
    let oddl_text_freq_value = eval_number_or_skip(&mut sheet, oddl_text_freq)
        .expect("ODDLPRICE should parse frequency=\"2,9\" under de-DE locale and truncate");
    assert_close(oddl_text_freq_value, baseline_oddlprice_value, 1e-9);

    // Yield functions should behave the same way.
    let baseline_oddfyield = "=LET(pr,ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,0),ODDFYIELD(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,pr,100,2,0))";
    let baseline_oddfyield_value = eval_number_or_skip(&mut sheet, baseline_oddfyield)
        .expect("ODDFYIELD baseline should evaluate");
    let oddf_text_freq = r#"=LET(pr,ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,0),ODDFYIELD(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,pr,100,"2,9",0))"#;
    let oddf_text_freq_value = eval_number_or_skip(&mut sheet, oddf_text_freq)
        .expect("ODDFYIELD should parse frequency=\"2,9\" under de-DE locale and truncate");
    assert_close(oddf_text_freq_value, baseline_oddfyield_value, 1e-9);

    let baseline_oddlyield = "=LET(pr,ODDLPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,0.0625,100,2,0),ODDLYIELD(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,pr,100,2,0))";
    let baseline_oddlyield_value = eval_number_or_skip(&mut sheet, baseline_oddlyield)
        .expect("ODDLYIELD baseline should evaluate");
    let oddl_text_freq = r#"=LET(pr,ODDLPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,0.0625,100,2,0),ODDLYIELD(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,pr,100,"2,9",0))"#;
    let oddl_text_freq_value = eval_number_or_skip(&mut sheet, oddl_text_freq)
        .expect("ODDLYIELD should parse frequency=\"2,9\" under de-DE locale and truncate");
    assert_close(oddl_text_freq_value, baseline_oddlyield_value, 1e-9);

    // basis="1,9" -> 1.9 -> trunc->1
    let baseline_basis_1 = "=ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,1)";
    let baseline_basis_1_value = eval_number_or_skip(&mut sheet, baseline_basis_1)
        .expect("ODDFPRICE baseline basis=1 should evaluate");
    let oddf_text_basis = r#"=ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,"1,9")"#;
    let oddf_text_basis_value = eval_number_or_skip(&mut sheet, oddf_text_basis)
        .expect("ODDFPRICE should parse basis=\"1,9\" under de-DE locale and truncate");
    assert_close(oddf_text_basis_value, baseline_basis_1_value, 1e-9);

    let baseline_oddfyield_basis_1 = "=LET(pr,ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,1),ODDFYIELD(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,pr,100,2,1))";
    let baseline_oddfyield_basis_1_value =
        eval_number_or_skip(&mut sheet, baseline_oddfyield_basis_1)
            .expect("ODDFYIELD baseline basis=1 should evaluate");
    let oddf_text_basis = r#"=LET(pr,ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,1),ODDFYIELD(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,pr,100,2,"1,9"))"#;
    let oddf_text_basis_value = eval_number_or_skip(&mut sheet, oddf_text_basis)
        .expect("ODDFYIELD should parse basis=\"1,9\" under de-DE locale and truncate");
    assert_close(
        oddf_text_basis_value,
        baseline_oddfyield_basis_1_value,
        1e-9,
    );
}

#[test]
fn odd_coupon_bond_functions_return_num_for_frequency_and_basis_out_of_i32_range() {
    let mut sheet = TestSheet::new();

    // Extremely large finite values should truncate to a value outside the i32 range used by the
    // engine and return #NUM! (not #VALUE!).
    match sheet.eval(
        "=ODDFPRICE(DATE(2020,3,1),DATE(2023,7,1),DATE(2020,1,1),DATE(2020,7,1),0.06,0.05,100,1E20,0)",
    ) {
        Value::Error(ErrorKind::Name) => return,
        Value::Error(ErrorKind::Num) => {}
        other => panic!("expected #NUM!, got {other:?}"),
    }

    match sheet.eval(
        "=ODDFPRICE(DATE(2020,3,1),DATE(2023,7,1),DATE(2020,1,1),DATE(2020,7,1),0.06,0.05,100,1,-1E20)",
    ) {
        Value::Error(ErrorKind::Name) => return,
        Value::Error(ErrorKind::Num) => {}
        other => panic!("expected #NUM!, got {other:?}"),
    }
}

#[test]
fn oddfprice_day_count_basis_relationships() {
    let system = ExcelDateSystem::EXCEL_1900;

    // Month-end / Feb dates chosen so 30/360 US vs EU differ (basis 0 vs 4).
    // issue=2019-01-31, settlement=2019-02-28, first_coupon=2019-03-31, maturity=2019-09-30
    let issue = serial(2019, 1, 31, system);
    let settlement = serial(2019, 2, 28, system);
    let first_coupon = serial(2019, 3, 31, system);
    let maturity = serial(2019, 9, 30, system);

    let rate = 0.05;
    let yld = 0.06;
    let redemption = 100.0;
    let frequency = 2;

    let mut prices = [0.0f64; 5];
    for basis in 0..=4 {
        let price = oddfprice(
            settlement,
            maturity,
            issue,
            first_coupon,
            rate,
            yld,
            redemption,
            frequency,
            basis,
            system,
        )
        .unwrap_or_else(|e| panic!("ODDFPRICE basis={basis} returned error: {e:?}"));
        assert!(
            price.is_finite(),
            "expected finite price for basis={basis}, got {price}"
        );
        prices[basis as usize] = price;
    }

    // basis=2 (Actual/360) vs basis=3 (Actual/365): E differs (360/f vs 365/f).
    assert!(
        (prices[2] - prices[3]).abs() > 1e-9,
        "expected basis=2 != basis=3, got {} vs {}",
        prices[2],
        prices[3]
    );
    // basis=0 (US 30/360) vs basis=2 (Actual/360): day-difference differs.
    assert!(
        (prices[0] - prices[2]).abs() > 1e-9,
        "expected basis=0 != basis=2, got {} vs {}",
        prices[0],
        prices[2]
    );
    // basis=0 vs basis=4 (US vs EU 30/360): 30/360 conventions diverge for month-end cases.
    assert!(
        (prices[0] - prices[4]).abs() > 1e-9,
        "expected basis=0 != basis=4, got {} vs {}",
        prices[0],
        prices[4]
    );
}

#[test]
fn oddlprice_day_count_basis_relationships() {
    let system = ExcelDateSystem::EXCEL_1900;

    // Month-end / Feb dates chosen so 30/360 US vs EU differ (basis 0 vs 4).
    // last_interest=2019-02-28, settlement=2019-03-15, maturity=2019-03-31
    // (Settlement must fall after last_interest for ODDL*.)
    let last_interest = serial(2019, 2, 28, system);
    let settlement = serial(2019, 3, 15, system);
    let maturity = serial(2019, 3, 31, system);

    let rate = 0.05;
    let yld = 0.06;
    let redemption = 100.0;
    let frequency = 2;

    let mut prices = [0.0f64; 5];
    for basis in 0..=4 {
        let price = oddlprice(
            settlement,
            maturity,
            last_interest,
            rate,
            yld,
            redemption,
            frequency,
            basis,
            system,
        )
        .unwrap_or_else(|e| panic!("ODDLPRICE basis={basis} returned error: {e:?}"));
        assert!(
            price.is_finite(),
            "expected finite price for basis={basis}, got {price}"
        );
        prices[basis as usize] = price;
    }

    // basis=2 (Actual/360) vs basis=3 (Actual/365): E differs (360/f vs 365/f).
    assert!(
        (prices[2] - prices[3]).abs() > 1e-9,
        "expected basis=2 != basis=3, got {} vs {}",
        prices[2],
        prices[3]
    );
    // basis=0 (US 30/360) vs basis=2 (Actual/360): day-difference differs.
    assert!(
        (prices[0] - prices[2]).abs() > 1e-9,
        "expected basis=0 != basis=2, got {} vs {}",
        prices[0],
        prices[2]
    );
    // basis=0 vs basis=4 (US vs EU 30/360): 30/360 conventions diverge for month-end cases.
    assert!(
        (prices[0] - prices[4]).abs() > 1e-9,
        "expected basis=0 != basis=4, got {} vs {}",
        prices[0],
        prices[4]
    );
}

#[test]
fn odd_first_coupon_basis_omitted_matches_explicit_zero() {
    let mut sheet = TestSheet::new();

    let price_with_basis =
        "=ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,0)";
    let price_without_basis =
        "=ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2)";

    let price_with_basis = match eval_number_or_skip(&mut sheet, price_with_basis) {
        Some(v) => v,
        None => return,
    };
    let price_without_basis = eval_number_or_skip(&mut sheet, price_without_basis)
        .expect("ODDFPRICE with omitted basis should return a number when ODDFPRICE is available");

    assert_close(price_without_basis, price_with_basis, 1e-9);

    let yield_with_basis =
        "=ODDFYIELD(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,98,100,2,0)";
    let yield_without_basis =
        "=ODDFYIELD(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,98,100,2)";

    let yield_with_basis = match eval_number_or_skip(&mut sheet, yield_with_basis) {
        Some(v) => v,
        None => return,
    };
    let yield_without_basis = eval_number_or_skip(&mut sheet, yield_without_basis)
        .expect("ODDFYIELD with omitted basis should return a number when ODDFYIELD is available");

    assert_close(yield_without_basis, yield_with_basis, 1e-10);
}

#[test]
fn odd_last_coupon_basis_omitted_matches_explicit_zero() {
    let mut sheet = TestSheet::new();

    let price_with_basis =
        "=ODDLPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,0.0625,100,2,0)";
    let price_without_basis =
        "=ODDLPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,0.0625,100,2)";

    let price_with_basis = match eval_number_or_skip(&mut sheet, price_with_basis) {
        Some(v) => v,
        None => return,
    };
    let price_without_basis = eval_number_or_skip(&mut sheet, price_without_basis)
        .expect("ODDLPRICE with omitted basis should return a number when ODDLPRICE is available");

    assert_close(price_without_basis, price_with_basis, 1e-9);

    let yield_with_basis =
        "=ODDLYIELD(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,98,100,2,0)";
    let yield_without_basis =
        "=ODDLYIELD(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,98,100,2)";

    let yield_with_basis = match eval_number_or_skip(&mut sheet, yield_with_basis) {
        Some(v) => v,
        None => return,
    };
    let yield_without_basis = eval_number_or_skip(&mut sheet, yield_without_basis)
        .expect("ODDLYIELD with omitted basis should return a number when ODDLYIELD is available");

    assert_close(yield_without_basis, yield_with_basis, 1e-10);
}

#[test]
fn odd_first_coupon_zero_yield_price_is_finite_and_roundtrips() {
    let mut sheet = TestSheet::new();

    sheet.set_formula(
        "A1",
        "=ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0,100,2,0)",
    );
    sheet.recalc();

    let price = match cell_number_or_skip(&sheet, "A1") {
        Some(v) => v,
        None => return,
    };
    assert!(
        price.is_finite(),
        "expected ODDFPRICE with yld=0 to be finite, got {price}"
    );

    // Optional `basis` is omitted => defaults to 0 (same result as explicit basis=0).
    sheet.set_formula(
        "B1",
        "=ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0,100,2)",
    );
    sheet.recalc();
    let price_omitted = cell_number_or_skip(&sheet, "B1")
        .expect("ODDFPRICE with omitted basis should evaluate to a number");
    assert_close(price_omitted, price, 1e-10);

    let recovered_yield = match eval_number_or_skip(
        &mut sheet,
        "=ODDFYIELD(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,A1,100,2,0)",
    ) {
        Some(v) => v,
        None => return,
    };

    assert_close(recovered_yield, 0.0, 1e-7);

    let recovered_yield_omitted = match eval_number_or_skip(
        &mut sheet,
        "=ODDFYIELD(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,B1,100,2)",
    ) {
        Some(v) => v,
        None => return,
    };
    assert_close(recovered_yield_omitted, 0.0, 1e-7);
}

#[test]
fn odd_last_coupon_zero_yield_price_is_finite_and_roundtrips() {
    let mut sheet = TestSheet::new();

    sheet.set_formula(
        "A1",
        "=ODDLPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,0,100,2,0)",
    );
    sheet.recalc();

    let price = match cell_number_or_skip(&sheet, "A1") {
        Some(v) => v,
        None => return,
    };
    assert!(
        price.is_finite(),
        "expected ODDLPRICE with yld=0 to be finite, got {price}"
    );

    // Optional `basis` is omitted => defaults to 0 (same result as explicit basis=0).
    sheet.set_formula(
        "B1",
        "=ODDLPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,0,100,2)",
    );
    sheet.recalc();
    let price_omitted = cell_number_or_skip(&sheet, "B1")
        .expect("ODDLPRICE with omitted basis should evaluate to a number");
    assert_close(price_omitted, price, 1e-10);

    let recovered_yield = match eval_number_or_skip(
        &mut sheet,
        "=ODDLYIELD(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,A1,100,2,0)",
    ) {
        Some(v) => v,
        None => return,
    };

    assert_close(recovered_yield, 0.0, 1e-7);

    let recovered_yield_omitted = match eval_number_or_skip(
        &mut sheet,
        "=ODDLYIELD(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,B1,100,2)",
    ) {
        Some(v) => v,
        None => return,
    };
    assert_close(recovered_yield_omitted, 0.0, 1e-7);
}
