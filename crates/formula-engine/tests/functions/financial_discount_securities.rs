use formula_engine::date::{ymd_to_serial, ExcelDate, ExcelDateSystem};
use formula_engine::functions::financial::{
    disc, intrate, pricedisc, pricemat, received, tbilleq, tbillprice, tbillyield, yielddisc,
    yieldmat,
};
use formula_engine::ExcelError;

fn assert_close(actual: f64, expected: f64, tol: f64) {
    assert!(
        (actual - expected).abs() <= tol,
        "expected {expected}, got {actual} (tol={tol})"
    );
}

fn serial_1900(y: i32, m: u8, d: u8) -> i32 {
    ymd_to_serial(ExcelDate::new(y, m, d), ExcelDateSystem::EXCEL_1900).unwrap()
}

#[test]
fn pricemat_matches_excel_example() {
    // https://support.microsoft.com/en-us/office/pricemat-function-52c3b4da-bc7e-476a-989f-a95f675cae77
    let settlement = serial_1900(2008, 2, 15);
    let maturity = serial_1900(2008, 4, 13);
    let issue = serial_1900(2007, 11, 11);
    let rate = 0.061;
    let yld = 0.061;
    let basis = 0;

    let price = pricemat(
        settlement,
        maturity,
        issue,
        rate,
        yld,
        basis,
        ExcelDateSystem::EXCEL_1900,
    )
    .unwrap();
    assert_close(price, 99.98, 0.02);
}

#[test]
fn yieldmat_matches_excel_example() {
    // https://support.microsoft.com/en-us/office/yieldmat-function-ba7d1809-0d33-4bcb-96c7-6c56ec62ef6f
    let settlement = serial_1900(2008, 3, 15);
    let maturity = serial_1900(2008, 11, 3);
    let issue = serial_1900(2007, 11, 8);
    let rate = 0.0625;
    let pr = 100.0123;
    let basis = 0;

    let yld = yieldmat(
        settlement,
        maturity,
        issue,
        rate,
        pr,
        basis,
        ExcelDateSystem::EXCEL_1900,
    )
    .unwrap();
    assert_close(yld, 0.061, 1e-4);
}

#[test]
fn intrate_matches_excel_example() {
    // https://support.microsoft.com/en-us/office/intrate-function-5cb34dde-a221-4cb6-b3eb-0b9e55e1316f
    let settlement = serial_1900(2008, 2, 15);
    let maturity = serial_1900(2008, 5, 15);
    let investment = 1_000_000.0;
    let redemption = 1_014_420.0;
    let basis = 2;

    let rate = intrate(
        settlement,
        maturity,
        investment,
        redemption,
        basis,
        ExcelDateSystem::EXCEL_1900,
    )
    .unwrap();
    assert_close(rate, 0.0577, 1e-4);
}

#[test]
fn received_matches_excel_example() {
    // https://support.microsoft.com/en-us/office/received-function-7a3f8b93-6611-4f81-8576-828312c9b5e5
    let settlement = serial_1900(2008, 2, 15);
    let maturity = serial_1900(2008, 5, 15);
    let investment = 1_000_000.0;
    let discount = 0.0575;
    let basis = 2;

    let amt = received(
        settlement,
        maturity,
        investment,
        discount,
        basis,
        ExcelDateSystem::EXCEL_1900,
    )
    .unwrap();
    assert_close(amt, 1_014_584.65, 0.02);
}

#[test]
fn disc_matches_excel_example() {
    // https://support.microsoft.com/en-us/office/disc-function-71fce9f3-3f05-4acf-a5a3-eac6ef4daa53
    let settlement = serial_1900(2018, 7, 1);
    // Note: the support article's example table lists 01/01/2048 but the published result
    // corresponds to a 20-year span (01/01/2038). Use 2038 here to match Excel's documented
    // output value.
    let maturity = serial_1900(2038, 1, 1);
    let pr = 97.975;
    let redemption = 100.0;
    let basis = 1;

    let discount = disc(
        settlement,
        maturity,
        pr,
        redemption,
        basis,
        ExcelDateSystem::EXCEL_1900,
    )
    .unwrap();
    assert_close(discount, 0.001038, 1e-6);
}

#[test]
fn pricedisc_matches_excel_example() {
    // https://support.microsoft.com/en-us/office/pricedisc-function-d06ad7c1-380e-4be7-9fd9-75e3079acfd3
    let settlement = serial_1900(2008, 2, 16);
    let maturity = serial_1900(2008, 3, 1);
    let discount = 0.0525;
    let redemption = 100.0;
    let basis = 2;

    let price = pricedisc(
        settlement,
        maturity,
        discount,
        redemption,
        basis,
        ExcelDateSystem::EXCEL_1900,
    )
    .unwrap();
    assert_close(price, 99.80, 0.02);
}

#[test]
fn yielddisc_matches_excel_example() {
    // https://support.microsoft.com/en-us/office/yielddisc-function-a9dbdbae-7dae-46de-b995-615faffaaed7
    let settlement = serial_1900(2008, 2, 16);
    let maturity = serial_1900(2008, 3, 1);
    let pr = 99.795;
    let redemption = 100.0;
    let basis = 2;

    let yld = yielddisc(
        settlement,
        maturity,
        pr,
        redemption,
        basis,
        ExcelDateSystem::EXCEL_1900,
    )
    .unwrap();
    assert_close(yld, 0.052823, 1e-6);
}

#[test]
fn tbillprice_tbillyield_tbilleq_match_excel_examples() {
    // https://support.microsoft.com/en-us/office/tbillprice-function-eacca992-c29d-425a-9eb8-0513fe6035a2
    // https://support.microsoft.com/en-us/office/tbillyield-function-6d381232-f4b0-4cd5-8e97-45b9c03468ba
    // https://support.microsoft.com/en-us/office/tbilleq-function-2ab72d90-9b4d-4efe-9fc2-0f81f2c19c8c
    let settlement = serial_1900(2008, 3, 31);
    let maturity = serial_1900(2008, 6, 1);

    let price = tbillprice(settlement, maturity, 0.09).unwrap();
    assert_close(price, 98.45, 0.02);

    let yld = tbillyield(settlement, maturity, 98.45).unwrap();
    assert_close(yld, 0.0914, 1e-4);

    let eq = tbilleq(settlement, maturity, 0.0914).unwrap();
    assert_close(eq, 0.0942, 1e-4);
}

#[test]
fn disc_pricedisc_roundtrip() {
    let settlement = serial_1900(2024, 1, 10);
    let maturity = serial_1900(2024, 7, 10);
    let redemption = 100.0;
    let discount = 0.045;
    let basis = 2;

    let price = pricedisc(
        settlement,
        maturity,
        discount,
        redemption,
        basis,
        ExcelDateSystem::EXCEL_1900,
    )
    .unwrap();
    let back = disc(
        settlement,
        maturity,
        price,
        redemption,
        basis,
        ExcelDateSystem::EXCEL_1900,
    )
    .unwrap();

    assert_close(back, discount, 1e-12);
}

#[test]
fn yielddisc_pricedisc_roundtrip_via_discount_conversion() {
    let settlement = serial_1900(2024, 1, 10);
    let maturity = serial_1900(2024, 7, 10);
    let redemption = 100.0;
    let basis = 2;
    let discount = 0.08;

    // Compute a discount-security price from a discount rate.
    let price = pricedisc(
        settlement,
        maturity,
        discount,
        redemption,
        basis,
        ExcelDateSystem::EXCEL_1900,
    )
    .unwrap();

    // Compute the yield on that price.
    let yld = yielddisc(
        settlement,
        maturity,
        price,
        redemption,
        basis,
        ExcelDateSystem::EXCEL_1900,
    )
    .unwrap();

    // Convert yield -> discount and recompute price.
    let year_frac = formula_engine::functions::date_time::yearfrac(
        settlement,
        maturity,
        basis,
        ExcelDateSystem::EXCEL_1900,
    )
    .unwrap();
    let discount_back = yld / (1.0 + yld * year_frac);

    let price_back = pricedisc(
        settlement,
        maturity,
        discount_back,
        redemption,
        basis,
        ExcelDateSystem::EXCEL_1900,
    )
    .unwrap();

    assert_close(price_back, price, 1e-10);
}

#[test]
fn tbillyield_tbillprice_roundtrip_via_discount_conversion() {
    let settlement = serial_1900(2024, 3, 1);
    let maturity = serial_1900(2024, 6, 1);
    let discount = 0.07;

    let price = tbillprice(settlement, maturity, discount).unwrap();
    let yld = tbillyield(settlement, maturity, price).unwrap();

    let dsm = (maturity - settlement) as f64;
    let discount_back = yld / (1.0 + yld * dsm / 360.0);
    let price_back = tbillprice(settlement, maturity, discount_back).unwrap();

    assert_close(price_back, price, 1e-10);
}

#[test]
fn pricemat_yieldmat_roundtrip() {
    // Pick dates where YEARFRAC(.,.,0) yields integer values so the two functions
    // should be consistent up to floating-point roundoff.
    let settlement = serial_1900(2020, 1, 1);
    let maturity = serial_1900(2021, 1, 1);
    let issue = serial_1900(2019, 1, 1);
    let rate = 0.05;
    let yld = 0.04;
    let basis = 0;

    let price = pricemat(
        settlement,
        maturity,
        issue,
        rate,
        yld,
        basis,
        ExcelDateSystem::EXCEL_1900,
    )
    .unwrap();
    let yld_back = yieldmat(
        settlement,
        maturity,
        issue,
        rate,
        price,
        basis,
        ExcelDateSystem::EXCEL_1900,
    )
    .unwrap();

    assert_close(yld_back, yld, 1e-12);
}

#[test]
fn tbilleq_long_branch_matches_bond_equivalent_yield_definition() {
    // Exercise the TBILLEQ branch used for maturities longer than 182 days.
    let settlement = serial_1900(2020, 1, 1);
    let maturity = settlement + 200;
    let discount = 0.05;

    let dsm = (maturity - settlement) as f64;
    let price_factor = 1.0 - discount * dsm / 360.0;
    let exponent = 365.0 / (2.0 * dsm);
    let expected = 2.0 * ((1.0 / price_factor).powf(exponent) - 1.0);

    let eq = tbilleq(settlement, maturity, discount).unwrap();
    assert_close(eq, expected, 1e-12);
}

#[test]
fn error_cases() {
    let system = ExcelDateSystem::EXCEL_1900;
    let settlement = serial_1900(2024, 1, 1);
    let maturity = serial_1900(2024, 1, 1);

    assert_eq!(
        intrate(settlement, maturity, 100.0, 101.0, 0, system),
        Err(ExcelError::Num)
    );
    assert_eq!(
        pricedisc(settlement, maturity, 0.05, 100.0, 0, system),
        Err(ExcelError::Num)
    );
    assert_eq!(
        yielddisc(settlement, maturity, 99.0, 100.0, 0, system),
        Err(ExcelError::Num)
    );

    assert_eq!(
        intrate(
            settlement,
            serial_1900(2024, 2, 1),
            100.0,
            101.0,
            99,
            system
        ),
        Err(ExcelError::Num)
    );

    // TBILL settlement must be before maturity.
    assert_eq!(tbillprice(settlement, maturity, 0.05), Err(ExcelError::Num));
    assert_eq!(tbillyield(settlement, maturity, 99.0), Err(ExcelError::Num));
    assert_eq!(tbilleq(settlement, maturity, 0.05), Err(ExcelError::Num));

    // TBILL maturity too long (> 365 days).
    let long_maturity = serial_1900(2025, 2, 1);
    assert_eq!(
        tbillprice(settlement, long_maturity, 0.05),
        Err(ExcelError::Num)
    );
    assert_eq!(
        tbillyield(settlement, long_maturity, 99.0),
        Err(ExcelError::Num)
    );
    assert_eq!(
        tbilleq(settlement, long_maturity, 0.05),
        Err(ExcelError::Num)
    );

    // TBILL functions require finite, positive discount/price inputs.
    let valid_maturity = serial_1900(2024, 7, 1);
    for discount in [0.0, -0.01, f64::INFINITY, f64::NAN] {
        assert_eq!(
            tbillprice(settlement, valid_maturity, discount),
            Err(ExcelError::Num),
            "TBILLPRICE(discount={discount:?}) should be #NUM!"
        );
        assert_eq!(
            tbilleq(settlement, valid_maturity, discount),
            Err(ExcelError::Num),
            "TBILLEQ(discount={discount:?}) should be #NUM!"
        );
    }
    for pr in [0.0, -1.0, f64::INFINITY, f64::NAN] {
        assert_eq!(
            tbillyield(settlement, valid_maturity, pr),
            Err(ExcelError::Num),
            "TBILLYIELD(pr={pr:?}) should be #NUM!"
        );
    }

    // Reject discounts that imply a non-positive bill price.
    assert_eq!(
        tbillprice(settlement, valid_maturity, 2.0),
        Err(ExcelError::Num)
    );
    assert_eq!(
        tbilleq(settlement, valid_maturity, 2.0),
        Err(ExcelError::Num)
    );
}
