use crate::date::ExcelDateSystem;
use crate::error::{ExcelError, ExcelResult};
use crate::functions::date_time;
use crate::functions::financial::iterative::{newton_raphson, EXCEL_ITERATION_TOLERANCE};

const MAX_ITER_ODD_YIELD_NEWTON: usize = 50;
const MAX_ITER_ODD_YIELD_BISECT: usize = 100;
const MAX_BRACKET_EXPANSIONS: usize = 100;
const MAX_COUPON_STEPS: usize = 50_000;
const YIELD_UPPER_CAP: f64 = 1.0e6;
const PRICE_RESIDUAL_TOLERANCE: f64 = 1.0e-6;

#[derive(Debug, Clone)]
struct BondEquation {
    /// Coupon compounding frequency (1, 2, or 4).
    freq: f64,
    /// Accrued interest at settlement (subtracted from dirty PV to return clean price).
    accrued_interest: f64,
    /// Remaining cashflows represented as `(t, amount)` where `t` is the number of coupon
    /// periods (may be fractional) between settlement and payment.
    payments: Vec<(f64, f64)>,
}

impl BondEquation {
    fn new(freq: f64, accrued_interest: f64, payments: Vec<(f64, f64)>) -> ExcelResult<Self> {
        if !(freq == 1.0 || freq == 2.0 || freq == 4.0) {
            return Err(ExcelError::Num);
        }
        if !accrued_interest.is_finite() {
            return Err(ExcelError::Num);
        }
        if payments.is_empty() {
            return Err(ExcelError::Num);
        }
        for (t, amt) in &payments {
            if !t.is_finite() || *t < 0.0 || !amt.is_finite() {
                return Err(ExcelError::Num);
            }
        }
        Ok(Self {
            freq,
            accrued_interest,
            payments,
        })
    }

    fn discount_base(&self, yld: f64) -> ExcelResult<f64> {
        if !yld.is_finite() {
            return Err(ExcelError::Num);
        }

        // Domain check:
        //
        // Excel's bond functions (and this engine's regular coupon bond functions in `bonds.rs`)
        // require the *per-period* discount base `1 + yld/freq` to be positive. This translates
        // to an annualized yield domain of `yld > -freq`.
        //
        // For the exact boundary `yld == -freq`, Excel returns `#DIV/0!` (the discount base
        // becomes 0). Any yield below that boundary is `#NUM!`.
        if yld == -self.freq {
            return Err(ExcelError::Div0);
        }
        if yld < -self.freq {
            return Err(ExcelError::Num);
        }

        let base = 1.0 + yld / self.freq;
        if base == 0.0 {
            return Err(ExcelError::Div0);
        }
        if base < 0.0 || !base.is_finite() {
            return Err(ExcelError::Num);
        }
        Ok(base)
    }

    fn price(&self, yld: f64) -> ExcelResult<f64> {
        let base = self.discount_base(yld)?;

        let mut pv = 0.0;
        for (t, amt) in &self.payments {
            if *amt == 0.0 {
                continue;
            }
            // PV_i = amt * base^(-t)
            let base_pow = base.powf(-*t);
            if !base_pow.is_finite() {
                return Err(ExcelError::Num);
            }
            let pv_i = *amt * base_pow;
            if !pv_i.is_finite() {
                return Err(ExcelError::Num);
            }
            pv += pv_i;
            if !pv.is_finite() {
                return Err(ExcelError::Num);
            }
        }

        let price = pv - self.accrued_interest;
        if price.is_finite() {
            Ok(price)
        } else {
            Err(ExcelError::Num)
        }
    }

    fn price_derivative(&self, yld: f64) -> ExcelResult<f64> {
        let base = self.discount_base(yld)?;

        let mut dpv = 0.0;
        for (t, amt) in &self.payments {
            if *amt == 0.0 || *t == 0.0 {
                continue;
            }
            let base_pow = base.powf(-*t);
            if !base_pow.is_finite() {
                return Err(ExcelError::Num);
            }

            // d/dy base^(-t) = (-t) * base^(-t-1) * (1/freq)
            //               = (-t) * base^(-t) / (freq * base)
            // So dPV_i = amt * (-t) * base^(-t) / (freq * base)
            let denom = self.freq * base;
            if denom == 0.0 {
                return Err(ExcelError::Div0);
            }
            if !denom.is_finite() {
                return Err(ExcelError::Num);
            }
            let dpv_i = -*amt * *t * base_pow / denom;
            if !dpv_i.is_finite() {
                return Err(ExcelError::Num);
            }
            dpv += dpv_i;
            if !dpv.is_finite() {
                return Err(ExcelError::Num);
            }
        }

        if dpv.is_finite() {
            Ok(dpv)
        } else {
            Err(ExcelError::Num)
        }
    }

    fn f(&self, yld: f64, pr: f64) -> Option<f64> {
        match self.price(yld) {
            Ok(price) => {
                let v = price - pr;
                if !v.is_finite() {
                    None
                } else {
                    Some(v)
                }
            }
            Err(_) => None,
        }
    }

    fn df(&self, yld: f64) -> Option<f64> {
        match self.price_derivative(yld) {
            Ok(dprice) => (dprice.is_finite() && dprice != 0.0).then_some(dprice),
            Err(_) => None,
        }
    }
}

fn normalize_frequency(frequency: i32) -> ExcelResult<f64> {
    Ok(f64::from(super::coupon_schedule::validate_frequency(
        frequency,
    )?))
}

fn validate_basis(basis: i32) -> ExcelResult<i32> {
    super::coupon_schedule::validate_basis(basis)
}

fn validate_finite(n: f64) -> ExcelResult<f64> {
    if n.is_finite() {
        Ok(n)
    } else {
        Err(ExcelError::Num)
    }
}

fn days_between(start: i32, end: i32, basis: i32, system: ExcelDateSystem) -> ExcelResult<f64> {
    validate_finite(super::coupon_schedule::days_between(start, end, basis, system)? as f64)
}

fn is_end_of_month(date: i32, system: ExcelDateSystem) -> ExcelResult<bool> {
    Ok(date_time::eomonth(date, 0, system)? == date)
}

fn coupon_date_with_eom(
    anchor: i32,
    months: i32,
    eom: bool,
    system: ExcelDateSystem,
) -> ExcelResult<i32> {
    if eom {
        date_time::eomonth(anchor, months, system)
    } else {
        date_time::edate(anchor, months, system)
    }
}

fn coupon_schedule_from_maturity(
    first_coupon: i32,
    maturity: i32,
    months_per_period: i32,
    system: ExcelDateSystem,
) -> ExcelResult<Vec<i32>> {
    let eom = is_end_of_month(maturity, system)?;

    // Generate coupon dates by stepping from maturity backward in fixed month increments.
    // This matches Excel's COUP* schedule behavior: the month-step anchor is the maturity date,
    // not the first coupon date (which may be clamped in shorter months).
    let mut dates_rev = Vec::new();

    // Hard cap to avoid pathological loops on invalid inputs. This must be large enough to cover
    // Excel's full date range (1900..=9999) at quarterly frequency.
    for k in 0..MAX_COUPON_STEPS {
        let months_back = i32::try_from(k).map_err(|_| ExcelError::Num)?;
        let offset = months_back
            .checked_mul(months_per_period)
            .ok_or(ExcelError::Num)?;
        let offset = -offset;

        let d = coupon_date_with_eom(maturity, offset, eom, system)?;
        if d < first_coupon {
            break;
        }
        dates_rev.push(d);
        if d == first_coupon {
            break;
        }
    }

    if dates_rev.is_empty() {
        return Err(ExcelError::Num);
    }

    dates_rev.reverse();
    if dates_rev[0] != first_coupon {
        return Err(ExcelError::Num);
    }
    let Some(&last) = dates_rev.last() else {
        debug_assert!(false, "coupon schedule should be non-empty after validation");
        return Err(ExcelError::Num);
    };
    if last != maturity {
        // This can only happen if `first_coupon > maturity` (validated elsewhere) or if date
        // stepping failed to hit maturity due to an inconsistent input schedule.
        return Err(ExcelError::Num);
    }

    Ok(dates_rev)
}

/// Coupon-period length `E` in days for the "regular" coupon period used by the odd-coupon bond
/// functions (ODDF* / ODDL*).
fn coupon_period_e(
    pcd: i32,
    ncd: i32,
    basis: i32,
    freq: f64,
    system: ExcelDateSystem,
) -> ExcelResult<f64> {
    let frequency = freq as i32;

    // `freq` is the number of coupon payments per year.
    // `freq` has already been validated as one of {1, 2, 4} by `normalize_frequency`.
    //
    // Reuse the shared COUP* helper (`coupon_schedule::coupon_period_e`) to keep `E` conventions
    // aligned with Excel's regular bond functions.
    //
    // In particular, for basis=4 (European 30E/360), Excel models `E` as a fixed `360/frequency`
    // (consistent with `COUPDAYS`), even though `DAYS360(PCD, NCD, TRUE)` can differ for some
    // end-of-month schedules involving February.
    super::coupon_schedule::coupon_period_e(pcd, ncd, frequency, basis, system)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::date::{ymd_to_serial, ExcelDate};

    #[test]
    fn odd_coupon_e_uses_coupondays_convention_for_basis_4() {
        // Regression test for basis=4 (30E/360):
        // - day-count quantities (A, DSC, etc.) use European DAYS360.
        // - but the modeled regular coupon period length `E` matches COUPDAYS, i.e. `360/frequency`.
        //
        // Using `DAYS360(PCD, NCD, TRUE)` for `E` breaks Excel parity for schedules involving
        // February (see `crates/formula-engine/tests/odd_coupon_oracle_regressions.rs`).
        let system = ExcelDateSystem::EXCEL_1900;
        let pcd = ymd_to_serial(ExcelDate::new(2019, 8, 30), system).unwrap();
        let ncd = ymd_to_serial(ExcelDate::new(2020, 2, 29), system).unwrap();
        let basis = 4;
        let frequency = 2;
        let freq = normalize_frequency(frequency).unwrap();

        // European 30/360 between these coupon dates differs from the modeled `E`.
        let days360 = super::super::coupon_schedule::days_between(pcd, ncd, basis, system).unwrap();
        assert_eq!(days360, 179);

        let e = coupon_period_e(pcd, ncd, basis, freq, system).unwrap();
        assert_eq!(e, 360.0 / (frequency as f64));
        assert_ne!(e as i64, days360);

        // COUPDAYS uses `coupon_schedule::coupon_period_e` and therefore matches.
        let settlement = ymd_to_serial(ExcelDate::new(2001, 5, 1), system).unwrap();
        let e_coupdays =
            super::super::coupon_schedule::coupdays(settlement, ncd, frequency, basis, system)
                .unwrap();
        assert_eq!(e_coupdays, 360.0 / (frequency as f64));
        assert_eq!(e, e_coupdays);
    }
}

fn oddf_equation(
    settlement: i32,
    maturity: i32,
    issue: i32,
    first_coupon: i32,
    rate: f64,
    redemption: f64,
    frequency: i32,
    basis: i32,
    system: ExcelDateSystem,
) -> ExcelResult<BondEquation> {
    let freq = normalize_frequency(frequency)?;
    let basis = validate_basis(basis)?;

    if !rate.is_finite() || !redemption.is_finite() {
        return Err(ExcelError::Num);
    }
    if rate < 0.0 {
        return Err(ExcelError::Num);
    }
    if redemption <= 0.0 {
        return Err(ExcelError::Num);
    }

    // Chronology constraints.
    //
    // NOTE: The CI "excel-oracle" dataset is currently a *synthetic baseline* generated from this
    // engine (not from real Excel). The oracle boundary corpus + unit tests therefore pin current
    // engine behavior for these boundary equalities. Verify real Excel behavior via
    // tools/excel-oracle/run-excel-oracle.ps1 before changing these rules.
    //
    // See:
    // - `tools/excel-oracle/odd_coupon_boundary_cases.json`
    // - `crates/formula-engine/tests/odd_coupon_date_boundaries.rs`
    // - `crates/formula-engine/tests/odd_coupon_oracle_regressions.rs`
    //
    // Allowed boundary equalities:
    // - `issue == settlement` (zero accrued interest)
    // - `settlement == first_coupon` (settlement on the first coupon date)
    // - `first_coupon == maturity` (single odd stub period paid at maturity)
    //
    // ODDF* accepts the boundary equalities `issue == settlement`, `settlement == first_coupon`,
    // and `first_coupon == maturity`, but still rejects `issue == first_coupon` and
    // `settlement == maturity`.
    //
    // Chronology:
    // - `first_coupon == maturity` is allowed (single odd stub period).
    // - `issue <= settlement <= first_coupon <= maturity`
    // - `issue < first_coupon` (reject `issue == first_coupon`)
    // - `settlement < maturity` (reject settlement on/after maturity)
    //
    // See:
    // - `crates/formula-engine/tests/odd_coupon_date_boundaries.rs`
    // - `crates/formula-engine/tests/functions/financial_oddcoupons.rs`
    if !(issue <= settlement
        && settlement <= first_coupon
        && first_coupon <= maturity
        && issue < first_coupon
        && settlement < maturity)
    {
        // Boundary behaviors are locked in `crates/formula-engine/tests/odd_coupon_date_boundaries.rs`.
        return Err(ExcelError::Num);
    }

    // Ensure inputs are representable dates in this system.
    let _ = crate::date::serial_to_ymd(issue, system)?;
    let _ = crate::date::serial_to_ymd(settlement, system)?;
    let _ = crate::date::serial_to_ymd(first_coupon, system)?;
    let _ = crate::date::serial_to_ymd(maturity, system)?;

    let months_per_period = 12 / frequency;
    let coupon_dates =
        coupon_schedule_from_maturity(first_coupon, maturity, months_per_period, system)?;
    let eom = is_end_of_month(maturity, system)?;

    // Compute day-count quantities:
    // - A: accrued days from issue to settlement
    // - DFC: days in the (odd) first accrual period (issue -> first_coupon)
    // - DSC: days from settlement to first_coupon
    let a = days_between(issue, settlement, basis, system)?;
    let dfc = days_between(issue, first_coupon, basis, system)?;
    let dsc = days_between(settlement, first_coupon, basis, system)?;

    // Defensive day-count guards (chronology checks above should ensure all are non-negative, and
    // DFC > 0). When settlement is on the first coupon date, `dsc` is 0.
    //
    // For 30/360 bases, these day-counts are not strictly tied to calendar day differences. In
    // particular, `A`/`DSC` can be 0 even when the underlying serial dates differ (due to the
    // day-of-month adjustments in DAYS360).
    if a < 0.0 || dfc <= 0.0 || dsc < 0.0 {
        return Err(ExcelError::Num);
    }

    // Regular coupon period length `E` (days).
    //
    // Excel models `E` as the length of the *regular* coupon period containing settlement,
    // i.e. the interval between the previous coupon date (PCD) and next coupon date (NCD).
    // For ODDF* functions, `NCD` is always `first_coupon` (since `settlement <= first_coupon`).
    //
    // `PCD` derivation is surprisingly basis-dependent for month-stepping schedules that involve
    // clamping (e.g. `Aug 30 -> Feb 29 -> Aug 29`):
    // - For basis=4 (European 30E/360), Excel steps backwards from `first_coupon` to determine
    //   `PCD` for the DAYS360_EU day-count used in `E` (see
    //   `tests/odd_coupon_oracle_regressions.rs`).
    // - For bases 1/2/3 (actual-day bases), Excel's behavior matches the maturity-anchored coupon
    //   schedule used for the cashflow dates.
    let prev_coupon = if basis == 4 {
        coupon_date_with_eom(first_coupon, -months_per_period, eom, system)?
    } else {
        let n = i32::try_from(coupon_dates.len()).map_err(|_| ExcelError::Num)?;
        let offset_prev = n
            .checked_mul(months_per_period)
            .ok_or(ExcelError::Num)?
            .checked_neg()
            .ok_or(ExcelError::Num)?;
        coupon_date_with_eom(maturity, offset_prev, eom, system)?
    };
    let e = coupon_period_e(prev_coupon, first_coupon, basis, freq, system)?;

    // Regular coupon payment per period.
    //
    // Excel's ODDFPRICE/ODDFYIELD return a price per $100 face value, so the coupon payment is
    // based on the $100 face value (and does not scale with the redemption amount).
    let c = 100.0 * rate / freq;
    validate_finite(c)?;

    let accrued_interest = c * (a / e);
    validate_finite(accrued_interest)?;

    // Fractional periods to first coupon.
    let t0 = dsc / e;
    validate_finite(t0)?;

    // Cashflows (see docs/financial-odd-coupon-bonds.md and `bonds_odd.rs`).
    let odd_first_coupon = c * (dfc / e);
    validate_finite(odd_first_coupon)?;

    let mut payments: Vec<(f64, f64)> = Vec::new();
    if payments.try_reserve_exact(coupon_dates.len()).is_err() {
        debug_assert!(
            false,
            "allocation failed (odd coupon payments={})",
            coupon_dates.len()
        );
        return Err(ExcelError::Num);
    }
    for (idx, date) in coupon_dates.iter().copied().enumerate() {
        let t = t0 + idx as f64;
        validate_finite(t)?;

        let amount = if date == maturity {
            if idx == 0 {
                // first_coupon == maturity: single odd coupon + redemption.
                redemption + odd_first_coupon
            } else {
                redemption + c
            }
        } else if idx == 0 {
            odd_first_coupon
        } else {
            c
        };
        validate_finite(amount)?;
        payments.push((t, amount));
    }

    BondEquation::new(freq, accrued_interest, payments)
}

fn oddl_equation(
    settlement: i32,
    maturity: i32,
    last_interest: i32,
    rate: f64,
    redemption: f64,
    frequency: i32,
    basis: i32,
    system: ExcelDateSystem,
) -> ExcelResult<BondEquation> {
    let freq = normalize_frequency(frequency)?;
    let basis = validate_basis(basis)?;

    if !rate.is_finite() || !redemption.is_finite() {
        return Err(ExcelError::Num);
    }
    if rate < 0.0 {
        return Err(ExcelError::Num);
    }
    if redemption <= 0.0 {
        return Err(ExcelError::Num);
    }

    // Excel-style chronology:
    //
    // - settlement < maturity
    // - last_interest < maturity
    //
    // See:
    // - `tools/excel-oracle/odd_coupon_boundary_cases.json`
    // - `crates/formula-engine/tests/odd_coupon_date_boundaries.rs`
    //
    // Settlement may be before, on, or after `last_interest` (see `bonds_odd.rs`). If it's before,
    // we PV the remaining regular coupons through `last_interest` plus the final odd stub cashflow
    // at maturity.
    if !(settlement < maturity) {
        return Err(ExcelError::Num);
    }
    if !(last_interest < maturity) {
        return Err(ExcelError::Num);
    }

    // Ensure inputs are representable dates in this system.
    let _ = crate::date::serial_to_ymd(last_interest, system)?;
    let _ = crate::date::serial_to_ymd(settlement, system)?;
    let _ = crate::date::serial_to_ymd(maturity, system)?;

    let months_per_period = 12 / frequency;
    let eom = is_end_of_month(last_interest, system)?;

    // Odd last period length (days) and coupon prorating.
    let dlm = days_between(last_interest, maturity, basis, system)?;
    if dlm <= 0.0 {
        return Err(ExcelError::Num);
    }

    // Regular coupon period length `E` (days).
    let prev_coupon = coupon_date_with_eom(last_interest, -months_per_period, eom, system)?;
    let e_last = coupon_period_e(prev_coupon, last_interest, basis, freq, system)?;

    // Regular coupon payment.
    //
    // Excel's ODDLPRICE/ODDLYIELD return a price per $100 face value, so the coupon payment is
    // based on the $100 face value (and does not scale with the redemption amount).
    let c = 100.0 * rate / freq;
    validate_finite(c)?;

    // Odd last coupon at maturity, prorated by DLM/E.
    let stub_periods = dlm / e_last;
    validate_finite(stub_periods)?;
    let odd_last_coupon = c * stub_periods;
    validate_finite(odd_last_coupon)?;
    let maturity_amount = redemption + odd_last_coupon;
    validate_finite(maturity_amount)?;

    if settlement >= last_interest {
        // Settlement inside the odd last coupon period.
        let a = days_between(last_interest, settlement, basis, system)?;
        let dsm = days_between(settlement, maturity, basis, system)?;
        if a < 0.0 || dsm <= 0.0 {
            return Err(ExcelError::Num);
        }

        let accrued_interest = c * (a / e_last);
        validate_finite(accrued_interest)?;

        let t = dsm / e_last;
        validate_finite(t)?;

        let mut payments: Vec<(f64, f64)> = Vec::new();
        if payments.try_reserve_exact(1).is_err() {
            debug_assert!(false, "allocation failed (odd last payments=1)");
            return Err(ExcelError::Num);
        }
        payments.push((t, maturity_amount));
        BondEquation::new(freq, accrued_interest, payments)
    } else {
        // Settlement before the last coupon date: PV remaining regular coupons through
        // `last_interest` plus the final odd stub cashflow at maturity.
        //
        // Find the regular coupon period containing settlement (PCD <= S < NCD) by stepping
        // backward from `last_interest`. We compute each coupon date from the fixed `last_interest`
        // anchor (not by iterative EDATE stepping) to avoid "EDATE drift" across short months.
        let mut k_found: Option<usize> = None;
        for k in 1..=MAX_COUPON_STEPS {
            let k_i32 = i32::try_from(k).map_err(|_| ExcelError::Num)?;
            let offset = k_i32
                .checked_mul(months_per_period)
                .ok_or(ExcelError::Num)?;
            let pcd = coupon_date_with_eom(last_interest, -offset, eom, system)?;
            if pcd <= settlement {
                k_found = Some(k);
                break;
            }
        }
        let k = k_found.ok_or(ExcelError::Num)?;
        let k_prev = k.checked_sub(1).ok_or(ExcelError::Num)?;
        let pcd_offset = i32::try_from(k)
            .map_err(|_| ExcelError::Num)?
            .checked_mul(months_per_period)
            .ok_or(ExcelError::Num)?;
        let ncd_offset = i32::try_from(k_prev)
            .map_err(|_| ExcelError::Num)?
            .checked_mul(months_per_period)
            .ok_or(ExcelError::Num)?;

        let pcd = coupon_date_with_eom(last_interest, -pcd_offset, eom, system)?;
        let ncd = coupon_date_with_eom(last_interest, -ncd_offset, eom, system)?;

        let e_settle = coupon_period_e(pcd, ncd, basis, freq, system)?;
        let a = days_between(pcd, settlement, basis, system)?;
        let dsc = days_between(settlement, ncd, basis, system)?;
        if a < 0.0 || dsc < 0.0 {
            return Err(ExcelError::Num);
        }

        let accrued_interest = c * (a / e_settle);
        validate_finite(accrued_interest)?;

        let frac = dsc / e_settle;
        validate_finite(frac)?;

        let n_reg = k;
        if n_reg == 0 {
            return Err(ExcelError::Num);
        }

        let expected = n_reg.saturating_add(1);
        let mut payments: Vec<(f64, f64)> = Vec::new();
        if payments.try_reserve_exact(expected).is_err() {
            debug_assert!(false, "allocation failed (odd last payments={expected})");
            return Err(ExcelError::Num);
        }
        for idx in 0..n_reg {
            let t = frac + idx as f64;
            validate_finite(t)?;
            payments.push((t, c));
        }

        // Maturity is `stub_periods` after `last_interest`.
        let t_last_interest = frac + (n_reg as f64 - 1.0);
        validate_finite(t_last_interest)?;
        let t_maturity = t_last_interest + stub_periods;
        validate_finite(t_maturity)?;
        payments.push((t_maturity, maturity_amount));

        BondEquation::new(freq, accrued_interest, payments)
    }
}

fn solve_odd_yield(pr: f64, equation: &BondEquation) -> ExcelResult<f64> {
    if !pr.is_finite() || pr <= 0.0 {
        return Err(ExcelError::Num);
    }

    let f = |y: f64| equation.f(y, pr);
    let df = |y: f64| equation.df(y);

    if let Some(y) = newton_raphson(0.1, MAX_ITER_ODD_YIELD_NEWTON, f, df) {
        if y.is_finite() && y > -equation.freq {
            if let Some(residual) = equation.f(y, pr) {
                if residual.abs() <= PRICE_RESIDUAL_TOLERANCE {
                    return Ok(y);
                }
            }
        }
    }

    // Deterministic fallback: monotonic bracketing + bisection.
    // For typical bond cashflows (positive redemption, non-negative coupon), price decreases with yield.
    let mut lo = -equation.freq + 1e-8;
    if !lo.is_finite() || lo <= -equation.freq {
        return Err(ExcelError::Num);
    }

    // Ensure the low end of the bracket has a positive residual. If the residual is negative at our
    // initial `lo`, move closer to the `-frequency` boundary (where price → +∞) until the residual
    // becomes positive or the evaluation overflows (treated as positive).
    let mut expansions = 0usize;
    loop {
        match equation.f(lo, pr) {
            Some(flo) => {
                if flo.abs() <= EXCEL_ITERATION_TOLERANCE {
                    return Ok(lo);
                }
                if flo > 0.0 {
                    break;
                }
            }
            None => break, // non-finite price => effectively +∞ residual
        }

        expansions += 1;
        if expansions > MAX_BRACKET_EXPANSIONS {
            return Err(ExcelError::Num);
        }

        let eps = lo + equation.freq;
        let next_eps = eps / 2.0;
        if next_eps <= 0.0 || !next_eps.is_finite() {
            return Err(ExcelError::Num);
        }
        lo = -equation.freq + next_eps;
        if lo <= -equation.freq {
            return Err(ExcelError::Num);
        }
    }

    let mut hi = 1.0;
    let mut fhi;
    loop {
        match equation.f(hi, pr) {
            Some(v) => {
                fhi = v;
                break;
            }
            None => {
                // Non-finite residual => treat as "too low yield" and increase `hi`.
                expansions += 1;
                if expansions > MAX_BRACKET_EXPANSIONS {
                    return Err(ExcelError::Num);
                }
                hi *= 2.0;
                if hi > YIELD_UPPER_CAP || !hi.is_finite() {
                    return Err(ExcelError::Num);
                }
                continue;
            }
        }
    }

    if fhi.abs() <= EXCEL_ITERATION_TOLERANCE {
        return Ok(hi);
    }

    while fhi >= 0.0 {
        expansions += 1;
        if expansions > MAX_BRACKET_EXPANSIONS {
            return Err(ExcelError::Num);
        }
        hi *= 2.0;
        if hi > YIELD_UPPER_CAP || !hi.is_finite() {
            return Err(ExcelError::Num);
        }
        match equation.f(hi, pr) {
            Some(v) => {
                fhi = v;
                if fhi.abs() <= EXCEL_ITERATION_TOLERANCE {
                    return Ok(hi);
                }
            }
            None => continue,
        }
    }

    for _ in 0..MAX_ITER_ODD_YIELD_BISECT {
        let mid = 0.5 * (lo + hi);
        match equation.f(mid, pr) {
            Some(fmid) => {
                if fmid.abs() <= EXCEL_ITERATION_TOLERANCE {
                    return Ok(mid);
                }
                if fmid > 0.0 {
                    lo = mid;
                } else {
                    hi = mid;
                }
            }
            None => {
                // A non-finite price implies we're too close to the `-frequency` boundary; treat it
                // as a positive residual and move the lower bracket upward.
                lo = mid;
            }
        }
    }

    let y = 0.5 * (lo + hi);
    if y.is_finite() && y > -equation.freq {
        Ok(y)
    } else {
        Err(ExcelError::Num)
    }
}

/// ODDFPRICE: price per 100 face value for a security with an odd (short/long) first coupon period.
pub fn oddfprice(
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
) -> ExcelResult<f64> {
    let eq = oddf_equation(
        settlement,
        maturity,
        issue,
        first_coupon,
        rate,
        redemption,
        frequency,
        basis,
        system,
    )?;
    eq.price(yld)
}

/// ODDFYIELD: yield of a security with an odd (short/long) first coupon period.
pub fn oddfyield(
    settlement: i32,
    maturity: i32,
    issue: i32,
    first_coupon: i32,
    rate: f64,
    pr: f64,
    redemption: f64,
    frequency: i32,
    basis: i32,
    system: ExcelDateSystem,
) -> ExcelResult<f64> {
    let eq = oddf_equation(
        settlement,
        maturity,
        issue,
        first_coupon,
        rate,
        redemption,
        frequency,
        basis,
        system,
    )?;
    solve_odd_yield(pr, &eq)
}

/// ODDLPRICE: price per 100 face value for a security with an odd (short/long) last coupon period.
pub fn oddlprice(
    settlement: i32,
    maturity: i32,
    last_interest: i32,
    rate: f64,
    yld: f64,
    redemption: f64,
    frequency: i32,
    basis: i32,
    system: ExcelDateSystem,
) -> ExcelResult<f64> {
    let eq = oddl_equation(
        settlement,
        maturity,
        last_interest,
        rate,
        redemption,
        frequency,
        basis,
        system,
    )?;
    eq.price(yld)
}

/// ODDLYIELD: yield of a security with an odd (short/long) last coupon period.
pub fn oddlyield(
    settlement: i32,
    maturity: i32,
    last_interest: i32,
    rate: f64,
    pr: f64,
    redemption: f64,
    frequency: i32,
    basis: i32,
    system: ExcelDateSystem,
) -> ExcelResult<f64> {
    let eq = oddl_equation(
        settlement,
        maturity,
        last_interest,
        rate,
        redemption,
        frequency,
        basis,
        system,
    )?;
    solve_odd_yield(pr, &eq)
}
