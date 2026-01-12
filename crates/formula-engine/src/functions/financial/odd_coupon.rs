use crate::date::ExcelDateSystem;
use crate::error::{ExcelError, ExcelResult};
use crate::functions::date_time;
use crate::functions::financial::iterative::{newton_raphson, EXCEL_ITERATION_TOLERANCE};

const MAX_ITER_ODD_YIELD_NEWTON: usize = 50;
const MAX_ITER_ODD_YIELD_BISECT: usize = 100;
const MAX_BRACKET_EXPANSIONS: usize = 100;
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

    fn price_and_derivative(&self, yld: f64) -> ExcelResult<(f64, f64)> {
        if !yld.is_finite() {
            return Err(ExcelError::Num);
        }

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

        let mut pv = 0.0;
        let mut dpv = 0.0;
        for (t, amt) in &self.payments {
            if *amt == 0.0 {
                continue;
            }
            // PV_i = amt * base^(-t)
            let base_pow = base.powf(-*t);
            if base_pow.is_nan() {
                return Err(ExcelError::Num);
            }
            let pv_i = *amt * base_pow;
            pv += pv_i;

            // d/dy base^(-t) = (-t) * base^(-t-1) * (1/freq)
            //               = (-t) * base^(-t) / (freq * base)
            // So dPV_i = amt * (-t) * base^(-t) / (freq * base)
            let denom = self.freq * base;
            if denom == 0.0 {
                return Err(ExcelError::Div0);
            }
            dpv += -*amt * *t * base_pow / denom;
        }

        let price = pv - self.accrued_interest;
        if price.is_nan() || dpv.is_nan() {
            return Err(ExcelError::Num);
        }
        Ok((price, dpv))
    }

    fn f(&self, yld: f64, pr: f64) -> Option<f64> {
        match self.price_and_derivative(yld) {
            Ok((price, _)) => {
                let v = price - pr;
                if v.is_nan() {
                    None
                } else {
                    Some(v)
                }
            }
            Err(_) => None,
        }
    }

    fn df(&self, yld: f64) -> Option<f64> {
        match self.price_and_derivative(yld) {
            Ok((_price, dprice)) => (dprice.is_finite() && dprice != 0.0).then_some(dprice),
            Err(_) => None,
        }
    }
}

fn normalize_frequency(frequency: i32) -> ExcelResult<f64> {
    match frequency {
        1 | 2 | 4 => Ok(frequency as f64),
        _ => Err(ExcelError::Num),
    }
}

fn validate_basis(basis: i32) -> ExcelResult<i32> {
    if (0..=4).contains(&basis) {
        Ok(basis)
    } else {
        Err(ExcelError::Num)
    }
}

fn validate_finite(n: f64) -> ExcelResult<f64> {
    if n.is_finite() {
        Ok(n)
    } else {
        Err(ExcelError::Num)
    }
}

fn days_between(start: i32, end: i32, basis: i32, system: ExcelDateSystem) -> ExcelResult<f64> {
    match basis {
        // 30/360 day count (basis 0 and 4).
        0 => Ok(date_time::days360(start, end, false, system)? as f64),
        4 => Ok(date_time::days360(start, end, true, system)? as f64),
        // Actual day count (basis 1/2/3).
        1 | 2 | 3 => Ok((end - start) as f64),
        _ => Err(ExcelError::Num),
    }
}

/// Coupon-period length `E` in days, following the same basis conventions as the regular bond
/// functions (`COUP*`, `PRICE`, `YIELD`).
fn coupon_period_e(
    pcd: i32,
    ncd: i32,
    basis: i32,
    freq: f64,
    _system: ExcelDateSystem,
) -> ExcelResult<f64> {
    let e = match basis {
        // US 30/360, Actual/360, European 30/360: fixed 360-day year.
        0 | 2 | 4 => 360.0 / freq,
        // Actual/365: fixed 365-day year.
        3 => 365.0 / freq,
        // Actual/Actual: actual days in the period.
        1 => (ncd - pcd) as f64,
        _ => return Err(ExcelError::Num),
    };
    if e == 0.0 {
        return Err(ExcelError::Div0);
    }
    validate_finite(e)
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

    // Excel-style chronology (common case): I < S < F <= M.
    if !(issue < settlement && settlement < first_coupon && first_coupon <= maturity) {
        return Err(ExcelError::Num);
    }
    if !(issue < maturity && settlement < maturity) {
        return Err(ExcelError::Num);
    }

    // Ensure inputs are representable dates in this system.
    let _ = crate::date::serial_to_ymd(issue, system)?;
    let _ = crate::date::serial_to_ymd(settlement, system)?;
    let _ = crate::date::serial_to_ymd(first_coupon, system)?;
    let _ = crate::date::serial_to_ymd(maturity, system)?;

    let months_per_period = 12 / frequency;
    let mut coupon_dates = Vec::new();
    let mut d = first_coupon;
    loop {
        if d > maturity {
            return Err(ExcelError::Num);
        }
        coupon_dates.push(d);
        if d == maturity {
            break;
        }
        d = date_time::edate(d, months_per_period, system)?;
    }

    // Compute day-count quantities:
    // - A: accrued days from issue to settlement
    // - DFC: days in the (odd) first accrual period (issue -> first_coupon)
    // - DSC: days from settlement to first_coupon
    let a = days_between(issue, settlement, basis, system)?;
    let dfc = days_between(issue, first_coupon, basis, system)?;
    let dsc = days_between(settlement, first_coupon, basis, system)?;

    if a < 0.0 || dfc <= 0.0 || dsc <= 0.0 {
        return Err(ExcelError::Num);
    }

    // Regular coupon period length `E` (days).
    let prev_coupon = date_time::edate(first_coupon, -months_per_period, system)?;
    let e = coupon_period_e(prev_coupon, first_coupon, basis, freq, system)?;

    // Regular coupon payment per period.
    let c = redemption * rate / freq;
    validate_finite(c)?;

    let accrued_interest = c * (a / e);
    validate_finite(accrued_interest)?;

    // Fractional periods to first coupon.
    let t0 = dsc / e;
    validate_finite(t0)?;

    // Cashflows (see docs/financial-odd-coupon-bonds.md and `bonds_odd.rs`).
    let odd_first_coupon = c * (dfc / e);
    validate_finite(odd_first_coupon)?;

    let mut payments = Vec::with_capacity(coupon_dates.len());
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

    if !(last_interest < settlement && settlement < maturity) {
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

    // Day-count quantities.
    let a = days_between(last_interest, settlement, basis, system)?;
    let dlm = days_between(last_interest, maturity, basis, system)?;
    let dsm = days_between(settlement, maturity, basis, system)?;
    if a < 0.0 || dlm <= 0.0 || dsm <= 0.0 {
        return Err(ExcelError::Num);
    }

    // Regular coupon period length `E` (days).
    let prev_coupon = date_time::edate(last_interest, -months_per_period, system)?;
    let e = coupon_period_e(prev_coupon, last_interest, basis, freq, system)?;

    // Regular coupon payment.
    let c = redemption * rate / freq;
    validate_finite(c)?;

    let accrued_interest = c * (a / e);
    validate_finite(accrued_interest)?;

    // Odd last coupon at maturity, prorated by DLM/E.
    let odd_last_coupon = c * (dlm / e);
    validate_finite(odd_last_coupon)?;
    let amount = redemption + odd_last_coupon;
    validate_finite(amount)?;

    // Fractional periods from settlement to maturity.
    let t = dsm / e;
    validate_finite(t)?;

    BondEquation::new(freq, accrued_interest, vec![(t, amount)])
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
    let mut flo = equation.f(lo, pr).ok_or(ExcelError::Num)?;
    if flo.abs() <= EXCEL_ITERATION_TOLERANCE {
        return Ok(lo);
    }

    let mut hi = 1.0;
    let mut fhi = equation.f(hi, pr).ok_or(ExcelError::Num)?;
    if fhi.abs() <= EXCEL_ITERATION_TOLERANCE {
        return Ok(hi);
    }

    let mut expansions = 0usize;
    while flo.signum() == fhi.signum() {
        expansions += 1;
        if expansions > MAX_BRACKET_EXPANSIONS {
            return Err(ExcelError::Num);
        }
        hi *= 2.0;
        if hi > YIELD_UPPER_CAP || !hi.is_finite() {
            return Err(ExcelError::Num);
        }
        fhi = equation.f(hi, pr).ok_or(ExcelError::Num)?;
        if fhi.abs() <= EXCEL_ITERATION_TOLERANCE {
            return Ok(hi);
        }
    }

    for _ in 0..MAX_ITER_ODD_YIELD_BISECT {
        let mid = 0.5 * (lo + hi);
        let fmid = equation.f(mid, pr).ok_or(ExcelError::Num)?;

        if fmid.abs() <= EXCEL_ITERATION_TOLERANCE {
            return Ok(mid);
        }

        if flo.signum() == fmid.signum() {
            lo = mid;
            flo = fmid;
        } else {
            hi = mid;
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
    let (price, _dprice) = eq.price_and_derivative(yld)?;
    if price.is_finite() {
        Ok(price)
    } else {
        Err(ExcelError::Num)
    }
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
    let (price, _dprice) = eq.price_and_derivative(yld)?;
    if price.is_finite() {
        Ok(price)
    } else {
        Err(ExcelError::Num)
    }
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
