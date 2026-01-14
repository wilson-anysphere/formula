use crate::date::ExcelDateSystem;
use crate::error::{ExcelError, ExcelResult};

use super::coupon_schedule::{
    coupon_pcd_ncd_num, coupon_period_e, days_between, validate_basis, validate_frequency,
    validate_serial,
};
use super::iterative::solve_root_newton_bisection;

const MAX_ITER_YIELD: usize = 100;

#[derive(Debug, Clone, Copy)]
struct CouponSchedule {
    /// A / E: days from previous coupon to settlement divided by days in coupon period.
    a_over_e: f64,
    /// DSC / E: days from settlement to next coupon divided by days in coupon period.
    d: f64,
    /// Number of coupon payments remaining (COUPNUM).
    n: i32,
}
fn coupon_schedule(
    settlement: i32,
    maturity: i32,
    frequency: i32,
    basis: i32,
    system: ExcelDateSystem,
) -> ExcelResult<CouponSchedule> {
    let (pcd, ncd, n) = coupon_pcd_ncd_num(settlement, maturity, frequency, system)?;

    // Coupon period length `E` (days).
    let e = coupon_period_e(pcd, ncd, frequency, basis, system)?;
    let a = days_between(pcd, settlement, basis, system)? as f64;
    let dsc = match basis {
        0 | 4 => e - a,
        _ => days_between(settlement, ncd, basis, system)? as f64,
    };

    if !a.is_finite() || !e.is_finite() || !dsc.is_finite() {
        return Err(ExcelError::Num);
    }
    if a < 0.0 || dsc < 0.0 || e <= 0.0 {
        return Err(ExcelError::Num);
    }

    Ok(CouponSchedule {
        a_over_e: a / e,
        d: dsc / e,
        n,
    })
}

/// Compute:
/// - dirty price (PV of cashflows, including accrued interest),
/// - `deriv_sum = Σ CF * t * PV_factor`, where `t` is measured in coupon periods,
/// - `g = 1 + yld/frequency`.
fn dirty_price_and_deriv_sum(
    coupon_payment: f64,
    redemption: f64,
    yld: f64,
    frequency: f64,
    d: f64,
    n: i32,
) -> Option<(f64, f64, f64)> {
    if !coupon_payment.is_finite()
        || !redemption.is_finite()
        || !yld.is_finite()
        || !frequency.is_finite()
    {
        return None;
    }
    if frequency <= 0.0 || n <= 0 {
        return None;
    }

    let per_yield = yld / frequency;
    if per_yield <= -1.0 {
        return None;
    }

    let g = 1.0 + per_yield;
    if !g.is_finite() || g == 0.0 {
        return None;
    }

    let ln1p = per_yield.ln_1p();
    if !ln1p.is_finite() {
        return None;
    }

    let inv_g = 1.0 / g;
    let mut discount = (-d * ln1p).exp(); // g^(-d)
    if !discount.is_finite() {
        return None;
    }

    let mut dirty = 0.0;
    let mut deriv_sum = 0.0;

    let n_usize = n as usize;
    for j in 0..n_usize {
        let t = d + (j as f64);
        let cf = if j + 1 == n_usize {
            coupon_payment + redemption
        } else {
            coupon_payment
        };

        dirty += cf * discount;
        deriv_sum += cf * t * discount;

        discount *= inv_g;
    }

    (dirty.is_finite() && deriv_sum.is_finite()).then_some((dirty, deriv_sum, g))
}

/// PRICE(settlement, maturity, rate, yld, redemption, frequency, [basis])
pub fn price(
    settlement: i32,
    maturity: i32,
    rate: f64,
    yld: f64,
    redemption: f64,
    frequency: i32,
    basis: i32,
    system: ExcelDateSystem,
) -> ExcelResult<f64> {
    if settlement >= maturity {
        return Err(ExcelError::Num);
    }
    validate_frequency(frequency)?;
    validate_basis(basis)?;
    validate_serial(settlement, system)?;
    validate_serial(maturity, system)?;

    if !rate.is_finite() || !yld.is_finite() || !redemption.is_finite() {
        return Err(ExcelError::Num);
    }
    if rate < 0.0 {
        return Err(ExcelError::Num);
    }
    if redemption <= 0.0 {
        return Err(ExcelError::Num);
    }

    let freq = frequency as f64;
    // Require 1 + yld/frequency > 0.
    if yld == -freq {
        return Err(ExcelError::Div0);
    }
    if yld < -freq {
        return Err(ExcelError::Num);
    }

    // Excel's `rate` is the annual coupon rate *per $100 face value* (not scaled by `redemption`,
    // which is the amount repaid per $100 face value at maturity).
    let coupon_payment = 100.0 * rate / freq;
    if !coupon_payment.is_finite() {
        return Err(ExcelError::Num);
    }

    let schedule = coupon_schedule(settlement, maturity, frequency, basis, system)?;
    let (dirty, _deriv_sum, _g) = dirty_price_and_deriv_sum(
        coupon_payment,
        redemption,
        yld,
        freq,
        schedule.d,
        schedule.n,
    )
    .ok_or(ExcelError::Num)?;

    let clean = dirty - coupon_payment * schedule.a_over_e;
    if clean.is_finite() {
        Ok(clean)
    } else {
        Err(ExcelError::Num)
    }
}

/// YIELD(settlement, maturity, rate, pr, redemption, frequency, [basis])
pub fn yield_rate(
    settlement: i32,
    maturity: i32,
    rate: f64,
    pr: f64,
    redemption: f64,
    frequency: i32,
    basis: i32,
    system: ExcelDateSystem,
) -> ExcelResult<f64> {
    if settlement >= maturity {
        return Err(ExcelError::Num);
    }
    validate_frequency(frequency)?;
    validate_basis(basis)?;
    validate_serial(settlement, system)?;
    validate_serial(maturity, system)?;

    if !rate.is_finite() || !pr.is_finite() || !redemption.is_finite() {
        return Err(ExcelError::Num);
    }
    if rate < 0.0 {
        return Err(ExcelError::Num);
    }
    if pr <= 0.0 {
        return Err(ExcelError::Num);
    }
    if redemption <= 0.0 {
        return Err(ExcelError::Num);
    }

    let freq = frequency as f64;
    // Excel's `rate` is the annual coupon rate *per $100 face value* (not scaled by `redemption`,
    // which is the amount repaid per $100 face value at maturity).
    let coupon_payment = 100.0 * rate / freq;
    if !coupon_payment.is_finite() {
        return Err(ExcelError::Num);
    }

    let schedule = coupon_schedule(settlement, maturity, frequency, basis, system)?;

    let f = |y: f64| {
        // Near the `-frequency` boundary, the dirty price can overflow (approaching +∞). Excel
        // still brackets a solution in that region; treat overflow as a large positive residual.
        match dirty_price_and_deriv_sum(coupon_payment, redemption, y, freq, schedule.d, schedule.n)
        {
            Some((dirty, _deriv_sum, _g)) => {
                let clean = dirty - coupon_payment * schedule.a_over_e;
                let fx = clean - pr;
                (fx.is_finite()).then_some(fx)
            }
            None => Some(f64::MAX),
        }
    };

    let df = |y: f64| {
        let (_dirty, deriv_sum, g) =
            dirty_price_and_deriv_sum(coupon_payment, redemption, y, freq, schedule.d, schedule.n)?;
        let derivative = -deriv_sum / (freq * g);
        (derivative.is_finite()).then_some(derivative)
    };

    // Excel does not expose an explicit guess for YIELD; it defaults to ~0.1.
    let guess = if rate > 0.0 { rate } else { 0.1 };
    let lower_bound = -freq + 1e-8; // yld must be > -frequency
    let upper_bound = 1.0e10;

    solve_root_newton_bisection(guess, lower_bound, upper_bound, MAX_ITER_YIELD, f, df)
        .ok_or(ExcelError::Num)
}

/// DURATION(settlement, maturity, coupon, yld, frequency, [basis])
pub fn duration(
    settlement: i32,
    maturity: i32,
    coupon: f64,
    yld: f64,
    frequency: i32,
    basis: i32,
    system: ExcelDateSystem,
) -> ExcelResult<f64> {
    if settlement >= maturity {
        return Err(ExcelError::Num);
    }
    validate_frequency(frequency)?;
    validate_basis(basis)?;
    validate_serial(settlement, system)?;
    validate_serial(maturity, system)?;

    if !coupon.is_finite() || !yld.is_finite() {
        return Err(ExcelError::Num);
    }
    if coupon < 0.0 {
        return Err(ExcelError::Num);
    }

    let freq = frequency as f64;
    if yld == -freq {
        return Err(ExcelError::Div0);
    }
    if yld < -freq {
        return Err(ExcelError::Num);
    }

    let coupon_payment = 100.0 * coupon / freq;
    let redemption = 100.0;

    let schedule = coupon_schedule(settlement, maturity, frequency, basis, system)?;
    let (dirty, deriv_sum, _g) = dirty_price_and_deriv_sum(
        coupon_payment,
        redemption,
        yld,
        freq,
        schedule.d,
        schedule.n,
    )
    .ok_or(ExcelError::Num)?;
    if dirty == 0.0 {
        return Err(ExcelError::Div0);
    }

    let dur = deriv_sum / (dirty * freq);
    if dur.is_finite() {
        Ok(dur)
    } else {
        Err(ExcelError::Num)
    }
}

/// MDURATION(settlement, maturity, coupon, yld, frequency, [basis])
pub fn mduration(
    settlement: i32,
    maturity: i32,
    coupon: f64,
    yld: f64,
    frequency: i32,
    basis: i32,
    system: ExcelDateSystem,
) -> ExcelResult<f64> {
    let dur = duration(settlement, maturity, coupon, yld, frequency, basis, system)?;
    let freq = frequency as f64;
    let g = 1.0 + yld / freq;
    if g == 0.0 {
        return Err(ExcelError::Div0);
    }
    let result = dur / g;
    if result.is_finite() {
        Ok(result)
    } else {
        Err(ExcelError::Num)
    }
}
