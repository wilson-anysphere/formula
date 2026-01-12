use crate::date::ExcelDateSystem;
use crate::error::{ExcelError, ExcelResult};
use crate::functions::date_time;

use super::coupon_schedule::{coupon_period_e, days_between, validate_basis, validate_frequency};

const MAX_COUPON_STEPS: usize = 50_000;

/// ACCRINTM(issue, settlement, rate, par, [basis])
///
/// Accrued interest for a security that pays interest at maturity.
pub fn accrintm(
    issue: i32,
    settlement: i32,
    rate: f64,
    par: f64,
    basis: i32,
    system: ExcelDateSystem,
) -> ExcelResult<f64> {
    if issue >= settlement {
        return Err(ExcelError::Num);
    }
    if !rate.is_finite() || rate < 0.0 {
        return Err(ExcelError::Num);
    }
    if !par.is_finite() || par <= 0.0 {
        return Err(ExcelError::Num);
    }
    validate_basis(basis)?;

    let yf = date_time::yearfrac(issue, settlement, basis, system)?;
    let result = par * rate * yf;
    if result.is_finite() {
        Ok(result)
    } else {
        Err(ExcelError::Num)
    }
}

/// ACCRINT(issue, first_interest, settlement, rate, par, frequency, [basis], [calc_method])
///
/// Accrued interest for a security that pays periodic interest.
pub fn accrint(
    issue: i32,
    first_interest: i32,
    settlement: i32,
    rate: f64,
    par: f64,
    frequency: i32,
    basis: i32,
    calc_method: bool,
    system: ExcelDateSystem,
) -> ExcelResult<f64> {
    if issue >= settlement || issue >= first_interest {
        return Err(ExcelError::Num);
    }
    if !rate.is_finite() || rate < 0.0 {
        return Err(ExcelError::Num);
    }
    if !par.is_finite() || par <= 0.0 {
        return Err(ExcelError::Num);
    }
    validate_basis(basis)?;

    let frequency = validate_frequency(frequency)?;
    let months = 12 / frequency;
    let coupon = par * rate / f64::from(frequency);
    if !coupon.is_finite() || coupon < 0.0 {
        return Err(ExcelError::Num);
    }

    // Coupon schedule is anchored at `first_interest` (not maturity), stepping by ±`months` via EDATE.
    let (pcd, ncd) = if settlement < first_interest {
        let pcd = date_time::edate(first_interest, -months, system)?;
        (pcd, first_interest)
    } else {
        // IMPORTANT: EDATE month-stepping is not invertible due to end-of-month clamping.
        // Compute each coupon date as an offset from `first_interest` to avoid day-of-month drift
        // (matching how Excel's COUP* functions behave when anchored at maturity).
        let mut pcd = first_interest;
        let mut ncd = date_time::edate(first_interest, months, system)?;
        let mut k: i32 = 0;

        for _ in 0..MAX_COUPON_STEPS {
            if settlement < ncd {
                break;
            }
            k = k.checked_add(1).ok_or(ExcelError::Num)?;
            pcd = ncd;

            let next_k = k.checked_add(1).ok_or(ExcelError::Num)?;
            let months_fwd = next_k.checked_mul(months).ok_or(ExcelError::Num)?;
            ncd = date_time::edate(first_interest, months_fwd, system)?;
        }

        if settlement >= ncd {
            return Err(ExcelError::Num);
        }
        (pcd, ncd)
    };

    // `calc_method` only affects the initial (issue -> first interest) stub period.
    // - For settlement < first_interest:
    //   - calc_method == FALSE (0): accrue from issue.
    //   - calc_method == TRUE (1): accrue from the start of the regular coupon period (PCD).
    // - For settlement >= first_interest: Excel accrues from PCD (standard since-last-coupon behavior),
    //   and calc_method is ignored.
    let accrual_start = if settlement < first_interest && calc_method {
        pcd
    } else if settlement < first_interest {
        issue
    } else {
        pcd
    };

    let a_start = days_between(accrual_start, settlement, basis, system)?;
    if a_start < 0 {
        return Err(ExcelError::Num);
    }

    let e = coupon_period_e(pcd, ncd, frequency, basis, system)?;
    if !e.is_finite() || e <= 0.0 {
        return Err(ExcelError::Num);
    }

    let result = coupon * (a_start as f64) / e;
    if result.is_finite() {
        Ok(result)
    } else {
        Err(ExcelError::Num)
    }
}
