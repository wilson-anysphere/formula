use crate::date::ExcelDateSystem;
use crate::error::{ExcelError, ExcelResult};
use crate::functions::date_time;

fn days_between(start: i32, end: i32, basis: i32, system: ExcelDateSystem) -> ExcelResult<i64> {
    match basis {
        0 => date_time::days360(start, end, false, system),
        4 => date_time::days360(start, end, true, system),
        1 | 2 | 3 => Ok(i64::from(end) - i64::from(start)),
        _ => Err(ExcelError::Num),
    }
}

fn months_per_period(frequency: i32) -> Option<i32> {
    match frequency {
        1 => Some(12),
        2 => Some(6),
        4 => Some(3),
        _ => None,
    }
}

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
    if !(0..=4).contains(&basis) {
        return Err(ExcelError::Num);
    }

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
    if !(0..=4).contains(&basis) {
        return Err(ExcelError::Num);
    }

    let months = months_per_period(frequency).ok_or(ExcelError::Num)?;
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

    let e = match basis {
        1 => {
            let e = i64::from(ncd) - i64::from(pcd);
            if e <= 0 {
                return Err(ExcelError::Num);
            }
            e as f64
        }
        0 | 2 | 4 => 360.0 / f64::from(frequency),
        3 => 365.0 / f64::from(frequency),
        _ => return Err(ExcelError::Num),
    };
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
