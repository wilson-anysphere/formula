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

fn coupon_date_leq(
    date: i32,
    first_interest: i32,
    months: i32,
    system: ExcelDateSystem,
) -> ExcelResult<i32> {
    // Coupon schedule is anchored on `first_interest` and repeats every `months`.
    // Find the last coupon date <= `date`.
    if date >= first_interest {
        let mut cur = first_interest;
        loop {
            let next = date_time::edate(cur, months, system)?;
            if next <= date {
                cur = next;
            } else {
                break;
            }
        }
        Ok(cur)
    } else {
        let mut cur = first_interest;
        loop {
            let prev = date_time::edate(cur, -months, system)?;
            if prev > date {
                cur = prev;
            } else {
                // `prev <= date < cur`
                return Ok(prev);
            }
        }
    }
}

fn coupon_period_bounds(
    date: i32,
    first_interest: i32,
    months: i32,
    system: ExcelDateSystem,
) -> ExcelResult<(i32, i32)> {
    let start = coupon_date_leq(date, first_interest, months, system)?;
    let end = date_time::edate(start, months, system)?;
    Ok((start, end))
}

fn accrued_interest_over_schedule(
    start: i32,
    end: i32,
    first_interest: i32,
    months: i32,
    coupon: f64,
    basis: i32,
    system: ExcelDateSystem,
) -> ExcelResult<f64> {
    if end < start {
        return Err(ExcelError::Num);
    }
    if end == start {
        return Ok(0.0);
    }

    let (mut period_start, mut period_end) = coupon_period_bounds(start, first_interest, months, system)?;
    let mut segment_start = start;
    let mut total = 0.0;

    loop {
        let segment_end = if end < period_end { end } else { period_end };

        let days_period = days_between(period_start, period_end, basis, system)?;
        if days_period <= 0 {
            return Err(ExcelError::Num);
        }
        let days_segment = days_between(segment_start, segment_end, basis, system)?;
        if days_segment < 0 {
            return Err(ExcelError::Num);
        }

        total += coupon * (days_segment as f64) / (days_period as f64);
        if !total.is_finite() {
            return Err(ExcelError::Num);
        }

        if segment_end == end {
            break;
        }

        period_start = period_end;
        period_end = date_time::edate(period_end, months, system)?;
        segment_start = period_start;
    }

    Ok(total)
}

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
    if !rate.is_finite() || rate <= 0.0 {
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
    if !rate.is_finite() || rate <= 0.0 {
        return Err(ExcelError::Num);
    }
    if !par.is_finite() || par <= 0.0 {
        return Err(ExcelError::Num);
    }
    if !(0..=4).contains(&basis) {
        return Err(ExcelError::Num);
    }

    let months = months_per_period(frequency).ok_or(ExcelError::Num)?;
    let coupon = par * rate / (frequency as f64);
    if !coupon.is_finite() {
        return Err(ExcelError::Num);
    }

    // `calc_method` controls whether interest accrues from `issue` (TRUE) or from the
    // last coupon date before settlement (FALSE). When settlement is on/before the
    // first interest date, there is no prior coupon payment, so both methods start
    // at `issue`.
    let accrual_start = if calc_method || settlement <= first_interest {
        issue
    } else {
        coupon_date_leq(settlement, first_interest, months, system)?
    };

    let result = accrued_interest_over_schedule(
        accrual_start,
        settlement,
        first_interest,
        months,
        coupon,
        basis,
        system,
    )?;
    if result.is_finite() {
        Ok(result)
    } else {
        Err(ExcelError::Num)
    }
}

