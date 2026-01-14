use crate::error::{ExcelError, ExcelResult};

pub fn sln(cost: f64, salvage: f64, life: f64) -> ExcelResult<f64> {
    if life == 0.0 {
        return Err(ExcelError::Div0);
    }
    if life < 0.0 {
        return Err(ExcelError::Num);
    }
    Ok((cost - salvage) / life)
}

pub fn syd(cost: f64, salvage: f64, life: f64, per: f64) -> ExcelResult<f64> {
    if life <= 0.0 {
        return Err(ExcelError::Num);
    }
    if per <= 0.0 || per > life {
        return Err(ExcelError::Num);
    }

    let syd = life * (life + 1.0) / 2.0;
    Ok((cost - salvage) * (life - per + 1.0) / syd)
}

pub fn ddb(
    cost: f64,
    salvage: f64,
    life: f64,
    period: f64,
    factor: Option<f64>,
) -> ExcelResult<f64> {
    let factor = factor.unwrap_or(2.0);
    if life <= 0.0 {
        return Err(ExcelError::Num);
    }
    if period <= 0.0 || period > life {
        return Err(ExcelError::Num);
    }
    if factor <= 0.0 {
        return Err(ExcelError::Num);
    }

    let mut accumulated = 0.0;
    let target_period = period.floor() as i32;
    for _ in 1..target_period {
        let dep = depreciation_step(cost, salvage, life, factor, accumulated);
        accumulated += dep;
    }

    Ok(depreciation_step(cost, salvage, life, factor, accumulated))
}

fn depreciation_step(cost: f64, salvage: f64, life: f64, factor: f64, accumulated: f64) -> f64 {
    let remaining = cost - accumulated;
    if remaining <= salvage {
        return 0.0;
    }

    let mut dep = remaining * factor / life;
    let max_dep = remaining - salvage;
    if dep > max_dep {
        dep = max_dep;
    }
    if dep < 0.0 {
        0.0
    } else {
        dep
    }
}

/// DB(cost, salvage, life, period, [month])
///
/// Fixed-declining balance depreciation. Matches Excel behavior:
/// - The depreciation rate is rounded to 3 decimals.
/// - The first period is prorated by `month/12` (default `month = 12`).
/// - If `month != 12`, Excel adds one extra period (`life + 1`) prorated by `(12-month)/12`.
pub fn db(cost: f64, salvage: f64, life: f64, period: f64, month: Option<f64>) -> ExcelResult<f64> {
    let month = month.unwrap_or(12.0);
    if !cost.is_finite()
        || !salvage.is_finite()
        || !life.is_finite()
        || !period.is_finite()
        || !month.is_finite()
    {
        return Err(ExcelError::Num);
    }

    if cost <= 0.0 || salvage < 0.0 || life <= 0.0 || period < 1.0 || !(1.0..=12.0).contains(&month)
    {
        return Err(ExcelError::Num);
    }

    // Excel truncates fractional periods.
    let target_period = period.floor() as i32;
    if target_period < 1 {
        return Err(ExcelError::Num);
    }

    // Excel returns #NUM! if `period` is greater than `life` (or `life + 1` when the first year
    // is partial, i.e. `month != 12`).
    let max_period = if month == 12.0 { life } else { life + 1.0 };
    if (target_period as f64) > max_period {
        return Err(ExcelError::Num);
    }

    // Depreciation rate with 3-decimal rounding.
    let ratio = salvage / cost;
    let rate_raw = 1.0 - ratio.powf(1.0 / life);
    if !rate_raw.is_finite() {
        return Err(ExcelError::Num);
    }
    let rate = (rate_raw * 1000.0).round() / 1000.0;

    let mut accumulated = 0.0;
    let mut dep = 0.0;
    for p in 1..=target_period {
        let remaining = cost - accumulated;
        if remaining <= salvage {
            dep = 0.0;
        } else if p == 1 {
            dep = cost * rate * (month / 12.0);
        } else if month != 12.0 && (p as f64) > life {
            // When the first year is partial (month < 12), Excel adds one extra depreciation
            // period for the remaining months in the last year.
            dep = remaining * rate * ((12.0 - month) / 12.0);
        } else {
            dep = remaining * rate;
        }

        // Ensure we never depreciate below the salvage value (and never return negative depreciation).
        let max_dep = remaining - salvage;
        if dep > max_dep {
            dep = max_dep;
        }
        if dep < 0.0 {
            dep = 0.0;
        }

        accumulated += dep;
    }

    Ok(dep)
}

/// VDB(cost, salvage, life, start_period, end_period, [factor], [no_switch])
///
/// Variable declining balance depreciation. Implements the Excel semantics:
/// - Supports fractional start/end periods by prorating the overlapping portion of each period.
/// - Uses DDB-style depreciation with an optional switch to straight-line when it produces a
///   larger depreciation amount (unless `no_switch` is TRUE / non-zero).
pub fn vdb(
    cost: f64,
    salvage: f64,
    life: f64,
    start: f64,
    end: f64,
    factor: Option<f64>,
    no_switch: Option<f64>,
) -> ExcelResult<f64> {
    let factor = factor.unwrap_or(2.0);
    let no_switch = no_switch.unwrap_or(0.0) != 0.0;

    if !cost.is_finite()
        || !salvage.is_finite()
        || !life.is_finite()
        || !start.is_finite()
        || !end.is_finite()
        || !factor.is_finite()
    {
        return Err(ExcelError::Num);
    }

    if cost <= 0.0 || salvage < 0.0 || life <= 0.0 || start < 0.0 || end <= start || factor <= 0.0 {
        return Err(ExcelError::Num);
    }

    // Excel requires end_period <= life.
    if end > life {
        return Err(ExcelError::Num);
    }

    let mut accumulated = 0.0;
    let mut out = 0.0;

    // Period indices are 1-based, but start/end are 0-based "time" offsets. Period `p` covers
    // the interval [p-1, p].
    let last_period = end.ceil() as i32;
    for p in 1..=last_period {
        let period_start = (p - 1) as f64;
        let period_end = p as f64;

        // Compute full-period depreciation for this period.
        let remaining = cost - accumulated;
        let dep_full = if remaining <= salvage {
            0.0
        } else {
            let ddb = depreciation_step(cost, salvage, life, factor, accumulated);
            if no_switch {
                ddb
            } else {
                let remaining_life = life - period_start;
                if remaining_life <= 0.0 {
                    0.0
                } else {
                    let sl = (remaining - salvage) / remaining_life;
                    if sl > ddb {
                        sl
                    } else {
                        ddb
                    }
                }
            }
        };

        // Add the overlapping fraction of this period to the output (Excel prorates partial periods).
        let overlap_start = start.max(period_start);
        let overlap_end = end.min(period_end);
        if overlap_end > overlap_start && dep_full != 0.0 {
            let frac = overlap_end - overlap_start;
            out += dep_full * frac;
        }

        accumulated += dep_full;
    }

    // Guard against small floating-point overshoots (and enforce the Excel contract).
    let max_total = cost - salvage;
    if out > max_total {
        out = max_total;
    }
    if out < 0.0 {
        out = 0.0;
    }

    Ok(out)
}
