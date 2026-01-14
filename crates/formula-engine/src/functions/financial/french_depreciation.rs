use crate::date::ExcelDateSystem;
use crate::error::{ExcelError, ExcelResult};
use crate::functions::date_time;

const ROUND_EPS: f64 = 1e-12;

fn round_down_to_int(value: f64) -> f64 {
    // Excel's AMOR* functions round depreciation amounts to an integer value.
    // This uses a small epsilon to avoid floating-point artifacts around exact integers.
    if value.is_sign_negative() {
        (value - ROUND_EPS).ceil()
    } else {
        (value + ROUND_EPS).floor()
    }
}

fn degressive_coefficient(life_years: f64) -> f64 {
    // Excel's AMORDEGRC uses a fixed coefficient based on the asset's life.
    // The life is implied by the rate: life = 1 / rate.
    //
    // Coefficient table (Excel docs):
    // - life < 3 years: 1.0
    // - 3..=4 years: 1.5
    // - 5..=6 years: 2.0
    // - > 6 years: 2.5
    if life_years < 3.0 {
        1.0
    } else if life_years < 5.0 {
        1.5
    } else if life_years <= 6.0 {
        2.0
    } else {
        2.5
    }
}

fn validate_common(
    cost: f64,
    date_purchased: i32,
    first_period: i32,
    salvage: f64,
    period: f64,
    rate: f64,
    basis: Option<i32>,
    system: ExcelDateSystem,
) -> ExcelResult<(f64, i32, i32)> {
    if !cost.is_finite() || !(cost > 0.0) {
        return Err(ExcelError::Num);
    }
    if !salvage.is_finite() || !(salvage >= 0.0) {
        return Err(ExcelError::Num);
    }
    if !rate.is_finite() || !(rate > 0.0) {
        return Err(ExcelError::Num);
    }
    if date_purchased > first_period {
        return Err(ExcelError::Num);
    }
    if !period.is_finite() {
        return Err(ExcelError::Num);
    }
    if period < 0.0 {
        return Err(ExcelError::Num);
    }

    let basis = basis.unwrap_or(0);
    let basis = super::coupon_schedule::validate_basis(basis)?;

    // Ensure the dates are representable under the date system and basis (yearfrac validates).
    let first_period_years = date_time::yearfrac(date_purchased, first_period, basis, system)?;
    if !first_period_years.is_finite() || first_period_years < 0.0 {
        return Err(ExcelError::Num);
    }

    let life_years = 1.0 / rate;
    if !life_years.is_finite() || life_years <= 0.0 {
        return Err(ExcelError::Num);
    }

    let truncated_period = period.trunc();
    if truncated_period < 0.0 || truncated_period > (i32::MAX as f64) {
        return Err(ExcelError::Num);
    }
    let period_i32 = truncated_period as i32;

    // Period indices are integers starting at 0. Period 0 is the first (possibly partial)
    // period between `date_purchased` and `first_period`.
    //
    // Subsequent periods are 1-year increments. The final period may be partial to ensure the
    // total time covered equals the life implied by `rate`.
    //
    // Excel errors (#NUM) when `period` is outside the implied asset life.
    let remaining_after_first = life_years - first_period_years;
    let max_period = if remaining_after_first <= 0.0 {
        0
    } else {
        // `ceil` gives the number of additional (full/partial) periods after period 0.
        let mut v = (remaining_after_first - ROUND_EPS).ceil();
        if v < 0.0 {
            v = 0.0;
        }
        if v > (i32::MAX as f64) {
            i32::MAX
        } else {
            v as i32
        }
    };
    if period_i32 > max_period {
        return Err(ExcelError::Num);
    }

    Ok((life_years, basis, period_i32))
}

/// AMORLINC(cost, date_purchased, first_period, salvage, period, rate, [basis])
pub fn amorlinc(
    cost: f64,
    date_purchased: i32,
    first_period: i32,
    salvage: f64,
    period: f64,
    rate: f64,
    basis: Option<i32>,
    system: ExcelDateSystem,
) -> ExcelResult<f64> {
    let (life_years, basis, period_i32) = validate_common(
        cost,
        date_purchased,
        first_period,
        salvage,
        period,
        rate,
        basis,
        system,
    )?;

    let first_period_years = date_time::yearfrac(date_purchased, first_period, basis, system)?;

    let mut remaining_life = life_years;
    let mut book_value = cost;

    for idx in 0..=period_i32 {
        let candidate_years = if idx == 0 { first_period_years } else { 1.0 };
        let period_years = candidate_years.min(remaining_life);
        let is_last_period = (remaining_life - period_years).abs() <= ROUND_EPS;

        let max_dep = book_value - salvage;
        let dep = if max_dep <= 0.0 {
            0.0
        } else if is_last_period {
            // Final period: adjust to end exactly at `salvage`.
            max_dep
        } else {
            let raw = cost * rate * period_years;
            let mut rounded = round_down_to_int(raw);
            if rounded > max_dep {
                rounded = max_dep;
            }
            if rounded < 0.0 {
                0.0
            } else {
                rounded
            }
        };

        if idx == period_i32 {
            return Ok(dep);
        }

        book_value -= dep;
        remaining_life -= period_years;
    }

    // With the period range validation above, the loop always returns before reaching here.
    Err(ExcelError::Num)
}

/// AMORDEGRC(cost, date_purchased, first_period, salvage, period, rate, [basis])
pub fn amordegrec(
    cost: f64,
    date_purchased: i32,
    first_period: i32,
    salvage: f64,
    period: f64,
    rate: f64,
    basis: Option<i32>,
    system: ExcelDateSystem,
) -> ExcelResult<f64> {
    let (life_years, basis, period_i32) = validate_common(
        cost,
        date_purchased,
        first_period,
        salvage,
        period,
        rate,
        basis,
        system,
    )?;

    let first_period_years = date_time::yearfrac(date_purchased, first_period, basis, system)?;

    let coef = degressive_coefficient(life_years);
    let degressive_rate = rate * coef;

    let mut remaining_life = life_years;
    let mut book_value = cost;
    let mut switched_to_linear = false;

    for idx in 0..=period_i32 {
        let candidate_years = if idx == 0 { first_period_years } else { 1.0 };
        let period_years = candidate_years.min(remaining_life);
        let is_last_period = (remaining_life - period_years).abs() <= ROUND_EPS;

        let max_dep = book_value - salvage;
        let dep = if max_dep <= 0.0 {
            0.0
        } else if is_last_period {
            max_dep
        } else {
            if !switched_to_linear {
                let linear_rate = 1.0 / remaining_life;
                if degressive_rate <= linear_rate {
                    switched_to_linear = true;
                }
            }

            let raw = if switched_to_linear {
                (book_value / remaining_life) * period_years
            } else {
                book_value * degressive_rate * period_years
            };
            let mut rounded = round_down_to_int(raw);
            if rounded > max_dep {
                rounded = max_dep;
            }
            if rounded < 0.0 {
                0.0
            } else {
                rounded
            }
        };

        if idx == period_i32 {
            return Ok(dep);
        }

        book_value -= dep;
        remaining_life -= period_years;
    }

    Err(ExcelError::Num)
}
