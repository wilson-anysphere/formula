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
