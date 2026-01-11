mod aggregates;
pub mod criteria;
mod random;
mod rounding;
mod trig;

pub use aggregates::{
    aggregate, averageif, averageifs, countifs, maxifs, minifs, subtotal, sumif, sumifs, sumproduct,
};
pub use random::{rand, randbetween};
pub use rounding::{
    ceiling, ceiling_math, ceiling_precise, floor, floor_math, floor_precise, iso_ceiling,
};
pub use trig::{acos, asin, atan, atan2, cos, sin, tan};

use crate::error::{ExcelError, ExcelResult};

/// PRODUCT(number1, [number2], ...)
pub fn product(values: &[f64]) -> ExcelResult<f64> {
    let mut acc = 1.0;
    for value in values {
        if !value.is_finite() {
            return Err(ExcelError::Num);
        }
        acc *= value;
        if !acc.is_finite() {
            return Err(ExcelError::Num);
        }
    }
    Ok(acc)
}

/// POWER(number, power)
pub fn power(number: f64, power: f64) -> ExcelResult<f64> {
    if !number.is_finite() || !power.is_finite() {
        return Err(ExcelError::Num);
    }
    if number == 0.0 && power < 0.0 {
        return Err(ExcelError::Div0);
    }

    if number < 0.0 && !is_effectively_integer(power) {
        return Err(ExcelError::Num);
    }

    let out = number.powf(power);
    if out.is_finite() {
        Ok(out)
    } else {
        Err(ExcelError::Num)
    }
}

fn is_effectively_integer(x: f64) -> bool {
    const TOL: f64 = 1.0e-10;
    (x - x.round()).abs() <= TOL
}

/// LN(number)
pub fn ln(number: f64) -> ExcelResult<f64> {
    if !(number > 0.0) || !number.is_finite() {
        return Err(ExcelError::Num);
    }
    let out = number.ln();
    if out.is_finite() {
        Ok(out)
    } else {
        Err(ExcelError::Num)
    }
}

/// LOG(number, [base])
pub fn log(number: f64, base: Option<f64>) -> ExcelResult<f64> {
    if !(number > 0.0) || !number.is_finite() {
        return Err(ExcelError::Num);
    }
    let base = base.unwrap_or(10.0);
    if !(base > 0.0) || base == 1.0 || !base.is_finite() {
        return Err(ExcelError::Num);
    }
    let out = number.ln() / base.ln();
    if out.is_finite() {
        Ok(out)
    } else {
        Err(ExcelError::Num)
    }
}

/// EXP(number)
pub fn exp(number: f64) -> ExcelResult<f64> {
    if !number.is_finite() {
        return Err(ExcelError::Num);
    }
    let out = number.exp();
    if out.is_finite() {
        Ok(out)
    } else {
        Err(ExcelError::Num)
    }
}

/// LOG10(number)
pub fn log10(number: f64) -> ExcelResult<f64> {
    log(number, None)
}

/// SQRT(number)
pub fn sqrt(number: f64) -> ExcelResult<f64> {
    if number < 0.0 || !number.is_finite() {
        return Err(ExcelError::Num);
    }
    let out = number.sqrt();
    if out.is_finite() {
        Ok(out)
    } else {
        Err(ExcelError::Num)
    }
}

/// PI()
#[must_use]
pub fn pi() -> f64 {
    std::f64::consts::PI
}
