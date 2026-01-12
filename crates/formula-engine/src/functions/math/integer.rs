use crate::error::{ExcelError, ExcelResult};

fn checked_out(out: f64) -> ExcelResult<f64> {
    if out.is_finite() {
        Ok(out)
    } else {
        Err(ExcelError::Num)
    }
}

pub(super) fn trunc_to_i64(number: f64) -> ExcelResult<i64> {
    if !number.is_finite() {
        return Err(ExcelError::Num);
    }
    let t = number.trunc();
    if t < (i64::MIN as f64) || t > (i64::MAX as f64) {
        return Err(ExcelError::Num);
    }
    Ok(t as i64)
}

pub(super) fn trunc_to_u64_nonnegative(number: f64) -> ExcelResult<u64> {
    if !number.is_finite() {
        return Err(ExcelError::Num);
    }
    let t = number.trunc();
    if t < 0.0 || t > (u64::MAX as f64) {
        return Err(ExcelError::Num);
    }
    Ok(t as u64)
}

fn round_half_away_from_zero(x: f64) -> f64 {
    let base = x.trunc();
    let frac = x.fract().abs();
    if frac < 0.5 {
        base
    } else {
        base + x.signum()
    }
}

/// MROUND(number, multiple)
pub fn mround(number: f64, multiple: f64) -> ExcelResult<f64> {
    if !number.is_finite() || !multiple.is_finite() {
        return Err(ExcelError::Num);
    }
    if number == 0.0 || multiple == 0.0 {
        return Ok(0.0);
    }
    if number.signum() * multiple.signum() < 0.0 {
        return Err(ExcelError::Num);
    }
    let q = number / multiple;
    let rounded_q = round_half_away_from_zero(q);
    checked_out(rounded_q * multiple)
}

/// EVEN(number)
pub fn even(number: f64) -> ExcelResult<f64> {
    if !number.is_finite() {
        return Err(ExcelError::Num);
    }
    let sign = if number.is_sign_negative() { -1.0 } else { 1.0 };
    let n = number.abs();
    let mut i = n.ceil();
    // Round to next even integer.
    if i % 2.0 != 0.0 {
        i += 1.0;
    }
    checked_out(sign * i)
}

/// ODD(number)
pub fn odd(number: f64) -> ExcelResult<f64> {
    if !number.is_finite() {
        return Err(ExcelError::Num);
    }
    let sign = if number.is_sign_negative() { -1.0 } else { 1.0 };
    let n = number.abs();
    let mut i = n.ceil();
    // Round to next odd integer.
    if i % 2.0 == 0.0 {
        i += 1.0;
    }
    checked_out(sign * i)
}

/// ISEVEN(number)
pub fn iseven(number: f64) -> ExcelResult<bool> {
    let t = trunc_to_i64(number)?;
    Ok(t % 2 == 0)
}

/// ISODD(number)
pub fn isodd(number: f64) -> ExcelResult<bool> {
    let t = trunc_to_i64(number)?;
    Ok(t % 2 != 0)
}

/// QUOTIENT(numerator, denominator)
pub fn quotient(numerator: f64, denominator: f64) -> ExcelResult<f64> {
    if !numerator.is_finite() || !denominator.is_finite() {
        return Err(ExcelError::Num);
    }
    if denominator == 0.0 {
        return Err(ExcelError::Div0);
    }
    checked_out((numerator / denominator).trunc())
}

fn gcd_u64(mut a: u64, mut b: u64) -> u64 {
    while b != 0 {
        let r = a % b;
        a = b;
        b = r;
    }
    a
}

/// GCD(number1, [number2], ...)
pub fn gcd(numbers: &[f64]) -> ExcelResult<f64> {
    let mut g: u64 = 0;
    for &n in numbers {
        let v = trunc_to_u64_nonnegative(n)?;
        g = gcd_u64(g, v);
        if g == 1 {
            // Early-out: 1 is the smallest possible GCD.
            break;
        }
    }
    Ok(g as f64)
}

/// LCM(number1, [number2], ...)
pub fn lcm(numbers: &[f64]) -> ExcelResult<f64> {
    if numbers.is_empty() {
        return Ok(0.0);
    }

    let mut acc: u64 = 1;
    for &n in numbers {
        let v = trunc_to_u64_nonnegative(n)?;
        if v == 0 {
            return Ok(0.0);
        }
        let g = gcd_u64(acc, v);
        let res = (acc as u128 / g as u128) * (v as u128);
        if res > (u64::MAX as u128) {
            return Err(ExcelError::Num);
        }
        acc = res as u64;
    }

    Ok(acc as f64)
}

/// SQRTPI(number)
pub fn sqrtpi(number: f64) -> ExcelResult<f64> {
    if !number.is_finite() || number < 0.0 {
        return Err(ExcelError::Num);
    }
    checked_out((std::f64::consts::PI * number).sqrt())
}

/// DELTA(number1, [number2])
pub fn delta(number1: f64, number2: f64) -> ExcelResult<f64> {
    if !number1.is_finite() || !number2.is_finite() {
        return Err(ExcelError::Num);
    }
    Ok(if number1 == number2 { 1.0 } else { 0.0 })
}

/// GESTEP(number, [step])
pub fn gestep(number: f64, step: f64) -> ExcelResult<f64> {
    if !number.is_finite() || !step.is_finite() {
        return Err(ExcelError::Num);
    }
    Ok(if number >= step { 1.0 } else { 0.0 })
}
