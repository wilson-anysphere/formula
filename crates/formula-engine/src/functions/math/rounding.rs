use crate::error::{ExcelError, ExcelResult};

fn round_to_multiple(number: f64, significance: f64, direction: RoundingDirection) -> ExcelResult<f64> {
    if !number.is_finite() || !significance.is_finite() {
        return Err(ExcelError::Num);
    }
    if number == 0.0 || significance == 0.0 {
        return Ok(0.0);
    }

    let step = significance.abs();
    let quotient = number / step;
    let q = match direction {
        RoundingDirection::TowardPositiveInfinity => quotient.ceil(),
        RoundingDirection::TowardNegativeInfinity => quotient.floor(),
    };
    let out = q * step;
    if out.is_finite() {
        Ok(out)
    } else {
        Err(ExcelError::Num)
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum RoundingDirection {
    TowardPositiveInfinity,
    TowardNegativeInfinity,
}

/// CEILING(number, significance)
pub fn ceiling(number: f64, significance: f64) -> ExcelResult<f64> {
    if number.signum() * significance.signum() < 0.0 {
        return Err(ExcelError::Num);
    }
    round_to_multiple(number, significance, RoundingDirection::TowardPositiveInfinity)
}

/// FLOOR(number, significance)
pub fn floor(number: f64, significance: f64) -> ExcelResult<f64> {
    if number.signum() * significance.signum() < 0.0 {
        return Err(ExcelError::Num);
    }
    round_to_multiple(number, significance, RoundingDirection::TowardNegativeInfinity)
}

/// CEILING.MATH(number, [significance], [mode])
pub fn ceiling_math(number: f64, significance: Option<f64>, mode: Option<f64>) -> ExcelResult<f64> {
    let step = significance.unwrap_or(1.0);
    let mode = mode.unwrap_or(0.0);
    let negative_away_from_zero = number < 0.0 && mode != 0.0;
    let direction = if negative_away_from_zero {
        RoundingDirection::TowardNegativeInfinity
    } else {
        RoundingDirection::TowardPositiveInfinity
    };
    round_to_multiple(number, step, direction)
}

/// FLOOR.MATH(number, [significance], [mode])
pub fn floor_math(number: f64, significance: Option<f64>, mode: Option<f64>) -> ExcelResult<f64> {
    let step = significance.unwrap_or(1.0);
    let mode = mode.unwrap_or(0.0);
    let negative_toward_zero = number < 0.0 && mode != 0.0;
    let direction = if negative_toward_zero {
        RoundingDirection::TowardPositiveInfinity
    } else {
        RoundingDirection::TowardNegativeInfinity
    };
    round_to_multiple(number, step, direction)
}

/// CEILING.PRECISE(number, [significance])
pub fn ceiling_precise(number: f64, significance: Option<f64>) -> ExcelResult<f64> {
    round_to_multiple(
        number,
        significance.unwrap_or(1.0).abs(),
        RoundingDirection::TowardPositiveInfinity,
    )
}

/// FLOOR.PRECISE(number, [significance])
pub fn floor_precise(number: f64, significance: Option<f64>) -> ExcelResult<f64> {
    round_to_multiple(
        number,
        significance.unwrap_or(1.0).abs(),
        RoundingDirection::TowardNegativeInfinity,
    )
}

/// ISO.CEILING(number, [significance])
pub fn iso_ceiling(number: f64, significance: Option<f64>) -> ExcelResult<f64> {
    ceiling_precise(number, significance)
}

