use crate::error::{ExcelError, ExcelResult};

fn checked_out(out: f64) -> ExcelResult<f64> {
    if out.is_finite() {
        Ok(out)
    } else {
        Err(ExcelError::Num)
    }
}

/// RADIANS(angle)
pub fn radians(angle: f64) -> ExcelResult<f64> {
    if !angle.is_finite() {
        return Err(ExcelError::Num);
    }
    checked_out(angle * std::f64::consts::PI / 180.0)
}

/// DEGREES(angle)
pub fn degrees(angle: f64) -> ExcelResult<f64> {
    if !angle.is_finite() {
        return Err(ExcelError::Num);
    }
    checked_out(angle * 180.0 / std::f64::consts::PI)
}

/// COT(number)
pub fn cot(number: f64) -> ExcelResult<f64> {
    if !number.is_finite() {
        return Err(ExcelError::Num);
    }
    let tan = number.tan();
    if tan == 0.0 {
        return Err(ExcelError::Div0);
    }
    checked_out(1.0 / tan)
}

/// CSC(number)
pub fn csc(number: f64) -> ExcelResult<f64> {
    if !number.is_finite() {
        return Err(ExcelError::Num);
    }
    let sin = number.sin();
    if sin == 0.0 {
        return Err(ExcelError::Div0);
    }
    checked_out(1.0 / sin)
}

/// SEC(number)
pub fn sec(number: f64) -> ExcelResult<f64> {
    if !number.is_finite() {
        return Err(ExcelError::Num);
    }
    let cos = number.cos();
    if cos == 0.0 {
        return Err(ExcelError::Div0);
    }
    checked_out(1.0 / cos)
}

/// ACOT(number)
///
/// Returns the inverse cotangent of `number`, in radians, in the range (0, PI).
pub fn acot(number: f64) -> ExcelResult<f64> {
    if !number.is_finite() {
        return Err(ExcelError::Num);
    }

    // Excel defines ACOT(0) = PI/2 and returns values in (0, PI).
    if number == 0.0 {
        return Ok(std::f64::consts::FRAC_PI_2);
    }

    let base = (1.0 / number).atan();
    if number > 0.0 {
        checked_out(base)
    } else {
        checked_out(base + std::f64::consts::PI)
    }
}
