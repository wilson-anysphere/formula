use crate::error::{ExcelError, ExcelResult};

fn checked_out(out: f64) -> ExcelResult<f64> {
    if out.is_finite() {
        Ok(out)
    } else {
        Err(ExcelError::Num)
    }
}

/// SIN(number)
pub fn sin(number: f64) -> ExcelResult<f64> {
    if !number.is_finite() {
        return Err(ExcelError::Num);
    }
    checked_out(number.sin())
}

/// COS(number)
pub fn cos(number: f64) -> ExcelResult<f64> {
    if !number.is_finite() {
        return Err(ExcelError::Num);
    }
    checked_out(number.cos())
}

/// TAN(number)
pub fn tan(number: f64) -> ExcelResult<f64> {
    if !number.is_finite() {
        return Err(ExcelError::Num);
    }
    checked_out(number.tan())
}

/// ASIN(number)
pub fn asin(number: f64) -> ExcelResult<f64> {
    if !number.is_finite() || number < -1.0 || number > 1.0 {
        return Err(ExcelError::Num);
    }
    checked_out(number.asin())
}

/// ACOS(number)
pub fn acos(number: f64) -> ExcelResult<f64> {
    if !number.is_finite() || number < -1.0 || number > 1.0 {
        return Err(ExcelError::Num);
    }
    checked_out(number.acos())
}

/// ATAN(number)
pub fn atan(number: f64) -> ExcelResult<f64> {
    if !number.is_finite() {
        return Err(ExcelError::Num);
    }
    checked_out(number.atan())
}

/// ATAN2(x_num, y_num)
///
/// Note: Excel's argument order is `(x_num, y_num)`, which is the reverse of many
/// standard library `atan2(y, x)` APIs. We forward to `atan2(y_num, x_num)` accordingly.
pub fn atan2(x_num: f64, y_num: f64) -> ExcelResult<f64> {
    if !x_num.is_finite() || !y_num.is_finite() {
        return Err(ExcelError::Num);
    }
    if x_num == 0.0 && y_num == 0.0 {
        return Err(ExcelError::Div0);
    }
    let mut out = y_num.atan2(x_num);
    // Excel documents the return range as (-PI, PI], excluding -PI; some `atan2`
    // implementations return -PI when y is -0 and x is negative.
    if out == -std::f64::consts::PI {
        out = std::f64::consts::PI;
    }
    checked_out(out)
}
