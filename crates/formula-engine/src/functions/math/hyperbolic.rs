use crate::error::{ExcelError, ExcelResult};

fn checked_out(out: f64) -> ExcelResult<f64> {
    if out.is_finite() {
        Ok(out)
    } else {
        Err(ExcelError::Num)
    }
}

/// SINH(number)
pub fn sinh(number: f64) -> ExcelResult<f64> {
    if !number.is_finite() {
        return Err(ExcelError::Num);
    }
    checked_out(number.sinh())
}

/// COSH(number)
pub fn cosh(number: f64) -> ExcelResult<f64> {
    if !number.is_finite() {
        return Err(ExcelError::Num);
    }
    checked_out(number.cosh())
}

/// TANH(number)
pub fn tanh(number: f64) -> ExcelResult<f64> {
    if !number.is_finite() {
        return Err(ExcelError::Num);
    }
    checked_out(number.tanh())
}

/// ASINH(number)
pub fn asinh(number: f64) -> ExcelResult<f64> {
    if !number.is_finite() {
        return Err(ExcelError::Num);
    }
    checked_out(number.asinh())
}

/// ACOSH(number)
pub fn acosh(number: f64) -> ExcelResult<f64> {
    if !number.is_finite() || number < 1.0 {
        return Err(ExcelError::Num);
    }
    checked_out(number.acosh())
}

/// ATANH(number)
pub fn atanh(number: f64) -> ExcelResult<f64> {
    if !number.is_finite() || number <= -1.0 || number >= 1.0 {
        return Err(ExcelError::Num);
    }
    checked_out(number.atanh())
}

/// COTH(number)
pub fn coth(number: f64) -> ExcelResult<f64> {
    if !number.is_finite() {
        return Err(ExcelError::Num);
    }
    let tanh = number.tanh();
    if tanh == 0.0 {
        return Err(ExcelError::Div0);
    }
    checked_out(1.0 / tanh)
}

/// CSCH(number)
pub fn csch(number: f64) -> ExcelResult<f64> {
    if !number.is_finite() {
        return Err(ExcelError::Num);
    }
    let sinh = number.sinh();
    if sinh == 0.0 {
        return Err(ExcelError::Div0);
    }
    checked_out(1.0 / sinh)
}

/// SECH(number)
pub fn sech(number: f64) -> ExcelResult<f64> {
    if !number.is_finite() {
        return Err(ExcelError::Num);
    }
    checked_out(1.0 / number.cosh())
}

/// ACOTH(number)
pub fn acoth(number: f64) -> ExcelResult<f64> {
    if !number.is_finite() || number.abs() <= 1.0 {
        return Err(ExcelError::Num);
    }
    // ACOTH(x) = ATANH(1/x) for |x|>1.
    atanh(1.0 / number)
}
