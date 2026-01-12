use crate::error::{ExcelError, ExcelResult};

/// Interest paid during a specific period of an investment with equal principal payments.
///
/// Excel semantics:
/// - `per` must be in `[1, nper]`.
/// - Non-finite inputs return `#NUM!`.
pub fn ispmt(rate: f64, per: f64, nper: f64, pv: f64) -> ExcelResult<f64> {
    if !rate.is_finite() || !per.is_finite() || !nper.is_finite() || !pv.is_finite() {
        return Err(ExcelError::Num);
    }
    if per < 1.0 || per > nper {
        return Err(ExcelError::Num);
    }
    if nper == 0.0 {
        return Err(ExcelError::Num);
    }

    // Excel's definition:
    // ISPMT = pv * rate * (per - 1) / nper - pv * rate
    //       = -pv * rate * (nper - per + 1) / nper
    let result = -pv * rate * (nper - per + 1.0) / nper;
    if result.is_finite() {
        Ok(result)
    } else {
        Err(ExcelError::Num)
    }
}

fn normalize_fraction_denom(fraction: f64) -> ExcelResult<i64> {
    if !fraction.is_finite() {
        return Err(ExcelError::Num);
    }

    // Excel truncates the `fraction` argument to an integer.
    let denom_f = fraction.trunc();
    if denom_f == 0.0 {
        return Err(ExcelError::Div0);
    }
    if denom_f < 0.0 {
        return Err(ExcelError::Num);
    }
    if denom_f > (i64::MAX as f64) {
        return Err(ExcelError::Num);
    }
    Ok(denom_f as i64)
}

fn denom_scale_10(denom: i64) -> ExcelResult<f64> {
    debug_assert!(denom > 0);

    // The scale is 10 ^ (number of decimal digits in denom).
    let mut d = denom;
    let mut digits: i32 = 0;
    while d > 0 {
        digits += 1;
        d /= 10;
    }
    let scale = 10f64.powi(digits);
    if scale.is_finite() && scale != 0.0 {
        Ok(scale)
    } else {
        Err(ExcelError::Num)
    }
}

/// Converts a dollar price expressed as a fraction into a dollar price expressed as a decimal.
///
/// Excel semantics:
/// - `fraction` is truncated to an integer; `fraction == 0` -> `#DIV/0!`, `fraction < 0` -> `#NUM!`.
/// - Non-finite inputs return `#NUM!`.
pub fn dollarde(fractional_dollar: f64, fraction: f64) -> ExcelResult<f64> {
    if !fractional_dollar.is_finite() || !fraction.is_finite() {
        return Err(ExcelError::Num);
    }

    let denom = normalize_fraction_denom(fraction)?;
    let scale = denom_scale_10(denom)?;

    let sign = if fractional_dollar.is_sign_negative() {
        -1.0
    } else {
        1.0
    };
    let abs = fractional_dollar.abs();
    let int_part = abs.trunc();
    let frac_part = abs - int_part;

    // Extract the fractional numerator encoded in the decimal digits.
    let numerator = (frac_part * scale).round();
    let result = sign * (int_part + numerator / (denom as f64));
    if result.is_finite() {
        Ok(result)
    } else {
        Err(ExcelError::Num)
    }
}

/// Converts a dollar price expressed as a decimal into a dollar price expressed as a fraction.
///
/// Excel semantics:
/// - `fraction` is truncated to an integer; `fraction == 0` -> `#DIV/0!`, `fraction < 0` -> `#NUM!`.
/// - Non-finite inputs return `#NUM!`.
pub fn dollarfr(decimal_dollar: f64, fraction: f64) -> ExcelResult<f64> {
    if !decimal_dollar.is_finite() || !fraction.is_finite() {
        return Err(ExcelError::Num);
    }

    let denom = normalize_fraction_denom(fraction)?;
    let scale = denom_scale_10(denom)?;

    let sign = if decimal_dollar.is_sign_negative() {
        -1.0
    } else {
        1.0
    };
    let abs = decimal_dollar.abs();
    let int_part = abs.trunc();
    let frac_part = abs - int_part;

    // Convert the decimal fractional part into a numerator in units of `denom`.
    let numerator = (frac_part * (denom as f64)).round();
    let result = sign * (int_part + numerator / scale);
    if result.is_finite() {
        Ok(result)
    } else {
        Err(ExcelError::Num)
    }
}
