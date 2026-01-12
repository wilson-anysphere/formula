use crate::error::{ExcelError, ExcelResult};
use crate::value::ErrorKind;

use super::integer::trunc_to_i64;

fn pow_i64(mut base: f64, exp: i64) -> ExcelResult<f64> {
    if !base.is_finite() {
        return Err(ExcelError::Num);
    }
    if exp == 0 {
        return Ok(1.0);
    }

    let mut e: u64 = if exp < 0 {
        if base == 0.0 {
            return Err(ExcelError::Div0);
        }
        base = 1.0 / base;
        // `abs(i64::MIN)` overflows.
        u64::try_from(exp.checked_neg().ok_or(ExcelError::Num)?).map_err(|_| ExcelError::Num)?
    } else {
        u64::try_from(exp).map_err(|_| ExcelError::Num)?
    };

    let mut acc = 1.0;
    while e > 0 {
        if e & 1 == 1 {
            acc *= base;
            if !acc.is_finite() {
                return Err(ExcelError::Num);
            }
        }
        e >>= 1;
        if e == 0 {
            break;
        }
        base *= base;
        if !base.is_finite() {
            // If base overflowed, any subsequent multiplication would overflow or become NaN.
            return Err(ExcelError::Num);
        }
    }
    Ok(acc)
}

/// SERIESSUM(x, n, m, coefficients)
pub fn seriessum(x: f64, n: f64, m: f64, coefficients: &[f64]) -> ExcelResult<f64> {
    if !x.is_finite() || !n.is_finite() || !m.is_finite() {
        return Err(ExcelError::Num);
    }

    let n = trunc_to_i64(n)?;
    let m = trunc_to_i64(m)?;

    let mut acc = 0.0;
    for (idx, &c) in coefficients.iter().enumerate() {
        if !c.is_finite() {
            return Err(ExcelError::Num);
        }
        let i = i64::try_from(idx).map_err(|_| ExcelError::Num)?;
        let step = m.checked_mul(i).ok_or(ExcelError::Num)?;
        let exp = n.checked_add(step).ok_or(ExcelError::Num)?;
        let term = c * pow_i64(x, exp)?;
        acc += term;
        if !acc.is_finite() {
            return Err(ExcelError::Num);
        }
    }
    Ok(acc)
}

/// SUMXMY2(array_x, array_y)
pub fn sumxmy2(xs: &[f64], ys: &[f64]) -> Result<f64, ErrorKind> {
    if xs.len() != ys.len() {
        return Err(ErrorKind::NA);
    }
    let mut acc = 0.0;
    for (&x, &y) in xs.iter().zip(ys.iter()) {
        if !x.is_finite() || !y.is_finite() {
            return Err(ErrorKind::Num);
        }
        let d = x - y;
        acc += d * d;
        if !acc.is_finite() {
            return Err(ErrorKind::Num);
        }
    }
    Ok(acc)
}

/// SUMX2MY2(array_x, array_y)
pub fn sumx2my2(xs: &[f64], ys: &[f64]) -> Result<f64, ErrorKind> {
    if xs.len() != ys.len() {
        return Err(ErrorKind::NA);
    }
    let mut acc = 0.0;
    for (&x, &y) in xs.iter().zip(ys.iter()) {
        if !x.is_finite() || !y.is_finite() {
            return Err(ErrorKind::Num);
        }
        acc += x * x - y * y;
        if !acc.is_finite() {
            return Err(ErrorKind::Num);
        }
    }
    Ok(acc)
}

/// SUMX2PY2(array_x, array_y)
pub fn sumx2py2(xs: &[f64], ys: &[f64]) -> Result<f64, ErrorKind> {
    if xs.len() != ys.len() {
        return Err(ErrorKind::NA);
    }
    let mut acc = 0.0;
    for (&x, &y) in xs.iter().zip(ys.iter()) {
        if !x.is_finite() || !y.is_finite() {
            return Err(ErrorKind::Num);
        }
        acc += x * x + y * y;
        if !acc.is_finite() {
            return Err(ErrorKind::Num);
        }
    }
    Ok(acc)
}
