use crate::error::{ExcelError, ExcelResult};

use super::integer::trunc_to_u64_nonnegative;

fn checked_out(out: f64) -> ExcelResult<f64> {
    if out.is_finite() {
        Ok(out)
    } else {
        Err(ExcelError::Num)
    }
}

/// FACT(number)
pub fn fact(number: f64) -> ExcelResult<f64> {
    let n = trunc_to_u64_nonnegative(number)?;
    // FACT(0)=1, and FACT values beyond 170 overflow IEEE doubles.
    if n > 170 {
        return Err(ExcelError::Num);
    }
    let mut acc = 1.0;
    for i in 2..=n {
        acc *= i as f64;
    }
    checked_out(acc)
}

/// FACTDOUBLE(number)
pub fn factdouble(number: f64) -> ExcelResult<f64> {
    let n = trunc_to_u64_nonnegative(number)?;
    if n <= 1 {
        return Ok(1.0);
    }
    let mut acc = 1.0;
    let mut i = n;
    while i > 1 {
        acc *= i as f64;
        if !acc.is_finite() {
            return Err(ExcelError::Num);
        }
        i = i.saturating_sub(2);
    }
    Ok(acc)
}

fn combin_u64(n: u64, k: u64) -> ExcelResult<f64> {
    if k > n {
        return Err(ExcelError::Num);
    }
    let k = k.min(n - k);
    if k == 0 {
        return Ok(1.0);
    }

    let mut acc = 1.0;
    // Multiplicative formula: C(n,k) = Π_{i=1..k} (n-k+i)/i
    for i in 1..=k {
        let num = (n - k + i) as f64;
        let den = i as f64;
        acc *= num / den;
        if !acc.is_finite() {
            return Err(ExcelError::Num);
        }
    }
    Ok(acc)
}

/// COMBIN(number, number_chosen)
pub fn combin(number: f64, number_chosen: f64) -> ExcelResult<f64> {
    let n = trunc_to_u64_nonnegative(number)?;
    let k = trunc_to_u64_nonnegative(number_chosen)?;
    combin_u64(n, k)
}

/// COMBINA(number, number_chosen)
pub fn combina(number: f64, number_chosen: f64) -> ExcelResult<f64> {
    let n = trunc_to_u64_nonnegative(number)?;
    let k = trunc_to_u64_nonnegative(number_chosen)?;
    if n == 0 {
        return if k == 0 {
            Ok(1.0)
        } else {
            Err(ExcelError::Num)
        };
    }
    let total = n
        .checked_add(k)
        .and_then(|v| v.checked_sub(1))
        .ok_or(ExcelError::Num)?;
    combin_u64(total, k)
}

/// PERMUT(number, number_chosen)
pub fn permut(number: f64, number_chosen: f64) -> ExcelResult<f64> {
    let n = trunc_to_u64_nonnegative(number)?;
    let k = trunc_to_u64_nonnegative(number_chosen)?;
    if k > n {
        return Err(ExcelError::Num);
    }
    if k == 0 {
        return Ok(1.0);
    }
    let mut acc = 1.0;
    for i in 0..k {
        acc *= (n - i) as f64;
        if !acc.is_finite() {
            return Err(ExcelError::Num);
        }
    }
    Ok(acc)
}

fn pow_u64(mut base: f64, mut exp: u64) -> ExcelResult<f64> {
    if !base.is_finite() {
        return Err(ExcelError::Num);
    }
    if exp == 0 {
        return Ok(1.0);
    }
    let mut acc = 1.0;
    while exp > 0 {
        if exp & 1 == 1 {
            acc *= base;
            if !acc.is_finite() {
                return Err(ExcelError::Num);
            }
        }
        exp >>= 1;
        if exp == 0 {
            break;
        }
        base *= base;
        if !base.is_finite() {
            // If the base overflowed, any remaining multiplication would overflow too (or produce NaN).
            return Err(ExcelError::Num);
        }
    }
    Ok(acc)
}

/// PERMUTATIONA(number, number_chosen)
pub fn permutationa(number: f64, number_chosen: f64) -> ExcelResult<f64> {
    let n = trunc_to_u64_nonnegative(number)?;
    let k = trunc_to_u64_nonnegative(number_chosen)?;
    if n == 0 {
        return Err(ExcelError::Num);
    }
    pow_u64(n as f64, k)
}

/// MULTINOMIAL(number1, [number2], ...)
pub fn multinomial(numbers: &[f64]) -> ExcelResult<f64> {
    let mut parts: Vec<u64> = Vec::new();
    if parts.try_reserve_exact(numbers.len()).is_err() {
        debug_assert!(
            false,
            "allocation failed (multinomial parts, len={})",
            numbers.len()
        );
        return Err(ExcelError::Num);
    }
    let mut total: u64 = 0;
    for &n in numbers {
        let v = trunc_to_u64_nonnegative(n)?;
        total = total.checked_add(v).ok_or(ExcelError::Num)?;
        parts.push(v);
    }
    if parts.is_empty() {
        return Ok(1.0);
    }

    let mut acc = 1.0;
    let mut remaining = total;
    for part in parts {
        let c = combin_u64(remaining, part)?;
        acc *= c;
        if !acc.is_finite() {
            return Err(ExcelError::Num);
        }
        remaining = remaining.saturating_sub(part);
    }
    Ok(acc)
}
