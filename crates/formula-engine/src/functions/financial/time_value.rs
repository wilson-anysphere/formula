use crate::error::{ExcelError, ExcelResult};
use crate::functions::financial::iterative::{newton_raphson, EXCEL_ITERATION_TOLERANCE};

const MAX_ITER_RATE: usize = 20;

fn normalize_type(typ: Option<f64>) -> f64 {
    match typ {
        Some(t) if t != 0.0 => 1.0,
        _ => 0.0,
    }
}

fn pow1p(rate: f64, nper: f64) -> Option<(f64, f64)> {
    // Returns ( (1+rate)^nper , (1+rate)^nper - 1 ) computed with numerically stable
    // primitives where possible.
    let ln1p = rate.ln_1p();
    if !ln1p.is_finite() {
        return None;
    }
    let exponent = nper * ln1p;
    let g_minus_1 = exponent.exp_m1();
    let g = g_minus_1 + 1.0;
    if !g.is_finite() || !g_minus_1.is_finite() {
        return None;
    }
    Some((g, g_minus_1))
}

/// Present value.
pub fn pv(rate: f64, nper: f64, pmt: f64, fv: Option<f64>, typ: Option<f64>) -> ExcelResult<f64> {
    let fv = fv.unwrap_or(0.0);
    let typ = normalize_type(typ);

    if rate == 0.0 {
        return Ok(-fv - pmt * nper);
    }

    if rate == -1.0 && nper != 0.0 {
        return Err(ExcelError::Div0);
    }

    let (g, g_minus_1) = pow1p(rate, nper).ok_or(ExcelError::Num)?;
    if g == 0.0 {
        return Err(ExcelError::Div0);
    }

    let annuity = g_minus_1 / rate;
    let pmt_factor = (1.0 + rate * typ) * annuity;

    Ok(-(fv + pmt * pmt_factor) / g)
}

/// Future value.
pub fn fv(rate: f64, nper: f64, pmt: f64, pv: Option<f64>, typ: Option<f64>) -> ExcelResult<f64> {
    let pv = pv.unwrap_or(0.0);
    let typ = normalize_type(typ);

    if rate == 0.0 {
        return Ok(-(pv + pmt * nper));
    }

    let (g, g_minus_1) = pow1p(rate, nper).ok_or(ExcelError::Num)?;
    let annuity = g_minus_1 / rate;
    let pmt_factor = (1.0 + rate * typ) * annuity;

    Ok(-(pv * g + pmt * pmt_factor))
}

/// Periodic payment.
pub fn pmt(rate: f64, nper: f64, pv: f64, fv: Option<f64>, typ: Option<f64>) -> ExcelResult<f64> {
    let fv = fv.unwrap_or(0.0);
    let typ = normalize_type(typ);

    if nper == 0.0 {
        return Err(ExcelError::Div0);
    }

    if rate == 0.0 {
        return Ok(-(pv + fv) / nper);
    }

    let (g, g_minus_1) = pow1p(rate, nper).ok_or(ExcelError::Num)?;
    let annuity = g_minus_1 / rate;
    let pmt_factor = (1.0 + rate * typ) * annuity;

    if pmt_factor == 0.0 {
        return Err(ExcelError::Div0);
    }

    Ok(-(pv * g + fv) / pmt_factor)
}

/// Number of payment periods.
pub fn nper(rate: f64, pmt: f64, pv: f64, fv: Option<f64>, typ: Option<f64>) -> ExcelResult<f64> {
    let fv = fv.unwrap_or(0.0);
    let typ = normalize_type(typ);

    if rate == 0.0 {
        if pmt == 0.0 {
            return if pv + fv == 0.0 {
                Ok(0.0)
            } else {
                Err(ExcelError::Num)
            };
        }
        return Ok(-(pv + fv) / pmt);
    }

    let ln1p = rate.ln_1p();
    if !ln1p.is_finite() || ln1p == 0.0 {
        return Err(ExcelError::Num);
    }

    if pmt == 0.0 {
        if pv == 0.0 {
            return Err(ExcelError::Num);
        }
        let g = -fv / pv;
        if g <= 0.0 {
            return Err(ExcelError::Num);
        }
        return Ok(g.ln() / ln1p);
    }

    let a = pmt * (1.0 + rate * typ) / rate;
    if pv + a == 0.0 {
        return Err(ExcelError::Num);
    }

    let g = (a - fv) / (pv + a);
    if g <= 0.0 {
        return Err(ExcelError::Num);
    }

    Ok(g.ln() / ln1p)
}

/// Interest rate per period.
pub fn rate(
    nper: f64,
    pmt: f64,
    pv: f64,
    fv: Option<f64>,
    typ: Option<f64>,
    guess: Option<f64>,
) -> ExcelResult<f64> {
    let fv = fv.unwrap_or(0.0);
    let typ = normalize_type(typ);
    let guess = guess.unwrap_or(0.1);

    if nper <= 0.0 {
        return Err(ExcelError::Num);
    }
    if guess <= -1.0 {
        return Err(ExcelError::Num);
    }

    let f = |r: f64| rate_equation(r, nper, pmt, pv, fv, typ);
    let df = |r: f64| numeric_derivative(r, &f);

    newton_raphson(guess, MAX_ITER_RATE, f, df).ok_or(ExcelError::Num)
}

fn rate_equation(rate: f64, nper: f64, pmt: f64, pv: f64, fv: f64, typ: f64) -> Option<f64> {
    if rate <= -1.0 {
        return None;
    }

    if rate == 0.0 {
        return Some(pv + pmt * nper + fv);
    }

    let (g, g_minus_1) = pow1p(rate, nper)?;
    let annuity = g_minus_1 / rate;
    let pmt_factor = (1.0 + rate * typ) * annuity;

    Some(pv * g + pmt * pmt_factor + fv)
}

fn numeric_derivative<F>(x: f64, f: &F) -> Option<f64>
where
    F: Fn(f64) -> Option<f64>,
{
    if x <= -1.0 {
        return None;
    }

    let mut h = (x.abs() * 1.0e-8).max(1.0e-6);

    // Keep x ± h inside the domain (> -1).
    if x - h <= -1.0 {
        h = (x + 1.0) * 0.5;
    }
    if h <= 0.0 {
        return None;
    }

    let f1 = f(x + h)?;
    let f0 = f(x - h)?;
    let df = (f1 - f0) / (2.0 * h);
    if !df.is_finite() || df == 0.0 {
        return None;
    }
    Some(df)
}

/// Interest payment for a given period.
pub fn ipmt(
    rate: f64,
    per: f64,
    nper: f64,
    pv: f64,
    fv_opt: Option<f64>,
    typ: Option<f64>,
) -> ExcelResult<f64> {
    let fv_value = fv_opt.unwrap_or(0.0);
    let typ = normalize_type(typ);

    if per < 1.0 || per > nper {
        return Err(ExcelError::Num);
    }

    if rate == 0.0 {
        return Ok(0.0);
    }

    let payment = pmt(rate, nper, pv, Some(fv_value), Some(typ))?;

    if typ == 1.0 && (per - 1.0).abs() <= EXCEL_ITERATION_TOLERANCE {
        return Ok(0.0);
    }

    if typ == 1.0 {
        let denom = 1.0 + rate;
        if denom == 0.0 {
            return Err(ExcelError::Div0);
        }
        let future_value = fv(rate, per - 1.0, payment, Some(pv), Some(1.0))?;
        Ok(future_value * rate / denom)
    } else {
        let future_value = fv(rate, per - 1.0, payment, Some(pv), Some(0.0))?;
        Ok(future_value * rate)
    }
}

/// Principal payment for a given period.
pub fn ppmt(
    rate: f64,
    per: f64,
    nper: f64,
    pv: f64,
    fv_opt: Option<f64>,
    typ: Option<f64>,
) -> ExcelResult<f64> {
    let fv_value = fv_opt.unwrap_or(0.0);
    let typ = normalize_type(typ);

    let payment = pmt(rate, nper, pv, Some(fv_value), Some(typ))?;
    let interest_payment = ipmt(rate, per, nper, pv, Some(fv_value), Some(typ))?;
    Ok(payment - interest_payment)
}
