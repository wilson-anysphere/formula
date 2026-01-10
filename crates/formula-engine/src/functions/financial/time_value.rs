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
    let df = |r: f64| rate_equation_derivative(r, nper, pmt, pv, fv, typ);

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

fn rate_equation_derivative(
    rate: f64,
    nper: f64,
    pmt: f64,
    pv: f64,
    _fv: f64,
    typ: f64,
) -> Option<f64> {
    if rate <= -1.0 {
        return None;
    }

    // Use the analytic derivative for Newton-Raphson. Handling `rate == 0`
    // specially avoids catastrophic cancellation in the `(g - 1) / rate` term.
    if rate == 0.0 {
        let df = nper * pv + pmt * (nper * (nper - 1.0) / 2.0 + typ * nper);
        return (df.is_finite() && df != 0.0).then_some(df);
    }

    let (g, g_minus_1) = pow1p(rate, nper)?;
    let one_plus_rate = 1.0 + rate;
    if one_plus_rate == 0.0 {
        return None;
    }

    // d/dr (1+r)^n = n*(1+r)^(n-1)
    let dg = nper * g / one_plus_rate;

    let annuity = g_minus_1 / rate;
    let rate_sq = rate * rate;
    if rate_sq == 0.0 {
        return None;
    }

    // d/dr ((g - 1) / r) = (dg*r - (g - 1)) / r^2
    let dannuity = (dg * rate - g_minus_1) / rate_sq;
    let dpmt_factor = typ * annuity + (1.0 + rate * typ) * dannuity;

    let df = pv * dg + pmt * dpmt_factor;
    (df.is_finite() && df != 0.0).then_some(df)
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
