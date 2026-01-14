use crate::error::{ExcelError, ExcelResult};

/// PDURATION(rate, pv, fv)
///
/// Number of periods required for an investment to reach `fv` from `pv` at constant `rate`.
///
/// Excel's formula is:
///   ln(fv/pv) / ln(1+rate)
pub fn pduration(rate: f64, pv: f64, fv: f64) -> ExcelResult<f64> {
    if !rate.is_finite() || !pv.is_finite() || !fv.is_finite() {
        return Err(ExcelError::Num);
    }
    if rate <= 0.0 {
        return Err(ExcelError::Num);
    }
    if pv <= 0.0 || fv <= 0.0 {
        return Err(ExcelError::Num);
    }

    let one_plus_rate = 1.0 + rate;
    if !one_plus_rate.is_finite() {
        return Err(ExcelError::Num);
    }
    let denom = one_plus_rate.ln();
    if !denom.is_finite() {
        return Err(ExcelError::Num);
    }
    if denom == 0.0 {
        return Err(ExcelError::Div0);
    }

    let ratio = fv / pv;
    if !(ratio > 0.0) || !ratio.is_finite() {
        return Err(ExcelError::Num);
    }
    let numer = ratio.ln();
    if !numer.is_finite() {
        return Err(ExcelError::Num);
    }

    let out = numer / denom;
    if out.is_finite() {
        Ok(out)
    } else {
        Err(ExcelError::Num)
    }
}
