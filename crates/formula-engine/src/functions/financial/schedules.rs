use crate::error::{ExcelError, ExcelResult};

/// Excel-compatible `FVSCHEDULE(principal, schedule)`.
///
/// Computes `principal * Π(1 + r_i)` for each rate in `schedule`.
pub fn fvschedule(principal: f64, schedule: &[f64]) -> ExcelResult<f64> {
    if !principal.is_finite() {
        return Err(ExcelError::Num);
    }
    if schedule.is_empty() {
        return Err(ExcelError::Value);
    }

    let mut acc = principal;
    for r in schedule {
        if !r.is_finite() {
            return Err(ExcelError::Num);
        }
        acc *= 1.0 + r;
        if !acc.is_finite() {
            return Err(ExcelError::Num);
        }
    }
    Ok(acc)
}
