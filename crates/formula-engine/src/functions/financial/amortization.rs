use crate::error::{ExcelError, ExcelResult};
use crate::functions::financial::iterative::EXCEL_ITERATION_TOLERANCE;

use super::time_value::{ipmt, ppmt};

fn validate_inputs(
    rate: f64,
    nper: f64,
    pv: f64,
    start: f64,
    end: f64,
    typ: f64,
) -> ExcelResult<()> {
    if !rate.is_finite()
        || !nper.is_finite()
        || !pv.is_finite()
        || !start.is_finite()
        || !end.is_finite()
        || !typ.is_finite()
    {
        return Err(ExcelError::Num);
    }

    if rate <= 0.0 || nper <= 0.0 || pv <= 0.0 {
        return Err(ExcelError::Num);
    }

    if start < 1.0 || end < start || end > nper {
        return Err(ExcelError::Num);
    }

    // Excel requires `type` to be 0 (end of period) or 1 (beginning of period).
    if typ != 0.0 && typ != 1.0 {
        return Err(ExcelError::Num);
    }

    // Period bounds are specified as numbers, but Excel treats them as integral periods.
    // Require inputs to be whole numbers (within tolerance) to avoid silently selecting
    // unexpected periods.
    let start_int = start.round();
    let end_int = end.round();
    if (start - start_int).abs() > EXCEL_ITERATION_TOLERANCE
        || (end - end_int).abs() > EXCEL_ITERATION_TOLERANCE
    {
        return Err(ExcelError::Num);
    }

    Ok(())
}

fn kahan_sum<I>(iter: I) -> ExcelResult<f64>
where
    I: IntoIterator<Item = ExcelResult<f64>>,
{
    let mut sum = 0.0;
    let mut c = 0.0;
    for term in iter {
        let x = term?;
        let y = x - c;
        let t = sum + y;
        c = (t - sum) - y;
        sum = t;
    }
    if sum.is_finite() {
        Ok(sum)
    } else {
        Err(ExcelError::Num)
    }
}

pub fn cumipmt(rate: f64, nper: f64, pv: f64, start: f64, end: f64, typ: f64) -> ExcelResult<f64> {
    validate_inputs(rate, nper, pv, start, end, typ)?;

    let start_period = start.round() as i64;
    let end_period = end.round() as i64;

    kahan_sum(
        (start_period..=end_period)
            .map(|per| ipmt(rate, per as f64, nper, pv, Some(0.0), Some(typ))),
    )
}

pub fn cumprinc(rate: f64, nper: f64, pv: f64, start: f64, end: f64, typ: f64) -> ExcelResult<f64> {
    validate_inputs(rate, nper, pv, start, end, typ)?;

    let start_period = start.round() as i64;
    let end_period = end.round() as i64;

    kahan_sum(
        (start_period..=end_period)
            .map(|per| ppmt(rate, per as f64, nper, pv, Some(0.0), Some(typ))),
    )
}
