use crate::error::{ExcelError, ExcelResult};
use crate::functions::financial::iterative::{newton_raphson, EXCEL_ITERATION_TOLERANCE};
use smallvec::SmallVec;

const MAX_ITER_IRR: usize = 20;
const MAX_ITER_XIRR: usize = 100;

pub fn npv(rate: f64, values: &[f64]) -> ExcelResult<f64> {
    if rate == -1.0 {
        return Err(ExcelError::Div0);
    }
    if rate < -1.0 {
        return Err(ExcelError::Num);
    }

    let mut sum = 0.0;
    for (i, v) in values.iter().enumerate() {
        let denom = (1.0 + rate).powi((i as i32) + 1);
        sum += *v / denom;
    }
    Ok(sum)
}

pub fn irr(values: &[f64], guess: Option<f64>) -> ExcelResult<f64> {
    if values.is_empty() {
        return Err(ExcelError::Num);
    }
    if !values.iter().any(|v| *v > 0.0) || !values.iter().any(|v| *v < 0.0) {
        return Err(ExcelError::Num);
    }

    let guess = guess.unwrap_or(0.1);
    if guess <= -1.0 {
        return Err(ExcelError::Num);
    }

    let f = |r: f64| irr_npv(values, r);
    let df = |r: f64| irr_npv_derivative(values, r);

    newton_raphson(guess, MAX_ITER_IRR, f, df).ok_or(ExcelError::Num)
}

/// Modified internal rate of return.
pub fn mirr(values: &[f64], finance_rate: f64, reinvest_rate: f64) -> ExcelResult<f64> {
    if values.len() < 2 {
        return Err(ExcelError::Div0);
    }

    if finance_rate == -1.0 || reinvest_rate == -1.0 {
        return Err(ExcelError::Div0);
    }
    if finance_rate < -1.0 || reinvest_rate < -1.0 {
        return Err(ExcelError::Num);
    }

    let has_positive = values.iter().any(|v| *v > 0.0);
    let has_negative = values.iter().any(|v| *v < 0.0);
    if !has_positive || !has_negative {
        return Err(ExcelError::Div0);
    }

    let n = values.len() as i32;
    let mut pv_neg = 0.0;
    let mut fv_pos = 0.0;

    for (t, v) in values.iter().enumerate() {
        let t = t as i32;
        if *v < 0.0 {
            pv_neg += *v / (1.0 + finance_rate).powi(t);
        } else if *v > 0.0 {
            fv_pos += *v * (1.0 + reinvest_rate).powi(n - 1 - t);
        }
    }

    if pv_neg == 0.0 || fv_pos == 0.0 {
        return Err(ExcelError::Div0);
    }

    let ratio = fv_pos / (-pv_neg);
    if ratio <= 0.0 || !ratio.is_finite() {
        return Err(ExcelError::Num);
    }

    let exponent = 1.0 / ((values.len() - 1) as f64);
    let result = ratio.powf(exponent) - 1.0;
    if result.is_finite() {
        Ok(result)
    } else {
        Err(ExcelError::Num)
    }
}

fn irr_npv(values: &[f64], rate: f64) -> Option<f64> {
    if rate <= -1.0 {
        return None;
    }
    let mut sum = 0.0;
    for (t, v) in values.iter().enumerate() {
        let denom = (1.0 + rate).powi(t as i32);
        sum += *v / denom;
    }
    if sum.is_finite() {
        Some(sum)
    } else {
        None
    }
}

fn irr_npv_derivative(values: &[f64], rate: f64) -> Option<f64> {
    if rate <= -1.0 {
        return None;
    }

    let mut sum = 0.0;
    for (t, v) in values.iter().enumerate().skip(1) {
        let power = (t as i32) + 1;
        let denom = (1.0 + rate).powi(power);
        sum += -(t as f64) * *v / denom;
    }
    if sum.is_finite() && sum != 0.0 {
        Some(sum)
    } else {
        None
    }
}

pub fn xnpv(rate: f64, values: &[f64], dates: &[f64]) -> ExcelResult<f64> {
    if values.len() != dates.len() {
        return Err(ExcelError::Num);
    }
    if values.is_empty() {
        return Err(ExcelError::Num);
    }
    if rate == -1.0 {
        return Err(ExcelError::Div0);
    }
    if rate < -1.0 {
        return Err(ExcelError::Num);
    }

    let base = dates[0];
    let mut sum = 0.0;
    for (v, d) in values.iter().zip(dates.iter()) {
        let years = (*d - base) / 365.0;
        let denom = (1.0 + rate).powf(years);
        sum += *v / denom;
    }
    Ok(sum)
}

pub fn xirr(values: &[f64], dates: &[f64], guess: Option<f64>) -> ExcelResult<f64> {
    if values.len() != dates.len() {
        return Err(ExcelError::Num);
    }
    if values.is_empty() {
        return Err(ExcelError::Num);
    }
    if !values.iter().any(|v| *v > 0.0) || !values.iter().any(|v| *v < 0.0) {
        return Err(ExcelError::Num);
    }

    let guess = guess.unwrap_or(0.1);
    if guess <= -1.0 {
        return Err(ExcelError::Num);
    }

    let base = dates[0];
    let mut exponents: SmallVec<[f64; 16]> = SmallVec::new();
    if exponents.try_reserve(dates.len()).is_err() {
        debug_assert!(false, "allocation failed (xirr exponents, len={})", dates.len());
        return Err(ExcelError::Num);
    }
    for d in dates {
        exponents.push((*d - base) / 365.0);
    }

    let f = |r: f64| xirr_npv(values, &exponents, r);
    let df = |r: f64| xirr_npv_derivative(values, &exponents, r);

    newton_raphson(guess, MAX_ITER_XIRR, f, df).ok_or(ExcelError::Num)
}

fn xirr_npv(values: &[f64], exponents: &[f64], rate: f64) -> Option<f64> {
    if rate <= -1.0 {
        return None;
    }
    let base = 1.0 + rate;
    if base <= 0.0 {
        return None;
    }

    let mut sum = 0.0;
    for (v, p) in values.iter().zip(exponents.iter()) {
        let denom = base.powf(*p);
        sum += *v / denom;
    }

    if sum.is_finite() {
        Some(sum)
    } else {
        None
    }
}

fn xirr_npv_derivative(values: &[f64], exponents: &[f64], rate: f64) -> Option<f64> {
    if rate <= -1.0 {
        return None;
    }
    let base = 1.0 + rate;
    if base <= 0.0 {
        return None;
    }

    let mut sum = 0.0;
    for (v, p) in values.iter().zip(exponents.iter()) {
        let denom = base.powf(*p + 1.0);
        sum += -p * *v / denom;
    }

    if sum.is_finite() && sum != 0.0 {
        Some(sum)
    } else {
        None
    }
}

#[allow(dead_code)]
fn _converged(previous: f64, next: f64) -> bool {
    (next - previous).abs() <= EXCEL_ITERATION_TOLERANCE
}
