use crate::value::ErrorKind;
use std::cmp::Ordering;

#[derive(Debug, Clone, Copy, PartialEq)]
pub enum RankMethod {
    Eq,
    Avg,
}

#[derive(Debug, Clone, Copy, PartialEq)]
pub enum RankOrder {
    Descending,
    Ascending,
}

fn sort_numbers(values: &mut [f64]) {
    values.sort_by(|a, b| a.total_cmp(b));
}

fn sum_kahan(values: &[f64]) -> f64 {
    let mut sum = 0.0;
    let mut c = 0.0;
    for &x in values {
        let y = x - c;
        let t = sum + y;
        c = (t - sum) - y;
        sum = t;
    }
    sum
}

fn mean(values: &[f64]) -> f64 {
    sum_kahan(values) / (values.len() as f64)
}

fn sum_squared_deviations(values: &[f64], mean: f64) -> f64 {
    let mut sum = 0.0;
    let mut c = 0.0;
    for &x in values {
        let d = x - mean;
        let term = d * d;
        let y = term - c;
        let t = sum + y;
        c = (t - sum) - y;
        sum = t;
    }
    sum
}

/// Returns (mean, sum of squared deviations) via a numerically stable two-pass algorithm.
fn variance_components(values: &[f64]) -> Result<(f64, f64), ErrorKind> {
    if values.is_empty() {
        return Err(ErrorKind::Div0);
    }
    let m = mean(values);
    let sse = sum_squared_deviations(values, m);
    if !m.is_finite() || !sse.is_finite() {
        return Err(ErrorKind::Num);
    }
    Ok((m, sse.max(0.0)))
}

pub fn var_p(values: &[f64]) -> Result<f64, ErrorKind> {
    if values.is_empty() {
        return Err(ErrorKind::Div0);
    }
    let (_mean, sse) = variance_components(values)?;
    Ok(sse / (values.len() as f64))
}

pub fn var_s(values: &[f64]) -> Result<f64, ErrorKind> {
    if values.len() < 2 {
        return Err(ErrorKind::Div0);
    }
    let (_mean, sse) = variance_components(values)?;
    Ok(sse / ((values.len() as f64) - 1.0))
}

pub fn stdev_p(values: &[f64]) -> Result<f64, ErrorKind> {
    Ok(var_p(values)?.sqrt())
}

pub fn stdev_s(values: &[f64]) -> Result<f64, ErrorKind> {
    Ok(var_s(values)?.sqrt())
}

pub fn median(values: &[f64]) -> Result<f64, ErrorKind> {
    if values.is_empty() {
        return Err(ErrorKind::Num);
    }
    let mut sorted = values.to_vec();
    sort_numbers(&mut sorted);
    let n = sorted.len();
    let mid = n / 2;
    if n % 2 == 1 {
        Ok(sorted[mid])
    } else {
        Ok((sorted[mid - 1] + sorted[mid]) / 2.0)
    }
}

pub fn mode_sngl(values: &[f64]) -> Result<f64, ErrorKind> {
    if values.is_empty() {
        return Err(ErrorKind::NA);
    }
    let mut sorted = values.to_vec();
    sort_numbers(&mut sorted);

    let mut best_count = 1usize;
    let mut best_value: Option<f64> = None;

    let mut current_value = sorted[0];
    let mut current_count = 1usize;

    for &x in sorted.iter().skip(1) {
        if x == current_value {
            current_count += 1;
            continue;
        }

        if current_count > best_count {
            best_count = current_count;
            best_value = Some(current_value);
        }

        current_value = x;
        current_count = 1;
    }

    if current_count > best_count {
        best_count = current_count;
        best_value = Some(current_value);
    }

    match (best_count, best_value) {
        (count, Some(v)) if count >= 2 => Ok(v),
        _ => Err(ErrorKind::NA),
    }
}

pub fn mode_mult(values: &[f64]) -> Result<Vec<f64>, ErrorKind> {
    if values.is_empty() {
        return Err(ErrorKind::NA);
    }
    let mut sorted = values.to_vec();
    sort_numbers(&mut sorted);

    let mut best_count = 1usize;
    let mut modes: Vec<f64> = Vec::new();

    let mut current_value = sorted[0];
    let mut current_count = 1usize;

    for &x in sorted.iter().skip(1) {
        if x == current_value {
            current_count += 1;
            continue;
        }

        match current_count.cmp(&best_count) {
            Ordering::Greater => {
                best_count = current_count;
                modes.clear();
                modes.push(current_value);
            }
            Ordering::Equal if current_count == best_count => modes.push(current_value),
            _ => {}
        }

        current_value = x;
        current_count = 1;
    }

    match current_count.cmp(&best_count) {
        Ordering::Greater => {
            best_count = current_count;
            modes.clear();
            modes.push(current_value);
        }
        Ordering::Equal if current_count == best_count => modes.push(current_value),
        _ => {}
    }

    if best_count < 2 {
        return Err(ErrorKind::NA);
    }

    Ok(modes)
}

pub fn large(values: &[f64], k: usize) -> Result<f64, ErrorKind> {
    if k == 0 || k > values.len() {
        return Err(ErrorKind::Num);
    }
    let mut sorted = values.to_vec();
    sort_numbers(&mut sorted);
    Ok(sorted[sorted.len() - k])
}

pub fn small(values: &[f64], k: usize) -> Result<f64, ErrorKind> {
    if k == 0 || k > values.len() {
        return Err(ErrorKind::Num);
    }
    let mut sorted = values.to_vec();
    sort_numbers(&mut sorted);
    Ok(sorted[k - 1])
}

pub fn percentile_inc(values: &[f64], k: f64) -> Result<f64, ErrorKind> {
    if !(0.0..=1.0).contains(&k) {
        return Err(ErrorKind::Num);
    }
    if values.is_empty() {
        return Err(ErrorKind::Num);
    }

    let mut sorted = values.to_vec();
    sort_numbers(&mut sorted);

    if sorted.len() == 1 {
        return Ok(sorted[0]);
    }

    let n_minus_1 = (sorted.len() - 1) as f64;
    let pos = k * n_minus_1;
    let lo = pos.floor() as usize;
    let hi = pos.ceil() as usize;
    let hi = hi.min(sorted.len() - 1);
    let frac = pos - (lo as f64);

    let base = sorted[lo];
    let next = sorted[hi];
    Ok(base + frac * (next - base))
}

pub fn percentile_exc(values: &[f64], k: f64) -> Result<f64, ErrorKind> {
    if !(0.0 < k && k < 1.0) {
        return Err(ErrorKind::Num);
    }
    if values.is_empty() {
        return Err(ErrorKind::Num);
    }

    let mut sorted = values.to_vec();
    sort_numbers(&mut sorted);

    let n = sorted.len() as f64;
    let pos = k * (n + 1.0);
    if pos < 1.0 || pos > n {
        return Err(ErrorKind::Num);
    }

    let idx = pos.floor() as usize; // 1-based
    let frac = pos - (idx as f64);
    if frac == 0.0 {
        return Ok(sorted[idx - 1]);
    }

    // pos is strictly within (1, n) when frac != 0.
    let lo = idx - 1;
    let hi = idx;
    let base = sorted[lo];
    let next = sorted[hi];
    Ok(base + frac * (next - base))
}

pub fn quartile_inc(values: &[f64], quart: i64) -> Result<f64, ErrorKind> {
    let k = match quart {
        0 => 0.0,
        1 => 0.25,
        2 => 0.5,
        3 => 0.75,
        4 => 1.0,
        _ => return Err(ErrorKind::Num),
    };
    percentile_inc(values, k)
}

pub fn quartile_exc(values: &[f64], quart: i64) -> Result<f64, ErrorKind> {
    let k = match quart {
        1 => 0.25,
        2 => 0.5,
        3 => 0.75,
        _ => return Err(ErrorKind::Num),
    };
    percentile_exc(values, k)
}

pub fn rank(number: f64, values: &[f64], order: RankOrder, method: RankMethod) -> Result<f64, ErrorKind> {
    if values.is_empty() {
        return Err(ErrorKind::NA);
    }

    let mut less = 0usize;
    let mut greater = 0usize;
    let mut equal = 0usize;
    for &x in values {
        match x.partial_cmp(&number).unwrap_or(Ordering::Equal) {
            Ordering::Less => less += 1,
            Ordering::Greater => greater += 1,
            Ordering::Equal => {
                if x == number {
                    equal += 1;
                }
            }
        }
    }

    let base = match order {
        RankOrder::Descending => greater as f64,
        RankOrder::Ascending => less as f64,
    };

    if matches!(method, RankMethod::Avg) && equal > 0 {
        // Average of the rank positions occupied by the duplicates.
        Ok(base + ((equal + 1) as f64) / 2.0)
    } else {
        Ok(base + 1.0)
    }
}

fn paired_means(xs: &[f64], ys: &[f64]) -> Result<(f64, f64), ErrorKind> {
    debug_assert_eq!(xs.len(), ys.len());
    if xs.is_empty() {
        return Err(ErrorKind::Div0);
    }
    let mean_x = mean(xs);
    let mean_y = mean(ys);
    if !mean_x.is_finite() || !mean_y.is_finite() {
        return Err(ErrorKind::Num);
    }
    Ok((mean_x, mean_y))
}

pub fn covariance_p(xs: &[f64], ys: &[f64]) -> Result<f64, ErrorKind> {
    if xs.len() != ys.len() {
        return Err(ErrorKind::NA);
    }
    if xs.is_empty() {
        return Err(ErrorKind::Div0);
    }
    let (mean_x, mean_y) = paired_means(xs, ys)?;

    let mut sum = 0.0;
    let mut c = 0.0;
    for (&x, &y) in xs.iter().zip(ys.iter()) {
        let term = (x - mean_x) * (y - mean_y);
        let yk = term - c;
        let t = sum + yk;
        c = (t - sum) - yk;
        sum = t;
    }

    let out = sum / (xs.len() as f64);
    if out.is_finite() {
        Ok(out)
    } else {
        Err(ErrorKind::Num)
    }
}

pub fn covariance_s(xs: &[f64], ys: &[f64]) -> Result<f64, ErrorKind> {
    if xs.len() != ys.len() {
        return Err(ErrorKind::NA);
    }
    if xs.len() < 2 {
        return Err(ErrorKind::Div0);
    }
    let (mean_x, mean_y) = paired_means(xs, ys)?;

    let mut sum = 0.0;
    let mut c = 0.0;
    for (&x, &y) in xs.iter().zip(ys.iter()) {
        let term = (x - mean_x) * (y - mean_y);
        let yk = term - c;
        let t = sum + yk;
        c = (t - sum) - yk;
        sum = t;
    }

    let out = sum / ((xs.len() as f64) - 1.0);
    if out.is_finite() {
        Ok(out)
    } else {
        Err(ErrorKind::Num)
    }
}

pub fn correl(xs: &[f64], ys: &[f64]) -> Result<f64, ErrorKind> {
    if xs.len() != ys.len() {
        return Err(ErrorKind::NA);
    }
    if xs.len() < 2 {
        return Err(ErrorKind::Div0);
    }

    let (mean_x, mean_y) = paired_means(xs, ys)?;

    let mut sxy = 0.0;
    let mut sx = 0.0;
    let mut sy = 0.0;

    // Compensated sums.
    let mut cxy = 0.0;
    let mut cx = 0.0;
    let mut cy = 0.0;

    for (&x, &y) in xs.iter().zip(ys.iter()) {
        let dx = x - mean_x;
        let dy = y - mean_y;

        let term_xy = dx * dy;
        let yk = term_xy - cxy;
        let t = sxy + yk;
        cxy = (t - sxy) - yk;
        sxy = t;

        let term_x = dx * dx;
        let yk = term_x - cx;
        let t = sx + yk;
        cx = (t - sx) - yk;
        sx = t;

        let term_y = dy * dy;
        let yk = term_y - cy;
        let t = sy + yk;
        cy = (t - sy) - yk;
        sy = t;
    }

    if sx == 0.0 || sy == 0.0 {
        return Err(ErrorKind::Div0);
    }

    let denom = (sx * sy).sqrt();
    if denom == 0.0 || !denom.is_finite() {
        return Err(ErrorKind::Div0);
    }

    let mut out = sxy / denom;
    if !out.is_finite() {
        return Err(ErrorKind::Num);
    }

    // Clamp minor floating-point overshoot.
    if out > 1.0 && out < 1.0 + 1e-12 {
        out = 1.0;
    } else if out < -1.0 && out > -1.0 - 1e-12 {
        out = -1.0;
    }
    Ok(out)
}
