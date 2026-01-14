use crate::value::ErrorKind;

#[derive(Debug, Default, Clone, Copy)]
struct KahanSum {
    sum: f64,
    c: f64,
}

impl KahanSum {
    fn add(&mut self, x: f64) {
        let y = x - self.c;
        let t = self.sum + y;
        self.c = (t - self.sum) - y;
        self.sum = t;
    }

    fn value(self) -> f64 {
        self.sum
    }
}

fn mean_and_scale(values: &[f64]) -> Result<(f64, f64), ErrorKind> {
    let mean = super::mean(values);
    if !mean.is_finite() {
        return Err(ErrorKind::Num);
    }

    let mut scale: f64 = 0.0;
    for &x in values {
        let d = x - mean;
        if !d.is_finite() {
            return Err(ErrorKind::Num);
        }
        scale = scale.max(d.abs());
    }
    if !scale.is_finite() {
        return Err(ErrorKind::Num);
    }
    Ok((mean, scale))
}

/// Returns the Kahan-summed standardized central moment sums (Σz², Σz³, Σz⁴), where
/// `z = (x - mean) / scale` and `scale = max |x - mean|`.
///
/// This scaling keeps the powers bounded (|z| ≤ 1), preventing overflow when the
/// raw values have large magnitude.
fn standardized_moment_sums(
    values: &[f64],
    mean: f64,
    scale: f64,
) -> Result<(f64, f64, f64), ErrorKind> {
    if scale == 0.0 {
        return Ok((0.0, 0.0, 0.0));
    }

    let inv_scale = 1.0 / scale;

    let mut s2 = KahanSum::default();
    let mut s3 = KahanSum::default();
    let mut s4 = KahanSum::default();

    for &x in values {
        let z = (x - mean) * inv_scale;
        if !z.is_finite() {
            return Err(ErrorKind::Num);
        }

        let z2 = z * z;
        let z3 = z2 * z;
        let z4 = z2 * z2;
        if !z2.is_finite() || !z3.is_finite() || !z4.is_finite() {
            return Err(ErrorKind::Num);
        }

        s2.add(z2);
        s3.add(z3);
        s4.add(z4);
    }

    let s2 = s2.value();
    let s3 = s3.value();
    let s4 = s4.value();

    if !s2.is_finite() || !s3.is_finite() || !s4.is_finite() {
        return Err(ErrorKind::Num);
    }

    Ok((s2.max(0.0), s3, s4.max(0.0)))
}

/// Excel-compatible sample skewness (bias-corrected).
///
/// Equivalent to `SKEW` in Excel.
pub fn skew(values: &[f64]) -> Result<f64, ErrorKind> {
    if values.len() < 3 {
        return Err(ErrorKind::Div0);
    }

    let n = values.len() as f64;

    let (mean, scale) = mean_and_scale(values)?;
    if scale == 0.0 {
        return Err(ErrorKind::Div0);
    }

    let (s2, s3, _s4) = standardized_moment_sums(values, mean, scale)?;
    if s2 == 0.0 {
        return Err(ErrorKind::Div0);
    }

    // SKEW = n * sqrt(n-1) / (n-2) * (Σz³ / (Σz²)^(3/2))
    let denom = s2 * s2.sqrt();
    if denom == 0.0 {
        return Err(ErrorKind::Div0);
    }
    if !denom.is_finite() {
        return Err(ErrorKind::Num);
    }

    let factor = n * (n - 1.0).sqrt() / (n - 2.0);
    if !factor.is_finite() {
        return Err(ErrorKind::Num);
    }

    let out = factor * s3 / denom;
    if out.is_finite() {
        Ok(out)
    } else {
        Err(ErrorKind::Num)
    }
}

/// Excel-compatible population skewness.
///
/// Equivalent to `SKEW.P` in Excel.
pub fn skew_p(values: &[f64]) -> Result<f64, ErrorKind> {
    if values.is_empty() {
        return Err(ErrorKind::Div0);
    }

    let n = values.len() as f64;

    let (mean, scale) = mean_and_scale(values)?;
    if scale == 0.0 {
        return Err(ErrorKind::Div0);
    }

    let (s2, s3, _s4) = standardized_moment_sums(values, mean, scale)?;
    if s2 == 0.0 {
        return Err(ErrorKind::Div0);
    }

    // SKEW.P = sqrt(n) * (Σz³ / (Σz²)^(3/2))
    let denom = s2 * s2.sqrt();
    if denom == 0.0 {
        return Err(ErrorKind::Div0);
    }
    if !denom.is_finite() {
        return Err(ErrorKind::Num);
    }

    let out = s3 * n.sqrt() / denom;
    if out.is_finite() {
        Ok(out)
    } else {
        Err(ErrorKind::Num)
    }
}

/// Excel-compatible excess kurtosis (bias-corrected).
///
/// Equivalent to `KURT` in Excel.
pub fn kurt(values: &[f64]) -> Result<f64, ErrorKind> {
    if values.len() < 4 {
        return Err(ErrorKind::Div0);
    }

    let n = values.len() as f64;

    let (mean, scale) = mean_and_scale(values)?;
    if scale == 0.0 {
        return Err(ErrorKind::Div0);
    }

    let (s2, _s3, s4) = standardized_moment_sums(values, mean, scale)?;
    if s2 == 0.0 {
        return Err(ErrorKind::Div0);
    }

    let n23 = (n - 2.0) * (n - 3.0);
    if n23 == 0.0 {
        return Err(ErrorKind::Div0);
    }
    if !n23.is_finite() {
        return Err(ErrorKind::Num);
    }

    // KURT = (n*(n+1)*(n-1) / ((n-2)*(n-3))) * (Σz⁴ / (Σz²)²)
    //        - 3*(n-1)² / ((n-2)*(n-3))
    let s2_sq = s2 * s2;
    if s2_sq == 0.0 {
        return Err(ErrorKind::Div0);
    }
    if !s2_sq.is_finite() {
        return Err(ErrorKind::Num);
    }

    let term1 = n * (n + 1.0) * (n - 1.0) * s4 / (n23 * s2_sq);
    if !term1.is_finite() {
        return Err(ErrorKind::Num);
    }

    let term2 = 3.0 * (n - 1.0) * (n - 1.0) / n23;
    if !term2.is_finite() {
        return Err(ErrorKind::Num);
    }

    let out = term1 - term2;
    if out.is_finite() {
        Ok(out)
    } else {
        Err(ErrorKind::Num)
    }
}
