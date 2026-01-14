use crate::value::ErrorKind;

use statrs::distribution::{Continuous, ContinuousCDF, Normal};

fn normal(mean: f64, standard_dev: f64) -> Result<Normal, ErrorKind> {
    if !mean.is_finite() || !standard_dev.is_finite() {
        return Err(ErrorKind::Num);
    }
    if standard_dev <= 0.0 {
        return Err(ErrorKind::Num);
    }
    Normal::new(mean, standard_dev).map_err(|_| ErrorKind::Num)
}

/// Excel-compatible `NORM.DIST`.
pub fn norm_dist(x: f64, mean: f64, standard_dev: f64, cumulative: bool) -> Result<f64, ErrorKind> {
    if !x.is_finite() {
        return Err(ErrorKind::Num);
    }
    let normal = normal(mean, standard_dev)?;
    let out = if cumulative {
        normal.cdf(x)
    } else {
        normal.pdf(x)
    };
    if out.is_finite() {
        Ok(out)
    } else {
        Err(ErrorKind::Num)
    }
}

/// Excel-compatible `NORM.S.DIST`.
pub fn norm_s_dist(z: f64, cumulative: bool) -> Result<f64, ErrorKind> {
    norm_dist(z, 0.0, 1.0, cumulative)
}

/// Excel-compatible `NORM.INV`.
pub fn norm_inv(probability: f64, mean: f64, standard_dev: f64) -> Result<f64, ErrorKind> {
    if !probability.is_finite() {
        return Err(ErrorKind::Num);
    }
    if probability <= 0.0 || probability >= 1.0 {
        return Err(ErrorKind::Num);
    }

    let normal = normal(mean, standard_dev)?;
    let out = normal.inverse_cdf(probability);
    if out.is_finite() {
        Ok(out)
    } else {
        Err(ErrorKind::Num)
    }
}

/// Excel-compatible `NORM.S.INV`.
pub fn norm_s_inv(probability: f64) -> Result<f64, ErrorKind> {
    norm_inv(probability, 0.0, 1.0)
}

/// Excel-compatible `PHI`.
///
/// Standard normal PDF at `x` (equivalent to `NORM.S.DIST(x, FALSE)`).
pub fn phi(x: f64) -> Result<f64, ErrorKind> {
    norm_s_dist(x, false)
}

/// Excel-compatible `GAUSS`.
///
/// Returns `NORM.S.DIST(z, TRUE) - 0.5`.
pub fn gauss(z: f64) -> Result<f64, ErrorKind> {
    let out = norm_s_dist(z, true)? - 0.5;
    if out.is_finite() {
        Ok(out)
    } else {
        Err(ErrorKind::Num)
    }
}
