use crate::value::ErrorKind;

#[derive(Debug, Clone)]
pub struct LinearRegressionResult {
    /// Coefficients for each predictor column, in the same order as the input X columns.
    pub slopes: Vec<f64>,
    /// Intercept term (0 when `const` is FALSE).
    pub intercept: f64,

    /// Standard error for each slope (present when `stats` is TRUE and df > 0).
    pub slope_standard_errors: Option<Vec<f64>>,
    /// Standard error for the intercept (present when `stats` is TRUE and df > 0).
    pub intercept_standard_error: Option<f64>,

    pub r_squared: f64,
    /// Standard error of the y estimate (sqrt(SSE / df)).
    pub standard_error_y: Option<f64>,
    pub f_statistic: Option<f64>,
    /// Residual degrees of freedom.
    pub df_resid: f64,
    /// Regression sum of squares.
    pub ss_regression: f64,
    /// Residual sum of squares.
    pub ss_resid: f64,
}

#[derive(Debug, Clone)]
pub struct ExponentialRegressionResult {
    /// Base coefficients (m values) for each predictor column, in the same order as the input X columns.
    pub bases: Vec<f64>,
    /// Multiplicative intercept term (b, 1 when `const` is FALSE).
    pub intercept: f64,

    pub base_standard_errors: Option<Vec<f64>>,
    pub intercept_standard_error: Option<f64>,

    pub r_squared: f64,
    pub standard_error_y: Option<f64>,
    pub f_statistic: Option<f64>,
    pub df_resid: f64,
    pub ss_regression: f64,
    pub ss_resid: f64,
}

struct LeastSquaresFit {
    beta: Vec<f64>, // length k
    r: Vec<f64>,    // k*k row-major upper triangular
    sse: f64,
}

fn checked_usize_mul(a: usize, b: usize) -> Result<usize, ErrorKind> {
    a.checked_mul(b).ok_or(ErrorKind::Num)
}

fn householder_qr_least_squares(
    mut a: Vec<f64>,
    mut b: Vec<f64>,
    n: usize,
    k: usize,
) -> Result<LeastSquaresFit, ErrorKind> {
    debug_assert_eq!(a.len(), n.saturating_mul(k));
    debug_assert_eq!(b.len(), n);

    if k == 0 {
        return Err(ErrorKind::Value);
    }
    if n < k {
        // Underdetermined.
        return Err(ErrorKind::Div0);
    }

    // Householder QR in-place (row-major).
    for j in 0..k {
        // x = a[j..n, j]
        let mut norm = 0.0f64;
        for i in j..n {
            norm = norm.hypot(a[i * k + j]);
        }
        if norm == 0.0 {
            // Column is all zeros; rank deficient.
            continue;
        }

        let x0 = a[j * k + j];
        let sign = if x0 >= 0.0 { 1.0 } else { -1.0 };
        // v = x; v0 += sign * norm
        let mut v = Vec::with_capacity(n - j);
        for i in j..n {
            v.push(a[i * k + j]);
        }
        v[0] += sign * norm;

        let mut v_norm_sq = 0.0f64;
        for &vi in &v {
            v_norm_sq += vi * vi;
        }
        if v_norm_sq == 0.0 || !v_norm_sq.is_finite() {
            return Err(ErrorKind::Num);
        }
        let beta = 2.0 / v_norm_sq;

        // Apply transform to remaining columns j..k-1
        for col in j..k {
            let mut dot = 0.0f64;
            for (idx, &vi) in v.iter().enumerate() {
                dot += vi * a[(j + idx) * k + col];
            }
            dot *= beta;
            for (idx, &vi) in v.iter().enumerate() {
                a[(j + idx) * k + col] -= dot * vi;
            }
        }

        // Apply to b.
        let mut dotb = 0.0f64;
        for (idx, &vi) in v.iter().enumerate() {
            dotb += vi * b[j + idx];
        }
        dotb *= beta;
        for (idx, &vi) in v.iter().enumerate() {
            b[j + idx] -= dotb * vi;
        }
    }

    // Extract R (upper triangular kxk)
    let mut r = vec![0.0f64; checked_usize_mul(k, k)?];
    for i in 0..k {
        for j in i..k {
            r[i * k + j] = a[i * k + j];
        }
    }

    // Solve R * beta = Q^T b (stored in b[0..k])
    let mut beta = vec![0.0f64; k];
    for i_rev in 0..k {
        let i = k - 1 - i_rev;
        let mut rhs = b[i];
        for j in (i + 1)..k {
            rhs -= r[i * k + j] * beta[j];
        }
        let diag = r[i * k + i];
        // Singular or ill-conditioned.
        if diag == 0.0 || !diag.is_finite() || diag.abs() <= 1e-12 {
            return Err(ErrorKind::Num);
        }
        beta[i] = rhs / diag;
    }

    // SSE from remaining components of Q^T b.
    let mut sse = 0.0f64;
    for i in k..n {
        let term = b[i];
        sse += term * term;
    }
    if !sse.is_finite() {
        return Err(ErrorKind::Num);
    }

    Ok(LeastSquaresFit { beta, r, sse })
}

fn invert_upper_triangular(r: &[f64], k: usize) -> Result<Vec<f64>, ErrorKind> {
    debug_assert_eq!(r.len(), k.saturating_mul(k));
    let mut inv = vec![0.0f64; checked_usize_mul(k, k)?];

    // Compute inverse of upper triangular matrix with back-substitution.
    for i_rev in 0..k {
        let i = k - 1 - i_rev;
        let diag = r[i * k + i];
        if diag == 0.0 || !diag.is_finite() {
            return Err(ErrorKind::Num);
        }
        inv[i * k + i] = 1.0 / diag;
        for j in (i + 1)..k {
            let mut sum = 0.0f64;
            for l in (i + 1)..=j {
                sum += r[i * k + l] * inv[l * k + j];
            }
            inv[i * k + j] = -sum / diag;
        }
    }
    Ok(inv)
}

fn diag_xtx_inv_from_r_inv(r_inv: &[f64], k: usize) -> Vec<f64> {
    debug_assert_eq!(r_inv.len(), k.saturating_mul(k));
    // (X^T X)^{-1} = R^{-1} * (R^{-1})^T, so diag is row-wise squared norms of R^{-1}.
    let mut diag = vec![0.0f64; k];
    for i in 0..k {
        let mut acc = 0.0f64;
        for j in i..k {
            let v = r_inv[i * k + j];
            acc += v * v;
        }
        diag[i] = acc;
    }
    diag
}

pub fn linest(
    y: &[f64],
    x: &[f64],
    n: usize,
    p: usize,
    include_intercept: bool,
    include_stats: bool,
) -> Result<LinearRegressionResult, ErrorKind> {
    if p == 0 {
        return Err(ErrorKind::Value);
    }
    if y.len() != n || x.len() != checked_usize_mul(n, p)? {
        return Err(ErrorKind::Ref);
    }
    if y.iter().any(|v| !v.is_finite()) || x.iter().any(|v| !v.is_finite()) {
        return Err(ErrorKind::Num);
    }

    let k = p + if include_intercept { 1 } else { 0 };
    if n < k {
        return Err(ErrorKind::Div0);
    }

    // Build design matrix A (n x k): [x | 1] if intercept.
    let mut a = Vec::new();
    a.try_reserve_exact(checked_usize_mul(n, k)?)
        .map_err(|_| ErrorKind::Num)?;
    for i in 0..n {
        let base = i * p;
        for j in 0..p {
            a.push(x[base + j]);
        }
        if include_intercept {
            a.push(1.0);
        }
    }

    let b = y.to_vec();
    let fit = householder_qr_least_squares(a, b, n, k)?;

    let mut slopes = Vec::with_capacity(p);
    slopes.extend_from_slice(&fit.beta[..p]);
    let intercept = if include_intercept { fit.beta[p] } else { 0.0 };

    // Total sum of squares.
    let sst = if include_intercept {
        let mean = y.iter().sum::<f64>() / (n as f64);
        let mut acc = 0.0f64;
        for &yi in y {
            let d = yi - mean;
            acc += d * d;
        }
        acc
    } else {
        y.iter().map(|v| v * v).sum()
    };

    if !sst.is_finite() {
        return Err(ErrorKind::Num);
    }

    let sse = fit.sse.max(0.0);
    let ssr = (sst - sse).max(0.0);

    let r_squared = if sst == 0.0 {
        if sse == 0.0 {
            1.0
        } else {
            0.0
        }
    } else {
        1.0 - (sse / sst)
    };
    if !r_squared.is_finite() {
        return Err(ErrorKind::Num);
    }

    let df_resid = (n as i64) - (k as i64);

    let (slope_standard_errors, intercept_standard_error, standard_error_y, f_statistic) =
        if include_stats && df_resid > 0 {
            let df_resid_f = df_resid as f64;
            let mse = sse / df_resid_f;
            if !mse.is_finite() {
                return Err(ErrorKind::Num);
            }

            let r_inv = invert_upper_triangular(&fit.r, k)?;
            let diag = diag_xtx_inv_from_r_inv(&r_inv, k);

            let mut ses = Vec::with_capacity(p);
            for i in 0..p {
                let se = (diag[i] * mse).sqrt();
                if !se.is_finite() {
                    return Err(ErrorKind::Num);
                }
                ses.push(se);
            }
            let se_intercept = if include_intercept {
                let se = (diag[p] * mse).sqrt();
                if !se.is_finite() {
                    return Err(ErrorKind::Num);
                }
                Some(se)
            } else {
                Some(0.0)
            };

            let se_y = (mse).sqrt();
            if !se_y.is_finite() {
                return Err(ErrorKind::Num);
            }

            let df_reg = if include_intercept {
                (k - 1) as f64
            } else {
                k as f64
            };
            let f = if df_reg == 0.0 || sse == 0.0 {
                None
            } else {
                let msr = ssr / df_reg;
                let f = msr / mse;
                if f.is_finite() {
                    Some(f)
                } else {
                    None
                }
            };

            (Some(ses), se_intercept, Some(se_y), f)
        } else {
            (None, None, None, None)
        };

    Ok(LinearRegressionResult {
        slopes,
        intercept,
        slope_standard_errors,
        intercept_standard_error,
        r_squared,
        standard_error_y,
        f_statistic,
        df_resid: df_resid as f64,
        ss_regression: ssr,
        ss_resid: sse,
    })
}

pub fn logest(
    y: &[f64],
    x: &[f64],
    n: usize,
    p: usize,
    include_intercept: bool,
    include_stats: bool,
) -> Result<ExponentialRegressionResult, ErrorKind> {
    if p == 0 {
        return Err(ErrorKind::Value);
    }
    if y.len() != n || x.len() != checked_usize_mul(n, p)? {
        return Err(ErrorKind::Ref);
    }
    if y.iter().any(|v| !v.is_finite()) || x.iter().any(|v| !v.is_finite()) {
        return Err(ErrorKind::Num);
    }

    // Transform y via ln(y). LOGEST/GROWTH require y > 0.
    let mut y_log = Vec::with_capacity(n);
    for &yi in y {
        if !(yi > 0.0) {
            return Err(ErrorKind::Num);
        }
        let t = yi.ln();
        if !t.is_finite() {
            return Err(ErrorKind::Num);
        }
        y_log.push(t);
    }

    let lin = linest(&y_log, x, n, p, include_intercept, include_stats)?;

    let mut bases = Vec::with_capacity(p);
    for &a in &lin.slopes {
        let m = a.exp();
        if !m.is_finite() {
            return Err(ErrorKind::Num);
        }
        bases.push(m);
    }

    let intercept = if include_intercept {
        let b = lin.intercept.exp();
        if !b.is_finite() {
            return Err(ErrorKind::Num);
        }
        b
    } else {
        1.0
    };

    let base_standard_errors = if include_stats {
        match &lin.slope_standard_errors {
            Some(ses) => {
                let mut out = Vec::with_capacity(p);
                for (m, se_a) in bases.iter().zip(ses.iter()) {
                    let se_m = m * se_a;
                    if !se_m.is_finite() {
                        return Err(ErrorKind::Num);
                    }
                    out.push(se_m);
                }
                Some(out)
            }
            None => None,
        }
    } else {
        None
    };

    let intercept_standard_error = if include_stats {
        match lin.intercept_standard_error {
            Some(se_c) => {
                let se_b = intercept * se_c;
                if !se_b.is_finite() {
                    return Err(ErrorKind::Num);
                }
                Some(se_b)
            }
            None => None,
        }
    } else {
        None
    };

    Ok(ExponentialRegressionResult {
        bases,
        intercept,
        base_standard_errors,
        intercept_standard_error,
        r_squared: lin.r_squared,
        standard_error_y: lin.standard_error_y,
        f_statistic: lin.f_statistic,
        df_resid: lin.df_resid,
        ss_regression: lin.ss_regression,
        ss_resid: lin.ss_resid,
    })
}
