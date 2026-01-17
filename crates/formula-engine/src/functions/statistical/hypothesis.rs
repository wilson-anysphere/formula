use crate::value::ErrorKind;

use statrs::distribution::{ChiSquared, ContinuousCDF, FisherSnedecor, Normal, StudentsT};

pub fn z_test(values: &[f64], x: f64, sigma: Option<f64>) -> Result<f64, ErrorKind> {
    if values.is_empty() {
        return Err(ErrorKind::Div0);
    }
    if !x.is_finite() {
        return Err(ErrorKind::Num);
    }

    let n = values.len() as f64;
    let mean = super::mean(values);
    if !mean.is_finite() {
        return Err(ErrorKind::Num);
    }

    let sigma = match sigma {
        Some(s) => {
            if !s.is_finite() {
                return Err(ErrorKind::Num);
            }
            if s < 0.0 {
                return Err(ErrorKind::Num);
            }
            if s == 0.0 {
                return Err(ErrorKind::Div0);
            }
            s
        }
        None => super::stdev_s(values)?,
    };
    if sigma == 0.0 {
        return Err(ErrorKind::Div0);
    }

    let denom = sigma / n.sqrt();
    if denom == 0.0 || !denom.is_finite() {
        return Err(ErrorKind::Div0);
    }

    let z = (mean - x) / denom;
    if !z.is_finite() {
        return Err(ErrorKind::Num);
    }

    let normal = Normal::new(0.0, 1.0).map_err(|_| ErrorKind::Num)?;
    let mut p = 1.0 - normal.cdf(z);
    if !p.is_finite() {
        return Err(ErrorKind::Num);
    }

    // Clamp minor floating error.
    if p < 0.0 && p > -1e-12 {
        p = 0.0;
    } else if p > 1.0 && p < 1.0 + 1e-12 {
        p = 1.0;
    }
    Ok(p)
}

pub fn t_test(xs: &[f64], ys: &[f64], tails: i64, test_type: i64) -> Result<f64, ErrorKind> {
    if tails != 1 && tails != 2 {
        return Err(ErrorKind::Num);
    }
    if test_type != 1 && test_type != 2 && test_type != 3 {
        return Err(ErrorKind::Num);
    }

    let (t, df) = match test_type {
        1 => {
            if xs.len() != ys.len() {
                return Err(ErrorKind::NA);
            }
            if xs.len() < 2 {
                return Err(ErrorKind::Div0);
            }
            // Compute diffs statistics without allocating an intermediate `Vec`.
            let n = xs.len() as f64;
            let mut sum = 0.0;
            let mut c = 0.0;
            for (&a, &b) in xs.iter().zip(ys.iter()) {
                let d = a - b;
                let y = d - c;
                let t = sum + y;
                c = (t - sum) - y;
                sum = t;
            }

            let mean_d = sum / n;
            if !mean_d.is_finite() {
                return Err(ErrorKind::Num);
            }

            let mut sse = 0.0;
            let mut c = 0.0;
            for (&a, &b) in xs.iter().zip(ys.iter()) {
                let d = a - b;
                let dev = d - mean_d;
                let term = dev * dev;
                let y = term - c;
                let t = sse + y;
                c = (t - sse) - y;
                sse = t;
            }
            if !sse.is_finite() {
                return Err(ErrorKind::Num);
            }
            let sse = sse.max(0.0);
            let sd_d = (sse / (n - 1.0)).sqrt();
            if sd_d == 0.0 {
                return Err(ErrorKind::Div0);
            }

            let se = sd_d / n.sqrt();
            if se == 0.0 || !se.is_finite() {
                return Err(ErrorKind::Div0);
            }
            let t = mean_d / se;
            (t, n - 1.0)
        }
        2 => {
            if xs.len() < 2 || ys.len() < 2 {
                return Err(ErrorKind::Div0);
            }
            let mean_x = super::mean(xs);
            let mean_y = super::mean(ys);
            if !mean_x.is_finite() || !mean_y.is_finite() {
                return Err(ErrorKind::Num);
            }

            let var_x = super::var_s(xs)?;
            let var_y = super::var_s(ys)?;
            if var_x == 0.0 || var_y == 0.0 {
                return Err(ErrorKind::Div0);
            }

            let n1 = xs.len() as f64;
            let n2 = ys.len() as f64;
            let df = (n1 + n2) - 2.0;
            if df <= 0.0 {
                return Err(ErrorKind::Div0);
            }
            let pooled = (((n1 - 1.0) * var_x) + ((n2 - 1.0) * var_y)) / df;
            if pooled == 0.0 || !pooled.is_finite() {
                return Err(ErrorKind::Div0);
            }

            let se = (pooled * (1.0 / n1 + 1.0 / n2)).sqrt();
            if se == 0.0 || !se.is_finite() {
                return Err(ErrorKind::Div0);
            }
            let t = (mean_x - mean_y) / se;
            (t, df)
        }
        3 => {
            if xs.len() < 2 || ys.len() < 2 {
                return Err(ErrorKind::Div0);
            }
            let mean_x = super::mean(xs);
            let mean_y = super::mean(ys);
            if !mean_x.is_finite() || !mean_y.is_finite() {
                return Err(ErrorKind::Num);
            }

            let var_x = super::var_s(xs)?;
            let var_y = super::var_s(ys)?;
            if var_x == 0.0 || var_y == 0.0 {
                return Err(ErrorKind::Div0);
            }

            let n1 = xs.len() as f64;
            let n2 = ys.len() as f64;
            let vx = var_x / n1;
            let vy = var_y / n2;
            let se2 = vx + vy;
            if se2 == 0.0 || !se2.is_finite() {
                return Err(ErrorKind::Div0);
            }
            let se = se2.sqrt();
            if se == 0.0 || !se.is_finite() {
                return Err(ErrorKind::Div0);
            }

            let t = (mean_x - mean_y) / se;

            let num = se2 * se2;
            let denom = (vx * vx) / (n1 - 1.0) + (vy * vy) / (n2 - 1.0);
            if denom == 0.0 || !denom.is_finite() {
                return Err(ErrorKind::Div0);
            }
            let df = num / denom;
            if df <= 0.0 || !df.is_finite() {
                return Err(ErrorKind::Num);
            }

            (t, df)
        }
        _ => {
            debug_assert!(false, "T.TEST type should have been validated: {test_type}");
            return Err(ErrorKind::Num);
        }
    };

    if !t.is_finite() || !df.is_finite() || df <= 0.0 {
        return Err(ErrorKind::Num);
    }

    let dist = StudentsT::new(0.0, 1.0, df).map_err(|_| ErrorKind::Num)?;
    let t_abs = t.abs();
    let mut p_one = 1.0 - dist.cdf(t_abs);
    if !p_one.is_finite() {
        return Err(ErrorKind::Num);
    }
    if p_one < 0.0 && p_one > -1e-12 {
        p_one = 0.0;
    } else if p_one > 1.0 && p_one < 1.0 + 1e-12 {
        p_one = 1.0;
    }

    let mut p = if tails == 1 { p_one } else { 2.0 * p_one };
    if p > 1.0 && p < 1.0 + 1e-12 {
        p = 1.0;
    }
    Ok(p)
}

pub fn f_test(xs: &[f64], ys: &[f64]) -> Result<f64, ErrorKind> {
    if xs.len() < 2 || ys.len() < 2 {
        return Err(ErrorKind::Div0);
    }

    let var_x = super::var_s(xs)?;
    let var_y = super::var_s(ys)?;
    if var_x == 0.0 || var_y == 0.0 {
        return Err(ErrorKind::Div0);
    }

    let f = var_x / var_y;
    if !f.is_finite() || f <= 0.0 {
        return Err(ErrorKind::Num);
    }

    let df1 = (xs.len() - 1) as f64;
    let df2 = (ys.len() - 1) as f64;
    let dist = FisherSnedecor::new(df1, df2).map_err(|_| ErrorKind::Num)?;

    let cdf = dist.cdf(f);
    if !cdf.is_finite() {
        return Err(ErrorKind::Num);
    }

    let mut p = 2.0 * cdf.min(1.0 - cdf);
    if !p.is_finite() {
        return Err(ErrorKind::Num);
    }
    if p < 0.0 && p > -1e-12 {
        p = 0.0;
    } else if p > 1.0 && p < 1.0 + 1e-12 {
        p = 1.0;
    }
    Ok(p)
}

pub fn chisq_test(
    actual: &[f64],
    expected: &[f64],
    rows: usize,
    cols: usize,
) -> Result<f64, ErrorKind> {
    if actual.len() != expected.len() {
        return Err(ErrorKind::NA);
    }
    if actual.is_empty() {
        return Err(ErrorKind::Num);
    }
    if rows == 0 || cols == 0 {
        return Err(ErrorKind::NA);
    }

    let df = (rows.saturating_sub(1) as f64) * (cols.saturating_sub(1) as f64);
    if df <= 0.0 {
        return Err(ErrorKind::Num);
    }

    let mut sum = 0.0;
    let mut c = 0.0;
    for (&a, &e) in actual.iter().zip(expected.iter()) {
        if !a.is_finite() || !e.is_finite() {
            return Err(ErrorKind::Num);
        }
        if e <= 0.0 {
            return Err(ErrorKind::Num);
        }
        let d = a - e;
        let term = (d * d) / e;
        let y = term - c;
        let t = sum + y;
        c = (t - sum) - y;
        sum = t;
    }

    if !sum.is_finite() {
        return Err(ErrorKind::Num);
    }

    let dist = ChiSquared::new(df).map_err(|_| ErrorKind::Num)?;
    let mut p = 1.0 - dist.cdf(sum);
    if !p.is_finite() {
        return Err(ErrorKind::Num);
    }
    if p < 0.0 && p > -1e-12 {
        p = 0.0;
    } else if p > 1.0 && p < 1.0 + 1e-12 {
        p = 1.0;
    }
    Ok(p)
}
