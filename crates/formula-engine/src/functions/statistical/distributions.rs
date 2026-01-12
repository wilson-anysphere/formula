use crate::value::ErrorKind;

use statrs::distribution::{
    Beta, ChiSquared, Continuous, ContinuousCDF, Exp, FisherSnedecor, Gamma, LogNormal, Normal,
    StudentsT, Weibull,
};

fn ensure_finite(x: f64) -> Result<f64, ErrorKind> {
    if x.is_finite() {
        Ok(x)
    } else {
        Err(ErrorKind::Num)
    }
}

fn ensure_positive(x: f64) -> Result<f64, ErrorKind> {
    let x = ensure_finite(x)?;
    if x > 0.0 {
        Ok(x)
    } else {
        Err(ErrorKind::Num)
    }
}

fn ensure_nonnegative(x: f64) -> Result<f64, ErrorKind> {
    let x = ensure_finite(x)?;
    if x >= 0.0 {
        Ok(x)
    } else {
        Err(ErrorKind::Num)
    }
}

fn ensure_probability(p: f64) -> Result<f64, ErrorKind> {
    let p = ensure_finite(p)?;
    if (0.0..=1.0).contains(&p) {
        Ok(p)
    } else {
        Err(ErrorKind::Num)
    }
}

fn inverse_cdf_nonnegative_bisect(
    target: f64,
    mut cdf: impl FnMut(f64) -> f64,
) -> Result<f64, ErrorKind> {
    if target == 0.0 {
        return Ok(0.0);
    }
    if target == 1.0 {
        // Quantile is +inf for continuous distributions with support on [0, +inf).
        return Err(ErrorKind::Num);
    }

    // Bracket the target probability.
    let mut low = 0.0;
    let mut high = 1.0;
    loop {
        let c = cdf(high);
        if !c.is_finite() {
            return Err(ErrorKind::Num);
        }
        if c >= target {
            break;
        }
        high *= 2.0;
        if !high.is_finite() {
            return Err(ErrorKind::Num);
        }
    }

    // Bisection to convergence.
    for _ in 0..128 {
        let mid = low + (high - low) / 2.0;
        let c = cdf(mid);
        if !c.is_finite() {
            return Err(ErrorKind::Num);
        }
        if c < target {
            low = mid;
        } else {
            high = mid;
        }
    }
    let out = low + (high - low) / 2.0;
    if out.is_finite() {
        Ok(out)
    } else {
        Err(ErrorKind::Num)
    }
}

fn inverse_cdf_unit_interval_bisect(
    target: f64,
    mut cdf: impl FnMut(f64) -> f64,
) -> Result<f64, ErrorKind> {
    if target == 0.0 {
        return Ok(0.0);
    }
    if target == 1.0 {
        return Ok(1.0);
    }

    let mut low = 0.0;
    let mut high = 1.0;
    for _ in 0..128 {
        let mid = low + (high - low) / 2.0;
        let c = cdf(mid);
        if !c.is_finite() {
            return Err(ErrorKind::Num);
        }
        if c < target {
            low = mid;
        } else {
            high = mid;
        }
    }
    let out = low + (high - low) / 2.0;
    if out.is_finite() {
        Ok(out)
    } else {
        Err(ErrorKind::Num)
    }
}

pub fn t_dist(x: f64, deg_freedom: f64, cumulative: bool) -> Result<f64, ErrorKind> {
    let x = ensure_finite(x)?;
    let deg_freedom = ensure_positive(deg_freedom)?;
    let dist = StudentsT::new(0.0, 1.0, deg_freedom).map_err(|_| ErrorKind::Num)?;
    let out = if cumulative { dist.cdf(x) } else { dist.pdf(x) };
    if out.is_finite() {
        Ok(out)
    } else {
        Err(ErrorKind::Num)
    }
}

pub fn t_dist_rt(x: f64, deg_freedom: f64) -> Result<f64, ErrorKind> {
    let x = ensure_finite(x)?;
    if x <= 0.0 {
        return Err(ErrorKind::Num);
    }
    let deg_freedom = ensure_positive(deg_freedom)?;
    let dist = StudentsT::new(0.0, 1.0, deg_freedom).map_err(|_| ErrorKind::Num)?;
    let out = dist.sf(x);
    if out.is_finite() {
        Ok(out)
    } else {
        Err(ErrorKind::Num)
    }
}

pub fn t_dist_2t(x: f64, deg_freedom: f64) -> Result<f64, ErrorKind> {
    let x = ensure_finite(x)?;
    if x <= 0.0 {
        return Err(ErrorKind::Num);
    }
    let deg_freedom = ensure_positive(deg_freedom)?;
    let dist = StudentsT::new(0.0, 1.0, deg_freedom).map_err(|_| ErrorKind::Num)?;
    let out = 2.0 * dist.sf(x);
    if out.is_finite() {
        Ok(out)
    } else {
        Err(ErrorKind::Num)
    }
}

pub fn t_inv(probability: f64, deg_freedom: f64) -> Result<f64, ErrorKind> {
    let probability = ensure_probability(probability)?;
    let deg_freedom = ensure_positive(deg_freedom)?;
    let dist = StudentsT::new(0.0, 1.0, deg_freedom).map_err(|_| ErrorKind::Num)?;
    let out = dist.inverse_cdf(probability);
    if out.is_finite() {
        Ok(out)
    } else {
        Err(ErrorKind::Num)
    }
}

pub fn t_inv_2t(probability: f64, deg_freedom: f64) -> Result<f64, ErrorKind> {
    let probability = ensure_probability(probability)?;
    let deg_freedom = ensure_positive(deg_freedom)?;
    let p = 1.0 - probability / 2.0;
    let dist = StudentsT::new(0.0, 1.0, deg_freedom).map_err(|_| ErrorKind::Num)?;
    let out = dist.inverse_cdf(p);
    if out.is_finite() {
        Ok(out)
    } else {
        Err(ErrorKind::Num)
    }
}

pub fn chisq_dist(x: f64, deg_freedom: f64, cumulative: bool) -> Result<f64, ErrorKind> {
    let x = ensure_nonnegative(x)?;
    let deg_freedom = ensure_positive(deg_freedom)?;
    let dist = ChiSquared::new(deg_freedom).map_err(|_| ErrorKind::Num)?;
    let out = if cumulative { dist.cdf(x) } else { dist.pdf(x) };
    if out.is_finite() {
        Ok(out)
    } else {
        Err(ErrorKind::Num)
    }
}

pub fn chisq_dist_rt(x: f64, deg_freedom: f64) -> Result<f64, ErrorKind> {
    let x = ensure_nonnegative(x)?;
    let deg_freedom = ensure_positive(deg_freedom)?;
    let dist = ChiSquared::new(deg_freedom).map_err(|_| ErrorKind::Num)?;
    let out = dist.sf(x);
    if out.is_finite() {
        Ok(out)
    } else {
        Err(ErrorKind::Num)
    }
}

pub fn chisq_inv(probability: f64, deg_freedom: f64) -> Result<f64, ErrorKind> {
    let probability = ensure_probability(probability)?;
    let deg_freedom = ensure_positive(deg_freedom)?;
    let dist = ChiSquared::new(deg_freedom).map_err(|_| ErrorKind::Num)?;
    inverse_cdf_nonnegative_bisect(probability, |x| dist.cdf(x))
}

pub fn chisq_inv_rt(probability: f64, deg_freedom: f64) -> Result<f64, ErrorKind> {
    let probability = ensure_probability(probability)?;
    let deg_freedom = ensure_positive(deg_freedom)?;
    let dist = ChiSquared::new(deg_freedom).map_err(|_| ErrorKind::Num)?;
    inverse_cdf_nonnegative_bisect(1.0 - probability, |x| dist.cdf(x))
}

pub fn f_dist(
    x: f64,
    deg_freedom1: f64,
    deg_freedom2: f64,
    cumulative: bool,
) -> Result<f64, ErrorKind> {
    let x = ensure_nonnegative(x)?;
    let deg_freedom1 = ensure_positive(deg_freedom1)?;
    let deg_freedom2 = ensure_positive(deg_freedom2)?;
    let dist = FisherSnedecor::new(deg_freedom1, deg_freedom2).map_err(|_| ErrorKind::Num)?;
    let out = if cumulative { dist.cdf(x) } else { dist.pdf(x) };
    if out.is_finite() {
        Ok(out)
    } else {
        Err(ErrorKind::Num)
    }
}

pub fn f_dist_rt(x: f64, deg_freedom1: f64, deg_freedom2: f64) -> Result<f64, ErrorKind> {
    let x = ensure_nonnegative(x)?;
    let deg_freedom1 = ensure_positive(deg_freedom1)?;
    let deg_freedom2 = ensure_positive(deg_freedom2)?;
    let dist = FisherSnedecor::new(deg_freedom1, deg_freedom2).map_err(|_| ErrorKind::Num)?;
    let out = dist.sf(x);
    if out.is_finite() {
        Ok(out)
    } else {
        Err(ErrorKind::Num)
    }
}

pub fn f_inv(probability: f64, deg_freedom1: f64, deg_freedom2: f64) -> Result<f64, ErrorKind> {
    let probability = ensure_probability(probability)?;
    let deg_freedom1 = ensure_positive(deg_freedom1)?;
    let deg_freedom2 = ensure_positive(deg_freedom2)?;
    let dist = FisherSnedecor::new(deg_freedom1, deg_freedom2).map_err(|_| ErrorKind::Num)?;
    inverse_cdf_nonnegative_bisect(probability, |x| dist.cdf(x))
}

pub fn f_inv_rt(probability: f64, deg_freedom1: f64, deg_freedom2: f64) -> Result<f64, ErrorKind> {
    let probability = ensure_probability(probability)?;
    let deg_freedom1 = ensure_positive(deg_freedom1)?;
    let deg_freedom2 = ensure_positive(deg_freedom2)?;
    let dist = FisherSnedecor::new(deg_freedom1, deg_freedom2).map_err(|_| ErrorKind::Num)?;
    inverse_cdf_nonnegative_bisect(1.0 - probability, |x| dist.cdf(x))
}

pub fn beta_dist(
    x: f64,
    alpha: f64,
    beta: f64,
    cumulative: bool,
    a: f64,
    b: f64,
) -> Result<f64, ErrorKind> {
    let x = ensure_finite(x)?;
    let alpha = ensure_positive(alpha)?;
    let beta = ensure_positive(beta)?;
    let a = ensure_finite(a)?;
    let b = ensure_finite(b)?;
    if !(a < b) {
        return Err(ErrorKind::Num);
    }

    if x < a || x > b {
        return Err(ErrorKind::Num);
    }
    let span = b - a;
    if !(span > 0.0) || !span.is_finite() {
        return Err(ErrorKind::Num);
    }
    let t = ((x - a) / span).clamp(0.0, 1.0);

    let dist = Beta::new(alpha, beta).map_err(|_| ErrorKind::Num)?;
    let out = if cumulative {
        dist.cdf(t)
    } else {
        dist.pdf(t) / span
    };
    if out.is_finite() {
        Ok(out)
    } else {
        Err(ErrorKind::Num)
    }
}

pub fn beta_inv(probability: f64, alpha: f64, beta: f64, a: f64, b: f64) -> Result<f64, ErrorKind> {
    let probability = ensure_probability(probability)?;
    let alpha = ensure_positive(alpha)?;
    let beta = ensure_positive(beta)?;
    let a = ensure_finite(a)?;
    let b = ensure_finite(b)?;
    if !(a < b) {
        return Err(ErrorKind::Num);
    }

    let dist = Beta::new(alpha, beta).map_err(|_| ErrorKind::Num)?;
    let t = inverse_cdf_unit_interval_bisect(probability, |x| dist.cdf(x))?;
    let out = a + t * (b - a);
    if out.is_finite() {
        Ok(out)
    } else {
        Err(ErrorKind::Num)
    }
}

pub fn gamma_dist(x: f64, alpha: f64, beta: f64, cumulative: bool) -> Result<f64, ErrorKind> {
    let x = ensure_nonnegative(x)?;
    let alpha = ensure_positive(alpha)?;
    let beta = ensure_positive(beta)?;

    // Excel's `beta` argument is the scale parameter. `statrs` models Gamma with a rate parameter.
    let rate = 1.0 / beta;
    if !rate.is_finite() || rate <= 0.0 {
        return Err(ErrorKind::Num);
    }
    let dist = Gamma::new(alpha, rate).map_err(|_| ErrorKind::Num)?;
    let out = if cumulative { dist.cdf(x) } else { dist.pdf(x) };
    if out.is_finite() {
        Ok(out)
    } else {
        Err(ErrorKind::Num)
    }
}

pub fn gamma_inv(probability: f64, alpha: f64, beta: f64) -> Result<f64, ErrorKind> {
    let probability = ensure_probability(probability)?;
    let alpha = ensure_positive(alpha)?;
    let beta = ensure_positive(beta)?;

    let rate = 1.0 / beta;
    if !rate.is_finite() || rate <= 0.0 {
        return Err(ErrorKind::Num);
    }
    let dist = Gamma::new(alpha, rate).map_err(|_| ErrorKind::Num)?;
    inverse_cdf_nonnegative_bisect(probability, |x| dist.cdf(x))
}

pub fn gamma_fn(number: f64) -> Result<f64, ErrorKind> {
    let number = ensure_finite(number)?;
    if number <= 0.0 && number.fract() == 0.0 {
        return Err(ErrorKind::Num);
    }

    let out = statrs::function::gamma::gamma(number);
    if out.is_finite() {
        Ok(out)
    } else {
        Err(ErrorKind::Num)
    }
}

pub fn gammaln(number: f64) -> Result<f64, ErrorKind> {
    let number = ensure_finite(number)?;
    if !(number > 0.0) {
        return Err(ErrorKind::Num);
    }

    let out = statrs::function::gamma::ln_gamma(number);
    if out.is_finite() {
        Ok(out)
    } else {
        Err(ErrorKind::Num)
    }
}

pub fn lognorm_dist(
    x: f64,
    mean: f64,
    standard_dev: f64,
    cumulative: bool,
) -> Result<f64, ErrorKind> {
    let x = ensure_finite(x)?;
    if x <= 0.0 {
        return Err(ErrorKind::Num);
    }
    let mean = ensure_finite(mean)?;
    let standard_dev = ensure_positive(standard_dev)?;

    let dist = LogNormal::new(mean, standard_dev).map_err(|_| ErrorKind::Num)?;
    let out = if cumulative { dist.cdf(x) } else { dist.pdf(x) };
    if out.is_finite() {
        Ok(out)
    } else {
        Err(ErrorKind::Num)
    }
}

pub fn lognorm_inv(probability: f64, mean: f64, standard_dev: f64) -> Result<f64, ErrorKind> {
    let probability = ensure_probability(probability)?;
    let mean = ensure_finite(mean)?;
    let standard_dev = ensure_positive(standard_dev)?;
    if probability == 0.0 {
        return Ok(0.0);
    }
    if probability == 1.0 {
        return Err(ErrorKind::Num);
    }

    // LOGNORM.INV(p, mean, std_dev) = exp(mean + std_dev * NORM.S.INV(p))
    let normal = Normal::new(0.0, 1.0).map_err(|_| ErrorKind::Num)?;
    let z = normal.inverse_cdf(probability);
    let out = (mean + standard_dev * z).exp();
    if out.is_finite() {
        Ok(out)
    } else {
        Err(ErrorKind::Num)
    }
}

pub fn expon_dist(x: f64, lambda: f64, cumulative: bool) -> Result<f64, ErrorKind> {
    let x = ensure_nonnegative(x)?;
    let lambda = ensure_positive(lambda)?;

    let dist = Exp::new(lambda).map_err(|_| ErrorKind::Num)?;
    let out = if cumulative { dist.cdf(x) } else { dist.pdf(x) };
    if out.is_finite() {
        Ok(out)
    } else {
        Err(ErrorKind::Num)
    }
}

pub fn weibull_dist(x: f64, alpha: f64, beta: f64, cumulative: bool) -> Result<f64, ErrorKind> {
    let x = ensure_nonnegative(x)?;
    let alpha = ensure_positive(alpha)?;
    let beta = ensure_positive(beta)?;

    let dist = Weibull::new(alpha, beta).map_err(|_| ErrorKind::Num)?;
    let out = if cumulative { dist.cdf(x) } else { dist.pdf(x) };
    if out.is_finite() {
        Ok(out)
    } else {
        Err(ErrorKind::Num)
    }
}

pub fn fisher(x: f64) -> Result<f64, ErrorKind> {
    let x = ensure_finite(x)?;
    if !(x > -1.0 && x < 1.0) {
        return Err(ErrorKind::Num);
    }
    let out = x.atanh();
    if out.is_finite() {
        Ok(out)
    } else {
        Err(ErrorKind::Num)
    }
}

pub fn fisherinv(y: f64) -> Result<f64, ErrorKind> {
    let y = ensure_finite(y)?;
    let out = y.tanh();
    if out.is_finite() {
        Ok(out)
    } else {
        Err(ErrorKind::Num)
    }
}

pub fn confidence_norm(alpha: f64, standard_dev: f64, size: i64) -> Result<f64, ErrorKind> {
    let alpha = ensure_finite(alpha)?;
    if !(alpha > 0.0 && alpha < 1.0) {
        return Err(ErrorKind::Num);
    }
    let standard_dev = ensure_positive(standard_dev)?;
    if size < 1 {
        return Err(ErrorKind::Num);
    }

    let n = size as f64;
    let dist = Normal::new(0.0, 1.0).map_err(|_| ErrorKind::Num)?;
    let z = dist.inverse_cdf(1.0 - alpha / 2.0);
    let out = z * standard_dev / n.sqrt();
    if out.is_finite() {
        Ok(out)
    } else {
        Err(ErrorKind::Num)
    }
}

pub fn confidence_t(alpha: f64, standard_dev: f64, size: i64) -> Result<f64, ErrorKind> {
    let alpha = ensure_finite(alpha)?;
    if !(alpha > 0.0 && alpha < 1.0) {
        return Err(ErrorKind::Num);
    }
    let standard_dev = ensure_positive(standard_dev)?;
    if size < 2 {
        return Err(ErrorKind::Num);
    }

    let n = size as f64;
    let df = (size - 1) as f64;
    let dist = StudentsT::new(0.0, 1.0, df).map_err(|_| ErrorKind::Num)?;
    let t = dist.inverse_cdf(1.0 - alpha / 2.0);
    let out = t * standard_dev / n.sqrt();
    if out.is_finite() {
        Ok(out)
    } else {
        Err(ErrorKind::Num)
    }
}
