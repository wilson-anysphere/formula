use crate::value::ErrorKind;

use statrs::distribution::{
    Binomial, Discrete, DiscreteCDF, Hypergeometric, NegativeBinomial, Poisson,
};

fn trunc_to_i64(n: f64) -> Result<i64, ErrorKind> {
    if !n.is_finite() {
        return Err(ErrorKind::Num);
    }
    let t = n.trunc();
    if t < (i64::MIN as f64) || t > (i64::MAX as f64) {
        return Err(ErrorKind::Num);
    }
    Ok(t as i64)
}

fn trunc_to_u64_nonneg(n: f64) -> Result<u64, ErrorKind> {
    let v = trunc_to_i64(n)?;
    if v < 0 {
        return Err(ErrorKind::Num);
    }
    Ok(v as u64)
}

fn validate_probability(p: f64) -> Result<f64, ErrorKind> {
    if !p.is_finite() || p < 0.0 || p > 1.0 {
        return Err(ErrorKind::Num);
    }
    Ok(p)
}

fn validate_probability_strict(p: f64) -> Result<f64, ErrorKind> {
    // Excel semantics differ across functions; some require a strictly positive
    // probability of success. We treat p == 0 as invalid here.
    if !p.is_finite() || p <= 0.0 || p > 1.0 {
        return Err(ErrorKind::Num);
    }
    Ok(p)
}

pub fn binom_dist(
    number_s: f64,
    trials: f64,
    probability_s: f64,
    cumulative: bool,
) -> Result<f64, ErrorKind> {
    let trials = trunc_to_u64_nonneg(trials)?;
    let k = trunc_to_i64(number_s)?;
    if k < 0 || k as u64 > trials {
        return Err(ErrorKind::Num);
    }

    let p = validate_probability(probability_s)?;

    // Handle degenerate probabilities explicitly because some distribution
    // implementations reject p == 0/1.
    if p == 0.0 {
        return Ok(if cumulative {
            1.0
        } else if k == 0 {
            1.0
        } else {
            0.0
        });
    }
    if p == 1.0 {
        return Ok(if cumulative {
            if (k as u64) >= trials {
                1.0
            } else {
                0.0
            }
        } else if (k as u64) == trials {
            1.0
        } else {
            0.0
        });
    }

    let dist = Binomial::new(p, trials).map_err(|_| ErrorKind::Num)?;
    let out = if cumulative {
        dist.cdf(k as u64)
    } else {
        dist.pmf(k as u64)
    };
    if out.is_finite() {
        Ok(out)
    } else {
        Err(ErrorKind::Num)
    }
}

pub fn binom_dist_range(
    trials: f64,
    probability_s: f64,
    number_s: f64,
    number_s2: Option<f64>,
) -> Result<f64, ErrorKind> {
    let trials_u = trunc_to_u64_nonneg(trials)?;
    let lo_i = trunc_to_i64(number_s)?;
    let hi_i = match number_s2 {
        Some(v) => trunc_to_i64(v)?,
        None => lo_i,
    };

    if lo_i < 0 || hi_i < 0 {
        return Err(ErrorKind::Num);
    }
    let lo = lo_i as u64;
    let hi = hi_i as u64;
    if lo > hi || lo > trials_u || hi > trials_u {
        return Err(ErrorKind::Num);
    }

    let p = validate_probability(probability_s)?;

    if p == 0.0 {
        // Degenerate at 0.
        let prob = if lo == 0 { 1.0 } else { 0.0 };
        return Ok(prob);
    }
    if p == 1.0 {
        // Degenerate at n.
        let prob = if lo <= trials_u && trials_u <= hi {
            1.0
        } else {
            0.0
        };
        return Ok(prob);
    }

    let dist = Binomial::new(p, trials_u).map_err(|_| ErrorKind::Num)?;
    let out = if lo == 0 {
        dist.cdf(hi)
    } else {
        dist.cdf(hi) - dist.cdf(lo - 1)
    };
    let out = out.max(0.0);
    if out.is_finite() {
        Ok(out)
    } else {
        Err(ErrorKind::Num)
    }
}

pub fn binom_inv(trials: f64, probability_s: f64, alpha: f64) -> Result<f64, ErrorKind> {
    let trials_u = trunc_to_u64_nonneg(trials)?;
    let p = validate_probability(probability_s)?;

    if !alpha.is_finite() || alpha < 0.0 || alpha > 1.0 {
        return Err(ErrorKind::Num);
    }

    if p == 0.0 {
        return Ok(0.0);
    }
    if p == 1.0 {
        return Ok(if alpha == 0.0 { 0.0 } else { trials_u as f64 });
    }

    let dist = Binomial::new(p, trials_u).map_err(|_| ErrorKind::Num)?;
    let out = dist.inverse_cdf(alpha) as f64;
    if out.is_finite() {
        Ok(out)
    } else {
        Err(ErrorKind::Num)
    }
}

pub fn poisson_dist(x: f64, mean: f64, cumulative: bool) -> Result<f64, ErrorKind> {
    let x_i = trunc_to_i64(x)?;
    if x_i < 0 {
        return Err(ErrorKind::Num);
    }
    if !mean.is_finite() || mean <= 0.0 {
        return Err(ErrorKind::Num);
    }

    let dist = Poisson::new(mean).map_err(|_| ErrorKind::Num)?;
    let out = if cumulative {
        dist.cdf(x_i as u64)
    } else {
        dist.pmf(x_i as u64)
    };
    if out.is_finite() {
        Ok(out)
    } else {
        Err(ErrorKind::Num)
    }
}

pub fn negbinom_dist(
    number_f: f64,
    number_s: f64,
    probability_s: f64,
    cumulative: bool,
) -> Result<f64, ErrorKind> {
    let f_i = trunc_to_i64(number_f)?;
    let s_i = trunc_to_i64(number_s)?;

    if f_i < 0 || s_i < 1 {
        return Err(ErrorKind::Num);
    }

    let p = validate_probability_strict(probability_s)?;

    // Degenerate at 0 failures when p == 1.
    if p == 1.0 {
        return Ok(if cumulative {
            1.0
        } else if f_i == 0 {
            1.0
        } else {
            0.0
        });
    }

    let dist = NegativeBinomial::new(s_i as f64, p).map_err(|_| ErrorKind::Num)?;
    let out = if cumulative {
        dist.cdf(f_i as u64)
    } else {
        dist.pmf(f_i as u64)
    };
    if out.is_finite() {
        Ok(out)
    } else {
        Err(ErrorKind::Num)
    }
}

pub fn hypgeom_dist(
    sample_s: f64,
    number_sample: f64,
    population_s: f64,
    number_pop: f64,
    cumulative: bool,
) -> Result<f64, ErrorKind> {
    let k_i = trunc_to_i64(sample_s)?;
    let draws_i = trunc_to_i64(number_sample)?;
    let succ_pop_i = trunc_to_i64(population_s)?;
    let pop_i = trunc_to_i64(number_pop)?;

    if pop_i <= 0 || draws_i < 0 || succ_pop_i < 0 || k_i < 0 {
        return Err(ErrorKind::Num);
    }
    let pop = pop_i as u64;
    let draws = draws_i as u64;
    let succ_pop = succ_pop_i as u64;
    let k = k_i as u64;

    if draws > pop || succ_pop > pop {
        return Err(ErrorKind::Num);
    }

    // Feasible k range.
    let min_k = draws.saturating_sub(pop.saturating_sub(succ_pop));
    let max_k = draws.min(succ_pop);
    if k < min_k || k > max_k {
        return Err(ErrorKind::Num);
    }

    let dist = Hypergeometric::new(pop, succ_pop, draws).map_err(|_| ErrorKind::Num)?;
    let out = if cumulative { dist.cdf(k) } else { dist.pmf(k) };
    if out.is_finite() {
        Ok(out)
    } else {
        Err(ErrorKind::Num)
    }
}

pub fn prob(
    x_values: &[f64],
    prob_values: &[f64],
    lower_limit: f64,
    upper_limit: Option<f64>,
) -> Result<f64, ErrorKind> {
    if x_values.len() != prob_values.len() {
        return Err(ErrorKind::NA);
    }
    if x_values.is_empty() {
        return Err(ErrorKind::Num);
    }
    if !lower_limit.is_finite() {
        return Err(ErrorKind::Num);
    }
    let upper = match upper_limit {
        Some(v) => {
            if !v.is_finite() {
                return Err(ErrorKind::Num);
            }
            v
        }
        None => lower_limit,
    };
    if upper < lower_limit {
        return Err(ErrorKind::Num);
    }

    // Validate probabilities and sum-to-1 invariant.
    let mut sum = 0.0;
    for &p in prob_values {
        if !p.is_finite() || p < 0.0 || p > 1.0 {
            return Err(ErrorKind::Num);
        }
        sum += p;
    }
    if !sum.is_finite() {
        return Err(ErrorKind::Num);
    }
    if (sum - 1.0).abs() > 1e-10 {
        return Err(ErrorKind::Num);
    }

    let mut out = 0.0;
    if upper_limit.is_none() {
        for (&x, &p) in x_values.iter().zip(prob_values.iter()) {
            if x == lower_limit {
                out += p;
            }
        }
    } else {
        for (&x, &p) in x_values.iter().zip(prob_values.iter()) {
            if x >= lower_limit && x <= upper {
                out += p;
            }
        }
    }

    if out.is_finite() {
        Ok(out)
    } else {
        Err(ErrorKind::Num)
    }
}
