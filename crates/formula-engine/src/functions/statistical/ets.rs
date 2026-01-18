use crate::date::ExcelDateSystem;
use crate::eval::MAX_MATERIALIZED_ARRAY_CELLS;
use crate::functions::date_time;
use crate::value::ErrorKind;
use smallvec::SmallVec;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum AggregationMethod {
    Average,
    Count,
    CountA,
    Max,
    Median,
    Min,
    Sum,
}

impl AggregationMethod {
    pub fn from_code(code: i64) -> Result<Self, ErrorKind> {
        match code {
            1 => Ok(Self::Average),
            2 => Ok(Self::Count),
            3 => Ok(Self::CountA),
            4 => Ok(Self::Max),
            5 => Ok(Self::Median),
            6 => Ok(Self::Min),
            7 => Ok(Self::Sum),
            _ => Err(ErrorKind::Num),
        }
    }
}

#[derive(Debug, Clone, Copy)]
pub enum TimelineStep {
    /// Timeline uses a fixed numeric step (`timeline[i] = start + step * i`).
    Fixed { start: f64, step: f64 },
    /// Timeline advances in fixed calendar-month steps (EDATE semantics).
    ///
    /// This is required to support common Excel use-cases like monthly/quarterly/yearly date
    /// timelines, where the difference in serial days varies (28/29/30/31).
    Month {
        start_serial: i32,
        months_step: i32,
        system: ExcelDateSystem,
    },
    /// Timeline advances in fixed calendar-month steps anchored to end-of-month (EOMONTH semantics).
    ///
    /// This is required to support common Excel timelines that use month-end dates
    /// (e.g. 2020-02-29, 2020-03-31, 2020-04-30...), where `EDATE` does not preserve month-end
    /// when the start date is clamped (e.g. Feb 29).
    MonthEnd {
        start_serial: i32,
        months_step: i32,
        system: ExcelDateSystem,
    },
}

#[derive(Debug, Clone)]
pub struct PreparedSeries {
    pub step: TimelineStep,
    pub values: Vec<f64>,
}

impl PreparedSeries {
    pub fn position(&self, target_date: f64) -> Result<f64, ErrorKind> {
        if !target_date.is_finite() {
            return Err(ErrorKind::Num);
        }
        match self.step {
            TimelineStep::Fixed { start, step } => {
                if !step.is_finite() || step == 0.0 {
                    return Err(ErrorKind::Num);
                }
                let pos = (target_date - start) / step;
                if pos.is_finite() {
                    Ok(pos)
                } else {
                    Err(ErrorKind::Num)
                }
            }
            TimelineStep::Month {
                start_serial,
                months_step,
                system,
            } => month_step_position(target_date, start_serial, months_step, system),
            TimelineStep::MonthEnd {
                start_serial,
                months_step,
                system,
            } => month_end_step_position(target_date, start_serial, months_step, system),
        }
    }

    pub fn last_pos(&self) -> f64 {
        (self.values.len().saturating_sub(1)) as f64
    }
}

#[derive(Debug, Clone)]
pub struct EtsFit {
    pub seasonality: usize,
    pub alpha: f64,
    pub beta: f64,
    pub gamma: f64,
    pub phi: f64,
    pub level: f64,
    pub trend: f64,
    /// Seasonal components indexed by `t % seasonality`.
    pub seasonals: Vec<f64>,
    pub mae: f64,
    pub rmse: f64,
    pub smape: f64,
    pub mase: f64,
    last_t: usize,
}

impl EtsFit {
    fn forecast_int(&self, h: usize) -> f64 {
        let seasonal = if self.seasonality > 1 {
            self.seasonals[(self.last_t + h) % self.seasonality]
        } else {
            0.0
        };

        let trend_term = if self.trend == 0.0 || h == 0 {
            0.0
        } else if (self.phi - 1.0).abs() < 1e-12 {
            self.trend * (h as f64)
        } else {
            // Sum_{i=1..h} phi^i = phi * (1 - phi^h) / (1 - phi)
            let phi = self.phi;
            let sum = phi * (1.0 - phi.powi(h as i32)) / (1.0 - phi);
            self.trend * sum
        };

        self.level + trend_term + seasonal
    }

    /// Forecast `steps_ahead` periods after the last observed point (`steps_ahead=1` is the next
    /// period). Fractional horizons are linearly interpolated between the adjacent integer steps.
    pub fn forecast(&self, steps_ahead: f64) -> Result<f64, ErrorKind> {
        if !steps_ahead.is_finite() {
            return Err(ErrorKind::Num);
        }
        if steps_ahead <= 0.0 {
            return Ok(self.forecast_int(0));
        }

        let h0 = steps_ahead.floor();
        let frac = steps_ahead - h0;
        let h0 = h0 as usize;
        if frac == 0.0 {
            return Ok(self.forecast_int(h0));
        }
        let a = self.forecast_int(h0);
        let b = self.forecast_int(h0 + 1);
        let out = a + frac * (b - a);
        if out.is_finite() {
            Ok(out)
        } else {
            Err(ErrorKind::Num)
        }
    }
}

pub fn prepare_series(
    values: &[f64],
    timeline: &[f64],
    data_completion: bool,
    aggregation: AggregationMethod,
    system: ExcelDateSystem,
) -> Result<PreparedSeries, ErrorKind> {
    if values.len() != timeline.len() {
        return Err(ErrorKind::NA);
    }
    if values.len() < 2 {
        return Err(ErrorKind::Num);
    }
    if values.len() > MAX_MATERIALIZED_ARRAY_CELLS {
        return Err(ErrorKind::Num);
    }

    let mut pairs: Vec<(f64, f64)> = Vec::new();
    if pairs.try_reserve_exact(values.len()).is_err() {
        debug_assert!(false, "ETS series allocation failed (pairs={})", values.len());
        return Err(ErrorKind::Num);
    }
    for (&t, &v) in timeline.iter().zip(values.iter()) {
        if !t.is_finite() || !v.is_finite() {
            return Err(ErrorKind::Num);
        }
        pairs.push((t, v));
    }

    pairs.sort_by(|(ta, _), (tb, _)| ta.total_cmp(tb));

    // Aggregate duplicates by timeline.
    let mut uniq_timeline: Vec<f64> = Vec::new();
    let mut uniq_values: Vec<f64> = Vec::new();
    if uniq_timeline.try_reserve_exact(pairs.len()).is_err()
        || uniq_values.try_reserve_exact(pairs.len()).is_err()
    {
        debug_assert!(
            false,
            "ETS unique-series allocation failed (pairs={})",
            pairs.len()
        );
        return Err(ErrorKind::Num);
    }

    let mut i = 0usize;
    while i < pairs.len() {
        let t = pairs[i].0;
        let mut end = i + 1;
        while end < pairs.len() && pairs[end].0 == t {
            end += 1;
        }
        let group_len = end - i;
        let mut group: Vec<f64> = Vec::new();
        if group.try_reserve_exact(group_len).is_err() {
            debug_assert!(false, "ETS group allocation failed (len={group_len})");
            return Err(ErrorKind::Num);
        }
        for k in i..end {
            group.push(pairs[k].1);
        }
        i = end;
        let agg = aggregate_group(&group, aggregation)?;
        uniq_timeline.push(t);
        uniq_values.push(agg);
    }

    if uniq_timeline.len() < 2 {
        return Err(ErrorKind::Num);
    }

    if let Ok(step) = base_step(&uniq_timeline) {
        let start = uniq_timeline[0];

        fn compute_multiple(t0: f64, t1: f64, step: f64) -> Result<usize, ErrorKind> {
            let diff = t1 - t0;
            let multiple_f = diff / step;
            let multiple_i64 = multiple_f.round() as i64;
            if multiple_i64 < 1 {
                return Err(ErrorKind::Num);
            }
            if multiple_i64 > MAX_MATERIALIZED_ARRAY_CELLS as i64 {
                return Err(ErrorKind::Num);
            }
            Ok(multiple_i64 as usize)
        }

        let mut total_points = 1usize;
        for idx in 0..uniq_timeline.len() - 1 {
            let multiple = compute_multiple(uniq_timeline[idx], uniq_timeline[idx + 1], step)?;
            total_points = total_points.checked_add(multiple).ok_or(ErrorKind::Num)?;
            if total_points > MAX_MATERIALIZED_ARRAY_CELLS {
                return Err(ErrorKind::Num);
            }
        }

        let mut out_values: Vec<f64> = Vec::new();
        if out_values.try_reserve_exact(total_points).is_err() {
            debug_assert!(
                false,
                "ETS expanded-series allocation failed (points={total_points})"
            );
            return Err(ErrorKind::Num);
        }

        for idx in 0..uniq_timeline.len() - 1 {
            let t0 = uniq_timeline[idx];
            let t1 = uniq_timeline[idx + 1];
            let v0 = uniq_values[idx];
            let v1 = uniq_values[idx + 1];
            out_values.push(v0);

            let multiple = compute_multiple(t0, t1, step)?;
            if multiple <= 1 {
                continue;
            }

            // Fill missing points between t0 and t1.
            for k in 1..multiple {
                let filled = if data_completion {
                    let frac = (k as f64) / (multiple as f64);
                    v0 + frac * (v1 - v0)
                } else {
                    0.0
                };
                out_values.push(filled);
            }
        }
        let Some(&last) = uniq_values.last() else {
            debug_assert!(false, "uniq_timeline.len() >= 2 implies uniq_values is non-empty");
            return Err(ErrorKind::Num);
        };
        out_values.push(last);

        return Ok(PreparedSeries {
            step: TimelineStep::Fixed { start, step },
            values: out_values,
        });
    }

    prepare_series_month_step(&uniq_timeline, &uniq_values, data_completion, system)
}

fn as_integer_day(serial: f64) -> Option<i32> {
    if !serial.is_finite() {
        return None;
    }
    let rounded = serial.round();
    if (serial - rounded).abs() > 1e-9 {
        return None;
    }
    if rounded < (i32::MIN as f64) || rounded > (i32::MAX as f64) {
        return None;
    }
    Some(rounded as i32)
}

fn month_index(serial: i32, system: ExcelDateSystem) -> Result<i32, ErrorKind> {
    let ymd = crate::date::serial_to_ymd(serial, system).map_err(|_| ErrorKind::Num)?;
    Ok(ymd.year * 12 + i32::from(ymd.month.saturating_sub(1)))
}

fn prepare_series_month_step(
    timeline: &[f64],
    values: &[f64],
    data_completion: bool,
    system: ExcelDateSystem,
) -> Result<PreparedSeries, ErrorKind> {
    if timeline.len() != values.len() || timeline.len() < 2 {
        return Err(ErrorKind::Num);
    }

    if timeline.len() > MAX_MATERIALIZED_ARRAY_CELLS {
        return Err(ErrorKind::Num);
    }

    let mut serials: Vec<i32> = Vec::new();
    if serials.try_reserve_exact(timeline.len()).is_err() {
        debug_assert!(
            false,
            "ETS month-step allocation failed (serials={})",
            timeline.len()
        );
        return Err(ErrorKind::Num);
    }
    for &t in timeline {
        serials.push(as_integer_day(t).ok_or(ErrorKind::Num)?);
    }

    // Compute the base step in calendar months.
    let mut month_ids: Vec<i32> = Vec::new();
    if month_ids.try_reserve_exact(serials.len()).is_err() {
        debug_assert!(
            false,
            "ETS month-step allocation failed (month_ids={})",
            serials.len()
        );
        return Err(ErrorKind::Num);
    }
    for &s in &serials {
        month_ids.push(month_index(s, system)?);
    }

    let mut months_step = i32::MAX;
    let mut diffs: Vec<i32> = Vec::new();
    if diffs
        .try_reserve_exact(month_ids.len().saturating_sub(1))
        .is_err()
    {
        debug_assert!(
            false,
            "ETS month-step allocation failed (diffs={})",
            month_ids.len().saturating_sub(1)
        );
        return Err(ErrorKind::Num);
    }
    for w in month_ids.windows(2) {
        let d = w[1] - w[0];
        if d <= 0 {
            return Err(ErrorKind::Num);
        }
        months_step = months_step.min(d);
        diffs.push(d);
    }
    if months_step <= 0 || months_step == i32::MAX {
        return Err(ErrorKind::Num);
    }
    for d in diffs {
        if d % months_step != 0 {
            return Err(ErrorKind::Num);
        }
    }

    let start_serial = serials[0];
    let start_month_id = month_ids[0];

    fn build_known<F>(
        serials: &[i32],
        month_ids: &[i32],
        values: &[f64],
        start_month_id: i32,
        months_step: i32,
        mut expected_for_months: F,
    ) -> Result<(std::collections::BTreeMap<usize, f64>, usize), ErrorKind>
    where
        F: FnMut(i32) -> Result<i32, ErrorKind>,
    {
        let mut known = std::collections::BTreeMap::<usize, f64>::new();
        let mut max_k = 0usize;
        for ((&serial, &month_id), &v) in serials.iter().zip(month_ids.iter()).zip(values.iter()) {
            let month_diff = month_id - start_month_id;
            if month_diff < 0 || month_diff % months_step != 0 {
                return Err(ErrorKind::Num);
            }
            let k_i32 = month_diff / months_step;
            let expected = expected_for_months(k_i32 * months_step)?;
            if expected != serial {
                return Err(ErrorKind::Num);
            }
            let k = k_i32 as usize;
            max_k = max_k.max(k);
            known.insert(k, v);
        }
        Ok((known, max_k))
    }

    // First try EDATE semantics (preserve day-of-month where possible).
    let (known, max_k, step) = match build_known(
        &serials,
        &month_ids,
        values,
        start_month_id,
        months_step,
        |months| date_time::edate(start_serial, months, system).map_err(|_| ErrorKind::Num),
    ) {
        Ok((known, max_k)) => (
            known,
            max_k,
            TimelineStep::Month {
                start_serial,
                months_step,
                system,
            },
        ),
        Err(_) => {
            // Fall back to EOMONTH semantics (month-end anchored). This covers month-end sequences
            // that do not follow EDATE after a clamped start date like Feb 29.
            let (known, max_k) = build_known(
                &serials,
                &month_ids,
                values,
                start_month_id,
                months_step,
                |months| {
                    date_time::eomonth(start_serial, months, system).map_err(|_| ErrorKind::Num)
                },
            )?;
            (
                known,
                max_k,
                TimelineStep::MonthEnd {
                    start_serial,
                    months_step,
                    system,
                },
            )
        }
    };

    // Expand to the full set of periods [0..=max_k], filling missing values according to
    // `data_completion` (interpolate or zero-fill).
    let total = max_k.saturating_add(1);
    if total > MAX_MATERIALIZED_ARRAY_CELLS {
        return Err(ErrorKind::Num);
    }
    let mut out: Vec<f64> = Vec::new();
    if out.try_reserve_exact(total).is_err() {
        debug_assert!(false, "ETS expanded-series allocation failed (points={total})");
        return Err(ErrorKind::Num);
    }
    out.resize(total, f64::NAN);
    for (&k, &v) in &known {
        out[k] = v;
    }

    if !data_completion {
        for v in &mut out {
            if v.is_nan() {
                *v = 0.0;
            }
        }
    } else {
        // Interpolate across gaps in period space.
        let mut prev_idx = None::<usize>;
        for idx in 0..out.len() {
            if !out[idx].is_nan() {
                if let Some(pi) = prev_idx {
                    let pv = out[pi];
                    let cv = out[idx];
                    let span = (idx - pi) as f64;
                    if span > 1.0 {
                        for k in (pi + 1)..idx {
                            let frac = (k - pi) as f64 / span;
                            out[k] = pv + frac * (cv - pv);
                        }
                    }
                }
                prev_idx = Some(idx);
            }
        }

        if out.iter().any(|v| v.is_nan()) {
            return Err(ErrorKind::Num);
        }
    }

    Ok(PreparedSeries { step, values: out })
}

fn month_step_position(
    target_date: f64,
    start_serial: i32,
    months_step: i32,
    system: ExcelDateSystem,
) -> Result<f64, ErrorKind> {
    if months_step <= 0 || !target_date.is_finite() {
        return Err(ErrorKind::Num);
    }

    let start_month_id = month_index(start_serial, system)?;
    let target_day = as_integer_day(target_date.floor()).ok_or(ErrorKind::Num)?;
    let target_month_id = month_index(target_day, system)?;
    let month_diff = target_month_id - start_month_id;

    // Initial guess based on year/month difference, then adjust for day-of-month clamping.
    let mut k = month_diff.div_euclid(months_step);

    // Ensure date_k <= target_date < date_{k+1}.
    for _ in 0..4 {
        let date_k = date_time::edate(start_serial, k * months_step, system)
            .map_err(|_| ErrorKind::Num)? as f64;
        if target_date < date_k {
            k -= 1;
            continue;
        }
        let date_k1 = date_time::edate(start_serial, (k + 1) * months_step, system)
            .map_err(|_| ErrorKind::Num)? as f64;
        if target_date >= date_k1 {
            k += 1;
            continue;
        }

        let denom = date_k1 - date_k;
        if denom == 0.0 || !denom.is_finite() {
            return Err(ErrorKind::Num);
        }
        let frac = (target_date - date_k) / denom;
        let pos = k as f64 + frac;
        if pos.is_finite() {
            return Ok(pos);
        } else {
            return Err(ErrorKind::Num);
        }
    }
    Err(ErrorKind::Num)
}

fn month_end_step_position(
    target_date: f64,
    start_serial: i32,
    months_step: i32,
    system: ExcelDateSystem,
) -> Result<f64, ErrorKind> {
    if months_step <= 0 || !target_date.is_finite() {
        return Err(ErrorKind::Num);
    }

    let start_month_id = month_index(start_serial, system)?;
    let target_day = as_integer_day(target_date.floor()).ok_or(ErrorKind::Num)?;
    let target_month_id = month_index(target_day, system)?;
    let month_diff = target_month_id - start_month_id;

    let mut k = month_diff.div_euclid(months_step);

    // Ensure date_k <= target_date < date_{k+1}.
    for _ in 0..4 {
        let date_k = date_time::eomonth(start_serial, k * months_step, system)
            .map_err(|_| ErrorKind::Num)? as f64;
        if target_date < date_k {
            k -= 1;
            continue;
        }
        let date_k1 = date_time::eomonth(start_serial, (k + 1) * months_step, system)
            .map_err(|_| ErrorKind::Num)? as f64;
        if target_date >= date_k1 {
            k += 1;
            continue;
        }

        let denom = date_k1 - date_k;
        if denom == 0.0 || !denom.is_finite() {
            return Err(ErrorKind::Num);
        }
        let frac = (target_date - date_k) / denom;
        let pos = k as f64 + frac;
        if pos.is_finite() {
            return Ok(pos);
        } else {
            return Err(ErrorKind::Num);
        }
    }
    Err(ErrorKind::Num)
}

pub fn detect_seasonality(values: &[f64]) -> Result<usize, ErrorKind> {
    let n = values.len();
    if n < 2 {
        return Err(ErrorKind::Num);
    }
    if n > MAX_MATERIALIZED_ARRAY_CELLS {
        return Err(ErrorKind::Num);
    }

    // Need at least two full seasons to estimate seasonality reliably.
    let max_m = (n / 2).min(8760);
    if max_m < 2 {
        return Ok(1);
    }

    // Detrend via simple linear regression to avoid selecting large "seasonality" on trending
    // series (autocorrelation stays high for many lags when the series has a strong trend).
    let mean_t = (n as f64 - 1.0) / 2.0;
    let mean_y = values.iter().sum::<f64>() / (n as f64);
    let mut cov = 0.0;
    let mut var_t = 0.0;
    for (i, &y) in values.iter().enumerate() {
        let t = i as f64;
        let dt = t - mean_t;
        cov += dt * (y - mean_y);
        var_t += dt * dt;
    }
    let slope = if var_t == 0.0 { 0.0 } else { cov / var_t };
    let intercept = mean_y - slope * mean_t;

    let mut residuals: Vec<f64> = Vec::new();
    if residuals.try_reserve_exact(n).is_err() {
        debug_assert!(false, "ETS residual allocation failed (n={n})");
        return Err(ErrorKind::Num);
    }
    for (i, &y) in values.iter().enumerate() {
        residuals.push(y - (intercept + slope * (i as f64)));
    }

    let mut best_m = 1usize;
    let mut best_corr = 0.0f64;

    // Compute correlations for candidate seasonalities. Prefer smaller lags when correlations are
    // within a small epsilon of the best value.
    const EPS: f64 = 1e-3;
    for m in 2..=max_m {
        // Require at least 2 seasons.
        if n < 2 * m {
            continue;
        }
        let (corr, ok) = autocorr(&residuals, m);
        if !ok {
            continue;
        }
        let corr = corr.abs();
        if corr > best_corr + EPS || ((corr - best_corr).abs() <= EPS && m < best_m) {
            best_corr = corr;
            best_m = m;
        }
    }

    // Heuristic: if we don't have a reasonably strong seasonal signal, fall back to no
    // seasonality (m=1). The threshold is intentionally conservative.
    if best_corr < 0.5 {
        Ok(1)
    } else {
        Ok(best_m)
    }
}

pub fn fit(values: &[f64], seasonality: usize) -> Result<EtsFit, ErrorKind> {
    if values.len() < 2 {
        return Err(ErrorKind::Num);
    }
    if seasonality == 0 || seasonality > 8760 {
        return Err(ErrorKind::Num);
    }
    if seasonality > 1 && values.len() < 2 * seasonality {
        return Err(ErrorKind::Num);
    }
    if values.iter().any(|v| !v.is_finite()) {
        return Err(ErrorKind::Num);
    }

    let (alpha, beta, gamma, phi) = optimize_params(values, seasonality);
    let sim = simulate(values, seasonality, alpha, beta, gamma, phi)?;

    Ok(sim)
}

fn aggregate_group(values: &[f64], method: AggregationMethod) -> Result<f64, ErrorKind> {
    if values.is_empty() {
        return Err(ErrorKind::Num);
    }
    match method {
        AggregationMethod::Average => Ok(values.iter().sum::<f64>() / (values.len() as f64)),
        AggregationMethod::Count | AggregationMethod::CountA => Ok(values.len() as f64),
        AggregationMethod::Max => Ok(values
            .iter()
            .copied()
            .fold(f64::NEG_INFINITY, |a, b| a.max(b))),
        AggregationMethod::Min => Ok(values.iter().copied().fold(f64::INFINITY, |a, b| a.min(b))),
        AggregationMethod::Sum => Ok(values.iter().sum::<f64>()),
        AggregationMethod::Median => {
            let mut sorted: Vec<f64> = Vec::new();
            if sorted.try_reserve_exact(values.len()).is_err() {
                debug_assert!(
                    false,
                    "ETS median allocation failed (len={})",
                    values.len()
                );
                return Err(ErrorKind::Num);
            }
            sorted.extend_from_slice(values);
            sorted.sort_by(|a, b| a.total_cmp(b));
            let mid = sorted.len() / 2;
            if sorted.len() % 2 == 1 {
                Ok(sorted[mid])
            } else {
                Ok((sorted[mid - 1] + sorted[mid]) / 2.0)
            }
        }
    }
}

fn base_step(timeline: &[f64]) -> Result<f64, ErrorKind> {
    if timeline.len() < 2 {
        return Err(ErrorKind::Num);
    }
    let mut step = f64::INFINITY;
    for w in timeline.windows(2) {
        let d = w[1] - w[0];
        if !(d > 0.0) || !d.is_finite() {
            return Err(ErrorKind::Num);
        }
        step = step.min(d);
    }
    if !(step > 0.0) || !step.is_finite() {
        return Err(ErrorKind::Num);
    }

    // Validate that every diff is an integer multiple of the smallest diff (within tolerance).
    // This supports missing timeline points (filled via data completion) but rejects irregular
    // spacing.
    const TOL: f64 = 1e-6;
    for w in timeline.windows(2) {
        let d = w[1] - w[0];
        let ratio = d / step;
        let rounded = ratio.round();
        if rounded < 1.0 {
            return Err(ErrorKind::Num);
        }
        if (ratio - rounded).abs() > TOL {
            return Err(ErrorKind::Num);
        }
    }
    Ok(step)
}

fn autocorr(values: &[f64], lag: usize) -> (f64, bool) {
    let n = values.len();
    if lag == 0 || lag >= n {
        return (0.0, false);
    }
    let mut sum_xy = 0.0;
    let mut sum_x2 = 0.0;
    let mut sum_y2 = 0.0;
    for t in lag..n {
        let x = values[t];
        let y = values[t - lag];
        sum_xy += x * y;
        sum_x2 += x * x;
        sum_y2 += y * y;
    }
    if sum_x2 == 0.0 || sum_y2 == 0.0 {
        return (0.0, false);
    }
    (sum_xy / (sum_x2 * sum_y2).sqrt(), true)
}

#[derive(Clone)]
struct Vertex {
    x: [f64; 4],
    f: f64,
}

fn optimize_params(values: &[f64], seasonality: usize) -> (f64, f64, f64, f64) {
    let dims = if seasonality > 1 { 4 } else { 3 };
    let x0 = if seasonality > 1 {
        [0.5, 0.1, 0.1, 1.0]
    } else {
        // (alpha, beta, phi) when there is no seasonality.
        [0.5, 0.1, 1.0, 0.0]
    };

    let mut simplex: SmallVec<[Vertex; 5]> = SmallVec::new();
    let f0 = objective(values, seasonality, &x0);
    simplex.push(Vertex { x: x0, f: f0 });

    let steps = [0.05, 0.05, 0.05, 0.02];
    for i in 0..dims {
        let mut x = x0;
        x[i] = clamp01(x[i] + steps[i]);
        let f = objective(values, seasonality, &x);
        simplex.push(Vertex { x, f });
    }

    const ALPHA: f64 = 1.0;
    const GAMMA: f64 = 2.0;
    const RHO: f64 = 0.5;
    const SIGMA: f64 = 0.5;

    const MAX_ITER: usize = 200;
    const FTOL: f64 = 1e-10;

    for _ in 0..MAX_ITER {
        simplex.sort_by(|a, b| a.f.total_cmp(&b.f));

        let best_f = simplex[0].f;
        let worst_f = simplex[dims].f;
        if (worst_f - best_f).abs() < FTOL {
            break;
        }

        // Centroid of best dims vertices (exclude worst).
        let mut centroid = [0.0; 4];
        for v in simplex.iter().take(dims) {
            for i in 0..dims {
                centroid[i] += v.x[i];
            }
        }
        for i in 0..dims {
            centroid[i] /= dims as f64;
        }

        let worst = simplex[dims].x;

        // Reflection
        let mut xr = [0.0; 4];
        for i in 0..dims {
            xr[i] = clamp01(centroid[i] + ALPHA * (centroid[i] - worst[i]));
        }
        let fr = objective(values, seasonality, &xr);

        if fr < simplex[0].f {
            // Expansion
            let mut xe = [0.0; 4];
            for i in 0..dims {
                xe[i] = clamp01(centroid[i] + GAMMA * (xr[i] - centroid[i]));
            }
            let fe = objective(values, seasonality, &xe);
            if fe < fr {
                simplex[dims] = Vertex { x: xe, f: fe };
            } else {
                simplex[dims] = Vertex { x: xr, f: fr };
            }
            continue;
        }

        if fr < simplex[dims - 1].f {
            simplex[dims] = Vertex { x: xr, f: fr };
            continue;
        }

        // Contraction
        let mut xc = [0.0; 4];
        if fr < simplex[dims].f {
            // Outside contraction
            for i in 0..dims {
                xc[i] = clamp01(centroid[i] + RHO * (xr[i] - centroid[i]));
            }
        } else {
            // Inside contraction
            for i in 0..dims {
                xc[i] = clamp01(centroid[i] + RHO * (worst[i] - centroid[i]));
            }
        }
        let fc = objective(values, seasonality, &xc);
        if fc < simplex[dims].f {
            simplex[dims] = Vertex { x: xc, f: fc };
            continue;
        }

        // Shrink
        let best = simplex[0].x;
        for i in 1..=dims {
            for j in 0..dims {
                simplex[i].x[j] = clamp01(best[j] + SIGMA * (simplex[i].x[j] - best[j]));
            }
            simplex[i].f = objective(values, seasonality, &simplex[i].x);
        }
    }

    simplex.sort_by(|a, b| a.f.total_cmp(&b.f));
    let best = &simplex[0].x;
    let alpha = best[0];
    let beta = best[1];
    let gamma = if seasonality > 1 { best[2] } else { 0.0 };
    let phi = if seasonality > 1 { best[3] } else { best[2] };
    (alpha, beta, gamma, phi)
}

fn objective(values: &[f64], seasonality: usize, x: &[f64; 4]) -> f64 {
    let (alpha, beta, gamma, phi) = if seasonality > 1 {
        (x[0], x[1], x[2], x[3])
    } else {
        (x[0], x[1], 0.0, x[2])
    };

    match simulate_sse(values, seasonality, alpha, beta, gamma, phi) {
        Ok(v) => v,
        Err(_) => f64::INFINITY,
    }
}

fn clamp01(x: f64) -> f64 {
    if x.is_nan() {
        return 0.5;
    }
    x.clamp(0.0, 1.0)
}

fn simulate_sse(
    values: &[f64],
    seasonality: usize,
    alpha: f64,
    beta: f64,
    gamma: f64,
    phi: f64,
) -> Result<f64, ErrorKind> {
    let sim = simulate(values, seasonality, alpha, beta, gamma, phi)?;
    Ok(sim.rmse * sim.rmse * (sim_error_count(values.len(), seasonality) as f64))
}

fn sim_error_count(n: usize, seasonality: usize) -> usize {
    if seasonality > 1 {
        n.saturating_sub(seasonality)
    } else {
        n.saturating_sub(1)
    }
}

fn simulate(
    values: &[f64],
    seasonality: usize,
    alpha: f64,
    beta: f64,
    gamma: f64,
    phi: f64,
) -> Result<EtsFit, ErrorKind> {
    let n = values.len();
    if n < 2 {
        return Err(ErrorKind::Num);
    }
    if seasonality == 0 {
        return Err(ErrorKind::Num);
    }
    if seasonality > 1 && n < 2 * seasonality {
        return Err(ErrorKind::Num);
    }

    let alpha = clamp01(alpha);
    let beta = clamp01(beta);
    let gamma = if seasonality > 1 { clamp01(gamma) } else { 0.0 };
    let phi = clamp01(phi);

    let mut level;
    let mut trend;
    let total = seasonality.max(1);
    let mut seasonals: Vec<f64> = Vec::new();
    if seasonals.try_reserve_exact(total).is_err() {
        debug_assert!(false, "ETS seasonal allocation failed (len={total})");
        return Err(ErrorKind::Num);
    }
    seasonals.resize(total, 0.0);

    if seasonality > 1 {
        let m = seasonality;
        let mean1 = values[..m].iter().sum::<f64>() / (m as f64);
        let mean2 = values[m..2 * m].iter().sum::<f64>() / (m as f64);
        level = mean1;
        trend = (mean2 - mean1) / (m as f64);
        for i in 0..m {
            seasonals[i] = values[i] - mean1;
        }
    } else {
        level = values[0];
        trend = values[1] - values[0];
    }

    let mut sum_abs = 0.0;
    let mut sum_sq = 0.0;
    let mut smape_sum = 0.0;
    let mut err_count = 0usize;

    if seasonality > 1 {
        let m = seasonality;
        for t in m..n {
            let idx = t % m;
            let prev_level = level;
            let prev_trend = trend;
            let prev_season = seasonals[idx];

            let forecast = prev_level + phi * prev_trend + prev_season;
            let e = values[t] - forecast;

            if e.is_finite() {
                sum_abs += e.abs();
                sum_sq += e * e;
                let denom = values[t].abs() + forecast.abs();
                if denom != 0.0 {
                    smape_sum += 2.0 * e.abs() / denom;
                }
                err_count += 1;
            }

            level =
                alpha * (values[t] - prev_season) + (1.0 - alpha) * (prev_level + phi * prev_trend);
            trend = beta * (level - prev_level) + (1.0 - beta) * phi * prev_trend;
            seasonals[idx] = gamma * (values[t] - level) + (1.0 - gamma) * prev_season;

            if !level.is_finite() || !trend.is_finite() || !seasonals[idx].is_finite() {
                return Err(ErrorKind::Num);
            }
        }
    } else {
        for t in 1..n {
            let prev_level = level;
            let prev_trend = trend;
            let forecast = prev_level + phi * prev_trend;
            let e = values[t] - forecast;

            if e.is_finite() {
                sum_abs += e.abs();
                sum_sq += e * e;
                let denom = values[t].abs() + forecast.abs();
                if denom != 0.0 {
                    smape_sum += 2.0 * e.abs() / denom;
                }
                err_count += 1;
            }

            level = alpha * values[t] + (1.0 - alpha) * (prev_level + phi * prev_trend);
            trend = beta * (level - prev_level) + (1.0 - beta) * phi * prev_trend;

            if !level.is_finite() || !trend.is_finite() {
                return Err(ErrorKind::Num);
            }
        }
    }

    if err_count == 0 {
        return Err(ErrorKind::Num);
    }

    let mae = sum_abs / (err_count as f64);
    let rmse = (sum_sq / (err_count as f64)).sqrt();
    let smape = smape_sum / (err_count as f64);

    let mase = mase(values, seasonality, mae)?;

    Ok(EtsFit {
        seasonality,
        alpha,
        beta,
        gamma,
        phi,
        level,
        trend,
        seasonals,
        mae,
        rmse,
        smape,
        mase,
        last_t: n - 1,
    })
}

fn mase(values: &[f64], seasonality: usize, mae: f64) -> Result<f64, ErrorKind> {
    let n = values.len();
    if n < 2 {
        return Err(ErrorKind::Num);
    }
    let lag = if seasonality > 1 { seasonality } else { 1 };
    if n <= lag {
        return Err(ErrorKind::Num);
    }

    let mut naive_sum = 0.0;
    let mut count = 0usize;
    for t in lag..n {
        naive_sum += (values[t] - values[t - lag]).abs();
        count += 1;
    }
    if count == 0 {
        return Err(ErrorKind::Num);
    }
    let naive_mae = naive_sum / (count as f64);
    if naive_mae == 0.0 {
        return Ok(if mae == 0.0 { 0.0 } else { f64::INFINITY });
    }
    Ok(mae / naive_mae)
}

/// Inverse CDF (quantile) for the standard normal distribution.
///
/// Uses the rational approximation by Peter J. Acklam with ~1e-9 absolute error.
pub fn norm_s_inv(p: f64) -> Result<f64, ErrorKind> {
    if !(0.0..=1.0).contains(&p) || !p.is_finite() {
        return Err(ErrorKind::Num);
    }
    if p == 0.0 || p == 1.0 {
        return Err(ErrorKind::Num);
    }

    // Coefficients in rational approximations.
    const A: [f64; 6] = [
        -3.969683028665376e+01,
        2.209460984245205e+02,
        -2.759285104469687e+02,
        1.383577518672690e+02,
        -3.066479806614716e+01,
        2.506628277459239e+00,
    ];
    const B: [f64; 5] = [
        -5.447609879822406e+01,
        1.615858368580409e+02,
        -1.556989798598866e+02,
        6.680131188771972e+01,
        -1.328068155288572e+01,
    ];
    const C: [f64; 6] = [
        -7.784894002430293e-03,
        -3.223964580411365e-01,
        -2.400758277161838e+00,
        -2.549732539343734e+00,
        4.374664141464968e+00,
        2.938163982698783e+00,
    ];
    const D: [f64; 4] = [
        7.784695709041462e-03,
        3.224671290700398e-01,
        2.445134137142996e+00,
        3.754408661907416e+00,
    ];

    const P_LOW: f64 = 0.02425;
    const P_HIGH: f64 = 1.0 - P_LOW;

    let x = if p < P_LOW {
        let q = (-2.0 * p.ln()).sqrt();
        (((((C[0] * q + C[1]) * q + C[2]) * q + C[3]) * q + C[4]) * q + C[5])
            / ((((D[0] * q + D[1]) * q + D[2]) * q + D[3]) * q + 1.0)
    } else if p <= P_HIGH {
        let q = p - 0.5;
        let r = q * q;
        (((((A[0] * r + A[1]) * r + A[2]) * r + A[3]) * r + A[4]) * r + A[5]) * q
            / (((((B[0] * r + B[1]) * r + B[2]) * r + B[3]) * r + B[4]) * r + 1.0)
    } else {
        let q = (-2.0 * (1.0 - p).ln()).sqrt();
        -(((((C[0] * q + C[1]) * q + C[2]) * q + C[3]) * q + C[4]) * q + C[5])
            / ((((D[0] * q + D[1]) * q + D[2]) * q + D[3]) * q + 1.0)
    };

    if x.is_finite() {
        Ok(x)
    } else {
        Err(ErrorKind::Num)
    }
}
