use crate::value::ErrorKind;

/// Excel-compatible histogram counts (`FREQUENCY`).
///
/// Returns a vector of length `bins.len() + 1`.
///
/// For each data value, we increment the first bin whose boundary is `>= value`
/// (treating `<=` as "belongs to the bin"), and fall back to the final bucket when
/// `value` is greater than every boundary (i.e. `value > max(bins)`).
pub fn frequency(data: &[f64], bins: &[f64]) -> Result<Vec<u64>, ErrorKind> {
    if data.iter().any(|v| !v.is_finite()) || bins.iter().any(|v| !v.is_finite()) {
        return Err(ErrorKind::Num);
    }

    let out_len = bins.len().saturating_add(1);
    let mut counts = Vec::<u64>::new();
    if counts.try_reserve_exact(out_len).is_err() {
        return Err(ErrorKind::Num);
    }
    counts.resize(out_len, 0);

    if bins.is_empty() {
        counts[0] = data.len() as u64;
        return Ok(counts);
    }

    let nondecreasing = bins.windows(2).all(|w| w[0] <= w[1]);
    if nondecreasing {
        for &x in data {
            // First index where bin >= x.
            let idx = bins.partition_point(|b| *b < x);
            counts[idx] = counts[idx].saturating_add(1);
        }
        return Ok(counts);
    }

    // Fallback for unsorted bins (Excel recommends ascending order, but we still want
    // deterministic behavior).
    for &x in data {
        let mut assigned = false;
        for (i, &b) in bins.iter().enumerate() {
            if x <= b {
                counts[i] = counts[i].saturating_add(1);
                assigned = true;
                break;
            }
        }
        if !assigned {
            counts[bins.len()] = counts[bins.len()].saturating_add(1);
        }
    }

    Ok(counts)
}
