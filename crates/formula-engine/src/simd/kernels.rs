use wide::f64x4;

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum CmpOp {
    Eq,
    Ne,
    Lt,
    Le,
    Gt,
    Ge,
}

#[derive(Clone, Copy, Debug, PartialEq)]
pub struct NumericCriteria {
    pub op: CmpOp,
    pub rhs: f64,
}

impl NumericCriteria {
    #[inline]
    pub const fn new(op: CmpOp, rhs: f64) -> Self {
        Self { op, rhs }
    }
}

#[inline]
pub fn sum_ignore_nan_f64(values: &[f64]) -> f64 {
    let (sum, _) = sum_count_ignore_nan_f64(values);
    sum
}

#[inline]
pub fn count_ignore_nan_f64(values: &[f64]) -> usize {
    let (_, count) = sum_count_ignore_nan_f64(values);
    count
}

/// Sum numbers while skipping NaNs. Returns `(sum, count_non_nan)`.
///
/// This implementation intentionally keeps the hot arithmetic in SIMD via `wide`,
/// while handling NaN filtering per-lane (branchy but still faster for large buffers).
pub fn sum_count_ignore_nan_f64(values: &[f64]) -> (f64, usize) {
    let mut acc = f64x4::from([0.0; 4]);
    let mut count = 0usize;

    let mut i = 0usize;
    while i + 4 <= values.len() {
        let mut lanes = [values[i], values[i + 1], values[i + 2], values[i + 3]];
        for lane in &mut lanes {
            if lane.is_nan() {
                *lane = 0.0;
            } else {
                count += 1;
            }
        }
        acc += f64x4::from(lanes);
        i += 4;
    }

    let acc_arr = acc.to_array();
    let mut sum = acc_arr[0] + acc_arr[1] + acc_arr[2] + acc_arr[3];

    for &v in &values[i..] {
        if v.is_nan() {
            continue;
        }
        sum += v;
        count += 1;
    }

    (sum, count)
}

pub fn min_ignore_nan_f64(values: &[f64]) -> Option<f64> {
    let mut acc = f64x4::from([f64::INFINITY; 4]);

    let mut i = 0usize;
    while i + 4 <= values.len() {
        let mut lanes = [values[i], values[i + 1], values[i + 2], values[i + 3]];
        for lane in &mut lanes {
            if lane.is_nan() {
                *lane = f64::INFINITY;
            }
        }
        let v = f64x4::from(lanes);
        acc = acc.min(v);
        i += 4;
    }

    let arr = acc.to_array();
    let mut best = arr[0].min(arr[1]).min(arr[2]).min(arr[3]);
    for &v in &values[i..] {
        if v.is_nan() {
            continue;
        }
        best = best.min(v);
    }

    if best.is_infinite() {
        None
    } else {
        Some(best)
    }
}

pub fn max_ignore_nan_f64(values: &[f64]) -> Option<f64> {
    let mut acc = f64x4::from([f64::NEG_INFINITY; 4]);

    let mut i = 0usize;
    while i + 4 <= values.len() {
        let mut lanes = [values[i], values[i + 1], values[i + 2], values[i + 3]];
        for lane in &mut lanes {
            if lane.is_nan() {
                *lane = f64::NEG_INFINITY;
            }
        }
        let v = f64x4::from(lanes);
        acc = acc.max(v);
        i += 4;
    }

    let arr = acc.to_array();
    let mut best = arr[0].max(arr[1]).max(arr[2]).max(arr[3]);
    for &v in &values[i..] {
        if v.is_nan() {
            continue;
        }
        best = best.max(v);
    }

    if best.is_infinite() {
        None
    } else {
        Some(best)
    }
}

pub fn count_if_f64(values: &[f64], criteria: NumericCriteria) -> usize {
    let mut count = 0usize;

    let mut i = 0usize;
    while i + 4 <= values.len() {
        let lanes = [values[i], values[i + 1], values[i + 2], values[i + 3]];
        for &v in &lanes {
            if v.is_nan() {
                continue;
            }
            if matches_criteria(v, criteria) {
                count += 1;
            }
        }
        i += 4;
    }

    for &v in &values[i..] {
        if v.is_nan() {
            continue;
        }
        if matches_criteria(v, criteria) {
            count += 1;
        }
    }

    count
}

pub fn sumproduct_ignore_nan_f64(a: &[f64], b: &[f64]) -> f64 {
    debug_assert_eq!(a.len(), b.len());

    let mut acc = f64x4::from([0.0; 4]);

    let mut i = 0usize;
    while i + 4 <= a.len() {
        let mut la = [a[i], a[i + 1], a[i + 2], a[i + 3]];
        let mut lb = [b[i], b[i + 1], b[i + 2], b[i + 3]];
        for (xa, xb) in la.iter_mut().zip(lb.iter_mut()) {
            if xa.is_nan() || xb.is_nan() {
                *xa = 0.0;
                *xb = 0.0;
            }
        }
        let va = f64x4::from(la);
        let vb = f64x4::from(lb);
        acc += va * vb;
        i += 4;
    }

    let arr = acc.to_array();
    let mut sum = arr[0] + arr[1] + arr[2] + arr[3];

    for (&x, &y) in a[i..].iter().zip(&b[i..]) {
        if x.is_nan() || y.is_nan() {
            continue;
        }
        sum += x * y;
    }
    sum
}

pub fn add_f64(out: &mut [f64], a: &[f64], b: &[f64]) {
    debug_assert_eq!(out.len(), a.len());
    debug_assert_eq!(out.len(), b.len());

    let mut i = 0usize;
    while i + 4 <= out.len() {
        let va = f64x4::from([a[i], a[i + 1], a[i + 2], a[i + 3]]);
        let vb = f64x4::from([b[i], b[i + 1], b[i + 2], b[i + 3]]);
        let vr = va + vb;
        let r = vr.to_array();
        out[i..i + 4].copy_from_slice(&r);
        i += 4;
    }
    for ((o, x), y) in out[i..].iter_mut().zip(&a[i..]).zip(&b[i..]) {
        *o = *x + *y;
    }
}

pub fn sub_f64(out: &mut [f64], a: &[f64], b: &[f64]) {
    debug_assert_eq!(out.len(), a.len());
    debug_assert_eq!(out.len(), b.len());

    let mut i = 0usize;
    while i + 4 <= out.len() {
        let va = f64x4::from([a[i], a[i + 1], a[i + 2], a[i + 3]]);
        let vb = f64x4::from([b[i], b[i + 1], b[i + 2], b[i + 3]]);
        let vr = va - vb;
        let r = vr.to_array();
        out[i..i + 4].copy_from_slice(&r);
        i += 4;
    }
    for ((o, x), y) in out[i..].iter_mut().zip(&a[i..]).zip(&b[i..]) {
        *o = *x - *y;
    }
}

pub fn mul_f64(out: &mut [f64], a: &[f64], b: &[f64]) {
    debug_assert_eq!(out.len(), a.len());
    debug_assert_eq!(out.len(), b.len());

    let mut i = 0usize;
    while i + 4 <= out.len() {
        let va = f64x4::from([a[i], a[i + 1], a[i + 2], a[i + 3]]);
        let vb = f64x4::from([b[i], b[i + 1], b[i + 2], b[i + 3]]);
        let vr = va * vb;
        let r = vr.to_array();
        out[i..i + 4].copy_from_slice(&r);
        i += 4;
    }
    for ((o, x), y) in out[i..].iter_mut().zip(&a[i..]).zip(&b[i..]) {
        *o = *x * *y;
    }
}

pub fn div_f64(out: &mut [f64], a: &[f64], b: &[f64]) {
    debug_assert_eq!(out.len(), a.len());
    debug_assert_eq!(out.len(), b.len());

    let mut i = 0usize;
    while i + 4 <= out.len() {
        let va = f64x4::from([a[i], a[i + 1], a[i + 2], a[i + 3]]);
        let vb = f64x4::from([b[i], b[i + 1], b[i + 2], b[i + 3]]);
        let vr = va / vb;
        let r = vr.to_array();
        out[i..i + 4].copy_from_slice(&r);
        i += 4;
    }
    for ((o, x), y) in out[i..].iter_mut().zip(&a[i..]).zip(&b[i..]) {
        *o = *x / *y;
    }
}

#[inline]
fn matches_criteria(v: f64, criteria: NumericCriteria) -> bool {
    match criteria.op {
        CmpOp::Eq => v == criteria.rhs,
        CmpOp::Ne => v != criteria.rhs,
        CmpOp::Lt => v < criteria.rhs,
        CmpOp::Le => v <= criteria.rhs,
        CmpOp::Gt => v > criteria.rhs,
        CmpOp::Ge => v >= criteria.rhs,
    }
}
